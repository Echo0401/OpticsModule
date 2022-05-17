# -*- coding: UTF-8 -*-

# -*- coding: UTF-8 -*-
import time
import serial
from enum import Enum
import serial
from serial.tools import list_ports

class CyclesCommand(Enum):
    """
    200 500 1500 cycles command
    """
    Two_hundred_command = ":0006000000C832\r\n"
    Five_hundred_command = ":0006000001f405\r\n"
    One_Five_hundred_command = ":0006000005dc19\r\n"

class ScanED(Enum):
    """
     Method Type 8-S_E1D1 10-S_E2D2 command
    """
    FAM = ":000600030800ef\r\n"
    ROX = ":000600030a00ed\r\n"
    HEX = ":010600030800ee\r\n"
    CY5 = ":010600030a00ec\r\n"

class StartMethod(Enum):
    """
    Start Method Command: 1 start
    """
    COM0 = ":000602000001f7\r\n"
    COM1 = ":010602000001f6\r\n"

class LEDCurrent(Enum):
    """
    LED1 Current LED2 Current command
    """
    FAM_50 = ":000600183200B0\r\n"
    FAM_240 = ":00060018F000F2\r\n"
    ROX_80 = ":00060019500091\r\n"
    HEX_20 = ":010600181400cd\r\n"
    HEX_40 = ":010600182800b9\r\n"
    HEX_80 = ":01060018500091\r\n"
    HEX_110 = ":010600186E0073\r\n"
    CY5_140 = ":010600198C0054\r\n"
    CY5_130 = ":0106001982005E\r\n"

class SleepTime(Enum):
    Tree = 3
    Six = 6
    Eighteen = 18

class CyclesPoints(Enum):
    Two_hundred = 200
    Five_hundred = 500
    One_Five_hundred = 1500

class SavePoints(Enum):
    Two_points = ":0003020000C833\r\n"
    Five_points = ":0003020001f406\r\n"    # 不可用
    One_Five_points = ":0003020005dc1a\r\n"


def cycle_count(start_demo, end_demo):
    """
    根据start method 开始和结束的cycle, 判断是否发送了指定的点数
    :param start_demo:
    :param end_demo:
    :return:
    """
    start_demo_num = list(start_demo[-1].decode('utf-8')[7:15])
    end_demo_num = list(list(end_demo[-1].decode('utf-8')[7:15]))

    total_start_demo = ''
    total_end_demo = ''
    for i in start_demo_num:
        total_start_demo += i
    cycle_start = int(total_start_demo, 16)

    for i in end_demo_num:
        total_end_demo += i
    cycle_end = int(total_end_demo, 16)

    return cycle_start, cycle_end

def command_serial(datapoints):
    """
    根据传入的点数，得到读取对应点数命令的列表
    :return:
    """
    address_list = []
    for points in range(datapoints):
        address_list.append(hex(513 + 2 * points)[2:])

    command_list = []
    for address in address_list:
        command_list.append('00030' + address + '0002')

    serial_command_list = []
    for command in command_list:
        row_command = command
        command = list(command)
        # 每两位合成一位， 并转成16进制
        command = [int(element + command[index + 1], 16) for index, element in enumerate(command) if index % 2 == 0]
        command = bytes(command)
        lrc = 0
        for b in command:
            lrc = (lrc + b) & 0xff
        lrc = ((lrc ^ 0xff) + 1) & 0xff
        lrc_num = "{:02x}".format(lrc)
        serial_command = ":" + row_command + lrc_num
        serial_command_list.append(serial_command)
    return serial_command_list

def counts_to_mv(counts_to_mv_list):
    """
    把counts数转化为mv
    :return:
    """
    # 过滤数据
    counts_to_mv_format_list = []
    conversion_formula_mv_list = []
    for counts_to_mv in counts_to_mv_list:
        if len(counts_to_mv) == 19:
            try:
                counts_to_mv.decode('utf-8')
                if counts_to_mv[0:9] == b':00030400':
                    counts_to_mv_format_list.append(counts_to_mv)
            except Exception as err:
                print(f"The readline can not format utf-8 and b':000304': {err}")

    for counts_to_mv in counts_to_mv_format_list:
        count_row_number = int(counts_to_mv[7:15], 16)
        conversion_formula_mv = count_row_number * 0.000298026
        conversion_formula_mv_list.append(round(conversion_formula_mv, 2))
    return conversion_formula_mv_list

class ScanEDTest(object):
    """
    稳定性、灵敏度、截止率
    FAM\ROX\HEX\CY5
    底层命令行的控制
    实例
    """
    def __init__(self):
        super(ScanEDTest, self).__init__()
        self.ser = None
        self.communicate_status = False

    def __del__(self):
        try:
            self.ser.close()
        except:
            pass

    def try_connect_optics_module(self):
        """
                确认光学模块连接的串口的端口号
                :return:
                """
        # 通信状态
        self.communicate_status = False
        try:
            # 获取所有串口
            com_ports = [item.device for item in list_ports.comports()]
            for com_port in com_ports:
                # print(com_port, "look for com in com ports.")
                try:
                    # 尝试心跳
                    self.ser = serial.Serial(com_port, 57600, timeout=0.5, writeTimeout=0.1)
                    self.ser.write(":0003018000106c\r\n".encode())
                    # print(f"command: {ser.readline()}")
                    if self.ser.read():
                        self.communicate_status = True
                        # print(f"communicate successful, port: {com_port}")
                        return True
                    else:
                        self.ser.close()
                except Exception as err:
                    # COM口异常，说明此端口，不是需要的端口
                    print(f"try connect port Exception1: {err}")
        except Exception as err:
            # COM口异常，说明此端口，不是需要的端口
            print(f"try connect port Exception2: {err}")

        return False

    def scantest(self, cyclescmmand, scaned, ledcurrent, startmethod, sleeptime, cyclespoints):
        """
            稳定性、灵敏度、截止率
            FAM\ROX\HEX\CY5
            底层命令行的控制
            :return:
            """
        if not self.communicate_status:
            print("Communicate status is False!")
            return

        # step1: cycles 200, 500, 1500
        self.ser.write(f"{cyclescmmand.value}\r\n".encode())
        time.sleep(0.1)
        # step2: Scan-E1D1
        self.ser.write(f"{scaned.value}\r\n".encode())
        time.sleep(0.1)
        # step3: E1D1 Current 50
        self.ser.write(f"{ledcurrent.value}\r\n".encode())
        time.sleep(0.1)

        self.ser.write(":000301000002fa\r\n".encode())
        cycle_command_before_start_method = self.ser.readlines()
        # print(f"ticket before start Method readlines: {cycle_command_before_start_method}")
        time.sleep(0.1)

        # step4: Start method
        self.ser.write(f"{startmethod.value}\r\n".encode())
        # 1000: 12s  (must be changed according to cycles)
        time.sleep(sleeptime.value)

        self.ser.write(":000301000002fa\r\n".encode())
        cycle_command_after_start_method = self.ser.readlines()
        # print(f"ticket before start Method readlines:{cycle_command_after_start_method} ")
        time.sleep(0.1)

        cycle_start, cycle_end = cycle_count(cycle_command_before_start_method, cycle_command_after_start_method)
        # print(f"cycle_start: {cycle_start}")
        # print(f"cycle_end: {cycle_end}")

        try:
            counts_to_mv_list = []
            cycles = cycle_end - cycle_start
            if cycles == cyclespoints.value:
                serial_command_list = command_serial(cycles)
                for serial_command in serial_command_list:
                    self.ser.write(f"{serial_command}\r\n".encode())
                    counts_to_mv_list.append(self.ser.readline())
                conversion_formula_mv_list = counts_to_mv(counts_to_mv_list)
                return conversion_formula_mv_list
            else:
                print("Maybe your time it too short to satisfy cycle request.")
        except Exception as err:
            print(f"The serial error is: {err}")

    # 稳定性——FAM
    def stabilityfam(self):
        stability_fam = self.scantest(cyclescmmand=CyclesCommand.One_Five_hundred_command,
                                 scaned=ScanED.FAM,
                                 ledcurrent=LEDCurrent.FAM_240,
                                 startmethod=StartMethod.COM0,
                                 sleeptime=SleepTime.Eighteen,
                                 cyclespoints=CyclesPoints.One_Five_hundred
                                 )
        return stability_fam
        # print(f"stability_fam: {stability_fam}")

    # 稳定性——ROX
    def stabilityrox(self):
        stability_rox = self.scantest(cyclescmmand=CyclesCommand.One_Five_hundred_command,
                                 scaned=ScanED.ROX,
                                 ledcurrent=LEDCurrent.ROX_80,
                                 startmethod=StartMethod.COM0,
                                 sleeptime=SleepTime.Eighteen,
                                 cyclespoints=CyclesPoints.One_Five_hundred
                                 )
        return stability_rox
        # print(f"stability_rox: {stability_rox}")
    # 稳定性——HEX
    def stabilityhex(self):
        # Why is same as the value of rox
        stability_hex = self.scantest(cyclescmmand=CyclesCommand.One_Five_hundred_command,
                                 scaned=ScanED.HEX,
                                 ledcurrent=LEDCurrent.HEX_110,
                                 startmethod=StartMethod.COM1,
                                 sleeptime=SleepTime.Eighteen,
                                 cyclespoints=CyclesPoints.One_Five_hundred
                                 )
        return stability_hex
        # print(f"stability_hex: {stability_hex}")
    # 稳定性——CY5
    def stabilitycy5(self):
        stability_cy5 = self.scantest(cyclescmmand=CyclesCommand.One_Five_hundred_command,
                                 scaned=ScanED.CY5,
                                 ledcurrent=LEDCurrent.CY5_140,
                                 startmethod=StartMethod.COM1,
                                 sleeptime=SleepTime.Eighteen,
                                 cyclespoints=CyclesPoints.One_Five_hundred
                                 )
        return stability_cy5
        # print(f"stability_cy5: {stability_cy5}")

    # 截止率——FAM
    def relativityfam(self):
        relativity_fam = self.scantest(cyclescmmand=CyclesCommand.Two_hundred_command,
                                 scaned=ScanED.FAM,
                                 ledcurrent=LEDCurrent.FAM_240,
                                 startmethod=StartMethod.COM0,
                                 sleeptime=SleepTime.Tree,
                                 cyclespoints=CyclesPoints.Two_hundred
                                 )
        return relativity_fam
        # print(f"relativity_fam: {relativity_fam}")

    # 截止率——ROX
    def relativityrox(self):
        relativity_rox = self.scantest(cyclescmmand=CyclesCommand.Two_hundred_command,
                                 scaned=ScanED.ROX,
                                 ledcurrent=LEDCurrent.ROX_80,
                                 startmethod=StartMethod.COM0,
                                 sleeptime=SleepTime.Tree,
                                 cyclespoints=CyclesPoints.Two_hundred
                                 )
        return relativity_rox
        # print(f"relativity_rox: {relativity_rox}")

    # 截止率——HEX
    def relativityhex(self):
        relativity_hex = self.scantest(cyclescmmand=CyclesCommand.Two_hundred_command,
                                 scaned=ScanED.HEX,
                                 ledcurrent=LEDCurrent.HEX_110,
                                  startmethod=StartMethod.COM1,
                                 sleeptime=SleepTime.Tree,
                                 cyclespoints=CyclesPoints.Two_hundred
                                 )
        return relativity_hex
        # print(f"relativity_hex: {relativity_hex}")

    # 截止率——CY5
    def relativitycy5(self):
        relativity_cy5 = self.scantest(cyclescmmand=CyclesCommand.Two_hundred_command,
                                 scaned=ScanED.CY5,
                                 ledcurrent=LEDCurrent.CY5_140,
                                 startmethod=StartMethod.COM1,
                                 sleeptime=SleepTime.Tree,
                                 cyclespoints=CyclesPoints.Two_hundred
                                 )
        return relativity_cy5
        # print(f"relativity_cy5: {relativity_cy5}")

    # 灵敏度——FAM
    def sensitivityfam(self):
        sensitivity_fam = self.scantest(cyclescmmand=CyclesCommand.Five_hundred_command,
                                  scaned=ScanED.FAM,
                                  ledcurrent=LEDCurrent.FAM_50,
                                  startmethod=StartMethod.COM0,
                                  sleeptime=SleepTime.Six,
                                  cyclespoints=CyclesPoints.Five_hundred
                                  )
        return sensitivity_fam
        # print(f"sensitivity_fam: {sensitivity_fam}")

    # 灵敏度——ROX
    def sensitivityrox(self):
        sensitivity_rox = self.scantest(cyclescmmand=CyclesCommand.Five_hundred_command,
                                  scaned=ScanED.ROX,
                                  ledcurrent=LEDCurrent.ROX_80,
                                  startmethod=StartMethod.COM0,
                                  sleeptime=SleepTime.Six,
                                  cyclespoints=CyclesPoints.Five_hundred
                                  )
        return sensitivity_rox
        # print(f"sensitivity_rox: {sensitivity_rox}")

    # 灵敏度——HEX
    def sensitivityhex(self):
        sensitivity_hex = self.scantest(cyclescmmand=CyclesCommand.Five_hundred_command,
                                  scaned=ScanED.HEX,
                                  ledcurrent=LEDCurrent.HEX_40,
                                  startmethod=StartMethod.COM1,
                                  sleeptime=SleepTime.Six,
                                  cyclespoints=CyclesPoints.Five_hundred
                                  )
        return sensitivity_hex
        # print(f"sensitivity_hex: {sensitivity_hex}")

    # 灵敏度——CY5
    def sensitivitycy5(self):
        sensitivity_cy5 = self.scantest(cyclescmmand=CyclesCommand.Five_hundred_command,
                                  scaned=ScanED.CY5,
                                  ledcurrent=LEDCurrent.CY5_140,
                                  startmethod=StartMethod.COM1,
                                  sleeptime=SleepTime.Six,
                                  cyclespoints=CyclesPoints.Five_hundred
                                  )
        return sensitivity_cy5
        # print(f"sensitivity_cy5: {sensitivity_cy5}")



if __name__ == '__main__':
    scanedtest = ScanEDTest()
    # scanedtest.stabilityfam()
    # scanedtest.stabilityrox()
    # scanedtest.stabilityhex()
    # scanedtest.stabilitycy5()
    # scanedtest.relativityfam()
    # scanedtest.relativityrox()
    # scanedtest.relativityhex()
    # scanedtest.relativitycy5()
    # scanedtest.sensitivityfam()
    # scanedtest.sensitivityrox()
    # scanedtest.sensitivityhex()
    # scanedtest.sensitivitycy5()

