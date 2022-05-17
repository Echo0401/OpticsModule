temp_channel_list = []
for channel in datapoints_report_dict['stability'].keys():
    temp_channel_list.append(channel)
if 'fam' not in temp_channel_list:
    self.label_relativity_fam.setText(f"平 均 值：  "
                                      f" {relativity_fam_avg}   结 果：  "
                                      f" 无 ")
else:
    stability_minus_relativity_fam = abs(
        round(datapoints_report_dict['stability']['fam'][0] - relativity_fam_avg, 1))
    self.label_relativity_fam.setText(f"差 值：  "
                                      f"{stability_minus_relativity_fam}   结 果：  "
                                      f"{self.pf(stability_minus_relativity_fam < 10)}")