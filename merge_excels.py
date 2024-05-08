import os
import pandas as pd
# .\env\Scripts\activate

def merge_one_image(input_dir):
    # open csv files in the input directory
    files = os.listdir(input_dir)
    files = [file for file in files if file.endswith(".csv")]

    for file in files: 
        if file == "area_measurements.csv":
            area_measurement = pd.read_csv(input_dir + "\\" + file)
            area = float(area_measurement["%Area"])
        elif file == "network_information.csv":
            network_information = pd.read_csv(input_dir + "\\" + file)
            network = network_information[["# Branches", "Average Branch Length", "Maximum Branch Length", "Longest Shortest Path"]]
            network_summary = network.describe()

            # fixing the count function
            for n in network:
                sum = network[n].sum()
                network_summary[n]["count"] = sum

            # create dataframe with columns "# Branches", "Average Branch Length", "Maximum Branch Length", "Longest Shortest Path" and their summary statistics
            net_summary = pd.DataFrame(columns=["# Branches", "Average Branch Length", "Maximum Branch Length", "Longest Shortest Path"])
            net_summary = network_summary
            # net_summary.loc[0] = ["mean", network_summary["# Branches"]["mean"], network_summary["Average Branch Length"]["mean"], network_summary["Maximum Branch Length"]["mean"], network_summary["Longest Shortest Path"]["mean"]]
        elif file == "particle_measurements.csv":
            particle_measurement = pd.read_csv(input_dir + "\\" + file, encoding='iso8859_9')
            particle = particle_measurement[["Vol. (pixels³)", "SA (pixels²)", "Encl. Vol. (pixels³)", "Thickness (pixels)", "Max Thickness (pixels)"]]
            particle_summary = particle.describe()

            # fixing the count function
            for p in particle:
                sum = particle[p].sum()
                particle_summary[p]["count"] = sum

            # create dataframe with columns "Vol. (pixels3)", "SA (pixels2)", "Encl. Vol. (pixels3)", "Thickness (pixels)", "Max Thickness (pixels)" and their summary statistics
            part_summary = pd.DataFrame(columns=["Vol. (pixels3)", "SA (pixels2)", "Encl. Vol. (pixels3)", "Thickness (pixels)", "Max Thickness (pixels)"])
            part_summary = particle_summary

    # merge the dataframes
    merged = pd.concat([particle, network], axis=1)
    merged2 = pd.concat([part_summary, net_summary], axis=1)

    # add area column to the merged dataframe, set only first row to the area value
    merged["%Area"] = None
    merged.loc[0, "%Area"] = area

    return merged, merged2

# get input directory address
input_dir = input("Enter Input directory")
files = os.listdir(input_dir)
sample_dict = {}

raw = []
summary = []
letter = files[0][0]
for cur_file in range(len(files)):
    num = files[cur_file][1]
    combo = letter + num
    if combo == "C2" or combo == 'C2':
        continue
    elif letter == files[cur_file][0]:
        print(letter + num)
        m, m2 = merge_one_image(input_dir + "\\" + files[cur_file])
        m.insert(0, "SAMPLE", letter + num)
        m2.insert(0, "SAMPLE", letter + num)
        raw.append(m)
        summary.append(m2)
    else:
        raw_pd = pd.concat(raw, axis=1)
        summary_pd = pd.concat(summary, axis=1)
        sample_dict[letter] = [raw_pd, summary_pd]
        letter = files[cur_file][0]
        m, m2 = merge_one_image(input_dir + "\\" + files[cur_file])
        m.insert(0, "SAMPLE", letter + num)
        m2.insert(0, "SAMPLE", letter + num)
        raw = [m]
        summary = [m2]

raw_pd = pd.concat(raw, axis=1)
summary_pd = pd.concat(summary, axis=1)
sample_dict[letter] = [raw_pd, summary_pd]

# # Assuming particle_summary, network_summary are DataFrame objects
with pd.ExcelWriter(input_dir + "\\" + "summary.xlsx") as writer:
    for key in sample_dict:
        sample_dict[key][0].to_excel(writer, sheet_name=key + 'Raw')
        sample_dict[key][1].to_excel(writer, sheet_name=key + 'Summary')