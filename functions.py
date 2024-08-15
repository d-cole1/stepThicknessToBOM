import os
from collections import Counter
import math
from pandas import read_excel


# HELPER FUNCTIONS
# Searches the selected directory and extracts step files into a list
def stp_finder(source_dir):
    stp_files = []  # Initialize an empty list to store step files
    extensions = (".stp", ".step", ".STP", ".STEP")  # Define the files extensions to look for
    for root, dirs, files in os.walk(source_dir):  # Walk through the directory
        for file in files:  # Iterate over the files in the directory
            if file.endswith(extensions):  # Check if the file has a step extension
                stp_files.append(os.path.join(root, file))  # Add the files to the list
    stp_files = [item.replace("\\", "/") for item in stp_files]  # Normalize the file paths
    return stp_files


# Reads the BOM file and returns a pd DataFrame
def extract_mre_bom(path):
    df = read_excel(path, sheet_name="BOM", dtype="str")
    return df


# Checks the format of the BOM DataFrame
def check_bom_format(df_arg):
    des_count = df_arg["DESCRIPTION"].count()
    num_count = df_arg["DOC NUMBER"].count()
    type_count = df_arg["DOC TYPE"].count()
    part_count = df_arg["DOC PART"].count()

    # Check if the BOM is formatted correctly
    return des_count == num_count and type_count == part_count  # Returns True if both cond. satisfied


# MAIN FUNCTIONS
def get_mode(x):
    if len(x) == 0:
        return None
    counts = Counter(x)
    mode = counts.most_common(1)[0]
    return mode[0], mode[1]


def get_thickness(file):
    cart_dic = {}
    ver_dic = {}
    edge_dic = {}

    with open(file, 'rb') as stp_file:
        for row in stp_file:
            tmp_str = row.decode("utf-8")
            if 'CARTESIAN_POINT' in tmp_str:
                tmp_str = tmp_str.replace('#', 'pnd_').replace(" ", "")
                cart_dic[tmp_str[:tmp_str.find("=")].strip()] = \
                    tmp_str[tmp_str.find(",(") + 2:tmp_str.find("))")].strip()

            if 'VERTEX_POINT' in tmp_str:
                tmp_str = tmp_str.replace('#', 'pnd_').replace(" ", "")
                ver_dic[tmp_str[:tmp_str.find("=")].strip()] = \
                    tmp_str[tmp_str.find("',") + 2:tmp_str.find(")")].strip()

    with open(file, 'rb') as stp_file:
        for row in stp_file:
            tmp_str = row.decode("utf-8")
            if 'EDGE_CURVE' in tmp_str:
                tmp_str = tmp_str.replace('#', 'pnd_').replace(" ", "")
                tmp_edge = tmp_str[tmp_str.find("',") + 2:tmp_str.find(",.")].split(",")

                start0 = cart_dic[ver_dic[tmp_edge[0].strip()]].split(',')
                end0 = cart_dic[ver_dic[tmp_edge[1].strip()]].split(',')

                end0 = [float(i) for i in end0]
                start0 = [float(i) for i in start0]

                dist = [round(a_i - b_i, 6) for a_i, b_i in zip(end0, start0)]
                dist = math.sqrt(sum(d ** 2 for d in dist))
                edge_dic[tmp_str[:tmp_str.find("=")]] = round(dist, 3)

    edge_lens = [length for length in edge_dic.values() if 0.5 < length < 20.0]

    mode_result = get_mode(edge_lens)

    return mode_result[0] if mode_result else None


def execute_func(window, values):
    try:
        bom_df = extract_mre_bom(values["excel"])

        if check_bom_format(bom_df):

            stp_files = stp_finder(values["source"])
            if not stp_files:
                window.write_event_value("Error", "no_steps_found")

            bom_df = extract_mre_bom(values["excel"])

            thickness_values = []
            found_file_names = []

            output_dir = os.path.join(values["source"])

            new_name = os.path.join(output_dir, os.path.basename(values["excel"]))

            for stp_file in stp_files:
                file_name = os.path.splitext(os.path.basename(stp_file))[0]
                is_file_in_df = bom_df.isin([file_name]).any().any()

                if is_file_in_df:  # '== True' omitted
                    found_file_names.append(file_name)
                    thickness = get_thickness(stp_file)
                    thickness_values.append(thickness)

            description_to_thickness = dict(zip(found_file_names, thickness_values))

            bom_df["Gage (Thickness)"] = bom_df["DESCRIPTION"].map(description_to_thickness)

            bom_df.to_excel(f"{new_name.strip('.xlsx')}_withThickness.xlsx", index=False, sheet_name="BOM")

            window.write_event_value("Done", None)

        else:
            window.write_event_value("Error", "invalid_BOM")

    except Exception as e:
        window.write_event_value("Error", str(e))
