import pandas as pd
import os

target_file_list = {}

BASE_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
excel_files_path = os.path.join(BASE_PATH, "excel_files")
target_file_path = os.path.join(BASE_PATH, "target", "Templete_File_chosen.xlsm")


def position_identify(target_file_path, excel_files_path):
    fileID = set()
    df = pd.read_excel(target_file_path, header=None)
    for index_row, row in df.iterrows():
        if row[0] == "●":
            fileID.add(index_row)

    format_requests = {}
    for fileid_index in fileID:
        row_content = [
            str(content) for content in df.iloc[fileid_index] if pd.notna(content)
        ]
        if len(row_content) >= 2:
            format_requests[row_content[1]] = row_content[2:]

    for filename, table_name in format_requests.items():
        excel_files = os.listdir(excel_files_path)
        matching_files = [str(file) for file in excel_files if filename in file]
        if not matching_files:
            print(f"警告: 未找到包含 {filename} 的文件")
            continue
        target_file = matching_files[0]
        if target_file not in target_file_list:
            target_file_list[target_file] = []
        target_file_list[target_file].extend(table_item for table_item in table_name)

    # print(target_file_list)
    return target_file_list


if __name__ == "__main__":
    file_table_identify(target_file_path, excel_files_path)
