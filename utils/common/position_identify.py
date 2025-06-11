import pandas as pd
import os


def position_identify(target_file_path, excel_files_path):
    target_file_list = {}
    fileID = set()
    filename_list={}
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
        if matching_files[0] not in filename_list:
            filename_list[matching_files[0]]=""
        filename_list[matching_files[0]]=filename  
        if not matching_files:
            print(f" {filename} が見つからないです。")
            continue
        target_file = matching_files[0]
        if target_file not in target_file_list:
            target_file_list[target_file] = []
        target_file_list[target_file].extend(table_item for table_item in table_name)
        
    return target_file_list,filename_list
