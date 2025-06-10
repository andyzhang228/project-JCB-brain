import os
import pandas as pd


def target_table_identify(target_file_path,format):
    format_column_index=None
    format_row_index=None
    df = pd.read_excel(target_file_path, header=None)
    for index_row, row in df.iterrows():
        for column_index,column in enumerate(row):
            if column == format:
                print(f"找到了在{index_row}行{column_index}列")
                format_row_index=index_row
                format_column_index=column_index
                break
    table_name=df.iloc[format_row_index,format_column_index+1]
    print(table_name)
    return table_name
    


    
if __name__ == "__main__":    
    target_file_path="/home/andy/JCBproject/scripts_merge/new_tool/project-JCB-brain/target/Format_list.xlsx"
    target_table_identify(target_file_path,"FCL05402_HED")