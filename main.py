import os
from datetime import datetime
from dotenv import load_dotenv
from utils.interface_extractor import interface_extractor_tool
from utils.file_extractor import file_extractor_tool
from utils.position_identify import position_identify
from utils.interface_content_format import interface_content_format
from utils.file_content_format import file_content_format
from utils.write_to_excel import write_to_excel
from utils.interface_table_identify import target_table_identify

load_dotenv()
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
excel_files_path = os.path.join(BASE_PATH, "excel_files")
target_file_path = os.path.join(BASE_PATH, "target", "Templete_File_chosen.xlsm")
template_path = os.path.join(BASE_PATH, "templates", "Template.xlsm")
target_table_identify_path=os.path.join(BASE_PATH,"target","Format_list.xlsx")

def generate_format(target_file_path):
    # 1.根据target_file_path中的"●"指定的一行取取得fileid和需要被提取的tablename
    target_file_list = position_identify(target_file_path, excel_files_path)
    if not target_file_list:
        raise FileNotFoundError(f"文件 {target_file} 不存在")

    # 创建output文件夹
    output_path = os.path.join(BASE_PATH, "output")
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    output_folder_path = os.path.join(
        output_path, datetime.now().strftime("%Y%m%d%H%M")
    )
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
    # 2.根据target_file_list中的fileid和tablename，提取指定表格的格式
    
    # print(f"target_file_list:{target_file_list}")
    for target_file, target_table_name in target_file_list.items():
        if "インターフェース" in str(target_file):
            interface_output_number = 1
            interface_or_file = "interface"
            print(target_table_name)
            
            output_interface_path = os.path.join(
                output_folder_path,
                f"interface_excel_{interface_output_number}.xlsx",
            )

            interface_target_sheet_name = os.getenv("INTERFACE_SHEET_NAME")
            # 2-1. 把各个fileid的文件名和其指定的tablename传入给extractor_tool，进行内容提取
            for table_name in target_table_name:
                formated_table_name=target_table_identify(target_table_identify_path,target_table_name)
                print(formated_table_name)
            #     extracted_content, _ = interface_extractor_tool(
            #         target_file, interface_target_sheet_name, table_name
            #     )
            #     formated_extracted_content = interface_content_format(
            #         extracted_content
            #     )
            #     # 3.把提取到的内容追加到template并保存到新的excel，导出至output_path
            #     write_to_excel(
            #         template_path,
            #         formated_extracted_content,
            #         output_interface_path,
            #         table_name,
            #         interface_or_file,
            #     )
            # interface_output_number += 1
        # if "ファイル仕様" in target_file:
        #     interface_or_file = "file"
        #     file_output_number = 1
        #     output_file_path = os.path.join(
        #         output_folder_path,
        #         f"file_excel_{file_output_number}.xlsx",
        #     )
        #     for table_name in target_table_name:
        #         file_target_sheet_name = f"ファイル仕様({table_name})"
        #         filter_by_type, table_header, header_columns = file_extractor_tool(
        #             target_file, file_target_sheet_name, table_name
        #         )

        #         formated_extracted_content = file_content_format(
        #             filter_by_type,
        #             target_file,
        #             file_target_sheet_name,
        #             header_columns,
        #             table_header,
        #         )
        #         write_to_excel(
        #             template_path,
        #             formated_extracted_content,
        #             output_file_path,
        #             table_name,
        #             interface_or_file,
        #         )
        #     file_output_number += 1
    return


if __name__ == "__main__":
    generate_format(target_file_path)
