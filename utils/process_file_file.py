import os
from utils.file_file.file_extractor import file_extractor_tool
from utils.file_file.file_content_format import file_content_format
from utils.common.write_to_excel import write_to_excel


def process_file_file(
    target_file,
    target_table_name,
    output_folder_path,
    template_path,
    new_file_name,
):
    interface_or_file = "file"
    output_file_path = os.path.join(
        output_folder_path,
        f"{new_file_name}_データ作成.xlsm",
    )
    
    for table_name in target_table_name:
        file_target_sheet_name = f"ファイル仕様({table_name})"
        filter_by_type, table_header, header_columns = file_extractor_tool(
            target_file, file_target_sheet_name, table_name
        )
        print(filter_by_type)

        formated_extracted_content = file_content_format(
            filter_by_type,
            target_file,
            file_target_sheet_name,
            header_columns,
            table_header,
        )
        write_to_excel(
            template_path,
            formated_extracted_content,
            output_file_path,
            table_name,
            interface_or_file,
        )
