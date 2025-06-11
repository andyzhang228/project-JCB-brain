import os
from utils.interface_file.interface_table_identify import target_table_identify
from utils.interface_file.interface_extractor import interface_extractor_tool
from utils.interface_file.interface_content_format import interface_content_format
from utils.common.write_to_excel import write_to_excel


def process_interface_file(
    target_file,
    target_table_name,
    output_folder_path,
    template_path,
    target_table_identify_path,
    new_file_name,
):
    interface_or_file = "interface"
    target_table_name = {
        table_name: target_table_identify(target_table_identify_path, table_name)
        for table_name in target_table_name
    }
    

    output_interface_path = os.path.join(
        output_folder_path,
        f"{new_file_name}_データ作成.xlsm",
    )

    interface_target_sheet_name = os.getenv("INTERFACE_SHEET_NAME")
    for table_name,table_name_value in target_table_name.items():
        extracted_content, _ = interface_extractor_tool(
            target_file, interface_target_sheet_name, table_name_value
        )
        formated_extracted_content = interface_content_format(extracted_content)
        write_to_excel(
            template_path,
            formated_extracted_content,
            output_interface_path,
            table_name,
            interface_or_file,
        )
