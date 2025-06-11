import os
from datetime import datetime
from dotenv import load_dotenv
from utils.common.position_identify import position_identify
from utils.process_interface_file import process_interface_file
from utils.process_file_file import process_file_file

load_dotenv()
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
excel_files_path = os.path.join(BASE_PATH, "excel_files")
target_file_path = os.path.join(BASE_PATH, "target", "Templete_File_chosen.xlsm")
template_path = os.path.join(BASE_PATH, "templates", "Template.xlsm")
target_table_identify_path = os.path.join(BASE_PATH, "target", "Format_list.xlsx")


def generate_format(target_file_path):
    """
    1. target_file_pathの「●」で指定された行からfileidと抽出対象のtablenameを取得
    2. target_file_listのfileidとtablenameに基づいて、指定されたテーブルのフォーマットを抽出
    3. 抽出した内容をtemplateに追加し、新しいexcelとしてoutput_pathに出力

    Args:
        target_file_path: FileID指定表のパス

    Returns:
        None
    """
    try:
        target_file_list, filename_list = position_identify(
            target_file_path, excel_files_path
        )
        if not target_file_list:
            raise FileNotFoundError(f" {target_file} が見つからないです。")
        output_path = os.path.join(BASE_PATH, "output")
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        # output_folder_path = os.path.join(
        #     output_path, datetime.now().strftime("%Y%m%d%H%M")
        # )
        # if not os.path.exists(output_folder_path):
        #     os.makedirs(output_folder_path)

        for target_file, target_table_name in target_file_list.items():
            if "インターフェース" in str(target_file):
                process_interface_file(
                    target_file,
                    target_table_name,
                    output_path,
                    template_path,
                    target_table_identify_path,
                    new_file_name=filename_list[target_file],
                )
            elif "ファイル仕様" in target_file:

                process_file_file(
                    target_file,
                    target_table_name,
                    output_path,
                    template_path,
                    new_file_name=filename_list[target_file],
                )
    except Exception as e:
        print(f"Error: {e}")
    return


if __name__ == "__main__":
    generate_format(target_file_path)
