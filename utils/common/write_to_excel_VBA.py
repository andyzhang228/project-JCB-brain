import os
import xlwings as xw


def write_to_excel_VBA(
    template_path, extracted_content, output_path, target_table_name, interface_or_file
):
    # ヘッダーマッピング辞書の作成 - ファイルフォーマット用
    if interface_or_file == "file":
        table_header_map = {
            "#": "#",
            f"FIELD NAME\n(項 目 名)": "項目名",
            f"SYMBOL NAME\n(シ ン ボ ル 名)": "シンボル名",
            f"ATTRIBUTES\n(属性)": "属性",
            f"NUMBER OF DIGITS\n(桁数)": "桁数",
            f"LENGTH\n(バイト数)": "バイト数",
            "OFFSET": "OFFSET",
            "備考": "説明",
        }

    elif interface_or_file == "interface":
        table_header_map = {
            "#": "No",
            f"FIELD NAME\n(項 目 名)": "項目名",
            f"SYMBOL NAME\n(シ ン ボ ル 名)": "項目ＩＤ",
            f"ATTRIBUTES\n(属性)": "タイプ",
            f"NUMBER OF DIGITS\n(桁数)": "Byte2",
            f"LENGTH\n(バイト数)": "Byte",
            f"OFFSET\n(オフセット)": "OFFSET",
            "備考": "処理・備考",
        }

    # keys_list = list(table_header_map.keys())
    # col_width = {key: 0 for key in keys_list}
    # new_table_header = list(table_header_map.keys())

    # thin_border = Border(
    #     left=Side(style="thin"),
    #     right=Side(style="thin"),
    #     top=Side(style="thin"),
    #     bottom=Side(style="thin"),
    # )
    print(output_path)
    if os.path.exists(output_path):
        xlsm_template = xw.Book(output_path)
    else:
        xlsm_template = xw.Book(template_path)
    # 重複書き込みを防ぎ、データを更新するため、シートが存在する場合は削除して再作成
    print(target_table_name)
    if target_table_name in xlsm_template.sheetnames:
        del xlsm_template[target_table_name]
