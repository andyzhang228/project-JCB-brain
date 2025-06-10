import pandas as pd


def interface_extractor_tool(
    target_file_name: str, target_sheet_name: str, target_table_name: str
):
    # print(target_file_name)
    target_file = "./excel_files/" + target_file_name
    df = pd.read_excel(target_file, sheet_name=target_sheet_name, header=None)

    target_table = False
    header_start_row_index = None
    header_end_row_index = None
    table_content_row = None
    table_header = []
    header_columns = {}

    for index_row, row in df.iterrows():
        for column in row:
            if column == target_table_name:
                target_name_row_index = index_row
                header_start_row_index = target_name_row_index + 1
                header_end_row_index = target_name_row_index + 2
                table_content_row = target_name_row_index + 3
                target_table = True
                break

    # ヘッダーの2行を走査し、ヘッダーと列インデックスを同時に収集
    for index_row, row in df.iloc[
        header_start_row_index : header_end_row_index + 1
    ].iterrows():
        for index_column, column in enumerate(row):
            # print(column)
            if column != "書式" and pd.notna(column):
                if column not in table_header:
                    table_header.append(column)
                    header_columns[column] = index_column
    # print(header_columns)

    filter_by_type = {filter_type: [] for filter_type in table_header}
    for header_item in table_header:
        if header_item == "Byte":
            col_index = header_columns[header_item] + 1
        elif header_item == "OFFSET":
            col_index = header_columns[header_item] + 1
        else:
            col_index = header_columns[header_item]  # 列インデックスを取得

        current_target = filter_by_type[header_item]

        for index_row, row in df.iloc[table_content_row:].iterrows():
            if all(
                pd.isna(row.iloc[i]) or row.iloc[i] == "" or row.iloc[i] == "NA"
                for i in range(len(df.columns))
            ):
                break
            else:
                if pd.isna(row.iloc[col_index]):  # 列インデックスを直接使用
                    current_target.append("")  # 空値の場合
                else:
                    # 元の空白数を保持
                    current_target.append(str(row.iloc[col_index]))
        current_target.pop()

    return filter_by_type, table_header
