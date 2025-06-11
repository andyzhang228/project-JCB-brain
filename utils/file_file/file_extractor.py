import pandas as pd


def file_extractor_tool(target_file_name, target_sheet_name, target_table_name):
    target_file = "./excel_files/" + target_file_name
    df = pd.read_excel(target_file, sheet_name=target_sheet_name, header=None)
    header_start_row_index = None
    header_end_row_index = None
    table_content_row = None
    table_header = []
    header_columns = {}

    for index_row, row in df.iterrows():
        for column in row:
            if column == target_table_name:
                target_name_row_index = index_row
                header_start_row_index = target_name_row_index + 3
                header_end_row_index = target_name_row_index + 4
                table_content_row = target_name_row_index + 5
                break

    # ヘッダーの2行を走査し、ヘッダーと列インデックスを同時に収集
    for index_row, row in df.iloc[
        header_start_row_index : header_end_row_index + 1
    ].iterrows():
        for index_column, column in enumerate(row):
            if (
                pd.notna(column)
                and column not in table_header
                and not ("相対位置" in column or "10進(Hex)" in column)
                and not column == "数"
            ):
                new_column = str(column).strip().replace(" ", "")
                if new_column == "バイト":
                    new_column = "バイト数"
                table_header.append(new_column)
                header_columns[new_column] = index_column

    filter_by_type = {filter_type: [] for filter_type in table_header}
    for key, value in header_columns.items():
        col_index = header_columns[key]  # 列インデックスを取得

        for index_row, row in df.iloc[table_content_row:].iterrows():
            if all(
                pd.isna(row.iloc[i]) or row.iloc[i] == "" or row.iloc[i] == "NA"
                for i in range(len(df.columns))
            ):
                break
            else:
                if pd.isna(row.iloc[col_index]):  # 列インデックスを直接使用
                    filter_by_type[key].append("")  # 空値の場合
                else:
                    # 元の空白数を保持
                    filter_by_type[key].append(str(row.iloc[col_index]))
    return filter_by_type, table_header, header_columns
