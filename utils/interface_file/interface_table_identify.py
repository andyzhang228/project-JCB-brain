import pandas as pd


def target_table_identify(target_file_path, format):
    try:
        format_column_index = None
        format_row_index = None
        target_table_name = None
        df = pd.read_excel(target_file_path, header=None)
        for index_row, row in df.iterrows():
            for column_index, column in enumerate(row):
                if column == format:
                    format_row_index = index_row
                    format_column_index = column_index
                    break
        target_table_name = df.iloc[format_row_index, format_column_index + 1]
        return target_table_name
    except Exception as e:
        print(f"エラー：{format}は{target_file_path}で見つかりませんでした")
        return None
