def G_content_extractor(
    target_row_list_by_loop,
    target_row_list_by_loop_content,
    header_columns,
    df,
    repeat_count,
):
    for index_row in target_row_list_by_loop:
        char_number = 9312
        loop_number = chr(char_number + repeat_count)  # 每次循环时递增
        for index_column, index_value in header_columns.items():
            if index_column == "項目名":
                content = str(df.iloc[index_row, index_value]) + loop_number
                target_row_list_by_loop_content[index_column].append(content)
            else:
                target_row_list_by_loop_content[index_column].append(
                    df.iloc[index_row, index_value]
                )
    return target_row_list_by_loop_content
