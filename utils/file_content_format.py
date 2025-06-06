from utils.file_extractor import file_extractor_tool
from utils.G_content_extractor import G_content_extractor
import pandas as pd


def file_content_format(
    filter_by_type, target_file_name, target_sheet_name, header_columns, table_header
):
    target_row_list_by_loop_content = {i: [] for i in table_header}
    number = 0
    G_index = 0
    for key, value in filter_by_type.items():
        if key == "属性":
            for i in range(len(value)):
                if value[i] == "G":
                    value[i] = ""
                    filter_by_type["バイト数"][i] = "0"
                elif "G*" in value[i]:
                    number = value[i].split("G*")[1]
                    # print(f"G*所在index:{i}")
                    G_index = i
                    # print(number)
                    target_file = "./excel_files/" + target_file_name
                    df = pd.read_excel(
                        target_file, sheet_name=target_sheet_name, header=None
                    )
                    for index_row, row in df.iterrows():
                        for index_column, column in enumerate(row):
                            if column == value[i]:
                                G_row_index = index_row
                                # print(f"G*所在index行:{G_row_index}")
                                G_column_index = index_column
                                # print(f"G*所在index列:{G_column_index}")
                                G_content_index = G_row_index + 1
                                # print(f"G*下行内容所在index:{G_content_index}")
                                break
                    level = df.iloc[G_row_index, header_columns["レベル"]]

                    # G*向下遍历，收集比其大的内容，并收集其内容
                    target_row_list_by_loop = []
                    for index_row, row in df.iloc[G_row_index + 1 :].iterrows():
                        if row[header_columns["レベル"]] <= level:
                            break
                        target_row_list_by_loop.append(index_row)

                    for repeat_count in range(int(number)):
                        G_content_extractor(
                            target_row_list_by_loop,
                            target_row_list_by_loop_content,
                            header_columns,
                            df,
                            repeat_count,
                        )
                else:
                    value[i] = value[i]

    for index, value in filter_by_type.items():
        del filter_by_type[index][G_index : G_index + len(target_row_list_by_loop) + 1]

    if target_row_list_by_loop_content:
        data_length = len(next(iter(target_row_list_by_loop_content.values())))

        for i in range(data_length):
            for index, value in target_row_list_by_loop_content.items():
                if index in filter_by_type:
                    filter_by_type[index].insert(G_index + i, value[i])

    filter_by_type["#"] = [
        i for i in range(1, len(next(iter(filter_by_type.values()))) + 1)
    ]
    filter_by_type["OFFSET"] = [
        0 for i in range(1, len(next(iter(filter_by_type.values()))) + 1)
    ]

    for num in range(len(filter_by_type["OFFSET"]) - 1):
        filter_by_type["OFFSET"][num + 1] = int(filter_by_type["OFFSET"][num]) + int(
            filter_by_type["バイト数"][num]
        )

    formated_filter_by_type = filter_by_type
    # print(formated_filter_by_type)
    return formated_filter_by_type
