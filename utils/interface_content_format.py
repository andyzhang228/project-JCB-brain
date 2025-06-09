import re


def interface_content_format(filter_by_type):
    comment_index = str("処理・備考")
    binary = str("Binary").lower()
    binary_jp = str("バイナリ")
    filter_by_type["Byte2"] = [int(i) for i in filter_by_type["Byte"]]
    for key, value in filter_by_type.items():
        if key == "タイプ":
            for i in range(len(value)):
                if value[i] == "半角" and (
                    "Binary" in filter_by_type["処理・備考"][i]
                    or "バイナリ" in filter_by_type["処理・備考"][i]
                ):
                    value[i] = "B"
                elif value[i] == "半角":
                    value[i] = "X"
                elif value[i] == "半角数字":
                    value[i] = "9"
                else:
                    value[i] = value[i]
        elif key == "Byte2":
            pattern = r"\b[A-Z]*\d+V\d+\b"
            for i in range(len(value)):
                text = filter_by_type["処理・備考"][i]
                match = re.search(pattern, text)
                if match:
                    value[i] = match.group()
                else:
                    value[i] = value[i]

    # print(filter_by_type)
    formated_filter_by_type = filter_by_type
    return formated_filter_by_type
