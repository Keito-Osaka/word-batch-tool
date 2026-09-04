import argparse
import json
import os
import sys
import traceback

import pandas as pd

import core


def write_json_to_stdout(value):
    """JSONをUTF-8で標準出力へ書き込む。"""
    data = json.dumps(
        value,
        ensure_ascii=False,
    ).encode("utf-8")

    sys.stdout.buffer.write(data)
    sys.stdout.buffer.write(b"\n")
    sys.stdout.buffer.flush()


def write_error_to_stderr(error):
    """例外情報をUTF-8のJSONとして標準エラーへ書き込む。"""
    error_data = json.dumps(
        {
            "error": str(error),
            "traceback": traceback.format_exc(),
        },
        ensure_ascii=False,
    ).encode("utf-8")

    sys.stderr.buffer.write(error_data)
    sys.stderr.buffer.write(b"\n")
    sys.stderr.buffer.flush()


def read_request_from_stdin():
    """Tauri側から渡されたUTF-8のJSONを読み取る。"""
    raw_input = sys.stdin.buffer.read()

    if not raw_input:
        raise ValueError("処理に必要な入力データが渡されていません。")

    try:
        return json.loads(raw_input.decode("utf-8"))
    except UnicodeDecodeError as error:
        raise ValueError(
            "入力データをUTF-8として読み込めませんでした。"
        ) from error
    except json.JSONDecodeError as error:
        raise ValueError(
            "入力データのJSON形式が正しくありません。"
        ) from error


def row_to_dict(row):
    """DataFrameの1行をJSON化可能な辞書へ変換する。"""
    return {
        str(key): "" if value is None else str(value)
        for key, value in row.to_dict().items()
    }


def validate_existing_file(file_path, label):
    """入力ファイルのパスと存在を確認する。"""
    if not file_path:
        raise ValueError(f"{label}が指定されていません。")

    if not os.path.isfile(file_path):
        raise FileNotFoundError(
            f"{label}が見つかりません。\n\n{file_path}"
        )


def validate_output_directory(output_path):
    """出力先を確認し、存在しない場合は作成する。"""
    if not output_path:
        raise ValueError("出力先フォルダが指定されていません。")

    os.makedirs(output_path, exist_ok=True)

    if not os.path.isdir(output_path):
        raise NotADirectoryError(
            f"出力先がフォルダではありません。\n\n{output_path}"
        )


def load_raw_dataframe(file_path):
    """ExcelまたはCSVを、行除外前のDataFrameとして読み込む。"""
    validate_existing_file(file_path, "置換データファイル")

    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        encodings = [
            "utf-8-sig",
            "utf-8",
            "cp932",
        ]
        last_error = None

        for encoding in encodings:
            try:
                dataframe = pd.read_csv(
                    file_path,
                    dtype=str,
                    encoding=encoding,
                )
                break
            except UnicodeDecodeError as error:
                last_error = error
        else:
            raise ValueError(
                "CSVの文字コードを判定できませんでした。"
            ) from last_error

        dataframe = dataframe.dropna(axis=1, how="all")
        dataframe.columns = core.deduplicate_columns(
            [str(column).strip() for column in dataframe.columns]
        )

    elif extension in [".xlsx", ".xlsm", ".xls"]:
        if extension in [".xlsx", ".xlsm"]:
            engine = "openpyxl"
        else:
            engine = "xlrd"

        dataframe = pd.read_excel(
            file_path,
            sheet_name=0,
            header=0,
            dtype=str,
            engine=engine,
        )

        dataframe = dataframe.dropna(axis=1, how="all")
        dataframe.columns = core.deduplicate_columns(
            [
                core.normalize_excel_column_name(column)
                for column in dataframe.columns
            ]
        )

    else:
        raise ValueError(
            "対応していないファイル形式です。"
            "CSVまたはExcelファイルを指定してください。"
        )

    dataframe = dataframe.loc[
        :,
        [
            column
            for column in dataframe.columns
            if str(column).strip() != ""
        ],
    ]

    if len(dataframe.columns) == 0:
        raise ValueError(
            "置換データに使用できる列がありません。"
        )

    return dataframe


def get_exclusion_reason(
    row,
    row_exclude_mode,
    target_column_number,
):
    """除外された行について、画面表示用の理由を作成する。"""
    values = list(row.values)
    columns = list(row.index)

    if row_exclude_mode == "selected_column_number_empty":
        try:
            column_number = int(target_column_number)
        except (TypeError, ValueError) as error:
            raise ValueError(
                "除外判定に使う列番号は半角数字で指定してください。"
            ) from error

        index = column_number - 1

        if index < 0 or index >= len(columns):
            return "指定した列番号がデータの列数を超えています"

        return f"「{columns[index]}」が空欄"

    empty_columns = []

    for index, value in enumerate(values):
        if (
            row_exclude_mode == "any_empty_except_first"
            and index == 0
        ):
            continue

        if core.is_empty_cell_for_row_exclusion(value):
            empty_columns.append(str(columns[index]))

    if empty_columns:
        displayed_columns = empty_columns[:3]
        reason = "、".join(
            f"「{column}」"
            for column in displayed_columns
        )

        if len(empty_columns) > len(displayed_columns):
            reason += "ほか"

        if row_exclude_mode == "all_empty_except_first":
            return f"{reason}がすべて空欄"

        return f"{reason}が空欄"

    return "現在の除外条件に一致"


def inspect_data(request):
    """置換データを解析し、使用行・除外行・件数を返す。"""
    file_path = request.get("data_path", "")

    dataframe = load_raw_dataframe(file_path)
    dataframe = core.clean_dataframe_values(dataframe)
    dataframe = core.format_bank_account_columns(dataframe)
    dataframe = core.format_amount_columns(
        dataframe,
        enabled=request.get(
            "format_amount_with_comma",
            True,
        ),
        include_keywords=request.get(
            "amount_include_keywords"
        ),
        exclude_keywords=request.get(
            "amount_exclude_keywords"
        ),
    )

    original_count = len(dataframe)
    first_column = dataframe.columns[0]

    example_mask = (
        dataframe[first_column]
        .astype(str)
        .str.contains("例", na=False)
    )

    example_count = int(example_mask.sum())
    candidate_dataframe = dataframe[~example_mask].copy()

    row_exclude_mode = request.get(
        "row_exclude_mode",
        core.DEFAULT_ROW_EXCLUDE_MODE,
    )
    target_column_number = request.get(
        "row_exclude_target_column_number",
        core.DEFAULT_ROW_EXCLUDE_TARGET_COLUMN_NUMBER,
    )

    included_rows = []
    excluded_rows = []

    for source_index, row in candidate_dataframe.iterrows():
        should_exclude = core.should_exclude_row_by_mode(
            row,
            row_exclude_mode,
            target_column_number,
        )

        converted_row = row_to_dict(row)

        if should_exclude:
            excluded_rows.append(
                {
                    "source_row_number": int(source_index) + 2,
                    "row": converted_row,
                    "reason": get_exclusion_reason(
                        row,
                        row_exclude_mode,
                        target_column_number,
                    ),
                }
            )
        else:
            included_rows.append(converted_row)

    return {
        "original_count": original_count,
        "example_count": example_count,
        "excluded_count": len(excluded_rows),
        "included_count": len(included_rows),
        "columns": [
            str(column)
            for column in dataframe.columns
        ],
        "included_rows": included_rows[:500],
        "excluded_rows": excluded_rows[:500],
    }


def generate_documents(request):
    """個別Wordファイルを生成する。"""
    template_path = request.get("template_path", "")
    data_path = request.get("data_path", "")
    output_path = request.get("output_path", "")

    validate_existing_file(
        template_path,
        "テンプレートWord",
    )
    validate_existing_file(
        data_path,
        "置換データファイル",
    )
    validate_output_directory(output_path)

    dataframe = core.prepare_dataframe(
        data_path,
        format_amount_with_comma=request.get(
            "format_amount_with_comma",
            True,
        ),
        amount_include_keywords=request.get(
            "amount_include_keywords"
        ),
        amount_exclude_keywords=request.get(
            "amount_exclude_keywords"
        ),
        row_exclude_mode=request.get(
            "row_exclude_mode",
            core.DEFAULT_ROW_EXCLUDE_MODE,
        ),
        row_exclude_target_column_number=request.get(
            "row_exclude_target_column_number",
            core.DEFAULT_ROW_EXCLUDE_TARGET_COLUMN_NUMBER,
        ),
    )

    core.generate_documents_from_template(
        template_path=template_path,
        df=dataframe,
        filename_keys=request.get(
            "filename_keys",
            [],
        ),
        output_dir_path=output_path,
        add_serial_number=request.get(
            "add_serial_number",
            False,
        ),
        serial_digits=request.get(
            "serial_digits",
            2,
        ),
    )

    return {
        "success": True,
        "output_path": output_path,
        "generated_count": len(dataframe),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "command",
        choices=[
            "inspect",
            "generate",
        ],
    )
    args = parser.parse_args()

    try:
        request = read_request_from_stdin()

        if args.command == "inspect":
            result = inspect_data(request)
        else:
            result = generate_documents(request)

        write_json_to_stdout(result)

    except Exception as error:
        write_error_to_stderr(error)
        sys.exit(1)


if __name__ == "__main__":
    main()
