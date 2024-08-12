#!/usr/bin/env python3
"""
Author: Wade Zhong wzhong@hso.com
Date: 2024-08-09 14:24:35
LastEditTime: 2024-08-09 16:46:40
LastEditors: Wade Zhong wzhong@hso.com
Description: 通过读取指定路径的Excel，得到笛卡尔积数据
FilePath: \JavaScriptWPSMacros\AllSheetsCombinationsJoin\Multiple-Config-Tables-Join.py
Copyright (c) 2024 by Wade Zhong wzhong@hso.com, All Rights Reserved. 
"""

# 方案一：报错  Unable to allocate 52.8 GiB for an array with shape (7081992000,) and data type int64
# import pandas as pd
# import itertools
# from openpyxl import load_workbook

# def read_excel_sheets(file_path):
#     # 读取Excel文件中的所有sheet
#     xls = pd.ExcelFile(file_path)
#     sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
#     return sheets

# def add_sheet_name_as_header(sheets):
#     # 将sheet的名字添加成第一行，当成标题
#     for sheet_name, df in sheets.items():
#         df.columns = [f"{sheet_name}_{col}" for col in df.columns]
#     return sheets

# def cartesian_product(df_list):
#     # 计算多个DataFrame的笛卡尔积
#     result = df_list[0]
#     for df in df_list[1:]:
#         result = result.merge(df, how='cross')
#     return result

# def main(file_path):
#     sheets = read_excel_sheets(file_path)
#     sheets = add_sheet_name_as_header(sheets)
#     df_list = list(sheets.values())
#     result_df = cartesian_product(df_list)

#     book = load_workbook(file_path)
#     with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
#         writer.book = book
#         if 'Result' in writer.book.sheetnames:
#             # 如果Result sheet已经存在，先清空数据
#             writer.book.remove(writer.book['Result'])
#             writer.book.create_sheet('Result')
#         result_df.to_excel(writer, sheet_name='Result', index=False)

# if __name__ == "__main__":
#     file_path = '价目表用例数据_Python.xlsm'  # 指定你的Excel文件路径
#     main(file_path)


# 方案二import pandas as pd
import pandas as pd
from itertools import product
import os
import time
from tqdm import tqdm
import math
import traceback


def read_excel(file_path):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            sheet_data = pd.read_excel(file_path, sheet_name=None, header=None)
            sheet_data = {
                sheet_name: data.values.tolist()
                for sheet_name, data in sheet_data.items()
                if sheet_name != "Result"
            }
            return sheet_data
        except PermissionError:
            if attempt < max_attempts - 1:
                print(
                    f"文件访问被拒绝。正在尝试重新访问...（尝试 {attempt + 1}/{max_attempts}）"
                )
                time.sleep(2)  # 等待2秒后重试
            else:
                raise


def calculate_total_combinations(sheet_data):
    lengths = {sheet_name: len(data) for sheet_name, data in sheet_data.items()}
    total_combinations = math.prod(lengths.values())
    lengths_str = ", ".join(
        [f"{sheet_name}: {length}" for sheet_name, length in lengths.items()]
    )
    calculation_logic = " * ".join(map(str, lengths.values()))
    print(
        f"每个 sheet 的长度: {lengths_str}\n计算逻辑: {calculation_logic}\n即将生成的笛卡尔积总数: {total_combinations}"
    )
    return total_combinations


def cartesian_product(sheet_data, total_combinations):
    sheets = list(sheet_data.keys())
    all_rows = [sheet_data[sheet] for sheet in sheets]

    # 准备结果，首先添加表头
    result = [
        sum(
            (
                tuple(sheet for _ in range(len(sheet_data[sheet][0])))
                for sheet in sheets
            ),
            (),
        )
    ]

    # 使用tqdm创建进度条
    with tqdm(total=total_combinations, desc="处理进度") as pbar:
        for row in product(*all_rows):
            result.append(sum((tuple(r) for r in row), ()))
            pbar.update(1)

    return result


def update_result_sheet(file_path, result_data):
    with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace") as writer:
        result_df = pd.DataFrame(result_data)
        result_df.to_excel(writer, sheet_name="Result", index=False)


def main(file_path):
    try:
        start_time = time.time()

        sheet_data = read_excel(file_path)
        if not sheet_data:
            print("Excel文件中没有找到有效的数据。")
            return

        total_combinations = calculate_total_combinations(sheet_data)

        result = cartesian_product(sheet_data, total_combinations)
        update_result_sheet(file_path, result)

        end_time = time.time()
        execution_time = end_time - start_time

        print(f"处理完成。结果已保存到 {file_path} 的 'Result' sheet 中。")
        print(f"总执行时间: {execution_time:.2f} 秒")
    except PermissionError:
        print("无法访问或保存文件。请确保文件未被其他程序打开，并且您有足够的权限。")
    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")
        traceback.print_exc()


if __name__ == "__main__":
    file_path = input("请输入Excel文件的路径: ")
    if not os.path.exists(file_path):
        print("文件不存在，请检查路径是否正确。")
    else:
        main(file_path)
