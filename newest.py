# -*- coding: utf-8 -*-
"""
Created on Fri Jul 21 19:03:05 2023

@author: Administrator
"""

import pandas as pd
import os
import shutil

# 读取 Excel 文件
excel_path = 'F:\\project1\\副本2017年债权债务明细账 - 副本2.xls'
df_excel = pd.read_excel(excel_path, sheet_name=None)

# 遍历每个工作表
for sheet_name, df_sheet in df_excel.items():
    # 提取两个特定列的数值并转为数字对
    col1 = 'Unnamed: 1'  # 第一个特定列名
    col2 = 'Unnamed: 4'  # 第二个特定列名

    # 检查特定列名是否存在于工作表中
    if col1 in df_sheet.columns and col2 in df_sheet.columns:
        # 筛选出不为 NaN 或汉字且为数值的数据
        valid_values_col1 = df_sheet[col1].dropna().loc[df_sheet[col1].apply(lambda x: isinstance(x, (int, float)) and not isinstance(x, bool))]
        valid_values_col2 = df_sheet[col2].dropna().loc[df_sheet[col2].apply(lambda x: isinstance(x, (int, float)) and not isinstance(x, bool))]
        # 创建和工作表同名的文件夹
        folder_name = f"{sheet_name}"
        os.makedirs(folder_name, exist_ok=True)
        # 遍历并处理这两个特定列的有效数值
        for value1, value2 in zip(valid_values_col1, valid_values_col2):
            # 进行你需要的操作，比如复制文件等
            print(f"工作表 '{sheet_name}' 中的有效数值对：{value1}, {value2}")
            num1_str = str(value1)
            num2_str = str(value2)
            # 在相应的文件夹中搜索并复制文件
            file_name= f"{num2_str}.pdf"  # 文件名格式示例
            source_path = os.path.join('F:\\project1\\广田2017-分拆凭证', f"{num1_str}") # 文件搜索路径
            target_path = os.path.join(folder_name, file_name)
            print(source_path)
            if os.path.exists(os.path.join(source_path, file_name)):
                shutil.copy(os.path.join(source_path, file_name), target_path)

# 完成操作后输出提示信息
print("文件复制完成！")
