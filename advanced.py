# -*- coding: utf-8 -*-
"""
Created on Mon Jul 17 22:38:16 2023

@author: Administrator
"""
import os
import pandas as pd

# Excel 文件路径
excel_path = 'F:\\project1\\副本2017年债权债务明细账 - 副本.xls'

# 读取 Excel 文件
excel_data = pd.read_excel(excel_path, sheet_name=None)

# 获取工作表名称列表
sheet_names = list(excel_data.keys())
# 将列名转换为列表
column_names = list(excel_data.column)
# 将行名转换为列表
row_names = list(excel_data.index)
# 遍历工作表名称列表，创建文件夹
for sheet_name in sheet_names:
    # 创建文件夹路径
    folder_path = os.path.join(os.getcwd(), sheet_name)
    
    # 创建文件夹
    os.makedirs(folder_path, exist_ok=True)

    # 可选：打印创建的文件夹路径
    print(f"Created folder: {folder_path}")

