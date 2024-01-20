# -*- coding: utf-8 -*-
"""
Created on Mon Jul 17 22:36:32 2023

@author: Administrator
"""

import os
import shutil
import pandas as pd

# 提取Excel文档中特定列内的数字值
def extract_numbers_from_excel(excel_path, column_name):
    df = pd.read_excel(excel_path)
    numbers = df[column_name].tolist()
    return [str(num) for num in numbers if pd.notnull(num)]

# 根据两个数字值查找特定路径内的文件，并复制同名文件
def find_and_copy_files(source_directory, destination_directory, numbers):
    for number in numbers:
        file_name = f"{number}.pdf"  # 假设要查找的文件为 .txt 文件
        source_path = os.path.join(source_directory, file_name)
        destination_path = os.path.join(destination_directory, file_name)
        if os.path.exists(source_path):
            shutil.copy(source_path, destination_path)

# 根据第一个数字值创建一个母文件夹，并将复制的文件粘贴至文件夹中
def create_folder_and_paste_files(destination_directory, folder_name):
    folder_path = os.path.join(destination_directory, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    files = os.listdir(destination_directory)
    for file in files:
        if os.path.isfile(os.path.join(destination_directory, file)):
            shutil.move(os.path.join(destination_directory, file), os.path.join(folder_path, file))

# Excel 文件路径
excel_path = 'F:\\project1\\副本2017年债权债务明细账 - 副本2.xls'

# 特定列名称
column_name ='Unnamed: 1'

# 数字值列表
numbers = extract_numbers_from_excel(excel_path, column_name)

# 源文件夹路径
source_directory = 'F:\\project1\\广田2017-分拆凭证'

# 目标文件夹路径
destination_directory = 'F:\\project1\\file1'

# 第一个数字值创建的母文件夹名称
folder_name = numbers[0]

# 根据两个数字值查找特定路径内的文件，并复制同名文件
find_and_copy_files(source_directory, destination_directory, numbers)

# 根据第一个数字值创建一个母文件夹，并将复制的文件粘贴至文件夹中
create_folder_and_paste_files(destination_directory, folder_name)
