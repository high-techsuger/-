# -*- coding: utf-8 -*-
"""
Created on Sat Jul 15 22:11:05 2023

@author: Administrator
"""

import os

# 指定要重命名文件的目录
directory = 'F:\project1\新建文件夹 (2)'


# 遍历目录中的文件
for filename in os.listdir(directory):
    # 构建旧的文件路径
    old_path = os.path.join(directory, filename)
    
    # 构建新的文件名
    new_name = filename.replace('0', '')
   # new_name = filename+'.pdf'
    # 构建新的文件路径
    new_path = os.path.join(directory, new_name)
    
    # 重命名文件
    os.rename(old_path, new_path)

