# -*- coding: utf-8 -*-
"""
Created on Wed May 26 22:53:24 2021

@author: Yz Wu
"""

from openpyxl import Workbook
# 创建一个工作簿对象
wb = Workbook()
#输入箱号的范围
for i in range(1,7) :
    wb.create_sheet(title='箱号'+str(i))
    
#删除多余的"Sheet"页面
del wb["Sheet"]
#箱号的下限，为了起名用
j=1
#工作簿名称
WorkbookName='箱号'+str(j)+'-'+str(i)+'.xlsx'
#保存工作薄
wb.save(WorkbookName)
# 最后关闭文件
wb.close()