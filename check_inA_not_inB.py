# -*- coding: utf-8 -*-
"""
Created on Mon Jun 17 16:50:37 2024

@author: liuchangjun

@email: 2857418430@qq.com
"""

import pandas as pd 
import xlwings as xw
import numpy as np

excel_path = 'D:/Hua/My_Own_Utilities/arxml/arxml_py/Mapping.xlsx'
app = xw.App(visible=False, add_book=False)
wb = app.books.open(excel_path)
sheet_ = wb.sheets[0]

df = pd.read_excel(excel_path)
df1 = list(df['A'])
df2 = df['B']
df2 = list(df2.dropna())

n = 0
not_found = []
flag = list(np.zeros(len(df1)))
for ind,i in enumerate(df1):
    for j in df2:
        if i in j:
            flag[ind] = 1
            n+=1
            break
        
        
sheet_.range('C2').options(transpose=True).value = flag
sheet_.range('A1').expand(mode='table').columns.autofit()
wb.save()
app.quit()