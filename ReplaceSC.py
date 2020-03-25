import os, re, openpyxl
import pandas as pd

path = os.getcwd()
files = os.listdir('testFiles/single')
print(files)

files_xls = [f for f in files if f[-4:] == 'xlsx']
print(files_xls)


df = pd.DataFrame()
for f in files_xls:
    data = pd.read_excel('testFiles/single/'+f, sheet_name='資料')
    data1 = pd.read_excel('testFiles/single/'+f, sheet_name='版本')
    data3 = pd.read_excel('testFiles/single/'+f, sheet_name='計算點')
    data4 = pd.read_excel('testFiles/single/'+f, sheet_name='比對')
    data5 = pd.read_excel('testFiles/single/'+f, sheet_name='剔除比對')
    for i in data.index:
            if 'SC' in str(data['B3 Name'][i]):
                data['B3 Name'][i]='B'+data['B3 Name'][i]


    writer_orig = pd.ExcelWriter('testFiles/single/'+f, engine='openpyxl')
    data1.to_excel(writer_orig, index=False, sheet_name='版本')
    data.to_excel(writer_orig, index=False, sheet_name='資料')
    data3.to_excel(writer_orig, index=False, sheet_name='計算點')
    data4.to_excel(writer_orig, index=False, sheet_name='比對')
    data5.to_excel(writer_orig, index=False, sheet_name='剔除比對')
    writer_orig.save()
