import os, re
import pandas as pd

path = os.getcwd()
files = os.listdir('C:/Users/Richard/Desktop/Database/嘉南AD_20200317')
#print(files)

files_xls = [f for f in files if f[-4:] == 'xlsx']
#print(files_xls)


df = pd.DataFrame()
for f in files_xls:
    data = pd.read_excel('C:/Users/Richard/Desktop/Database/嘉南AD_20200317/'+f, sheet_name='資料')
    for i in data.index:
            if 'SC' in str(data['B3 Name'][i]):
                print(f)
                break
    df = df.append(data)
