import os, re, sys
import pandas as pd
print('Please enter which sheet you want to search. (Default: 資料)')
inpu1=input()
print('Please enter which column you want to search. (Default: B3 Name)')
inpu2=input()
print('Please enter what you want to seach. (Default: SC)')
inpu3=input()

path = os.getcwd()
files = os.listdir('./search')
#print(files)

inpu1 = '資料' if inpu1 == '' else inpu1
inpu2 = 'B3 Name' if inpu2 == '' else inpu2
inpu3 = 'SC' if inpu3 == '' else inpu3


files_xls = [f for f in files if f[-4:] == 'xlsx']
#print(files_xls)


df = pd.DataFrame()
for f in files_xls:
    data = pd.read_excel('./search/'+f, sheet_name=inpu1)
    for i in data.index:
            if inpu3 in str(data[inpu2][i]):
                print(f, ': [',inpu2,',',i+2,']' )
