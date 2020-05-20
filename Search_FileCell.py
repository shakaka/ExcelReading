import os, re, sys
import pandas as pd


path = os.getcwd()
files = os.listdir('./search')
#print(files)
print('Please enter which sheet you want to search. (Default: 資料)')
inpu1=input()
print('Please enter which column you want to search. (Default: B3 Name)')
inpu2=input()
print('Please enter what you want to seach. (Default: SC)')
inpu3=input()
inpu1 = '資料' if inpu1 == '' else inpu1
inpu2 = 'B3 Name' if inpu2 == '' else inpu2
inpu3 = 'SC' if inpu3 == '' else inpu3
print('Search in '+inpu1+' sheet and '+inpu2+' column and for'+inpu3)

files_xls = [f for f in files if f[-4:] == 'xlsx']
for f in files_xls:
    data = pd.read_excel('./search/'+f, sheet_name=inpu1)
    for i in data.index:
            if inpu3 in str(data[inpu2][i]):
                print(f, ': [',inpu2,',',i+2,']' )
