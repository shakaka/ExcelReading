import os
import pandas as pd

path = os.getcwd()
files = os.listdir('testFiles')
print(files)

files_xls = [f for f in files if f[-4:] == 'xlsx']
print(files_xls)


df = pd.DataFrame()
for f in files_xls:
    data = pd.read_excel('testFiles/'+f)
    df = df.append(data)
print(df)
