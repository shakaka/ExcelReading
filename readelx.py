import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('testFiles/S35西濱臨.xlsx', sheet_name='資料')


for i in df.index:
    print(df['B2 Name'][i])
