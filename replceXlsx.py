import pandas as pd
import xlsxwriter, openpyxl
from openpyxl import load_workbook


from pandas import ExcelWriter
from pandas import ExcelFile

df1 = pd.read_excel('testFiles/basicTest.xlsx', sheet_name='report')
df2 = pd.read_excel('testFiles/basicTest.xlsx', sheet_name='report2')
df3 = pd.read_excel('testFiles/basicTest.xlsx', sheet_name='report21')


for i in df2.index:
    print(df2['name'][i])

for i in df2.index:
    df2['name'][i]='B'+df2['name'][i]
    print(df2['name'][i])

writer_orig = pd.ExcelWriter('testFiles/basicTest.xlsx', engine='openpyxl')
df1.to_excel(writer_orig, index=False, sheet_name='report')
df2.to_excel(writer_orig, index=False, sheet_name='report2')
df3.to_excel(writer_orig, index=False, sheet_name='report21')
writer_orig.save()

"""
book = load_workbook('Masterfile.xlsx')
writer = pandas.ExcelWriter('Masterfile.xlsx', engine='openpyxl')
writer.book = book

writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
data_filtered.to_excel(writer, "Main", cols=['Diff1', 'Diff2'])
writer.save()
"""
