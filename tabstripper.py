import xlrd
import os
import openai
import xlsxwriter

def func(value):
    return ''.join(value.splitlines())

dataframe = xlrd.open_workbook("finalGeneratedArticles.xlsx")
dataframe = dataframe.sheet_by_index(0)

workbook = xlsxwriter.Workbook('unformattedtext.xlsx')
worksheet = workbook.add_worksheet()


for i in range(0, 120):
    value = dataframe.cell_value(i, 0)
    string = value.strip()
    worksheet.write(i, 0, func(string))
workbook.close()