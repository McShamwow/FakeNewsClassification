##import packages
import xlrd
import os
import xlsxwriter

##create a function to split the tabs and replaces with space
def func(value):
    return ''.join(value.splitlines())

##open the final articles excel file 
dataframe = xlrd.open_workbook("finalGeneratedArticles.xlsx")
dataframe = dataframe.sheet_by_index(0)

##open a new excel file to store the formatted articles
workbook = xlsxwriter.Workbook('unformattedtext.xlsx')
worksheet = workbook.add_worksheet()

##for each article in the range 
for i in range(0, 120):
    ##select the values in the final articles
    value = dataframe.cell_value(i, 0)
    ##strip the starting and ending new lines
    string = value.strip()
    ##run the article through the function and write to the formatted articles excel file
    worksheet.write(i, 0, func(string))
##close the workbook and commit changes
workbook.close()
