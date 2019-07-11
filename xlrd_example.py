import xlrd
import pprint

wb = xlrd.open_workbook('data/src/sample.xlsx')

print(type(wb))

print(wb.sheet_names())
