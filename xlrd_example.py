# xls, xlsxの読み込みをおこなるライブラリ
import xlrd
# 結果を見やすくするもの
import pprint

# Bookオブジェクト
wb = xlrd.open_workbook('data/src/sample.xlsx')
print(type(wb))
print(wb.sheet_names())

# sheetクラス
sheets = wb.sheets()
print(type(sheets))
print(type(sheets[0]))

sheet = wb.sheet_by_name('sheet1')
print(type(sheet))

cell = sheet.cell(1, 2)
print(cell)
print(type(cell))
print(cell.value)

print(sheet.cell_value(1, 2))

col = sheet.col(1)

print(col)
print(type(col[0]))

col_values = sheet.col_values(1)
print(col_values)
print(sheet.row_values(1))

