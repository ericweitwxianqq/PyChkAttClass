
import xlrd

from datetime import date,datetime

file = 'f:\pythonstudy\work\北京9月考勤数据.xls'

#def read_excel():
wb = xlrd.open_workbook(file)

print(wb.sheet_names())

sheet1 = wb.sheet_by_index(0)

#print(sheet1)
print(sheet1.name,sheet1.nrows,sheet1.ncols)