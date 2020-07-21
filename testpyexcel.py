

import xlrd
from datetime import date,datetime




wb = xlrd.open_workbook('f:\PythonStudy\work\北京9月考勤数据.xls')

sheet_names = wb.sheet_names()
sheets_count = len(sheet_names)

#常用单元格中的数据类型： 0 empty,1 string（text）, 2 number, 3 date, 4 boolean, 5 error， 6 blank
sheet_names = wb.sheet_names()
#print (len(sheet_names),sheet_names)

ws = wb.sheet_by_index(0)
wl = ws.row_values(1)
print (wl[3])
NStr = wl[3]

#dt = datetime.strptime('2017-11-11 18:58:39','%Y-%m-%d %H:%M:%S')
#print (dt)

#print (datetime.strptime(NStr, '%Y-%m-%d %H:%M:%S') )

DateTime = datetime.strptime(NStr, '%Y-%m-%d %H:%M:%S')

print (DateTime.timetuple())

#print (datetime.now().strftime('%b-%d-%y %H:%M:%S'))
#print (datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
#print ('timestamp',datetime.now().timestamp())
#print ('timetuple',datetime.now().timetuple())
#print (datetime.now())



#print (datetime.time.strptime("2017-11-11 18:58:39", "%Y-%m-%d %H:%M:%S"))
#print (datetime.time.strptime(wl[3](), "%Y-%m-%d %H:%M:%S") )

#format_time = datetime.time.strptime("2017-11-11 18:58:39", "%Y-%m-%d %H:%M:%S") 
#print time.mktime(format_time)



#format_time = datetime.time.strptime(wl[3], "%Y-%m-%d %H:%M:%S") 

#print (format_time)

#print datetime.time.mktime(format_time)

'''
for sn in sheet_names:
    ws = wb.sheet_by_name(sn)
    print (sn,ws.nrows,ws.ncols)

 
    rx = ws.nrows
    cx = ws.ncols
    r = 0
    c = 0

    print (ws.row(0))
    print (ws.row(1))
    #print (ws.row_slice(0))
    #print (ws.row_slice(1))
    print (ws.row_types(0))
    print (ws.row_types(1))
    print (ws.row_values(0))
    print (ws.row_values(1))
    print (ws.row_len(0))
    print (ws.row_len(1))

    
    #while r < rx:

        #nrow = ws.row(r)
        #print (ws.row(r))
        #print (ws.row_slice(r))

        
        #while c < cx:
        #    print (ws.row(r)[c])
        #    c += 1
        
    #    r += 1
'''

'''
i = 0
while i < sheets_count:
    print (sheet_names[i])
    i += 1
'''
#get an array of sheets' names 
#names = wb.sheet_names()
#ws = wb.sheet_by_index(0)
#print(ws.name, ws.nrows, ws.ncols)
#print(names)
#print(len(names))

'''
ws = wb.sheet_by_name('Sheet1')
print(ws.name, ws.nrows, ws.ncols)
ws = wb.sheet_by_name('Sheet2')
print(ws.name, ws.nrows, ws.ncols)
'''


