

import tkinter as tk
from tkinter import filedialog

window = tk.Tk()
#设置窗口名称
window.title('睿控考勤数据处理转换工具')
#设置窗口大小 (长 x 宽)
window.geometry('500x400')

l = tk.Label(window, text='       ', bg= 'green')

tk.Label(window, text='on the window', bg='red', font=('Arial', 16)).pack()  
# 和前面部件分开创建和放置不同，其实可以创建和放置一步完成
 
'''
frame = tk.Frame(window)

frame.pack()
frame_l = tk.Frame(frame)# 第二层frame，左frame，长在主frame上
frame_r = tk.Frame(frame)# 第二层frame，右frame，长在主frame上
frame_l.pack(side='left')
frame_r.pack(side='right')
 
# 第7步，创建三组标签，为第二层frame上面的内容，分为左区域和右区域，用不同颜色标识
tk.Label(frame_l, text='on the frame_l1', bg='green').pack()
tk.Label(frame_l, text='on the frame_l2', bg='green').pack()
tk.Label(frame_l, text='on the frame_l3', bg='green').pack()
tk.Label(frame_r, text='on the frame_r1', bg='yellow').pack()
tk.Label(frame_r, text='on the frame_r2', bg='yellow').pack()
tk.Label(frame_r, text='on the frame_r3', bg='yellow').pack()
'''

#输入 raw 档案






'''
from tkinter import *
# 导入ttk
from tkinter import ttk
# 导入filedialog
from tkinter import filedialog
class App:
    def __init__(self, master):
        self.master = master
        self.initWidgets()
    def initWidgets(self):
        # 创建7个按钮，并为之绑定事件处理函数
        ttk.Button(self.master, text='打开单个文件',
            command=self.open_file # 绑定open_file方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='打开多个文件',
            command=self.open_files # 绑定open_files方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='获取单个打开文件的文件名',
            command=self.open_filename # 绑定open_filename方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='获取多个打开文件的文件名',
            command=self.open_filenames # 绑定open_filenames方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='获取保存文件',
            command=self.save_file # 绑定save_file方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='获取保存文件的文件名',
            command=self.save_filename # 绑定save_filename方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        ttk.Button(self.master, text='打开路径',
            command=self.open_dir # 绑定open_dir方法
            ).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
    def open_file(self):
        # 调用askopenfile方法获取单个打开的文件
        print(filedialog.askopenfile(title='打开单个文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def open_files(self):
        # 调用askopenfile方法获取多个打开的文件
        print(filedialog.askopenfiles(title='打开多个文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def open_filename(self):
        # 调用askopenfilename方法获取单个文件的文件名
        print(filedialog.askopenfilename(title='打开单个文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def open_filenames(self):
        # 调用askopenfilenames方法获取多个文件的文件名
        print(filedialog.askopenfilenames(title='打开多个文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def save_file(self):
        # 调用asksaveasfile方法保存文件
        print(filedialog.asksaveasfile(title='保存文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def save_filename(self):
        # 调用asksaveasfilename方法获取保存文件的文件名
        print(filedialog.asksaveasfilename(title='保存文件',
            filetypes=[("文本文件", "*.txt"), ('Python源文件', '*.py')], # 只处理的文件类型
            initialdir='g:/')) # 初始目录
    def open_dir(self):
        # 调用askdirectory方法打开目录
        print(filedialog.askdirectory(title='打开目录',
            initialdir='g:/')) # 初始目录
root = Tk()
root.title("文件对话框测试")
App(root)
root.mainloop()
'''

'''
import os
from time import sleep
from tkinter import *
class Dirlist():
    def __init__(self,initdir=None,path=None,root=None):# 在主界面生成Dirlist对象时，把变量和控件对象传递过来。参数initdir为初始文件路径
        # 主界面中：变量：path = StringVar()   控件：filename = Entry(frame1,textvariable = path) 即在主界面中生成对象Dirlist('路径',path=path,root=filename,)

        self.top2 = Tk()
        self.top2.title('睿控考勤资料判读-请选择原始数据文件')
        self.path = path
        self.root = root

        #这里声明了Tk的一个变量cwd，用于保存当前所在的目录名，并用Label控件对象dirl显示出来（在调用函数doLS时显示出来）
        self.cwd = StringVar(self.top2)
        self.dirl = Label(self.top2,fg = 'blue')
        self.dirl.pack()

        # 第一个框架dirfm，应用核心部分，里面放置Listbox控件dirs，和Scrollbar滚动条。并通过使用List的bind()方法，将鼠标双击事件绑定
        self.dirfm = Frame(self.top2)
        self.dirsb = Scrollbar(self.dirfm)
        self.dirsb.pack(side=RIGHT,fill=Y)
        self.dirs = Listbox(self.dirfm,height=15,width=100,yscrollcommand=self.dirsb.set)

        #通过使用List的bind()方法，将鼠标双击事件绑定，并调用setDirAndGo函数
        self.dirs.bind('<Double-1>',self.setDirAndGo)

        # 下面实现单击时，将所选文件路径加名字更新到下方输入框控件中，不能用self.dirs.bind('<Button-1>', self.setDirn)绑定单击事件，会出错
        self.dirs.bind("<<ListboxSelect>>", self.setDirn)
        self.dirsb.config(command=self.dirs.yview)
        self.dirs.pack(side=LEFT,fill=BOTH)
        self.dirfm.pack()
        #界面下方输入框，随着单击更新内容，并在点击打开按钮时，打开输入框的内容目录
        self.dirn = Entry(self.top2,width=50,textvariable=self.cwd)
        #绑定回车事件，即当光标在输入框时，回车调用doLS函数
        self.dirn.bind('<Return>',self.doLS)
        self.dirn.pack()

        #第二个框架bfm，放置按钮
        self.bfm = Frame(self.top2)
        self.open = Button(self.bfm,text='打开',command=self.doLS,activeforeground='white',activebackground='blue')
        self.ls = Button(self.bfm,text='确认',command=self.result,activeforeground='white',activebackground='green')

        #没有用到退出按钮，故注释掉了
        # self.quit = Button(self.bfm,text='退出',command=self.top2.quit,activeforeground='white',activebackground='red')
        self.open.pack(side=LEFT)
        self.ls.pack(side=LEFT)
        # self.quit.pack(side=LEFT)
        self.bfm.pack()

        #若存在初始文件路径，则更新界面
        if initdir:
            self.cwd.set(initdir)
            self.top2.update()
            self.doLS()

    #当点击确认按钮时调用，将当前所在的目录名和文件名，更新到主界面，并关闭当前界面（在被调用的情况下点击确认可正常关闭）
    def result(self):
        tdir = self.cwd.get()
        print ('tdir=',tdir)
        #self.path.set(tdir)#更新主界面变量
        #self.root.update()#更新主界面
        #self.top2.destroy()#关闭当前界面

    #单击列表中内容时，调用此函数，更新下方输入框内容
    def setDirn(self,ev=None):
        t = self.dirs.get(self.dirs.curselection())
        print(t)
        text=os.getcwd()+'\\'+t  #文件目录和文件名
        self.dirn.delete(0, END)
        self.dirn.insert(INSERT, text)

    #双击时调用，双击时，设置背景色为红色，并调用doLS函数打开所选文件
    def setDirAndGo(self,ev=None):
        self.last = self.cwd.get()
        self.dirs.config(selectbackground='red')
        check = self.dirs.get(self.dirs.curselection())
        if not check:
            check = os.curdir
        self.cwd.set(check)
        self.doLS()

    #实现更新目录的核心函数
    def doLS(self,ev=None):
        error = ''
        tdir = self.cwd.get()
        if not tdir:tdir=os.curdir
        #若路径输入错误，或者打开的是文件，而不是目录，则更新错误提示信息
        if not os.path.exists(tdir):
            error = os.getcwd()+'\\'+tdir + '：未找到文件'
        elif not os.path.isdir(tdir):
            error = os.getcwd()+'\\'+tdir + '：未找到目录'
        if error:
            self.cwd.set(error)
            self.top2.update()
            sleep(1)
            if not (hasattr(self,'last') and self.last):
                self.last = os.curdir
            self.cwd.set(os.curdir)
            self.dirs.config(selectbackground='LightSkyBlue')
            self.dirn.config(text=os.getcwd()+'\\'+tdir)
            self.top2.update()
            return

        self.cwd.set(os.getcwd()+'\\'+tdir)
        self.top2.update()
        dirlist = os.listdir(tdir)#os.listdir() 方法用于返回指定的文件夹包含的文件或文件夹的名字的列表。
        dirlist.sort()
        os.chdir(tdir)#os.chdir() 方法用于改变当前工作目录到指定的路径。

        #更新界面上方标签内容
        self.dirl.config(text=os.getcwd())
        self.top2.update()

        self.dirs.delete(0,END)
        self.dirs.insert(END,os.pardir)#os.chdir(os.pardir)  切换到上级目录 即将上级目录.. 插入到dirs对象中

        #把选定目录的文件或文件夹的名字的列表依次插入到dirs对象中
        for eachFile in dirlist:
            self.dirs.insert(END,eachFile)
            self.cwd.set(os.curdir)
            self.dirs.config(selectbackground='LightSkyBlue')

if __name__ =='__main__':
    #设定初始目录为桌面
    d = Dirlist(r'.')
    mainloop()
'''
#########################################################################


import pandas as pd

import time
from datetime import datetime
import calendar


'''

def ifLate(office, name, dtm):
    timeNum = int(dtm.strftime("%H%M%S"))
    lateTime = 84000 # 08:40:00
    lateOffTime = 94000 # 09:40:00
    retStr = 'AM '
    if office == '北京':
        lateTime = 91000 # 09:10:00
        lateOffTime = 101000 # 10:10:00
    if timeNum > lateOffTime:
        retStr += '迟超60缺席'
        #retStr += dtm.strftime(" %H:%M:%S")
    elif timeNum > lateTime:
        retStr += '迟到'
        #retStr += dtm.strftime(" %H:%M:%S")
    else:
        retStr += '正常'
    retStr += dtm.strftime(" %H:%M:%S")
    retStr += '+'
    return retStr

def ifEarly(office, name, dtm): #return '正常', '早退', '事假半天', ’加班'

    timeNum = int(dtm.strftime("%H%M%S"))

    # datetime to time时间戳
    tmStr = dtm.strftime("%Y-%m-%d ")
              
    earlyTime = 173000 
    earlyOffTime = 163000
    overTime = 203000
    
    if office == '北京':
        tmStr += "18:00:00"
        earlyTime = 180000
        earlyOffTime = 170000
        overTime = 210000
    else:
         tmStr += "17:30:00"

    #print ('tmStr=',tmStr)
    

    retStr = 'PM '
    dtmStdOffDtm = datetime.strptime(tmStr, "%Y-%m-%d %H:%M:%S")

    if timeNum < earlyOffTime:
        retStr += '早退超60缺席'

        #retStr += dtm.strftime(" %H:%M:%S")
    elif timeNum < earlyTime:
        retStr += '早退'
        #retStr += dtm.strftime(" %H:%M:%S")
    elif timeNum > overTime:
        retStr += '加班'
        dltm = dtm - dtmStdOffDtm

        print ('dltm=',dltm,dtm,dtmStdOffDtm,retStr)
        print ('%d:%d:%d',int(dltm.seconds / 3600), int(dltm.seconds % 3600 / 60), int(dltm.seconds % 60))
        #retStr += '%d:%d:%d',dltm.seconds / 3600, dltm.seconds % 3600 / 60, dltm.seconds % 60
        
    else:
        retStr += '正常'

    retStr += dtm.strftime(" %H:%M:%S")
    return retStr

def ifAttendanceTwice(office, name, fstDtm, lstDtm):
    retStr = ''
    fstTime = int(fstDtm.strftime("%H%M%S"))
    lstTime = int(lstDtm.strftime("%H%M%S"))
    if fstTime > 120000:
        retStr += 'AM无打卡+'
    else:
        #判断上午状态
        retStr += ifLate(office, name, fstDtm)

    if lstTime < 120000:
        retStr += 'PM无打卡'
    else:
        #判断下午状态
        retStr += ifEarly(office, name, lstDtm)
     
    return retStr

    
def ifAttendanceOnce(office, name, fstDtm):
    retStr = ''
    fstTime = int(fstDtm.strftime("%H%M%S"))
    if fstTime > 120000:
        #判断下午状态
        retStr += 'AM无打卡+'
        retStr += ifEarly(office, name, fstDtm)
    else:
        #判断上午状态
        retStr += ifLate(office, name, fstDtm)
        retStr += 'PM无打卡'
    
    return retStr

'''    
    
    
#def time_cmp(first_time, second_time):
#    print(first_time)
#    print(second_time)
#    return int(time.strftime("%H%M%S", first_time)) - int(time.strftime("%H%M%S", second_time))

#print(time_cmp(time.gmtime(), time.strptime("09:35:10", "%H:%M:%S")))

'''
#check if the file exist
import os

os.chdir('f:\PythonStudy\work')
if os.path.exists('北京9月考勤数据.xls'):
#if os.path.exists('f:\PythonStudy\work\北京9月考勤数据.xls'):
    print ('档案存在')
else:
    print ('file not found!')

#attend_name = ["王艳萍", "刘涛", "刘江山", "吴晟"]  
#attend_code = [1, 2, 12, 15]
#att_name_dict = {"attend_code": attend_code,  
#        "attend_name": attend_name
#        }

'''
'''
#北京同事代码索引
BJ_att_name_arr = [
    [1 ,'王艳萍'],
    [2 ,'刘涛'],
    [12 ,'刘江山'],
    [15 ,'吴晟'],
    [16 ,'俞府钟'],
    [24 ,'杨琳'  ],
    [25 ,'高杰'  ],
    [32 ,'汪秀文'],
    [34 ,'杨帆'  ],
    [35 ,'邓爱芬'],
    [40 ,'刘冬'  ],
    [41 ,'宋彦彬']   
]


BJ_atname = pd.DataFrame(BJ_att_name_arr, columns = ['code', 'name']) #指定栏位名称
print (BJ_atname)
'''

BJ_name_dict = {
    1  : '王艳萍' ,
    2  : '刘涛'   ,
#    6  : '未知姓名',
    12 : '刘江山',
    13 : '未知姓名13号',
    15 : '吴晟'  ,
    16 : '俞府钟',
    24 : '杨琳'  ,
    25 : '高杰'  ,
    32 : '汪秀文',
    33 : '未知姓名33号',
    34 : '杨帆'  ,
    35 : '邓爱芬',
    36 : '未知姓名36号',
    37 : '张栋'  ,
    39 : '未知姓名39号',
    40 : '刘冬'  ,
    41 : '宋彦彬'   
}






#attend_name = ["王艳萍", "刘涛", "刘江山", "吴晟"]  
#attend_code = [1, 2, 12, 15]
#att_name_dict = {"attend_code": attend_code,  
#        "attend_name": attend_name
#        }



#def get_lateOBtm(nameStr = '')
#    return nameStr


'''
#处理日期时间
#import time
#当前时间戳，适宜日期运算，但UNIX/Windows只适用于1970-2038年之间
#ticks = time.time()
#print ('当前时间戳为：',ticks)
# time tuple, struct_time
#    0    tm_year    2019
#    1    tm_mon     1-12
#    2    tm_mday    1-31
#    3    tm_hour    0-23
#    4    tm_min     0-59
#    5    tm_sec     0-61 (60 or 61 是润秒)
#    6    tm_wday    0-6  (0是周一)
#    7    tm_yday    1-366(儒略历)
#    8    tm_isdst   -1, 0, 1, -1 是决定是否为夏令时的旗帜

#from datetime import datetime as dt
#t = dt(2019, 10, 5)
#print (t)
#print (dt.today())
#print (dt.now())

from datetime import datetime


now = datetime.now()
print ('年:{}, 月:{} ,日:{}'.format(now.year, now.month, now.day))

diff = datetime(2019,3,4,17) - datetime(2019,2,8,15)
print (type(diff))
print (diff)
print ('历经了{}天，{}秒'.format(diff.days, diff.seconds))
'''
# datetime -> str
#    str(datetime_obj)
#    datetime.strftime()
# str -> datetime
#    datetime.strptime() #需要指定时间表示形式
#    dateutil.parser.parse() #可以解析绝大部分时间表示形式
#    pd.to_datetime() #可以处理缺失值和空字符串
#
'''
dt_obj = datetime(2019,11,26)
str_obj = str(dt_obj)
print (type(str_obj))
print (str_obj)
str_obj2 = dt_obj.strftime('%y-%m-%d')
print(str_obj2)
'''


#rawFrame = pd.read_excel(rawFileName,sheetName = 'sheet1')
#print (rawFrame)

#rf1 = pd.read_excel(rawFileName,sheetName = 'sheet1')
#print (rf1.dtypes)
#print (rf1.head())  #看前5行
#print (rf1.tail())
#print (rf1.index)
#print (rf1.columns)
#print (rf1.iloc[0,0])
#print (rf1.shape) #return rows, columns
#print (rf1.info)
#print (rf1.describe())



##############################################################################
# Rewrite this program with OOP                                              #
# ############################################################################    

#西安迟到判定
# AM 8:40后打卡
#     西安 李亚明 AM 9:10后打卡
#     西安 史红杰 AM 9:10后打卡
#西安早退判定
# 17:30前打卡
#     西安 李亚明 18:00前打卡
#     西安 史红杰 18:00前打卡
#     西安 王歆 16:30前打卡 2020年3月18日(含)前有效
#北京迟到判定
# AM 9:10后打卡
#     北京 刘涛 AM 9:40后打卡
#     北京 王艳萍 AM 8:40后打卡
#北京早退判定
# 18:00前打卡
#     北京 刘涛 18:30前打卡
#     北京 王艳萍 17:30前打卡
#事假半天判定 ：迟到或早退在 30（含）分钟-60（不含）分钟之间
#旷职的判定   ：迟到或早退在 60（含）分钟以上
#加班的判定   ：周内延时工作超过3小时（含），周末延时工作超过4小时（含）

ExceptionName_Dict = {    
    '王艳萍' : ['八点半'    ,datetime.strptime('2099-1-1','%Y-%m-%d')],
    '李亚明' : ['九点'      ,datetime.strptime('2099-1-1','%Y-%m-%d')],
    '史红杰' : ['九点'      ,datetime.strptime('2099-1-1','%Y-%m-%d')],
    '刘涛'   : ['九点半'    ,datetime.strptime('2099-1-1','%Y-%m-%d')],
    '王歆'   : ['八点半哺乳',datetime.strptime('2020-3-8','%Y-%m-%d')]
}

OfficeRules_Dict = {
    '西安' : '八点半',
    '北京' : '九点'
}

Rules_Dict = {
    '八点半' : [
        int(datetime.strptime('8:40:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('17:30:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '17:30:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('9:30:00','%H:%M:%S').strftime("%H%M%S")) , # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('16:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('20:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班        
    ], 
    '九点' : [
        int(datetime.strptime('9:10:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('18:00:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '18:00:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('10:00:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('21:00:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ],
     '九点半' : [
        int(datetime.strptime('9:40:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('18:30:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '18:30:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('10:30:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('21:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ],
    '八点半哺乳' : [
        int(datetime.strptime('8:40:00','%H:%M:%S').strftime("%H%M%S")) , # 上班
        int(datetime.strptime('16:30:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '16:30:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('9:30:00','%H:%M:%S').strftime("%H%M%S")) , # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('15:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('19:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班        
    ],
    '九点哺乳' : [
        int(datetime.strptime('9:10:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '17:00:00','%H:%M:%S',                      # 下班 用于计算加班时间
        int(datetime.strptime('10:00:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('20:00:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ]
}
onDutyTmIdx, offDutyTmIdx, offDutySTRIdx = 0, 1, 2
onDutyLate60mTmIdx, offDutyEarly60mTmIdx, overTimeTmIdx = 3, 4, 5

class RkEmployeeRuls:
    onDutyTm   = offDutyTm = onDutyLate60mTm = offDutyEarly60mTm = overTimeTm = 0
    offDutySTR = '17:30:00'
    name = ''

    # 期限在本月分的特殊上班时间检查
    spcCheckFg = False
    spcDate = datetime.strptime('2099-1-1','%Y-%m-%d')

    office = ''
    def __init__(self, officeName = '', nam='', dtmMonthYear=0):
    
        if nam == '':
            raise RuntimeError('RkEmployeeRuls: People name not defineded!')
        self.name = nam
        self.office = officeName    

        def setRules(rk):
            print ('rk=',rk)
            self.onDutyTm, self.offDutyTm, self.offDutySTR, self.onDutyLate60mTm, self.offDutyEarly60mTm, self.overTimeTm = Rules_Dict[rk]
        
        if self.office in OfficeRules_Dict.keys():
            #print (self.office,OfficeRules_Dict[self.office])
            setRules(OfficeRules_Dict[self.office])
        else:
            raise RuntimeError('RkEmployeeRuls: Branch name not defineded!')

        print ('RKemployeeRule name=',self.name)
        if self.name in ExceptionName_Dict.keys():
            rKey, rDate = ExceptionName_Dict[self.name]
            #这里只检查今天，是不对的，应该是比对考勤资料的日期
            print ('rKey=',rKey,rDate.date(), datetime.today().date())
            if rDate.date() >= datetime.today().date():
                print(rKey)
                setRules(rKey)      

    def ifLate(self, dtm): # return 时间+'正常', '迟到', '迟超60分缺席'

        timeNum = int(dtm.strftime("%H%M%S"))
        retStr = 'AM '
        if timeNum > self.onDutyLate60mTm:
            retStr += '迟超60缺席'
        elif timeNum > self.onDutyTm:
            retStr += '迟到'
        else:
            retStr += '正常'
        retStr += dtm.strftime(" %H:%M:%S")
        retStr += '+'
        return retStr

    def ifEarly(self, dtm): #return 时间+'正常', '早退', '早退超60分缺席', ’加班'

        timeNum = int(dtm.strftime("%H%M%S"))

        # datetime to time时间戳
        tmStr = dtm.strftime("%Y-%m-%d ")
        tmStr += self.offDutySTR          
                    
        retStr = 'PM '
        dtmStdOffDtm = datetime.strptime(tmStr, "%Y-%m-%d %H:%M:%S")

        if timeNum < self.offDutyEarly60mTm:
            retStr += '早退超60缺席'
        elif timeNum < self.offDutyTm:
            retStr += '早退'
            #retStr += dtm.strftime(" %H:%M:%S")
        elif timeNum > self.overTimeTm:
            retStr += '加班'
            dltm = dtm - dtmStdOffDtm

            #print ('dltm=',dltm,dtm,dtmStdOffDtm,retStr)
            retStr += (str(int(dltm.seconds / 3600))+':'+str(int(dltm.seconds % 3600 / 60))+':'+str(int(dltm.seconds % 60)))
            #retStr += '%d:%d:%d',dltm.seconds / 3600, dltm.seconds % 3600 / 60, dltm.seconds % 60
        
        else:
            retStr += '正常'

        retStr += dtm.strftime(" %H:%M:%S")
        return retStr 

    def AttenTwice(self, fstDtm, lstDtm):
        retStr = ''
        #print (fstDtm, lstDtm)
        fstTime = int(fstDtm.strftime("%H%M%S"))
        lstTime = int(lstDtm.strftime("%H%M%S"))
        if fstDtm.weekday() == 5 or fstDtm.weekday() == 6: #周六 or 周日
            if (fstDtm.weekday() == 5):
               retStr += '周六'
            else:
               retStr += '周日'
            retStr += '加班'
            dltm = lstDtm - fstDtm
            retStr += '从'+fstDtm.strftime(" %H:%M:%S")+'到'+lstDtm.strftime(" %H:%M:%S")+'共'
            retStr += (str(int(dltm.seconds / 3600))+':'+str(int(dltm.seconds % 3600 / 60))+':'+str(int(dltm.seconds % 60)))
        else:
            if fstTime > 120000:
                retStr += 'AM无打卡+'
            else:
                #判断上午状态
                retStr += self.ifLate(fstDtm)

            if lstTime < 120000:
                retStr += 'PM无打卡'
            else:
                #判断下午状态
                retStr += self.ifEarly(lstDtm)
     
        return retStr

    
    def AttenOnce(self, fstDtm):
        retStr = ''
        fstTime = int(fstDtm.strftime("%H%M%S"))
        if fstDtm.weekday() == 5 or fstDtm.weekday() == 6: #周六 or 周日
            if (fstDtm.weekday() == 5):
                retStr += '周六'
            else:
                retStr += '周日'
            retStr += '加班只打卡一次'+fstDtm.strftime(" %H:%M:%S")+'无法计算'           
        elif fstTime > 120000:
            #判断下午状态
            retStr += 'AM无打卡+'
            retStr += self.ifEarly(fstDtm)
        else:
            #判断上午状态
            retStr += self.ifLate(fstDtm)
            retStr += 'PM无打卡'
    
        return retStr




def readExcelAndConvert(rawFileName):
    print ('XA read_excel:',rawFileName)

    try:
        rf = pd.read_excel(rawFileName,sheet_name=0)
        
    except:
        raise RuntimeError('XA open excel file error!')

    print ('rf.types',rf.dtypes)

    officeStr = '' # 办公室子串 '北京' 或 '西安'
    dtmKeyStr = '' # 用于档案关键字搜寻
    numKeyStr = '' # 用于档案关键字搜寻
   
    if rawFileName.find('北京') >= 0:
        officeStr = '北京'
        dtmKeyStr = '日期时间'
        numKeyStr = '考勤号码'
        #print ('inside 1 officeStr=',officeStr)
    elif rawFileName.find('打卡记录') >= 0:
        officeStr = '西安'
        dtmKeyStr = '打卡时间'
        numKeyStr = '工号'    
        #print ('inside 2 officeStr=',officeStr)
    else:
        raise RuntimeError('officeStr cannot be idendified from filename!')
    
    print ('officeStr=',officeStr)

    #先按照 部门、工号 or 考勤号码、人员、日期排序
    
    rf = rf.sort_values(by = ['部门',numKeyStr,'姓名',dtmKeyStr], ascending = [True, True, True, True])
    
    # 先整表查找，生成一个数据列，用于后续判读
    # [人次计数, 个人第一笔数据位置, 个人数据数, [日期别数据数的列表]]

    MjrIdx = 0 # Major index 总指标
    nIdx = 0   # name index 同一人的第一笔数据(用于区别不同的人)
    nCnt = 1   # name count 同一人在本月的数据数
    ndCnt = 1   # day count 同一人在同一天天的数据数
    peopleCnt = 0 #几个人

    dataList = []
    subList = []

    firstDataFg = True
    outputList = []
    outputFileNameStr = ''
    num_days = 0

    for MjrIdx in range(0, len(rf)-1): 
        kno = rf.iloc[MjrIdx]['姓名']
        nxtKno = rf.iloc[MjrIdx + 1]['姓名']

        dtm    = datetime.strptime(rf.iloc[MjrIdx  ][dtmKeyStr], "%Y-%m-%d %H:%M:%S")
        nxtDtm = datetime.strptime(rf.iloc[MjrIdx+1][dtmKeyStr], "%Y-%m-%d %H:%M:%S")
        day    = int(   dtm.strftime("%d"))
        nxtDay = int(nxtDtm.strftime("%d"))

        if firstDataFg == True:
            firstDataFg = False
            outputFileNameStr = officeStr+'{0}年{1}月考勤导出数据'
            outputFileNameStr = outputFileNameStr.format(dtm.year, dtm.month)
            fstColList = ['部门','姓名',numKeyStr]
            #datetime.datetime(year=2014, month=1, day=1)
            num_days = calendar.monthrange(dtm.year,dtm.month)[1]
            for day in range(1, num_days+1):
                fstColList.append((datetime(year=dtm.year,month=dtm.month,day=day)).strftime("%Y-%m-%d %a"))
                
            #istdaysStrL = [(datetime(year=dtm.year,month=dtm.month,day=day)).strftime("%Y-%m-%d %a") for day in range(1, num_days+1)]
            #fstRow.append(daysStrList)   

            #for i in range(calendar.monthrange(dtm.year,dtm.month)[1]):
            #    fstRow.append((dtm + datetime.timedelta(days=+i)).strftime("%Y-%m-%d %a"))
            print (fstColList)
            #outputList.append(fstColList)
            print (outputList)
            print (outputFileNameStr)

        if kno == nxtKno: #同一人
            nCnt += 1
            if day == nxtDay: #同一人同一天
                ndCnt +=1
            else:             #同一人不同一天
                subList.append(ndCnt)
                #print ('同一人不同一天%d',ndCnt)
                ndCnt = 1
        else: #不同人
            subList.append(ndCnt)
            aList = [peopleCnt, MjrIdx, nIdx, nCnt, subList]
            dataList.append(aList)
            print (peopleCnt, MjrIdx, nIdx, nCnt, subList) 

            nIdx += nCnt
            nCnt = 1
            ndCnt = 1
            subList = []
            peopleCnt += 1

    subList.append(ndCnt)
    aList = [peopleCnt, MjrIdx, nIdx, nCnt, subList]
    dataList.append(aList)
    print (peopleCnt, MjrIdx, nIdx, nCnt, subList) 
    #print (dataList)
    
    outputList.append(fstColList)

    for i in range(0, len(dataList)): #依人次计数
        curList = dataList[i]
        idx = curList[2]           #个人第一笔data位置
        #cnt = curList[3]           #个人当月资料总数
        dList = curList[4]         #个人data段计数List

        kno = rf.iloc[idx][numKeyStr]
        dep = rf.iloc[idx]['部门']
        nam = rf.iloc[idx]['姓名']
  
        #print ('换人', idx, kno, dep, nam)
        if officeStr == '北京':
            try:
                nameStr = BJ_name_dict[kno] #由考勤号码取得姓名
                if nameStr != nam:          #若姓名没有登录，则填入姓名
                    nam = nameStr
            except:
                print ('由',kno,'考勤号码查不出姓名:')
        
            #print ('换人2'+nam, idx)
        
        curColList = fstColList.copy()
        curColList[0], curColList[1], curColList[2],  = dep, nam, kno
        for xdx in range(3, num_days+3):
            curColList[xdx] = '无打卡'
        #print (curColList)

        emRules = RkEmployeeRuls(officeStr,nam)
        fstIdx = idx
        for j in range(0, len(dList)): #依每人在当月的资料段计数
            dCnt = dList[j] #当段data数
      
            dtmStr = rf.iloc[fstIdx][dtmKeyStr]
            fstDtm = datetime.strptime(dtmStr, "%Y-%m-%d %H:%M:%S")
            lstIdx = fstIdx + (dCnt - 1)
            #print ('idxFstLst',idx, fstIdx, lstIdx, dCnt)
            if dCnt > 1:
                dtmStr = rf.iloc[lstIdx][dtmKeyStr]
                lstDtm = datetime.strptime(dtmStr, "%Y-%m-%d %H:%M:%S")  
                #print (nam, fstDtm, lstDtm) 
                retStr = emRules.AttenTwice(fstDtm, lstDtm)
                #print (nam, str, fstDtm, lstDtm)
            else:
                retStr = emRules.AttenOnce(fstDtm)
                #print (nam, str, fstDtm)
        
            print (nam, fstDtm.strftime("%Y-%m-%d %a "),retStr)
            curColList[fstDtm.day+2] = retStr
            #ans = input()
            fstIdx += dCnt    

        print (curColList)
        outputList.append(curColList)
        print (outputList)
        ans = input()
    
    # write an output excel file
    wdf = pd.DataFrame(outputList)
    outputFileNameStr += '.xlsx'

    wdf.to_excel(outputFileNameStr)



##############################################################################
def open_filename():

    #empRules = RkEmployeeRuls('北京','刘涛')
    #print ('empRules',empRules)

    # 调用askopenfilename方法获取单个文件的文件名
    filepath = filedialog.askopenfilename(
        title='打开单个西安打卡或北京考勤xls档案',
        filetypes=[
            ("西安打卡", "打卡记录*.xls") ,
            ("北京考勤", "北*.xls")                     
        ], # 只处理的文件类型
        initialdir='.'  # 初始目录
    ) 
    if filepath != '':
        print (filepath)
        readExcelAndConvert(filepath)
        

l.pack()
counter = 0
def do_job():
    global counter
    l.config(text = 'do '+ str(counter))
    open_filename()
    counter += 1

menubar = tk.Menu(window)

filemenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=filemenu)
filemenu.add_command(label='Open', command=do_job)
filemenu.add_command(label='Exit', command=window.quit)

window.config(menu=menubar)

window.mainloop()


#print (rf.iloc[0,])
#print (rf.iloc[1,])
#print (rf.iloc[2,])


#输出档案


'''
key = [
    '部门', '姓名',
         '1', '2', '3', '4', '5', '6', '7', '8', '9', 
    '10','11', '12', '13', '14', '15', '16', '17', '18', '19', 
    '20','21', '22', '23', '24', '25', '26', '27', '28', '29', 
    '30', '31'
]

kdata = [
    str
]

oDataRow = {
    
}


of = pd.DateFrame()

'''

