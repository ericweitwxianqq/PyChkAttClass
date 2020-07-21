


import pandas as pd

import time
from datetime import datetime
import calendar



    
    
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
        int(datetime.strptime('9:40:00','%H:%M:%S').strftime("%H%M%S")) , # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('16:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('20:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班        
    ], 
    '九点' : [
        int(datetime.strptime('9:10:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('18:00:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '18:00:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('10:10:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('21:00:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ],
     '九点半' : [
        int(datetime.strptime('9:40:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('18:30:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '18:30:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('10:40:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('21:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ],
    '八点半哺乳' : [
        int(datetime.strptime('8:40:00','%H:%M:%S').strftime("%H%M%S")) , # 上班
        int(datetime.strptime('16:30:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '16:30:00',                                 # 下班 用于计算加班时间
        int(datetime.strptime('9:40:00','%H:%M:%S').strftime("%H%M%S")) , # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('15:30:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('19:30:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班        
    ],
    '九点哺乳' : [
        int(datetime.strptime('9:10:00','%H:%M:%S').strftime("%H%M%S")),  # 上班
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # 下班
                              '17:00:00','%H:%M:%S',                      # 下班 用于计算加班时间
        int(datetime.strptime('10:10:00','%H:%M:%S').strftime("%H%M%S")), # AM 迟 超60分 算 旷职/请假 半天
        int(datetime.strptime('17:00:00','%H:%M:%S').strftime("%H%M%S")), # PM 早退 超60分 算 旷职/请假 半天
        int(datetime.strptime('20:00:00','%H:%M:%S').strftime("%H%M%S"))  # 平日超3小时算加班
    ]
}
onDutyTmIdx, offDutyTmIdx, offDutySTRIdx = 0, 1, 2
onDutyLate60mTmIdx, offDutyEarly60mTmIdx, overTimeTmIdx = 3, 4, 5

class RkEmployeeRuls:
    onDutyTm   = offDutyTm = onDutyLate60mTm = offDutyEarly60mTm = overTimeTm = 0
    offDutySTR = ''
    name = ''

    # 期限在本月分的特殊上班时间启动检查
    spcCheckFg = False
    spcDate = datetime.strptime('2099-1-1','%Y-%m-%d')

    office = ''
    def __init__(self, officeName = '', nam='', dtmMonthYear=datetime(2019,12,27)):
    
        if nam == '':
            raise RuntimeError('RkEmployeeRuls: People name not defineded!')
        self.name = nam
        self.office = officeName    
            
        if self.office in OfficeRules_Dict.keys():
            #print (self.office,OfficeRules_Dict[self.office])
            self.setRules(OfficeRules_Dict[self.office])
        else:
            raise RuntimeError('RkEmployeeRuls: Branch name not defineded!')

        print ('RKemployeeRule name=',self.name)
        if self.name in ExceptionName_Dict.keys():
            rKey, rDate = ExceptionName_Dict[self.name]
            dtmMonthYear.date().replace(day=1)
            if rDate.date() >= dtmMonthYear.date():
                self.setRules(rKey) 
            #print (dtmMonthYear.year, rDate.year, dtmMonthYear.month, rDate.month)    
            if ((dtmMonthYear.year == rDate.year) and (dtmMonthYear.month == rDate.month)):
                self.spcCheckFg = True
                print ('spcCheckFg = True')
                self.spcDate = rDate

            #这里只检查今天，是不对的，应该是比对考勤资料的日期
            #print ('rKey=',rKey,rDate.date(), datetime.today().date())
            #if rDate.date() >= datetime.today().date():
            #    print(rKey)
            #    self.setRules(rKey)      

    def setRules(self, rk):
        print ('rk=',rk, self.name)
        self.onDutyTm, self.offDutyTm, self.offDutySTR, self.onDutyLate60mTm, self.offDutyEarly60mTm, self.overTimeTm = Rules_Dict[rk]
    
    def checkDateRule(self,dtm):
        #if (self.spcCheckFg == True):
        #    print (self.spcDate.date(), dtm.date())

        if ((self.spcCheckFg == True) and (self.spcDate.date() < dtm.date())):
            self.spcCheckFg = False
            print ('spcCheckFg = False')
            self.setRules(OfficeRules_Dict[self.office])

    def ifLate(self, dtm): # return 时间+'正常', '迟到', '迟超60分缺席'
        self.checkDateRule(dtm)
        timeNum = int(dtm.strftime("%H%M%S"))
        retStr = 'AM '
        if timeNum > self.onDutyLate60mTm:
            retStr += '迟超60m缺席'
        elif timeNum > self.onDutyTm:
            retStr += '迟到'
        else:
            retStr += '正常'
        #retStr += dtm.strftime(" %H:%M:%S")
        retStr += '\n'
        return retStr

    def ifEarly(self, dtm): #return 时间+'正常', '早退', '早退超60分缺席', ’加班'
        self.checkDateRule(dtm)
        timeNum = int(dtm.strftime("%H%M%S"))

        # datetime to time时间戳
        tmStr = dtm.strftime("%Y-%m-%d ")
        tmStr += self.offDutySTR          
                    
        retStr = 'PM '
        dtmStdOffDtm = datetime.strptime(tmStr, "%Y-%m-%d %H:%M:%S")

        if timeNum < self.offDutyEarly60mTm:
            retStr += '早退超60m缺席'
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

        #retStr += dtm.strftime(" %H:%M:%S")
        retStr += '\n'
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
                retStr += 'AM无打卡\n'
            else:
                #判断上午状态
                retStr += self.ifLate(fstDtm)

            if lstTime < 120000:
                retStr += 'PM无打卡\n'
            else:
                #判断下午状态
                retStr += self.ifEarly(lstDtm)
            retStr += fstDtm.strftime("%H:%M:%S")+lstDtm.strftime("\n%H:%M:%S")
     
        return retStr

    
    def AttenOnce(self, fstDtm):
        retStr = ''
        fstTime = int(fstDtm.strftime("%H%M%S"))
        if fstDtm.weekday() == 5 or fstDtm.weekday() == 6: #周六 or 周日
            if (fstDtm.weekday() == 5):
                retStr += '周六'
            else:
                retStr += '周日'
            retStr += '加班只打一次卡无法计算\n'
            #retStr += '加班只打卡一次'+fstDtm.strftime(" %H:%M:%S")+'无法计算'           
        elif fstTime > 120000:
            #判断下午状态
            retStr += 'AM无打卡\n'
            retStr += self.ifEarly(fstDtm)
        else:
            #判断上午状态
            retStr += self.ifLate(fstDtm)
            retStr += 'PM无打卡\n'

        retStr += fstDtm.strftime("%H:%M:%S")
        return retStr

class RkAttenTableToFile:
    officeStr = '' # 办公室子串 '北京' 或 '西安'
    dtmKeyStr = '' # 用于档案关键字搜寻
    numKeyStr = '' # 用于档案关键字搜寻
    
    curMonthDtm = datetime(2019,12,27)
    num_days = 0
    
    def scanDFSetDataList(self, rf = pd.DataFrame()):
        # 先整表查找，生成一个数据列，用于后续判读
        # [人次计数, 个人第一笔数据位置, 个人数据数, [日期别数据数的列表]]
        mjrIdx = 0 # Major index 总指标
        nIdx = 0   # name index 同一人的第一笔数据(用于区别不同的人)
        nCnt = 1   # name count 同一人在本月的数据数
        ndCnt = 1   # day count 同一人在同一天的数据数
        peopleCnt = 0 #几个人

        dataList = []
        subList = []
        for mjrIdx in range(0, len(rf)-1): 
            kno = rf.iloc[mjrIdx]['姓名']
            nxtKno = rf.iloc[mjrIdx + 1]['姓名']

            dtm    = datetime.strptime(rf.iloc[mjrIdx  ][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S")
            nxtDtm = datetime.strptime(rf.iloc[mjrIdx+1][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S")
            fstDay = int(   dtm.strftime("%d"))
            nxtDay = int(nxtDtm.strftime("%d"))
        
            if kno == nxtKno: #同一人
                nCnt += 1
                if fstDay == nxtDay: #同一人同一天
                    ndCnt +=1
                else:             #同一人不同一天
                    subList.append(ndCnt)
                    #print ('同一人不同一天%d',ndCnt)
                    ndCnt = 1
            else: #不同人
                subList.append(ndCnt)
                aList = [peopleCnt, mjrIdx, nIdx, nCnt, subList]
                dataList.append(aList)
                print (peopleCnt, mjrIdx, nIdx, nCnt, subList) 

                nIdx += nCnt
                nCnt = 1
                ndCnt = 1
                subList = []
                peopleCnt += 1
        
            #print ('AFTER nCnt+ndCnt', dtm, nxtDtm, fstDay, nxtDay, nCnt, ndCnt)
            #ans = input()

        subList.append(ndCnt)
        aList = [peopleCnt, mjrIdx, nIdx, nCnt, subList]
        dataList.append(aList)
        print (peopleCnt, mjrIdx, nIdx, nCnt, subList)

        return dataList

    def buildOutputList(self, dataList=[], rf = pd.DataFrame()):
        
        if dataList == []:
            return []

        outputList = []

        # initial first column
        fstColList = ['部门','姓名',self.numKeyStr]
        
        for day in range(1, self.num_days+1):
            fstColList.append((datetime(year=self.curMonthDtm.year,month=self.curMonthDtm.month,day=day)).strftime("%Y-%m-%d %a"))
        print (fstColList) 

        outputList.append(fstColList)

        for i in range(0, len(dataList)): #依人次计数
            curList = dataList[i]
            idx = curList[2]           #个人第一笔data位置
            #cnt = curList[3]           #个人当月资料总数
            dList = curList[4]         #个人data段计数List

            kno = rf.iloc[idx][self.numKeyStr]
            dep = rf.iloc[idx]['部门']
            nam = rf.iloc[idx]['姓名']
  
            #print ('换人', idx, kno, dep, nam)
            if self.officeStr == '北京':
                try:
                    nameStr = BJ_name_dict[kno] #由考勤号码取得姓名
                    if nameStr != nam:          #若姓名没有登录，则填入姓名
                        nam = nameStr
                except:
                    print ('由',kno,'考勤号码查不出姓名:')
        
                #print ('换人2'+nam, idx)
        
            curColList = fstColList.copy()
            curColList[0], curColList[1], curColList[2],  = dep, nam, kno
            for xdx in range(3, self.num_days+3):
                curColList[xdx] = '无打卡'
            #print (curColList)

            emRules = RkEmployeeRuls(self.officeStr,nam)
            fstIdx = idx # = curList[2]           #个人第一笔data位置
            for j in range(0, len(dList)): #依每人在当月的资料段计数
                dCnt = dList[j] #当段data数
      
                fstDtm = datetime.strptime(rf.iloc[fstIdx][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S")
                lstIdx = fstIdx + (dCnt - 1)
            
                lstDtm = datetime.strptime(rf.iloc[lstIdx][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S") 
                #print ('idxFstLst',idx, fstIdx, lstIdx, dCnt)
                if dCnt > 1:
                    f = int(fstDtm.strftime("%H%M%S"))
                    l = int(lstDtm.strftime("%H%M%S"))
                    #print ('BEFOR compare',nam, fstDtm, lstDtm) 
                    for i in range (fstIdx,fstIdx+dCnt):
                        dtm = datetime.strptime(rf.iloc[i][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S")
                        d = int(dtm.strftime("%H%M%S"))
                        #print('d f l', d, f, l, f-d, d-l)
                        if (f - d) > 0:
                            fstDtm = dtm
                            f = int(fstDtm.strftime("%H%M%S"))
                        if (d - l) > 0:
                            lstDtm = dtm
                            l = int(lstDtm.strftime("%H%M%S"))
                
                    #print ('AFTER compare',nam, fstDtm, lstDtm)
                    retStr = emRules.AttenTwice(fstDtm, lstDtm)
                
                else:
                    retStr = emRules.AttenOnce(fstDtm)
                    #print ('AFTER ATTEN ONCE', nam, retStr, fstDtm)

                #ans = input()
                #print (nam, fstDtm.strftime("%Y-%m-%d %a "),retStr)
                curColList[fstDtm.day+2] = retStr
                #ans = input()
                fstIdx += dCnt    

            #print (curColList)
            outputList.append(curColList)

            #print (outputList)
            #ans = input()
        return outputList

    def wrAttenTableToExcel(self, outputList = []):
        # initial output file name
        outputFileNameStr = self.officeStr+'{0}年{1}月考勤数据导出表格'
        outputFileNameStr = outputFileNameStr.format(self.curMonthDtm.year, self.curMonthDtm.month)
        #print (outputFileNameStr)
        
        # write an output excel file
    
        wdf = pd.DataFrame(outputList)
        now_time = datetime.now()
        timeStampStr = now_time.strftime("%Y年%m月%d日%H%M%S")
        outputFileNameStr += timeStampStr
        outputFileNameStr += '.xlsx'

        return wdf.to_excel(outputFileNameStr)


    def __init__(self, rawFileName=''):

        if rawFileName.find('北京') >= 0:
            self.officeStr = '北京'
            self.dtmKeyStr = '日期时间'
            self.numKeyStr = '考勤号码'
        elif rawFileName.find('打卡记录') >= 0:
            self.officeStr = '西安'
            self.dtmKeyStr = '打卡时间'
            self.numKeyStr = '工号'    
        else:
            raise RuntimeError('officeStr cannot be idendified from filename!')

        try:
            rf = pd.read_excel(rawFileName,sheet_name=0)
        except:
            raise RuntimeError('open excel file error!')
        
        #先按照 部门、工号 or 考勤号码、人员、日期排序
        #rf = rf.sort_values(by = ['姓名',self.dtmKeyStr], ascending = [True, True])
        rf = rf.sort_values(by = ['部门',self.numKeyStr,'姓名',self.dtmKeyStr], ascending = [True, True, True, True])

        self.curMonthDtm = datetime.strptime(rf.iloc[0][self.dtmKeyStr], "%Y-%m-%d %H:%M:%S") 
        self.num_days = calendar.monthrange(self.curMonthDtm.year,self.curMonthDtm.month)[1]

        idxList = self.scanDFSetDataList(rf)
        outputList = self.buildOutputList(idxList, rf)
        return self.wrAttenTableToExcel(outputList)
       






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
        initialdir='F:/PythonStudy/work/2019年6-10月打卡数据-西安-20191122/'  # 初始目录
    ) 
    if filepath != '':
        print (filepath)
        #readExcelAndConvert(filepath)
        return RkAttenTableToFile(filepath)
         
counter = 0

def do_job():
    global counter
    l.config(text = 'do '+ str(counter))
    try:
        retcode = open_filename()
    except:
        raise RuntimeError('open_filename error!')
    
    counter += 1


import tkinter as tk
from tkinter import filedialog

window = tk.Tk()
#设置窗口名称
window.title('睿控考勤数据处理转换工具')
#设置窗口大小 (长 x 宽)
window.geometry('500x400')

l = tk.Label(window, text='       ', bg= 'green')

tk.Label(window, text='on the window', bg='red', font=('Arial', 16)).pack()  

l.pack()
menubar = tk.Menu(window)

filemenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=filemenu)
filemenu.add_command(label='Open', command=do_job)
filemenu.add_command(label='Exit', command=window.quit)

window.config(menu=menubar)

window.mainloop()





