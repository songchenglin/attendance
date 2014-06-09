# -*- coding: utf-8 -*-
import xlrd
import xlwt
import re
import string

class StaffInfo():
    name = ''
    serial = ''
    depart = ''
    attendanceInfo = []
    def __init__(self,dic):
        self.name = dic['name']
        self.serial = dic['serial']
        self.depart = dic['depart']
    def SetAttendanceInfo(self,lists):
        #reObj = re.compile('(\d{2})\s*(\w{1})')
        for list in lists:
            dic = {}.fromkeys(('date', 'week', 'start', 'end'))
            dic['date'] = list['date']
            dic['start'] = list['start']
            dic['end'] = list['end']
            self.attendanceInfo.append(dic)
    def Get(self,list):
        return self.infoList
    def ShowDepartInfo(self):
        print 'serial:%s,name:%s,depart:%s' % (self.serial, self.name, self.depart)
    def ShowAttendanceInfo(self):
        for info in self.attendanceInfo:
            print info['date'],info['start'],info['end']

class AttendanceTable():
    fname = u'标准报表.xls'
    excel = None
    sheets = []
    sheets_count = 0
    sheets_names = []
    def __init__(self):
        try:
            self.excel = xlrd.open_workbook(self.fname)
        except Exception,e:
            print str(e)
        self.sheets = self.excel.sheets()
        self.sheets_count = self.excel.nsheets
        self.sheets_names = self.excel.sheet_names()

    def GetStaffLists(self):
        lists = []
        sheet = self.sheets[0]
        nrows = sheet.nrows
        for row in range(4,nrows):
            info = {}.fromkeys(('serial', 'name', 'depart'))
            info['serial'] = sheet.cell(row,0).value
            while len(info['serial']) < 8: #Begin of '0' maybe ignored,so add it
                info['serial'] = '0' + info['serial']
            info['name'] = sheet.cell(row, 1).value
            info['depart'] = sheet.cell(row, 2).value
            lists.append(info)
        return lists
    def GetAttendInfo(self,serial):
        #some constant
        week_start = 1
        week_end = 3
        weekend_start = 10
        weekend_end = 12
        #constant end
        sheetName = self.FindSheetName(serial)
        serialIsFound = False
        lists = []
        while serial[0] == '0' and serial != '0':
            serial = serial[1:]
        sheet = self.excel.sheet_by_name(sheetName)
        #self.PrintSheet(sheet) #ok
        nrows = sheet.nrows
        ncols = sheet.ncols
        for col in range(0,ncols):
            val = sheet.cell(3,col).value
            if val == serial:
                serialIsFound = True
                break
        if serialIsFound is not True:
            raise StandardError
        #col-9 is data start col
        col = col-9
        for row in range(11,nrows):
            dic = {}.fromkeys(('date', 'start', 'end'))
            dic['date'] = sheet.cell(row,col).value
            dic['start'] = sheet.cell(row,col+week_start).value
            dic['end'] = sheet.cell(row,col+week_start).value
            if dic['start'] == '' or dic['end'] == '':
                dic['start'] = sheet.cell(row,col+weekend_start).value
                dic['end'] = sheet.cell(row,col+weekend_end).value
            #print dic['date'],dic['start'],dic['end']
            lists.append(dic)
        return lists

    def FindSheetName(self,serial):
        while serial[0] == '0' and serial != '0':
            serial = serial[1:] #delete first '0'
        for sheetName in self.sheets_names:
            if sheetName.find(serial) < 0:
                continue
            else:
                print 'Found sheet:%s' % sheetName
                return sheetName
        print 'No sheet found!!!'
        return None

    def PrintSheet(self,sheet):
        nrows = sheet.nrows
        print 'nrows=%d' % nrows
        for i in range(nrows):
            values =  sheet.row_values(i)
            for value in values:
                print value,
            print ''

class SaveAttendce():
    fname = ''
    sheet = ''
    excel = None
    def __init__(self,fname):
        titles = (u'工号',u'姓名',u'考勤日期',u'星期',u'属性',u'上班打卡时间',u'下班打卡时间')
        self.fname = fname
        self.excel = xlwt.Workbook()
        self.sheet = self.excel.add_sheet('sheet1')
        col = 0
        for title in titles:
            self.sheet.write(0, col, title)
            col += 1
        self.excel.save(self.fname)
    def SaveInfo(self,lists):
        i = 1
        for list in range(1,len(lists)+1):
            self.sheet.write(i,0,list['serial'])
            i += 1

if __name__ == '__main__':
    staffLists = []  #Save everyone's information Object
    attendance = AttendanceTable()

    #Get everyone's information,and put it to a list
    lists = attendance.GetStaffLists()
    for dic in lists:
        staff = StaffInfo(dic)
        staffLists.append(staff)
    print 'staffLists len = %d' % len(staffLists)
    #Get everyone's attendance information,and put to staff ojbect
    for staff in staffLists:
        infolists = attendance.GetAttendInfo(staff.serial)
        staff.SetAttendanceInfo(infolists)
        staff.ShowDepartInfo()
        staff.ShowAttendanceInfo()
    fn = u'西安考勤申报.xls'
    s = SaveAttendce(fn)
__author__ = 'songchenglin'
