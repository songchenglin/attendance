# -*- coding: utf-8 -*-
import xlrd
import re
import string

class StaffInfo():
    name = ''
    serial = ''
    depart = ''
    infoList = []
    def __init__(self,dic):
        self.name = dic['name']
        self.serial = dic['serial']
        self.depart = dic['depart']
        #print 'Staff Info:%s,%s,%s' % (self.serial,self.name,self.depart)
    def SetInfo(self,list):
        dic = {}.fromkeys(('date','week','first','second'))
        dic['date'] = list[0]
        dic['first'] = list[1]
        dic['second'] = list[2]
        self.infoList.append(dic)
    def GetInfo(self,list):
        return self.infoList
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
            info = {}.fromkeys(('serial','name','depart'))
            info['serial'] = sheet.cell(row,0).value
            if len(info['serial']) != 8: #Begin of '0' maybe ignored,so add it
                info['serial'] = '0' + info['serial']
            info['name'] = sheet.cell(row,1).value
            info['depart'] = sheet.cell(row,2).value
            lists.append(info)
        return lists
        #self.PrintSheet(sheet)
    def GetAttendInfo(self,serial):
        sheetName = self.FindSheeName(serial)
    def FindSheetName(self,serial):
        if serial[0] == '0':
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
    def test(self):
        print 'sheet_count = %d' % self.sheets_count
        self.FindSheetName('11109062')
        #for name in self.sheets_names:
        #    print name,
        #print ''
if __name__ == '__main__':
    staffLists = []  #Save everyone's information Object
    attendance = AttendanceTable()
    attendance.test()

    #Get everyone's information
    lists = attendance.GetStaffLists()
    for dic in lists:
        staff = StaffInfo(dic)
        staffLists.append(staff)
    print 'staffLists len = %d' % len(staffLists)

__author__ = 'songchenglin'
