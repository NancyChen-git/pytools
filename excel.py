#!/usr/bin/env python
#-*-coding:utf-8-*-
#Author:nan.chen
#Version 1.0
             
class Excel:
    def __init__(self, bookname = '新建表格.xls', sheetname = 'Sheet1', mode = 'r', start_row = 0, workbook = '', worksheet = ''):
        self.bookname = bookname
        self.sheetname = sheetname
        self.mode = mode
        self.start_row = 0
        self.workbook = workbook
        self.worksheet = worksheet
  
    @classmethod    
    def open(cls,bookname = '新建表格.xls', sheetname = 'Sheet1', mode = 'r', start_row = 0):
        if mode == 'r':
            import xlrd
            workbook = xlrd.open_workbook(bookname)
            worksheet = workbook.sheet_by_name(sheetname)
        elif mode == 'w':
            import xlwt
            workbook = xlwt.Workbook(encoding="utf-8")
            worksheet = workbook.add_sheet(sheetname) 
        else:
            sys.exit("Mode should be w or r!")
        return cls(bookname, sheetname, mode, start_row, workbook, worksheet)

    def writeline(self, line):
        if type(line) == str:
            line = line.split("\t")
        for i in range(0,len(line)):
            self.worksheet.write(self.start_row,i,line[i]) 
        self.start_row += 1

    def save(self):
        self.workbook.save(self.bookname) 
    
    def readline(self):
        for row in self.worksheet.get_rows():
            row = map(lambda x:x.value, row)
            yield row

