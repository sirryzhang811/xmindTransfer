# -*- coding:utf-8 -*-


import json
import xlwt

class xlsPrac():

    def __init__(self):
        pass

    def xlsWrite(self, filePath):
        self.filePath = filePath
        workbook = xlwt.Workbook(encoding='utf-8')   # 创建workbook对象
        worksheet = workbook.add_sheet('sheet1')    # 创建工作表
        worksheet.write(0, 0, 'hello word')   # 往表格填写内容
        fileName = self.filePath + 'test.xls'
        workbook.save(fileName)   # 保存表格


if __name__ == '__main__':
    xlsPrac().xlsWrite('./')