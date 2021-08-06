# -*- coding:utf-8 -*-

from xmind import xmind_to_dict
from xmind import xmind_to_json
import json
import xlwt


# TODO
# xmind文件转成字典
filePath = "C:\\Users\\admin\Desktop\\testXmind.xmind"
out = xmind_to_dict(filePath)[0]['topic']
print(json.dumps(out, indent=2, ensure_ascii=False))

# TODO
# 创建excel
workbook = xlwt.Workbook(encoding='utf-8')   # 创建workbook对象
worksheet = workbook.add_sheet('sheet1')    # 创建工作表
worksheet.write(0, 0, 'hello word')   # 往表格填写内容
workbook.save('C:\\Users\\admin\Desktop\\test.xls')   # 保存表格
