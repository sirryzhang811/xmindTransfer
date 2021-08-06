# xlwt操作excel的常见方法与问题
***

### 前言
Python操作excel的包

### 一、安装
```
pip install xlwt
```

### 二、简单使用xlwt
```angular2html
import xlwt #导入模块
workbook = xlwt.Workbook(encoding='utf-8') #创建workbook 对象
worksheet = workbook.add_sheet('sheet1') #创建工作表sheet
worksheet.write(0, 0, 'hello') #往表中写内容,第一个参数行,第二个参数列,第三个参数内容
workbook.save('students.xls') #保存表为students.xls
```

### 三、修改文字样式
```angular2html
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1')
#设置字体样式
font = xlwt.Font()
#字体
font.name = 'Time New Roman'
#加粗
font.bold = True
#下划线
font.underline = True
#斜体
font.italic = True
 
#创建style
style = xlwt.XFStyle()
style.font = font
#根据样式创建workbook
worksheet.write(0, 1, 'world', style)
workbook.save('students.xls')
```

# 四、合并单元格
```angular2html
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1')
#通过worksheet调用merge()创建合并单元格
#第一个和第二个参数单表行合并,第三个和第四个参数列合并,
 
#合并第0列到第2列的单元格
worksheet.write_merge(0, 0, 0, 2, 'first merge')
 
#合并第1行第2行第一列的单元格
worksheet.write_merge(0, 1, 0, 0, 'first merge')
 
workbook.save('students.xls')
```