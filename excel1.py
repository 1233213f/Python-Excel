import openpyxl

# 打开工作簿
wb = openpyxl.load_workbook('GRB射线数据记录——后续补充.xlsx')

# 获取所有表的表名
sheets_names = wb.sheetnames
print(sheets_names)  # 结果: 工作表的名字，一列表（字符串）形式呈现 ['Sheet1', 'Sheet2', 'Sheet3']

# 获取活动表对应的表对象(表对象就是Worksheet类的对象)
active_sheet = wb.active
print(active_sheet)  # 结果：<Worksheet "表1"> 当前工作的表

# 根据表名获取工作簿中指定的表
sheet2 = wb['Sheet2']
print(sheet2)  # 结果：<Worksheet "表2">

# 根据表对象获取表的名字
sheet_name1 = active_sheet.title
sheet_name2 = sheet2.title

print(sheet_name1, sheet_name2)  # 结果：Sheet3 Sheet2