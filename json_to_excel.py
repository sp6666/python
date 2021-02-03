import json
import os

import openpyxl
import typeof

# 拿到json文件
filename = os.path.abspath('.') + '/gameData.json'

# 获取json数据
f = open(filename, 'r', encoding='utf-8', errors='ignore')
rows = json.load(f)
f.close()
print(rows)

data = []
for excelName in rows:
    # 拿到所有文件
    book = openpyxl.Workbook()
    file = rows[excelName]
    k = 0
    for sheetName in file:
        # 拿到所有sheet
        sheet = book.create_sheet(sheetName, k)  # 添加一个sheet页
        k += 1
        # 读到数组里
        titles = ['ID']
        allValues = []
        types = ['int']
        for v in file[sheetName]:
            elements = file[sheetName][v]
            values = [v]
            for key in elements:
                if v == '1':
                    titles.append(key)
                    types.append(typeof.typeof(elements[key]))
                values.append(elements[key])
            allValues.append(values)
        # 写表
        sheet.append(titles)
        sheet.append(types)
        for i in range(len(allValues)):
            sheet.append(allValues[i])

    newPath = os.path.abspath('.') + '/newExcel'
    if not os.path.exists(newPath):
        os.mkdir(newPath)
    newExcel = newPath + '/' + excelName + '.xlsx'
    f2 = open(newExcel, 'a+', encoding='utf-8', errors='ignore')
    f2.close()
    book.save(newExcel)



