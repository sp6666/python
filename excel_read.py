# 导入读取excel库


import xlrd
import os


# 安装的 xlrd-2.0.1，该版本只支持.xls文件

def read_excel():
    # 打开文件
    path = os.path.abspath('.')
    path2 = path + '/excel/role.xls'
    print(path2)
    wb = xlrd.open_workbook(path2)
    # 获取所有sheet的名字
    print(wb.sheet_names())

    length = len(wb.sheet_names())
    sheets = wb.sheets()
    print(sheets)
    sheet = wb.sheet_by_index(0)
    rowNum = sheet.nrows
    colNum = sheet.ncols
    print(rowNum, colNum)


read_excel()
