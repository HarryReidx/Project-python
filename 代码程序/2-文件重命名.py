import os
import xlrd

fileName = r"C:\Users\K\Desktop\abc.xls"
sheetName = "s1"
try:
    data= xlrd.open_workbook(fileName)
    table= data.sheet_by_name(sheetName)
    i=0
    while i<table.nrows:
        oldname=os.path.join(table.cell_value(i, 0),table.cell_value(i, 4))
        newname=os.path.join(table.cell_value(i, 0),table.cell_value(i, 5))
        if os.path.exists(oldname):
            os.rename(oldname,newname)
        else:
            print(oldname,"文件不存在")
        i=i+1
except Exception as e:
    print("---- 检查文件路径:" + fileName + " ----")
    print(e.args)