import shutil
import xlrd
import os

data= xlrd.open_workbook(r"D:\演示材料\案例 报告转格式\文件移动路径.XLS")
table= data.sheet_by_name("My Worksheet")
i=0
while i<table.nrows:
    path1=os.path.join(table.cell_value(i, 0),table.cell_value(i, 4))
    path2=os.path.join(table.cell_value(i, 5))

    if os.path.exists(path1):
        if os.path.exists(path2)==False:
            os.makedirs(path2)
        shutil.copy(path1, path2)
    else:
        print(path1,"文件不存在")
    print(i)
    i=i+1



