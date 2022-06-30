import os
import xlrd
import shutil

data= xlrd.open_workbook(r"C:\Users\os\Desktop\文件目录.XLS")
table= data.sheet_by_name("My Worksheet")
i=0
while i<table.nrows:
    path1=os.path.join(table.cell_value(i, 0),table.cell_value(i, 4))
    path2=os.path.join(table.cell_value(i, 5))

    if os.path.exists(path1):
        shutil.copy(path1, path2)
    else:
        print(path1,"文件不存在")

    i=i+1