import os
import xlrd


data= xlrd.open_workbook(r"C:\Users\os\Desktop\文件目录.XLS")
table= data.sheet_by_name("My Worksheet")
i=0
while i<table.nrows:
    path=os.path.join(table.cell_value(i, 5),table.cell_value(i, 6),table.cell_value(i, 7))

    if os.path.exists(path)==False:
        os.makedirs(path)

    i=i+1
