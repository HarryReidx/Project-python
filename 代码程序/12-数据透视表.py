import os
import xlwings as xw
import pandas as pd
import re
app = xw.App(visible = True, add_book = True)



workbook = app.books.open(r"C:\Users\DELL\Desktop\南投本部：凭证序时簿2020.1-12 - 透视表用.xls")
worksheet = workbook.sheets(1)
worksheet2=workbook.sheets.add("数据透视数据")

info=worksheet.used_range
nrows = info.last_cell.row
ncolumns = info.last_cell.column
print("最大列数"+str(ncolumns)+"最大行数"+str(nrows))
n1=13 #筛选关键字所在列
print(worksheet.range(1, n1).value)
new_worksheet = workbook.sheets.add('筛选数据')
m=1
i=2
for n in range(1,ncolumns+1):
    new_worksheet.range(1, n).value=worksheet.range(1, n).value

nlist=[]
while m<nrows+1:
    if re.compile(r".*管理费用.*").search(str(worksheet.range(m, n1).value)) != None:
        nlist.append(m-2)
        # for n in range(1, ncolumns + 1):
        #     new_worksheet.range(i, n).value = worksheet.range(m, n).value
        i=i+1
    m=m+1
    print(m)


table = worksheet.range('A1').expand('table').options(pd.DataFrame).value
product = table.iloc[nlist]
new_worksheet.range('A1').value = product


FX="借方"
values = new_worksheet.range('A1').expand('table').options(pd.DataFrame).value
# values[FX] = values[FX].astype('float')

pivottable1 = pd.pivot_table(values, values=FX, index='期间', columns='科目名称', aggfunc='sum', fill_value=0,
                            margins=True, margins_name='总计')

#index为纵轴指标，columns为横轴指标，通过调换这两个名称可以改变排列方式，引号内借方金额，月份，科目名称等指标按序时账第一行的名称填列，需对应，如需对两个指标进行统计，参考下面的方式，通过增加代码的方式可以实现一次出来多个数据透视表
# pivottable1 = pd.pivot_table(values, values=["借方金额","贷方金额"], index='月份', columns='科目名称', aggfunc={"借方金额":'sum',"贷方金额":'sum'}, fill_value=0,
#                              margins=True, margins_name='总计')

worksheet2.range('A1').value = pivottable1



# FX="贷方金额"
# values = new_worksheet.range('A1').expand('table').options(pd.DataFrame).value
# values[FX] = values[FX].astype('float')
# # result = values.groupby('科目名称').sum()
# pivottable2 = pd.pivot_table(values, values=FX, index='月份', columns='对方一级科目', aggfunc='sum', fill_value=0,
#                             margins=True, margins_name='总计')
# worksheet2.range('A20').value = pivottable2



workbook.save(r"C:\Users\DELL\Desktop\南投本部：凭证序时簿2020.1-12 - 透视表用-new.xls")
#下面两行注释掉了，为了实现程序运行完后，新能成的Excel是打开状态的，可以直接粘到底稿里
# workbook.close()
# app.quit()
