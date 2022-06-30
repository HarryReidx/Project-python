import xlwings as xw
import random
import re
import glob

app = xw.App(visible = True, add_book = True)
wb2=app.books.open(r'C:\Users\K\Desktop\合并TB-3.6\TB-科技园合并-2021. 3.8-期初.XLS') #TB模板位置
sht2=wb2.sheets["试算平衡表-期初"]
i=8
# 上期用1  本期用2
filearray=[]
filelocation=glob.glob(r"C:\Users\K\Desktop\合并TB-3.6\子公司TB\*XLS") #久其导出报表所在文件夹
for filename in filelocation:
    filearray.append(filename)
    list1=filename.split("\\")
    name1=list1[-1]
    name2=name1[:-4]
    print(name2)
    sht2.range(5, i).value=name2
    wb1=app.books.open(filename)
    sht1=wb1.sheets["Z01 资产负债表(企财01表)"]
    # worksheet.write(t, i, table.cell_value(j, n))
    t=7
    j=6
    n=4
    while t<65:
        sht2.range(t, i).value = sht1.range(j, n).value
        t=t+1
        j=j+1
    t=65
    j=82
    sht2.range(t, i).value = sht1.range(j, n).value
    t = 72
    j = 6
    n = 9
    while t < 149:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1

    sht1 = wb1.sheets["Z02 利润表(企财02表)"]
    t = 153
    j = 4
    n = 4
    while t<194:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1
    j = 5
    n = 8
    while t<229:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1

    sht1 = wb1.sheets["Z04 所有者权益变动表(企财04表)"]
    t = 230
    j = 9
    n = 27
    while t<263:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1

    sht1 = wb1.sheets["Z03 现金流量表(企财03表)"]
    t = 267
    j = 4
    n = 4
    while t < 297:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1
    j = 5
    n = 8
    while t < 326:
        sht2.range(t, i).value = sht1.range(j, n).value
        t = t + 1
        j = j + 1

    i=i+1
    wb1.close()

wb2.save(r'C:\Users\K\Desktop\合并TB-3.6\TB-科技园合并-2021. 3.8-期末.XLS') #新生成的TB位置及命名方式
wb2.close()
app.quit()






