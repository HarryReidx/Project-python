import xlwings as xw
import re
import glob

app = xw.App(visible=True, add_book=True)


wb2=app.books.open(r'D:\百度企业网盘同步\BaiduBusinessSync\我的文件\王小芳组2021年度审计资料\2021年审计报告\合并变单户报表模板.xlsx') #TB模板位置


# 合并报表工作簿，母公司工作簿在256行输入
wb1 = app.books.open(r'D:\百度企业网盘同步\BaiduBusinessSync\我的文件\王小芳组2021年度审计资料\2021年审计报告\久其导表\常州铁道高等职业技术学校.XLS')
sht1 = wb1.sheets["FMDM 封面代码"]
sht2 = wb2.sheets["资产负债表"]

sht2.range(4, 1).value ="编制单位："+ sht1.range(3, 2).value
name=sht1.range(3, 2).value
sht1 = wb1.sheets["Z01 资产负债表(企财01表)"]

#上年资产取数
n2=5  #n取粘数的列
n1=5
m2=6 #m取行
while m2<21:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2=m2+1
sht2.range(m2, n2).value = sht1.range(m2, n1).value   #m=21时，取应收股利
m2=m2+1
while m2<31:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2=m2+1
m2=53
while m2 < 63:
    sht2.range(m2, n2).value = sht1.range(m2-20, n1).value
    m2 = m2 + 1
sht2.range(63, n2).value = sht1.range(49, n1).value   #固定资产
sht2.range(64, n2).value = sht1.range(43, n1).value   #固定资产原价
sht2.range(65, n2).value = sht1.range(44, n1).value   #累计折旧
sht2.range(66, n2).value = sht1.range(46, n1).value   #固定资产减值准备
sht2.range(67, n2).value = sht1.range(50, n1).value   #在建工程
m2 = 68
while m2 < 79:
    sht2.range(m2, n2).value = sht1.range(m2-15, n1).value
    m2 = m2 + 1
sht2.range(79, n2).value = sht1.range(82, n1).value   #资产总计


#期初资产取数
n2=4  #n取粘数的列
n1=4
m2=6 #m取行
while m2<21:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2=m2+1
sht2.range(m2, n2).value = sht1.range(m2, n1).value   #m=21时，取应收股利
m2=m2+1
while m2<31:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2=m2+1
m2=53
while m2 < 63:
    sht2.range(m2, n2).value = sht1.range(m2-20, n1).value
    m2 = m2 + 1
sht2.range(63, n2).value = sht1.range(49, n1).value   #固定资产
sht2.range(64, n2).value = sht1.range(43, n1).value   #固定资产原价
sht2.range(65, n2).value = sht1.range(44, n1).value   #累计折旧
sht2.range(66, n2).value = sht1.range(46, n1).value   #固定资产减值准备
sht2.range(67, n2).value = sht1.range(50, n1).value   #在建工程
m2 = 68
while m2 < 79:
    sht2.range(m2, n2).value = sht1.range(m2-15, n1).value
    m2 = m2 + 1
sht2.range(79, n2).value = sht1.range(82, n1).value   #资产总计


# 本年资产取数
n2 = 3 # n取粘数的列
n1 = 3
m2 = 7  # m取行
while m2 < 21:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1
sht2.range(m2, n2).value = sht1.range(m2, n1).value  # m=21时，取应收股利
m2=m2+1
while m2 < 31:
    sht2.range(m2, n2).value = sht1.range(m2 + 1, n1).value
    m2 = m2 + 1
m2 = 53
while m2 < 63:
    sht2.range(m2, n2).value = sht1.range(m2 - 20, n1).value
    m2 = m2 + 1
sht2.range(63, n2).value = sht1.range(49, n1).value  # 固定资产
sht2.range(64, n2).value = sht1.range(43, n1).value  # 固定资产原价
sht2.range(65, n2).value = sht1.range(44, n1).value  # 累计折旧
sht2.range(66, n2).value = sht1.range(46, n1).value  # 固定资产减值准备
sht2.range(67, n2).value = sht1.range(50, n1).value  # 在建工程
m2 = 68
while m2 < 79:
    sht2.range(m2, n2).value = sht1.range(m2 - 15, n1).value
    m2 = m2 + 1
sht2.range(79, n2).value = sht1.range(82, n1).value  # 资产总计

sht2 = wb2.sheets["资产负债表 (续)"]
# 上年负债取数
n2 = 5  # n取粘数的列
n1 = 10
m2 = 7  # m取行
while m2 < 28:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1
sht2.range(m2, n2).value = sht1.range(m2, n1).value  # m=28时，取应付股利
m2 = m2 + 1
while m2 < 43:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 51:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1
m2 = m2 + 1
while m2 < 80:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1

# 期初负债取数
n2 = 4  # n取粘数的列
n1 = 9
m2 = 7  # m取行
while m2 < 28:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1
sht2.range(m2, n2).value = sht1.range(m2, n1).value  # m=28时，取应付股利
m2 = m2 + 1
while m2 < 43:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 51:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1
m2 = m2 + 1
while m2 < 80:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1

# 本年负债取数
n2 = 3  # n取粘数的列
n1 = 8
m2 = 7  # m取行
while m2 < 28:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1
sht2.range(m2, n2).value = sht1.range(m2, n1).value  # m=21时，取应付股利
m2 = m2 + 1
while m2 < 43:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 51:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1
m2 = m2 + 1
while m2 < 80:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1


#利润表取数
sht2 = wb2.sheets["合并利润表"]
sht1 = wb1.sheets["Z02 利润表(企财02表)"]
# 取上年数
n2 = 4  # n取粘数的列
n1 = 4
m2 = 6  # m取行
while m2 < 8:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2 = m2 + 1
while m2 < 13:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 42:
    sht2.range(m2, n2).value = sht1.range(m2 + 3, n1).value
    m2 = m2 + 1

n1 = 8
while m2 < 77:
    sht2.range(m2, n2).value = sht1.range(m2-37, n1).value
    m2 = m2 + 1

# 取本年数
n2 = 3  # n取粘数的列
n1 = 3
m2 = 6  # m取行
while m2 < 8:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2 = m2 + 1
while m2 < 13:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 42:
    sht2.range(m2, n2).value = sht1.range(m2 + 3, n1).value
    m2 = m2 + 1

n1 = 7
while m2 < 77:
    sht2.range(m2, n2).value = sht1.range(m2-37, n1).value
    m2 = m2 + 1

# 现金流量表取数
sht2 = wb2.sheets["合并现金流量表"]
sht1 = wb1.sheets["Z03 现金流量表(企财03表)"]
# 取上年数
n2 = 4  # n取粘数的列
n1 = 4
m2 = 7  # m取行
while m2 < 35:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1

n1 = 8
while m2 < 64:
    sht2.range(m2, n2).value = sht1.range(m2 - 30, n1).value
    m2 = m2 + 1

# 取本年数
n2 = 3  # n取粘数的列
n1 = 3
m2 = 7  # m取行
while m2 < 35:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1

n1 = 7
while m2 < 64:
    sht2.range(m2, n2).value = sht1.range(m2 - 30, n1).value
    m2 = m2 + 1

# 所有者权益变动表取数
sht2 = wb2.sheets["合并所有者权益变动表"]
sht1 = wb1.sheets["Z04 所有者权益变动表(企财04表)"]
# 取本年数
m2 = 9  # m取行
while m2 < 42:
    n2 = 2  # n取粘数的列
    while n2<16:
        sht2.range(m2, n2).value = sht1.range(m2, n2+1).value
        n2=n2+1
    m2 = m2 + 1
#取上年数
sht2 = wb2.sheets["合并所有者权益变动表 (续)"]
m2 = 9  # m取行
while m2 < 42:
    n2 = 2  # n取粘数的列
    while n2 < 16:
        sht2.range(m2, n2).value = sht1.range(m2, n2 + 15).value
        n2 = n2 + 1
    m2 = m2 + 1

wb1.close()
wb1 = app.books.open(r'D:\百度企业网盘同步\BaiduBusinessSync\我的文件\王小芳组2021年度审计资料\2021年审计报告\久其导表\常州铁道高等职业技术学校(1232040046728393XP0).XLS')
sht2 = wb2.sheets["资产负债表"]

sht1 = wb1.sheets["Z01 资产负债表(企财01表)"]

#上年资产取数
n2=6  #n取粘数的列
n1=5
m2=6 #m取行
while m2<21:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2=m2+1
sht2.range(m2, n2).value = sht1.range(m2, n1).value   #m=21时，取应收股利
m2=m2+1
while m2<31:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2=m2+1
m2=53
while m2 < 63:
    sht2.range(m2, n2).value = sht1.range(m2-20, n1).value
    m2 = m2 + 1
sht2.range(63, n2).value = sht1.range(49, n1).value   #固定资产
sht2.range(64, n2).value = sht1.range(43, n1).value   #固定资产原价
sht2.range(65, n2).value = sht1.range(44, n1).value   #累计折旧
sht2.range(66, n2).value = sht1.range(46, n1).value   #固定资产减值准备
sht2.range(67, n2).value = sht1.range(50, n1).value   #在建工程
m2 = 68
while m2 < 79:
    sht2.range(m2, n2).value = sht1.range(m2-15, n1).value
    m2 = m2 + 1
sht2.range(79, n2).value = sht1.range(82, n1).value   #资产总计

sht2 = wb2.sheets["资产负债表 (续)"]
# 上年负债取数
n2 = 6  # n取粘数的列
n1 = 10
m2 = 7  # m取行
while m2 < 28:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1
sht2.range(m2, n2).value = sht1.range(m2, n1).value  # m=28时，取应付股利
m2 = m2 + 1
while m2 < 43:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 51:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1
m2 = m2 + 1
while m2 < 80:
    sht2.range(m2, n2).value = sht1.range(m2+3, n1).value
    m2 = m2 + 1

#利润表取数
sht2 = wb2.sheets["母公司利润表"]
sht1 = wb1.sheets["Z02 利润表(企财02表)"]
# 取上年数
n2 = 4  # n取粘数的列
n1 = 4
m2 = 6  # m取行
while m2 < 8:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2 = m2 + 1
while m2 < 13:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 42:
    sht2.range(m2, n2).value = sht1.range(m2 + 3, n1).value
    m2 = m2 + 1

n1 = 8
while m2 < 77:
    sht2.range(m2, n2).value = sht1.range(m2-37, n1).value
    m2 = m2 + 1

# 取本年数
n2 = 3  # n取粘数的列
n1 = 3
m2 = 6  # m取行
while m2 < 8:
    sht2.range(m2, n2).value = sht1.range(m2-1, n1).value
    m2 = m2 + 1
while m2 < 13:
    sht2.range(m2, n2).value = sht1.range(m2+1, n1).value
    m2 = m2 + 1
while m2 < 42:
    sht2.range(m2, n2).value = sht1.range(m2 + 3, n1).value
    m2 = m2 + 1

n1 = 7
while m2 < 77:
    sht2.range(m2, n2).value = sht1.range(m2-37, n1).value
    m2 = m2 + 1

# 现金流量表取数
sht2 = wb2.sheets["母公司现金流量表"]
sht1 = wb1.sheets["Z03 现金流量表(企财03表)"]
# 取上年数
n2 = 4  # n取粘数的列
n1 = 4
m2 = 7  # m取行
while m2 < 35:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1

n1 = 8
while m2 < 64:
    sht2.range(m2, n2).value = sht1.range(m2 - 30, n1).value
    m2 = m2 + 1

# 取本年数
n2 = 3  # n取粘数的列
n1 = 3
m2 = 7  # m取行
while m2 < 35:
    sht2.range(m2, n2).value = sht1.range(m2 - 1, n1).value
    m2 = m2 + 1

n1 = 7
while m2 < 64:
    sht2.range(m2, n2).value = sht1.range(m2 - 30, n1).value
    m2 = m2 + 1

# 所有者权益变动表取数
sht2 = wb2.sheets["母公司所有者权益变动表"]
sht1 = wb1.sheets["Z04 所有者权益变动表(企财04表)"]
# 取本年数
m2 = 9  # m取行
while m2 < 42:
    n2 = 2  # n取粘数的列
    while n2<16:
        sht2.range(m2, n2).value = sht1.range(m2, n2+1).value
        n2=n2+1
    m2 = m2 + 1
#取上年数
sht2 = wb2.sheets["母公司所有者权益变动表 (续)"]
m2 = 9  # m取行
while m2 < 42:
    n2 = 2  # n取粘数的列
    while n2 < 16:
        sht2.range(m2, n2).value = sht1.range(m2, n2 + 15).value
        n2 = n2 + 1
    m2 = m2 + 1





wb2.save(r"D:\百度企业网盘同步\BaiduBusinessSync\我的文件\王小芳组2021年度审计资料\2021年审计报告\0203生成报表\2-{}2021年度财务报表.xlsx".format(name))
wb2.close()
wb1.close()

app.quit()
