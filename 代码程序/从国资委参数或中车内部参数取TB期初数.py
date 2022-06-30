import xlwings as xw
import random
import re

app = xw.App(visible = True, add_book = True)
# 上期用1  本期用2
wb1=app.books.open(r'D:\2021年度审计\TB模板\集团内部参数2020.xls')
wb2=app.books.open(r'D:\2021年度审计\TB模板\TB（CRRC00-中国中车） -1116.xls')
sht1=wb1.sheets["Z01 资产负债表(企财01表)"]
sht2=wb2.sheets["未审资产负债表"]
m1=3
m2=5
t2=4
n2=6


while n2<18:
    n1=n2
    sht2.range(n2, m2).value =  sht1.range(n1, m1).value
    sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

n2=23
while n2<43:
    if not n2 in [19,31,32]:
        n1=n2-2
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
        sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

n2=43    #固定资产原值
n1=42
sht2.range(n2, m2).value =  sht1.range(n1, m1).value
sht2.range(n2, t2).value = sht1.range(n1, m1).value

n2=44    #固定资产累计折旧
n1=43
sht2.range(n2, m2).value =  sht1.range(n1, m1).value
sht2.range(n2, t2).value = sht1.range(n1, m1).value


n2=46    #固定资产减值准备
n1=44
sht2.range(n2, m2).value =  sht1.range(n1, m1).value
sht2.range(n2, t2).value = sht1.range(n1, m1).value


n2=48    #固定资产清理
sht2.range(n2, m2).value =  sht1.range(41, m1).value-sht1.range(42, m1).value+sht1.range(43, m1).value+sht1.range(44, m1).value
sht2.range(n2, t2).value = sht1.range(41, m1).value-sht1.range(42, m1).value+sht1.range(43, m1).value+sht1.range(44, m1).value

n2=51    #在建工程，由于在建工程是浮动行，目前暂时无法取工程物资，都放到在建工程一行，如有工程物资，手动调整
sht2.range(n2, m2).value =  sht1.range(45, m1).value
sht2.range(n2, t2).value = sht1.range(45, m1).value

n2=53
while n2<63:
    n1=n2-7
    sht2.range(n2, m2).value =  sht1.range(n1, m1).value
    sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1


m1=7
n2=71
while n2<91:
    n1=n2-65
    sht2.range(n2, m2).value =  sht1.range(n1, m1).value
    sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

n2=95
while n2<108:
    if not n2 in [100,101]:
        n1=n2-67
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
        sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

n2=111
while n2<146:
    if not n2 in [117,118,119,120,127,136,144]:
        n1=n2-69
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
        sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1


sht1=wb1.sheets["QCF32 其他应收款"]
sht2.range(20, m2).value = sht1.range(4, 2).value
sht2.range(20, t2).value = sht1.range(4, 2).value
sht2.range(21, m2).value = sht1.range(5, 2).value
sht2.range(21, t2).value = sht1.range(5, 2).value
sht2.range(22, m2).value = sht1.range(6, 2).value
sht2.range(22, t2).value = sht1.range(6, 2).value

sht1=wb1.sheets["QCF111 其他应付款"]
sht2.range(92, m2).value = sht1.range(4, 2).value
sht2.range(92, t2).value = sht1.range(4, 2).value
sht2.range(93, m2).value = sht1.range(5, 2).value
sht2.range(93, t2).value = sht1.range(5, 2).value
sht2.range(94, m2).value = sht1.range(6, 2).value
sht2.range(94, t2).value = sht1.range(6, 2).value

sht1=wb1.sheets["QCF127 长期应付款"]
sht2.range(109, m2).value = sht1.range(4, 5).value
sht2.range(109, t2).value = sht1.range(4, 5).value
sht2.range(110, m2).value = sht1.range(5, 5).value
sht2.range(110, t2).value = sht1.range(5, 5).value



sht1=wb1.sheets["Z02 利润表(企财02表)"]
sht2=wb2.sheets["未审利润表"]
m1=3
m2=4
n2=7
sht2.range(n2, m2).value =  sht1.range(6, m1).value  #未区分主营业务收入和其他业务收入，将营业收入放在了主营业务收入
n2=14
sht2.range(n2, m2).value =  sht1.range(11, m1).value  #未区分主营业务成本和其他业务成本，将营业收入放在了主营业务成本
n2=16
while n2<45:
    if not n2 in [12,13,42]:
        n1=n2-4
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value

    n2=n2+1

m1=7
n2=45
while n2<74:
    if not n2 in [46,49,48,50,52,55,56,57,63,74,75,76]:
        n1=n2-40
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
    n2=n2+1

sht1=wb1.sheets["Z04 所有者权益变动表(企财04表)"]
n2=81
m1=13
while n2<112:
    if not n2 in [82,83,85,86,89,90,91,94,95,97,88,93,96,97,106,107,108]:
        n1 = n2 - 72
        sht2.range(n2, m2).value = sht1.range(n1, m1).value
    n2 = n2 + 1


sht1=wb1.sheets["Z03 现金流量表(企财03表)"]
sht2=wb2.sheets["未审现金流量表"]
m1=3
m2=4
n2=6
while n2<32:
    if not n2 in [20,31,32]:
        n1=n2
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value

    n2=n2+1

m1=7
n2=34
while n2<62:
    if not n2 in [39,45,46,52,57,58,60]:
        n1=n2-29
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
    n2=n2+1


sht1=wb1.sheets["Z04 所有者权益变动表(企财04表)"]
sht2=wb2.sheets["审定所有者权益变动表"]
m2=16 #实收资本
for n2 in [7,10,15,16,17,18,31,33,34,38]:
    n1=n2+2
    m1=m2-13
    sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [17,18,19]: #其他权益工具
    for n2 in[7,10,16,31,40]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value


for m2 in [20] :#资本公积
    for n2 in[7,10,15,16,17,18,31,33,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [21]: #库存股
    for n2 in[7,10,18,31,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [22]: #其他综合收益
    for n2 in[7,10,18,31,36,37,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [23] :#专项储备
    for n2 in[7,10,18,20,21,31,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value


for m2 in [24] :#盈余公积
    for n2 in[7,10,18,31,34,35,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [25]: #一般风险准备
    for n2 in[7,10,18,31,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

for m2 in [28]: #少数股东权益
    for n2 in[7,10,15,16,17,18,20,21,30,31,38]:
        n1=n2+2
        m1=m2-13
        sht2.range(n2, m2).value = sht1.range(n1, m1).value

wb2.save(r'D:\2021年度审计\TB模板\TB（CRRC00-中国中车）-new -1116.xls')
wb2.close()
wb1.close()
app.quit()






