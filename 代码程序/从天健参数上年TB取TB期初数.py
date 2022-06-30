import xlwings as xw
import random
import re

app = xw.App(visible = True, add_book = True)
# 上期用1  本期用2
wb1=app.books.open(r'D:\2021年度审计\TB模板\石家庄印钞.xls')
wb2=app.books.open(r'D:\2021年度审计\TB模板\TB（CRRC00-中国中车） -1116.xls')
sht1=wb1.sheets["Z01 资产负债表(企财01表)"]
sht2=wb2.sheets["未审资产负债表"]
m1=3
m2=5
t2=4
n2=6


while n2<63:
    if not n2 in [19,31,32,45,47,49,50]:
        n1=n2
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
        sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

m1=7
n2=71
while n2<147:
    if not n2 in [91,100,101,108,117,118,119,120,127,136,144,146,147]:
        n1=n2-65
        sht2.range(n2, m2).value =  sht1.range(n1, m1).value
        sht2.range(n2, t2).value = sht1.range(n1, m1).value
    n2=n2+1

sht1=wb1.sheets["Z02 利润表(企财02表)"]
sht2=wb2.sheets["未审利润表"]
m1=3
m2=4
n2=7
while n2<45:
    if not n2 in [12,13,42]:
        n1=n2
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






