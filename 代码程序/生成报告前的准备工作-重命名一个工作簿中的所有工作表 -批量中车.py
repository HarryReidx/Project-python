import xlwings as xw
import re
import pandas as pd
import glob

app = xw.App(visible=True, add_book=True)


filearray=[]
filelocation=glob.glob(r"E:\久其导表\合并表\*XLS") #久其导出报表所在文件夹
for filename in filelocation:
    workbook = app.books.open(filename)
    bookname=filename.split("\\")[-1]
    print(bookname)
    worksheet = workbook.sheets
    for i in worksheet:
        name=i.name
        if name == "QCF10 受限制的货币资金明细":
            m = 4
            while m < 24:
                n = 1
                while n < 4:
                    if len(str(i.range(m, n).value))==0:
                    #     i.range(m, n).value=i.range(m, n).value
                    # else:
                        i.range(m, n).value = 0
                    n = n + 1

                m = m + 1

            i.range(4, 6).value = "浮动行合计"
            i.range(3, 6).value = "合计"
            i.range(f"G3:H4").number_format = "#,##0.00"
            n0=4
            while n0<24:
                    if  i.range(n0, 1).value=="合    计":
                        m0=2
                        i.range(3, 7).value = i.range(n0, m0).value
                        # if i.range(n0, m0).value<>0
                        qt = i.range(n0, m0).value-i.range(4, m0).value-i.range(5, m0).value-i.range(6, m0).value-i.range(7, m0).value-i.range(8, m0).value-i.range(9, m0).value-i.range(11, m0).value
                        i.range(4, 7).value =qt
                        m0=3
                        i.range(3, 8).value = i.range(n0, m0).value
                        qt = i.range(n0, m0).value - i.range(4, m0).value - i.range(5, m0).value - i.range(6,m0).value - i.range(7, m0).value - i.range(8, m0).value - i.range(9, m0).value - i.range(11, m0).value
                        i.range(4, 8).value = qt
                    n0=n0+1
            area = i.range("F:H")
            area.column_width = 25

        if name == "QCF48 合同资产情况":
            m = 5
            while m < 7:
                n = 2
                while n < 8:
                    if len(str(i.range(m, n).value))==0:
                    #     i.range(m, n).value=i.range(m, n).value
                    # else:
                        i.range(m, n).value = 0
                    n = n + 1
                m = m + 1

            i.range("K4").value = "非质保金合计"
            i.range(f"L4:Q4").number_format = "#,##0.00"
            n0=2
            while n0<8:
                i.range(4,n0+10).value= i.range(5, n0).value- i.range(6, n0).value
                n0=n0+1
            area = i.range("K:Q")
            area.column_width = 25


        if name == "QCF52 其他流动资产":
            m = 5
            while m < 11:
                n = 2
                while n < 8:
                    if len(str(i.range(m, n).value)) == 0:
                        #     i.range(m, n).value=i.range(m, n).value
                        # else:
                        i.range(m, n).value = 0
                    n = n + 1
                m = m + 1

            i.range("K4").value = "浮动行合计"
            i.range(f"L4:Q4").number_format = "#,##0.00"
            n0 = 2
            while n0 < 8:
                i.range(4, n0 + 10).value = i.range(5, n0).value - i.range(6, n0).value-i.range(7, n0).value- i.range(8, n0).value- i.range(9, n0).value- i.range(10, n0).value
                n0 +=1
            area = i.range("K:Q")
            area.column_width = 25

        if name == "QCF90 递延所得税资产和递延所得税负债不以抵销后的净额列示":
            i.api.Rows(6).Insert()
            i.range(6,2).value="期末余额"
            i.range(6, 3).value= "期末暂时性差异"
            i.range(6, 4).value= "期初余额"
            i.range(6, 5).value= "期初暂时性差异"
            table = i.range('A6:E24').options(pd.DataFrame).value
            # table.columns= ["期末余额", "期末暂时性差异", "期初余额","期初暂时性差异"]
            result = table.sort_values(by=["期末余额", "期初余额"], ascending=False)
            # product = table[table["单位名称"] == "——"]
            # print(table)
            i.range("H25").value = i.range("A25").value
            i.range("I25").value = i.range("B25").value
            i.range("J25").value = i.range("C25").value
            i.range("K25").value = i.range("D25").value
            i.range("L25").value = i.range("E25").value

            i.range(f"I6:L25").number_format = "#,##0.00"
            i.range('H6').value = result
            area = i.range("H:L")
            area.column_width = 25

            n0=25

            while n0 < 100:
                if i.range(n0, 1).value == "二、递延所得税负债":
                    # print(n0)
                    i.api.Rows(n0).Insert()
                    i.range(n0, 1).value = "项目"
                    i.range(n0, 2).value = "期末余额"
                    i.range(n0, 3).value = "期末暂时性差异"
                    i.range(n0, 4).value = "期初余额"
                    i.range(n0, 5).value = "期初暂时性差异"
                    # table = i.range("A31:E42").options(pd.DataFrame).value
                    table = i.range((n0,1),(n0+11,5)).options(pd.DataFrame).value
                    # table.columns= ["期末余额", "期末暂时性差异", "期初余额","期初暂时性差异"]
                    result = table.sort_values(by=["期末余额", "期初余额"], ascending=False)
                    # product = table[table["单位名称"] == "——"]
                    # print(table)
                    i.range("o18").value = i.range(n0+12,1).value
                    i.range("p18").value = i.range(n0+12,2).value
                    i.range("q18").value = i.range(n0+12,3).value
                    i.range("r18").value = i.range(n0+12,4).value
                    i.range("s18").value = i.range(n0+12,5).value

                    i.range(f"P7:S18").number_format = "#,##0.00"
                    i.range('O6').value = result
                    area = i.range("O:S")
                    area.column_width = 25
                    break
                n0 =n0+1

        if name == "QCF95 其他非流动资产":
            m = 5
            info = i.used_range
            nrows = info.last_cell.row
            ncolumns = info.last_cell.column
            while m < nrows+1:
                n = 2
                while n < 8:
                    if len(str(i.range(m, n).value)) == 0:
                        #     i.range(m, n).value=i.range(m, n).value
                        # else:
                        i.range(m, n).value = 0
                    n = n + 1
                m = m + 1
            i.range("K4").value = "合计"
            i.range("K5").value = "浮动行合计"
            i.range(f"L4:Q5").number_format = "#,##0.00"
            n0 = 2
            # print(nrows)
            while n0 < 8:
                i.range(4, n0 + 10).value=i.range(nrows, n0).value
                # print( i.range(4, n0 + 10).value)
                i.range(5, n0 + 10).value = i.range(4, n0+10).value - i.range(5, n0).value - i.range(6, n0).value - i.range(
                    7, n0).value - i.range(8, n0).value
                n0 += 1
            area = i.range("K:Q")
            area.column_width = 25

        if name=="QCF110 应交税费":
            i.range(36, 1).value="其他税费合计"
            for m0 in range(2,6):
                n0=15
                i.range(36, m0).value=0
                i.range(36, m0).number_format ="#,##0.00"
                # i.range(36, m0).type =float(2)
                while n0<27:
                    i.range(36, m0).value +=i.range(n0, m0).value
                    n0 +=1

        if name == "QCF142_A3 管理费用":
            # i.autofit(axis="c")
            i.range(21, 4).value ="其中：年度决算审计费用"
            i.range(50, 4).value = "其他"
            table = i.range('A3').expand('table').options(pd.DataFrame).value
            product = table[table["单位名称"] == "——"]
            result = product.sort_values(by=["本期发生额","上期发生额"], ascending=False)
            # result = product.sort_values(by="本期发生额", ascending=False)
            i.range(f"I3:J50").number_format = "#,##0.00"
            i.range('H3').value = result
            i.range(50, 2).value = 0
            for m0 in range(9, 11):
                n0=5
                if i.range(4, m0).value==None:
                    i.range(50, m0).value =0
                else:
                    i.range(50, m0).value = i.range(4, m0).value
                while n0 < 15:
                    if i.range(n0, m0).value== None:
                        i.range(n0, m0).value=0
                    i.range(50, m0).value -= i.range(n0, m0).value
                    n0 += 1
            area=i.range("H:J")
            area.column_width=25

        if name == "QCF142_A2 销售费用":
            # i.autofit(axis="c")
            i.range(31, 4).value = "其他"
            table = i.range('A3').expand('table').options(pd.DataFrame).value
            product = table[table["单位名称"] == "——"]
            result = product.sort_values(by=["本期发生额", "上期发生额"], ascending=False)
            # result = product.sort_values(by="本期发生额", ascending=False)
            i.range(f"I3:J50").number_format = "#,##0.00"
            i.range('H3').value = result
            i.range(50, 2).value = 0
            for m0 in range(9, 11):
                n0 = 5
                if i.range(4, m0).value == None:
                    i.range(50, m0).value = 0
                else:
                    i.range(50, m0).value = i.range(4, m0).value
                while n0 < 15:
                    if i.range(n0, m0).value== None:
                        i.range(n0, m0).value=0
                    i.range(50, m0).value -= i.range(n0, m0).value
                    n0 += 1
            area = i.range("H:J")
            area.column_width = 25

        if name == "QCF142_A4 研发费用":
            # i.autofit(axis="c")
            i.range(27, 4).value = "其他"
            table = i.range('A3').expand('table').options(pd.DataFrame).value
            product = table[table["单位名称"] == "——"]
            result = product.sort_values(by=["本期发生额", "上期发生额"], ascending=False)
            # result = product.sort_values(by="本期发生额", ascending=False)
            i.range(f"I3:J50").number_format = "#,##0.00"
            i.range('H3').value = result
            i.range(50, 2).value = 0
            for m0 in range(9, 11):
                n0 = 5
                if i.range(4, m0).value == None:
                    i.range(50, m0).value = 0
                else:
                    i.range(50, m0).value = i.range(4, m0).value
                while n0 < 15:
                    if i.range(n0, m0).value== None:
                        i.range(n0, m0).value=0
                    i.range(50, m0).value -= i.range(n0, m0).value
                    n0 += 1
            area = i.range("H:J")
            area.column_width = 25

        if name == "QCF142_A5 财务费用":
            i.range(29, 1).value = "其他财务费用合计"
            for m0 in range(2, 4):
                n0 = 20
                i.range(29, m0).value = 0
                i.range(29, m0).number_format = "#,##0.00"
                # i.range(36, m0).type =float(2)
                while n0 < 24:
                    i.range(29, m0).value += i.range(n0, m0).value
                    n0 += 1

        if name == "QCF149 营业外收入":
            info = i.used_range
            nrows = info.last_cell.row
            ncolumns = info.last_cell.column
            m = 4
            while m < nrows+1:
                n = 1
                while n < ncolumns+1:
                    if len(str(i.range(m, n).value)) == 0:
                        #     i.range(m, n).value=i.range(m, n).value
                        # else:
                        i.range(m, n).value = 0
                    n = n + 1
                m = m + 1
            i.api.Rows(4).Delete()
            table = i.range('A3').expand('table').options(pd.DataFrame).value
            # table.columns=["本期发生额","上期发生额","计入当期非经常性损益的金额"]
            # print(table.shape)
            product = table.iloc[[0,5,6,7,8,9,10,11]]
            result = product.sort_values(by=["本期发生额", "上期发生额"], ascending=False)
            # result = product.sort_val ues(by="本期发生额", ascending=False)
            i.range(f"I3:K50").number_format = "#,##0.00"
            i.range('H3').value = result
            i.range('H12').value=i.range('A16').value
            i.range('I12').value = i.range('B16').value
            i.range('J12').value = i.range('C16').value
            i.range('K12').value = i.range('D16').value
            i.range('H13').value = i.range(nrows-1,1).value
            i.range('I13').value = i.range(nrows-1,2).value
            i.range('J13').value = i.range(nrows-1,3).value
            i.range('K13').value = i.range(nrows-1,4).value
            i.range('H14').value = "本期计入当期损益的政府补助金额"
            sht=workbook.sheets["QCF142"]
            if len(str(sht.range('B8').value))==0 or sht.range('B8').value==None:
                sht.range('B8').value=0
            i.range('I14').value= i.range('B13').value+sht.range('B8').value

            area = i.range("H:K")
            area.column_width = 25

        if name == "QCF151 营业外支出":
            info = i.used_range
            nrows = info.last_cell.row
            ncolumns = info.last_cell.column
            m = 4
            while m < nrows + 1:
                n = 1
                while n < ncolumns + 1:
                    if len(str(i.range(m, n).value)) == 0:
                        #     i.range(m, n).value=i.range(m, n).value
                        # else:
                        i.range(m, n).value = 0
                    n = n + 1
                m = m + 1
            i.api.Rows(4).Delete()
            table = i.range('A3').expand('table').options(pd.DataFrame).value
            # table.columns = ["本期发生额", "上期发生额", "计入当期非经常性损益的金额"]
            # print(table.shape)
            product = table.iloc[[0, 5, 6, 7, 8, 9, 10, 11,12,13,14,15]]
            result = product.sort_values(by=["本期发生额", "上期发生额"], ascending=False)
            # result = product.sort_val ues(by="本期发生额", ascending=False)
            i.range(f"I3:K50").number_format = "#,##0.00"
            i.range('H3').value = result
            i.range('H16').value = i.range('A20').value
            i.range('I16').value = i.range('B20').value
            i.range('J16').value = i.range('C20').value
            i.range('K16').value = i.range('D20').value
            i.range('H17').value = i.range(nrows-1, 1).value
            i.range('I17').value = i.range(nrows-1, 2).value
            i.range('J17').value = i.range(nrows-1, 3).value
            i.range('K17').value = i.range(nrows-1, 4).value
            area = i.range("H:K")
            area.column_width = 25


        if name !="Z01 资产负债表(企财01表)" and name !="Z02 利润表(企财02表)" and name !="Z03 现金流量表(企财03表)" and name !="Z04 所有者权益变动表(企财04表)":
            newname=name.split()[0]
            i.name =newname
            info = i.used_range
            nrows = info.last_cell.row
            ncolumns = info.last_cell.column
            m=1
            while m<nrows+1:
                n=1
                while n < ncolumns + 1:
                    if i.range(m, n).value==0:
                        i.range(m, n).value =None
                    n=n+1

                m=m+1
        # print(name)




    workbook.save(r"E:\生成附注\{}".format(bookname))
    workbook.close()


app.quit()
