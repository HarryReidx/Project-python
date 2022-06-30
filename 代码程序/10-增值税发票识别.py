
import datetime
import glob
import os
import sys
import fitz
import numpy as np
import pandas as pd
from aip import AipOcr
import xlwt


APP_ID = "待补充，百度智能云账号"
API_KEY = "待补充，百度智能云账号"
SECRET_KEY = "待补充，百度智能云账号"
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
options={}

def pdf2image(pdfpath,image_path):
    with fitz.open(pdfpath) as pdfDoc:
        for pg in range(pdfDoc.pageCount):
            page = pdfDoc[pg]
            rotate = int(0)
            zoom_x = 3
            zoom_y = 3
            mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
            pix = page.getPixmap(matrix=mat, alpha=False)
            pix.writePNG(image_path)  # 将图片写入指定的文件夹内
            print('转化图片完成')

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('My Worksheet')
i=1
worksheet.write(0,i-1,"pdf文件名称")
worksheet.write(0,i,"图片文件名称")
worksheet.write(0,i+1,"发票种类")
worksheet.write(0,i+2,"发票名称")
worksheet.write(0,i+3,"发票代码")
worksheet.write(0,i+4,"发票号码")
worksheet.write(0,i+5,"开票日期")
worksheet.write(0,i+6,"购方名称")
worksheet.write(0,i+7,"购方纳税人识别号")
worksheet.write(0,i+8,"购方地址及电话")
worksheet.write(0,i+9,"购方开户行及账号")
worksheet.write(0,i+10,"销售方名称")
worksheet.write(0,i+11,"销售方纳税人识别号")
worksheet.write(0,i+12,"销售方地址及电话")
worksheet.write(0,i+13,"销售方开户行及账号")
worksheet.write(0,i+14,"合计金额")
worksheet.write(0,i+15,"合计税额")
worksheet.write(0,i+16,"价税合计(小写)")
worksheet.write(0,i+17,"货物名称")




filearray=[]
filelocation=glob.glob(r"C:\Users\DELL\Desktop\新建文件夹\*.pdf")
m=1
for filename in filelocation:
    print(filename)
    filearray.append(filename)
    imgname=filename[:-3]+"png"
    print(imgname)
    pdf2image(filename, imgname)
    ocr_result=client.vatInvoice(open(imgname, 'rb').read(),options=options)
    fp=ocr_result['words_result']
    print(fp)
    # print(fp["AmountInWords"])
    worksheet.write(m, i - 1, filename)
    worksheet.write(m, i, imgname)
    worksheet.write(m, i + 1, fp["InvoiceType"])
    worksheet.write(m, i + 2, fp["InvoiceTypeOrg"])
    worksheet.write(m, i + 3, fp["InvoiceCode"])
    worksheet.write(m, i + 4, fp["InvoiceNum"])
    worksheet.write(m, i + 5, fp["InvoiceDate"])
    worksheet.write(m, i + 6, fp["PurchaserName"])
    worksheet.write(m, i + 7, fp["PurchaserRegisterNum"])
    worksheet.write(m, i + 8, fp["PurchaserAddress"])
    worksheet.write(m, i + 9, fp["PurchaserBank"])
    worksheet.write(m, i + 10, fp["SellerName"])
    worksheet.write(m, i + 11, fp["SellerRegisterNum"])
    worksheet.write(m, i + 12, fp["SellerAddress"])
    worksheet.write(m, i + 13, fp["SellerBank"])
    worksheet.write(m, i + 14, fp["TotalAmount"])
    worksheet.write(m, i + 15, fp["TotalTax"])
    worksheet.write(m, i + 16, fp["AmountInFiguers"])
    worksheet.write(m, i + 17, str(fp["CommodityName"]))
    m=m+1



workbook.save(r"C:\Users\DELL\Desktop\发票识别.XLS")
