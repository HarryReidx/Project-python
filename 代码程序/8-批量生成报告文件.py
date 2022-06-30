from win32com import client
# from comtypes.client import CreateObject
import glob
import os
from pathlib import Path
from PyPDF2 import PdfFileReader,PdfFileMerger
import xlrd




def excel2pdf(filepath, name, pdfname):
    exceldir = filepath
    # 指定Excel类型
    excel = client.DispatchEx("Excel.Application")
    # 使用Excel软件打开a.xls
    file = excel.Workbooks.Open(f"{exceldir}\{name}", False)
    # 文件另存为当前目录下的pdf文件
    file.ExportAsFixedFormat(0, f"{exceldir}\{pdfname}")
    # 关闭文件
    file.Close()
    # 结束excel应用程序进程
    excel.Quit()




def word2pdf(filepath, wordname, pdfname):
    worddir = filepath
    # 指定Word类型
    word = client.DispatchEx("Word.Application")
    # 使用Word软件打开a.doc
    file = word.Documents.Open(f"{worddir}\{wordname}", ReadOnly=1)
    # 文件另存为当前目录下的pdf文件
    file.ExportAsFixedFormat(f"{worddir}\{pdfname}",17,Item=7, CreateBookmarks=0)
    # file.ExportAsFixedFormat(f"{worddir}\{pdfname}", FileFormat=17, Item=7, CreateBookmarks=0)
    # 关闭文件
    file.Close()
    # 结束word应用程序进程
    word.Quit()


data = xlrd.open_workbook(r"C:\Users\os\Desktop\文件目录.XLS")
table = data.sheet_by_name("My Worksheet")
i=0
while i<200:

    filearray=[]
    # path=Path(os.path.join(table.cell_value(i, 0)))
    path=table.cell_value(i, 0)+"\*"
    # print(path)
    hpdfname=table.cell_value(i, 1)
    filelocation=glob.glob(path)
    for filename in filelocation:
        if filename.find("pdf")>-1:
            # print(filename.find("pdf"))
            os.remove(filename)
            print("删除"+filename)
    filelocation=glob.glob(path)
    for filename in filelocation:
        print(filename)
        filearray.append(filename)
        filepath = os.path.dirname(filename) # 获取文件路径
        name=os.path.basename(filename)  # 获取含后缀文件名
        a = len(os.path.splitext(filename)[1]) #后缀长度
        pdfname= os.path.basename(filename)[:-a]+".pdf"
        gs=os.path.splitext(filename)[1]  #后缀名称

        if gs==".docx" or  gs==".doc":
            if name.find("计划")<0 and name.find("总结")<0:
            # print(filename, pdfname)
                if name.find("~$") < 0:
                    word2pdf(filepath, name, pdfname)
                    print("成功转换为pdf")
            else:
                print("计划总结不转换")


        if gs == ".xls" or gs == ".xlsx":
            if name.find("~$")<0:
                excel2pdf(filepath, name, pdfname)
                print("成功转换为pdf")

    src_folder=Path(filepath)

    des_file=Path(os.path.join(filepath, hpdfname))
    if not des_file.parent.exists():
        des_file.parent.mkdir(parents=True)
    file_list = list(src_folder.glob("*.pdf"))
    merger = PdfFileMerger()
    outputPages = 0
    for pdf in file_list:
        if os.path.basename(filename) !=hpdfname:
            inputfile = PdfFileReader(str(pdf))
            # inputfile= inputfile.payload + "&province=" + str(pro).encode("utf-8").decode("latin1")
            # inputfile= inputfile.encode("utf-8").decode("latin1")
            merger.append(inputfile)
            pageCount = inputfile.getNumPages()
            print(f'{pdf.name} 页数：{pageCount}')
            outputPages += pageCount

    merger.write(str(des_file))
    merger.close()
    print(f'\n{table.cell_value(i, 1)}合并后的总页数：{outputPages}')
    # print(i)
    i=i+1

