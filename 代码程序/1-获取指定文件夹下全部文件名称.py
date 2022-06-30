import os
import xlwt

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('My Worksheet')
src_dir=r"C:\projdata\A\电子底稿" #指定文件夹路径，根据想要获取文件名称的文件夹修改
# 源文件目录地址
def list_all_files(rootdir):
    import os
    _files=[]

    # 列出文件夹下所有的目录与文件
    list_file=os.listdir(rootdir)

    for i in range(0, len(list_file)):

        # 构造路径
        path=os.path.join(rootdir, list_file[i])

        # 判断路径是否是一个文件目录或者文件
        # 如果是文件目录，继续递归

        if os.path.isdir(path):
            _files.extend(list_all_files(path))
        if os.path.isfile(path):
            _files.append(path)
    return _files

files = list_all_files(src_dir)
i=0
for file in files:
    a=len(os.path.splitext(file)[1])
    worksheet.write(i, 0, os.path.dirname(file)) #获取文件路径
    worksheet.write(i, 1, os.path.basename(file)[:-a]) #获取文件名
    worksheet.write(i, 3, os.path.splitext(file)[1]) #分离文件名与扩展名
    worksheet.write(i, 4, os.path.basename(file))
    i=i+1


workbook.save(r"C:\Users\K\Desktop\文件目录.XLS")   #生成文件保存路径


