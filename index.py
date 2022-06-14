import xlrd
import os
import re
#file_path = './files/'
file_path = 'F:/WechatFiles/WeChat Files/wxid_7uwuevbvm3iv21/FileStorage/MsgAttach/22c54f9f4f35259c421aa6a3156520d9/File/2022-06/合格/'
excel_path = input('excel:')
file_path = input('文件夹：') + '\\'

def getNameList():
    xlsx = xlrd.open_workbook('./电子合格.xls')

    # 通过sheet名查找：xlsx.sheet_by_name("sheet1")
    # 通过索引查找：xlsx.sheet_by_index(3)
    table = xlsx.sheet_by_index(0)

    # 获取单个表格值 (2,1)表示获取第3行第2列单元格的值
    # value = table.cell_value(2, 1)
    # print("第3行2列值为",value)

    # 获取表格行数
    nrows = table.nrows
    print("表格一共有", nrows, "行")

    # 获取第4列所有值（列表生成式）
    name_list = [str(table.cell_value(i, 0)) for i in range(0, nrows)]
    print("第1列所有的值：", name_list)
    return name_list

# 获取 文件列表 和 待命名名称列表
file_list = os.listdir(file_path)
print(file_list)
file_list.sort(key=lambda x:int(re.search(r'(\d+$)', x.split('.')[0]).group(1))) #对‘.’进行切片，并取列表的第一个值（左边的文件名）转化整数型。(其实可以用正则，匹配出文件名中的数字部分，然后转为整数型，然后排序就对了)
name_list = getNameList()
print("待处理的文件：", file_list)

for i in range(len(file_list)):

    # 设置旧文件名（就是路径+文件名）
    oldname = file_path+file_list[i]

    # 设置新文件名
    newname = file_path+name_list[i]+'.pdf'

    # 用os模块中的rename方法对文件改名
    os.rename(oldname, newname)
    # print(oldname,'======>',newname)
