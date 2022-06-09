import xlrd
import os
path = './files/'

def getNameList():
    xlsx = xlrd.open_workbook('./list.xlsx')

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
    name_list = [str(table.cell_value(i, 0)) for i in range(1, nrows)]
    print("第1列所有的值：", name_list)
    return name_list

# 获取 文件列表 和 待命名名称列表
file_list = os.listdir(path)
name_list = getNameList()

for i in range(len(file_list)):

    # 设置旧文件名（就是路径+文件名）
    oldname = path+file_list[i]

    # 设置新文件名
    newname = path+name_list[i]+'.txt'

    # 用os模块中的rename方法对文件改名
    os.rename(oldname, newname)
    # print(oldname,'======>',newname)
