from pathlib import Path, PurePath

import xlrd
import xlwt

# 指定要合并excel的路径
src_path = './test_files/调查问卷'
# 指定合并完成的路径
dist_file = "./test_files/调查问卷/结果.xlsx"

# 取得该目录下所有的xlsx格式文件
p = Path(src_path)
files = [x for x in p.iterdir() if PurePath(x).match('*.xlsx')]

# 准备一个列表存放读取结果
content = []

# 对每一个文件进行重复处理
for file in files:
    # 用文件名作为每个用户标识
    username = file.stem
    data = xlrd.open_workbook(file)
    table = data.sheets()[0]
    # 取得每一项的结果
    answer1 = table.cell_value(rowx=4, colx=4)
    answer2 = table.cell_value(rowx=10, colx=4)
    temp = f'{username},{answer1},{answer2}'
    # 合并为一行先存储起来
    content.append(temp.split(','))
    print(temp)

# 准备写入文件的表头
table_header = ['员工姓名', '第一题', '第二题']

workbook = xlwt.Workbook(encoding='utf-8')
xlsheet = workbook.add_sheet("统计结果")

# 写入表头
row = 0
col = 0
for cell_header in table_header:
    xlsheet.write(row, col, cell_header)
    col += 1

# 向下移动一行
row += 1
# 取出每一行内容
for line in content:
    col = 0
    # 取出每个单元格内容
    for cell in line:
        # 写入内容
        xlsheet.write(row, col, cell)
        # 向右移动一个单元格
        col += 1
    # 向下移动一行
    row += 1

# 保持最终结果
workbook.save(dist_file)