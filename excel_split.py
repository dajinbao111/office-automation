from pathlib import Path, PurePath

import xlrd
import xlwt

# 工资单文件
salary_file = './test_files/工资单/工资.xlsx'
# 拆分文件保持路径
dist_dir = "./test_files/工资单/"
data = xlrd.open_workbook(Path(salary_file))
table = data.sheets()[0]
# 取得表头
salary_header = table.row_values(rowx=0, start_colx=0, end_colx=None)


# 定义写入文件函数
def write_to_file(filename, cnt):
    workbook = xlwt.Workbook(encoding='utf-8')
    xlsheet = workbook.add_sheet("工资明细")

    row = 0
    for line in cnt:
        col = 0
        for cell in line:
            xlsheet.write(row, col, cell)
            col += 1
        row += 1
    workbook.save(PurePath(salary_file).with_name(filename).with_suffix(".xlsx"))


employee_number = table.nrows
for line in range(1, employee_number):
    content = table.row_values(rowx=line, start_colx=0, end_colx=None)
    # 将表头和员工数量重新组成一个新的文件
    new_content = []
    # 增加表头到要写的内容中
    new_content.append(salary_header)
    # 增加员工工资到要写入的内容中
    new_content.append(content)

    write_to_file(filename=content[1], cnt=new_content)
