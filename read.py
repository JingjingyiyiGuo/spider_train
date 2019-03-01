import xlrd
import re

data = xlrd.open_workbook('./表.xlsx')

table = data.sheet_by_name(u'需求2所有车次')#通过名称获取

# col_array = table.col_values(0)
row_num = table.nrows
# print(len(col_array))
# print(row_num)
# print(str(col_array[0:5]))
# print(col_array[1].ctype)
pattern_url = 'http://search.huochepiao.com/chaxun/resultc.asp?txtCheci=D2&cc.x=0&cc.y=0'
for i in range(1, row_num):
    ctype = table.cell(i, 0).ctype
    value = table.cell(i, 0).value
    if ctype == 2:
        value = str(int(value))
        print(value)
        new_url = re.sub(r'D2', value, pattern_url, count=1, flags=0)
        print(new_url)