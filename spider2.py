#!/usr/bin/python3
# coding: utf-8
import requests
from pyquery import PyQuery as pq
import xlwt
import xlrd
import re
import time
wbk = xlwt.Workbook()
sheet_1 = wbk.add_sheet('sheet 1')
sheet_2 = wbk.add_sheet('sheet 2')
sheet_1.write(0, 0, '车次')
sheet_1.write(0, 1, '站次')
sheet_1.write(0, 2, '站名')
sheet_1.write(0, 3, '到达时间')
sheet_1.write(0, 4, '开车时间')
sheet_1.write(0, 5, '运行时间')
sheet_1.write(0, 6, '里程')
sheet_1.write(0, 7, '硬卧 上/中/下')
sheet_1.write(0, 8, '软卧 上/下')
sheet_1.write(0, 9, '二等/一等软座')

data = xlrd.open_workbook('./表.xlsx')
table = data.sheet_by_name(u'需求2所有车次')   #通过名称获取
row_num = table.nrows
pattern_url = 'http://search.huochepiao.com/chaxun/resultc.asp?txtCheci=D2&cc.x=0&cc.y=0'
header = {'user-agent': 'Mozilla/5.0'}


sheet_1_row_index = 1

for i in range(8779, row_num):
    # time.sleep(0.01)
    ctype = table.cell(i, 0).ctype
    value = table.cell(i, 0).value
    if ctype == 2:
        value = str(int(value))
    print(value)
    new_url = re.sub(r'D2', value, pattern_url, count=1, flags=0)
    print(new_url)
    main_html = requests.get(new_url, headers = header)
    main_html.encoding = 'gb2312'
    main_html = main_html.text
    doc = pq(main_html)
    trs = doc('tr')
    for tr in trs:
        # print('tr')
        # print(len(tr))
        if tr[0].text_content() == '车次':
            continue

        # 车次 站次 站名 到达时间 开车时间         运行时间      里程 硬卧上/中/下  软卧上/下 二等/一等座
        # 车次 站次 站名	到达时间	开车时间	停留时间 运行时间	 天数 里程 硬卧上/中/下  软卧上/下 二等/一等座  （D C G）
        # 车次 站次 站名	到达时间	开车时间	停留时间 运行时间	 天数 里程 硬座         软座价 	 硬卧上/中/下  软卧上/下
        if len(tr) == 12:
            for i in range(5):
                sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
            sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
            sheet_1.write(sheet_1_row_index, 6, tr[8].text_content())
            sheet_1.write(sheet_1_row_index, 7, tr[9].text_content())
            sheet_1.write(sheet_1_row_index, 8, tr[10].text_content())
            sheet_1.write(sheet_1_row_index, 9, tr[11].text_content())
            sheet_1_row_index += 1
        elif len(tr) == 13:
            for i in range(5):
                sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
            sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
            sheet_1.write(sheet_1_row_index, 6, tr[8].text_content())
            sheet_1.write(sheet_1_row_index, 7, tr[11].text_content())
            sheet_1.write(sheet_1_row_index, 8, tr[12].text_content())
            sheet_1.write(sheet_1_row_index, 9, tr[9].text_content())
            sheet_1_row_index += 1
        else:
            continue
        print(len(tr))
    wbk.save('./test2.xls')

