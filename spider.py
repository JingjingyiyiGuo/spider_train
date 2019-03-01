#!/usr/bin/python3
# coding: utf-8
import requests
from pyquery import PyQuery as pq
import xlwt
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

main_url = 'http://www.huochepiao.com/tiaozheng/'
main_html = requests.get(main_url)
main_html.encoding = 'gb2312'
main_html = main_html.text
main_doc = pq(main_html)
main_trs = main_doc('tr')
print(len(main_trs))  # 2560 / 2557; 顶端有三行其他的

sheet_1_row_index = 1
for index, main_tr in enumerate(list(main_trs.items())):
    main_tds = main_tr('td')
    if len(main_tds) < 7:
        continue

    train_type = main_tds[0].text_content()
    print(index)
    print(main_tds)
    print(train_type)

    if main_tds[4].text_content() == '调整变动':
        url = main_tds[5].find('a').items()[0][1]
    elif main_tds[4].text_content() == '新增开行':
        url = 'http://www.huochepiao.com/' + main_tds[6].find('a').items()[0][1]
    else:
        continue

    html = requests.get(url)
    html.encoding = 'gb2312'
    html = html.text
    doc = pq(html)
    trs = doc('tr')
    for tr in trs:
        # print(train_type)
        # print(len(tr))
        if tr[0].text_content() == '车次':
            continue

        if train_type[0] == 'G':  # 11
            # 车次	站次	途经站	到达时间	开车时间	停留时间	运行时间	里程	天数	二等座	一等座
            if len(tr) == 11:
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, '0/0/0')
                sheet_1.write(sheet_1_row_index, 8, '0/0')
                sheet_1.write(sheet_1_row_index, 9, tr[9].text_content() + '/' + tr[10].text_content())
                sheet_1_row_index += 1
            else:
                continue
            print(len(tr))
        elif train_type[0] == 'D':  # 10
            # 车次	站次	站名	到达时间	开车时间	运行时间	里程	硬卧上/中/下	软卧上/下	二等/一等软座  # 按照这个来爬取的
            # 车次	站次	途经站	到达时间	开车时间	停留时间	运行时间	里程	特等/商务座	软卧上/下	二等/一等座
            if len(tr) == 10:
                for i in range(10):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1_row_index += 1
            elif len(tr) == 11:
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, '0/0/0')
                sheet_1.write(sheet_1_row_index, 8, tr[9].text_content())
                sheet_1.write(sheet_1_row_index, 9, tr[10].text_content())
                sheet_1_row_index += 1
            else:
                continue
            print(len(tr))
        elif train_type[0] == 'K':  # 11 / 12
            # 车次	站次	站名	到达时间	开车时间	运行时间	里程	硬座	软座价	硬卧上/中/下	软卧上/下
            # 车次	站次	途经站	到达时间	开车时间	停留时间	运行时间	里程	硬座	软座价	硬卧上/中/下	软卧上/下
            if len(tr) == 11:
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, tr[9].text_content())
                sheet_1.write(sheet_1_row_index, 8, tr[10].text_content())
                sheet_1.write(sheet_1_row_index, 9, '0/0')
                sheet_1_row_index += 1
            if len(tr) == 12:
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, tr[10].text_content())
                sheet_1.write(sheet_1_row_index, 8, tr[11].text_content())
                sheet_1.write(sheet_1_row_index, 9, '0/0')
                sheet_1_row_index += 1
            print(len(tr))
        elif train_type[0] == 'C':
            # 车次	站次	站名	到达时间	开车时间	运行时间	里程	硬卧上/中/下	软卧上/下	二等/一等软座  # 按照这个来爬取的
            # 车次	站次	途经站	到达时间	开车时间	停留时间	运行时间	里程	硬卧上/中/下	二等/一等座
            if len(tr) == 10 and tr[5].text_content() == '运行时间':
                for i in range(10):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1_row_index += 1
            elif len(tr) == 10 and tr[5].text_content() == '停留时间':
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, tr[8].text_content())
                sheet_1.write(sheet_1_row_index, 8, '0/0')
                sheet_1.write(sheet_1_row_index, 9, tr[9].text_content())
                sheet_1_row_index += 1
            print(len(tr))
        else:
            # 车次	站次	站名	到达时间	开车时间	运行时间	里程	硬卧上/中/下	软卧上/下	二等/一等软座  # 按照这个来爬取的
            # 车次	站次	站名	到达时间	开车时间	运行时间	里程	硬座	软座价	硬卧上/中/下	软卧上/下
            # 车次	站次	途经站	到达时间	开车时间	停留时间	运行时间	里程	硬座	软座价	硬卧上/中/下	软卧上/下
            if len(tr) == 11 and tr[5].text_content() == '运行时间':
                for i in range(7):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 7, tr[9].text_content())
                sheet_1.write(sheet_1_row_index, 8, tr[10].text_content())
                sheet_1.write(sheet_1_row_index, 9, '0/0')
                sheet_1_row_index += 1
            if len(tr) == 12 and tr[5].text_content() == '停留时间':
                for i in range(5):
                    sheet_1.write(sheet_1_row_index, i, tr[i].text_content())
                sheet_1.write(sheet_1_row_index, 5, tr[6].text_content())
                sheet_1.write(sheet_1_row_index, 6, tr[7].text_content())
                sheet_1.write(sheet_1_row_index, 7, tr[10].text_content())
                sheet_1.write(sheet_1_row_index, 8, tr[11].text_content())
                sheet_1.write(sheet_1_row_index, 9, '0/0')
                sheet_1_row_index += 1
            print(len(tr))

wbk.save('./tmp.xls')  # 保存
