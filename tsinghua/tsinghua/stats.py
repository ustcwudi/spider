import xlrd
import json
import os
import urllib.request
from xlutils.copy import copy


def stat(file):
    readbook = xlrd.open_workbook(file)
    writebook = copy(readbook)
    if len(readbook.sheets()) == 2:
        writebook.add_sheet('stats')
    writesheet = writebook.get_sheet(2)
    writesheet.write(0, 0, 'url')
    writesheet.write(0, 1, 'view1')
    writesheet.write(0, 2, 'view2')
    sheet = readbook.sheet_by_index(0)
    list = []
    id_list = []
    nrows = sheet.nrows
    for row in range(1, nrows):
        list.append(sheet.cell(row, 0).value)
        id_list.append(sheet.cell(row, 7).value.replace('/', '_'))
    i = 0
    for url in list:
        print(url)
        i += 1
        if not os.path.exists('files/stats/'+id_list[i-1]+".json"):
            opener = urllib.request.build_opener()
            opener.addheaders = [
                ('User-agent', 'Opera/9.80 (Android 2.3.4; Linux; Opera Mobi/build-1107180945; U; en-GB) Presto/2.8.149 Version/11.10')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(
                url + '/stats', 'files/stats/'+id_list[i-1]+".json")
        with open('files/stats/'+id_list[i-1]+".json", "r", encoding='utf-8') as f:
            dict = json.load(f)
            # 0 url
            writesheet.write(i, 0, url)
            # 1 view
            content = ''
            for value in dict['chart']['elements'][0]['values']:
                content += str(value['value'])+'\n'
            writesheet.write(i, 1, content)
            # 2 view
            content = ''
            for value in dict['chart']['elements'][1]['values']:
                content += str(value['value'])+'\n'
            writesheet.write(i, 2, content)
    writebook.save(file)


for num in range(1, 2):
    print(num)
    stat('files/xls/2071-1050-13-'+str(num)+'.xls')
