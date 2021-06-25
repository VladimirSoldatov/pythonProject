import os
import re
from datetime import datetime
from datetime import timedelta
import openpyxl

import pandas as pd


def max_min(html_text):
    pd_html = pd.DataFrame({'index': html.index, 'Name': html.values})
    a = pd_html.drop('index', axis=1)

    for i in range(a.shape[0]):
        if type(a.Name[i]) == str:
            for y in range(10):
                if ' ' + str(y) + ':' in a.Name[i]:
                    a.Name[i].replace(' ' + str(y) + ':', ' 0' + str(y) + ':')
                else:
                    continue
        if type(a.Name[i]) == str:

            if '0000' in a.Name[i]:
                a.Name[i] = pd.NA
            elif '[' in a.Name[i]:
                a.Name[i] = a.Name[i].replace(';', '')
                a.Name[i] = datetime.strptime(a.Name[i].replace('[', '').replace(']', ''), '%d.%m.%Y %H:%M:%S:%f')
            else:
                a.Name[i] = pd.NA
        else:
            a.Name[i] = pd.NA
    a = a.dropna()
    print(min(a.Name).strftime("%d.%m.%Y %H:%M:%S:%f"))
    print(max(a.Name).strftime("%d.%m.%Y %H:%M:%S:%f"))
    # print(begin.strftime("%d.%m.%Y %H:%M:%S:%f"))
    # print(end.strftime("%d.%m.%Y %H:%M:%S:%f"))
    return min(a.Name), max(a.Name)


def read_xml(html_text, sheet, begin, end):
    a = html_text
    p = 0
    p1 = 0
    p2 = 0
    p3 = []
    for i in range(a.shape[0]):
        if type(a.Name[i]) == str:
            if '0000' in a.Name[i]:
                b = a.Name[i].split("  ")
                if int(b[0]) == 10:
                    a.Name[i - 1] += b[2].strip()
                    a.Name[i] = pd.NA
                elif int(b[0]) == 20:
                    a.Name[i - 2] += b[2]
                    a.Name[i] = pd.NA
                else:
                    a.Name[i] = b[2]

                if b[2].split(" ")[0:2] == ['00', '04'] and b[1].split(" ")[5:7] == ['06', '01']:
                    p += 1
                if b[2].split(" ")[0:2] == ['00', '60'] and b[1].split(" ")[5:7] == ['06', '01']:
                    p1 += 1
                if b[2].split(" ")[0:2] == ['11', '08'] and b[1].split(" ")[5:7] == ['06', '01']:
                    p2 += 1
                if '01' in b[1]:
                    aa = [' ']

                    p3.append(b[1:3])
            elif '[' in a.Name[i]:
                a.Name[i] = datetime.strptime(a.Name[i].replace('[', '').replace(']', ''), '%d.%m.%Y %H:%M:%S:%f')

    print('Первый этап закончен')
    x = 0
    if sheet == 'Спутник':
        delta = timedelta(hours=1)
        begin += delta
        end += delta
    count = 0
    datetime_count = 0
    for i in range(a.shape[0]):
        if type(a.Name[i]) == datetime:
            if a.Name[i] < begin or a.Name[i] > end:
                a.Name[i] = pd.NA
                y = 1
                if i + y + 1 < len(a.Name):
                    while type(a.Name[i + y]) != datetime:
                        a.Name[i + y] = pd.NA
                        if i + y + 1 < len(a.Name):
                            y += 1
                        else:
                            break
                else:
                    print('Список закончен')
                    break

            else:
                datetime_count += 1
                # print('Дата ' + str(datetime_count) + ' добавлена')
        elif type(a.Name[i]) == str:
            a.Name[i] = len(str(a.Name[i]).split(' '))
            x += a.Name[i]
            count += 1
            # print('Данные ' + str(count) + ' добавлены добавлены')
        # print('Обработано записей ' + str(i))
    print('Завершен расчет')
    a = a.dropna(how='all')
    a = a.reindex()
    print('Переиндексация завершена')
    print(sheet)
    print(str(x) + ' байт')
    print(str(count) + ' пакетов')
    print(str(datetime_count) + ' дат')
    print(a.shape)
    print(begin.strftime("%d.%m.%Y %H:%M:%S:%f"))
    print(end.strftime("%d.%m.%Y %H:%M:%S:%f"))
    a.to_excel(sheet + ".xlsx")
    print("Перепад даления 2 -\t\t\t", p, " раз запрошен")
    print("Загазованность служебная -\t", p1, " раз запрошен")
    print("GSM канал -\t\t\t\t\t", p2, " раз запрошен")

    for item in p3:
        print(item)


def html_reader(path_text):
    if 'csv' in path_text:
        html_text = pd.read_csv(path_text)
    else:
        html_text = pd.read_table(path_text)
    html_text.columns = ['Name']

    if 'поток' in path_text:
        for i in range(html_text.shape[0]):
            p = re.compile(r'<.*?>')
            html_text.Name[i] = p.sub('', html_text.Name[i])

    html_text = html_text.Name[html_text.Name.str.contains('00000') | html_text.Name.str.contains(']')]
    print('html read')

    return html_text


def read_xml_free(html_text, sheet):
    a = html_text
    for i in range(a.shape[0]):
        if type(a.Name[i]) == str:
            if '0000' in a.Name[i]:
                b = a.Name[i].split("  ")
                if int(b[0]) == 10:
                    a.Name[i - 1] += b[2].strip()
                    a.Name[i] = pd.NA
                elif int(b[0]) == 20:
                    a.Name[i - 2] += b[2]
                    a.Name[i] = pd.NA
                else:
                    a.Name[i] = b[2]

    print('Первый этап закончен')
    x = 0

    count = 0
    datetime_count = 0
    for i in range(a.shape[0]):

        # print('Дата ' + str(datetime_count) + ' добавлена')
        if type(a.Name[i]) == str:
            a.Name[i] = len(str(a.Name[i]).split(' '))
            x += a.Name[i]
            count += 1
            # print('Данные ' + str(count) + ' добавлены добавлены')
        # print('Обработано записей ' + str(i))
    print('Завершен расчет')
    a = a.dropna(how='all')
    a = a.reindex()
    print('Переиндексация завершена')
    print(sheet)
    print(str(x) + ' байт')
    print(str(count) + ' пакетов')
    print(str(datetime_count) + ' дат')
    print(a.shape)
    # print(begin.strftime("%d.%m.%Y %H:%M:%S:%f"))
    # print(end.strftime("%d.%m.%Y %H:%M:%S:%f"))
    a.to_excel(sheet + ".xlsx")


# read_xml(pathXls, 'Спутник', begin_time, end_time)
# read_xml(pathXls, 'GPRS', begin_time, end_time)
# read_xml(pathXls, 'Лист2')

# begin_time = datetime.strptime('19.04.2021 10:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('20.04.2021 09:35:16:000', '%d.%m.%Y %H:%M:%S:%f')
# begin_time = datetime.strptime('20.04.2021 10:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('21.04.2021 04:30:00:000', '%d.%m.%Y %H:%M:%S:%f')
path = 'C:\\Users\\СолдатовВВ\\Desktop\Logs\\'
name_file = 'report_10_06_2021_08_30.html'

# name_file = 'Logs2\\Отчёт о потоках пакетов27.html'
path += name_file
if 'report' in path:
    type_name = 'GPRS'
else:
    type_name = 'Спутник'
html = html_reader(path)
html_clone = html.copy()
pd_html = pd.DataFrame({'index': html_clone.index, 'Name': html_clone.values})
html_clone = pd_html.drop('index', axis=1)
begin_time = datetime.strptime('10.06.2021 16:59:03:361', '%d.%m.%Y %H:%M:%S:%f')
end_time = datetime.strptime('10.06.2021 20:43:57:579', '%d.%m.%Y %H:%M:%S:%f')
# print(html_clone)
# max_min(html_clone)
# print(html_clone)
read_xml(html_clone, type_name, begin_time, end_time)
# read_xml_free(html_clone, type_name)
