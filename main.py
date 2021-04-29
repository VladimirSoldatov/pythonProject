import os
from datetime import datetime
from datetime import timedelta
import xlwt
import pandas as pd


def read_xml(path, sheet, begin, end):
    a = pd.read_excel(path, sheet_name=sheet, header=None)
    if len(a.columns) == 2:
        a.columns = ['Name', 'Value']
        a.drop(axis=1, columns='Value')
    else:
        a.columns = ['Name']
    for i in range(a.shape[0]):
        if type(a.Name[i]) == str:
            for y in range(10):
                if ' ' + str(y) + ':' in a.Name[i]:
                    a.Name[i].replace(' ' + str(y) + ':', ' 0' + str(y) + ':')
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

            elif '[' in a.Name[i]:
                a.Name[i] = datetime.strptime(a.Name[i].replace('[', '').replace(']', ''), '%d.%m.%Y %H:%M:%S:%f')
        else:
            a.Name[i] = pd.NA
    print('Первый этап закончен')
    x = 0
    if sheet == 'GPRS':
        delta = timedelta(hours=1)
        begin -= delta
        end -= delta
    count = 0
    datetime_count = 0
    for i in range(a.shape[0]):
        if type(a.Name[i]) == datetime:
            if a.Name[i] < begin or a.Name[i] > end:
                a.Name[i] = pd.NA
                y = 1
                while type(a.Name[i + y]) != datetime:
                    a.Name[i + y] = pd.NA
                    if i + y + 1 < len(a.Name):
                        y += 1
                    else:
                        break
            else:
                datetime_count += 1
        if type(a.Name[i]) == str:
            a.Name[i] = len(str(a.Name[i]).split(' '))
            x += a.Name[i]
            count += 1

    a = a.dropna(how='any')
    a.reindex()
    a.to_excel(sheet + ".xlsx")
    print(sheet)
    print(str(x) + ' байт')
    print(str(count) + ' пакетов')
    print(str(datetime_count) + ' дат')
    print(a.shape)
    print(begin)
    print(end)


begin_time = datetime.strptime('23.04.2021 09:03:16:000', '%d.%m.%Y %H:%M:%S:%f')
end_time = datetime.strptime('24.04.2021 06:08:00:000', '%d.%m.%Y %H:%M:%S:%f')
pathXls = 'C:\\Users\\СолдатовВВ\\Desktop\\Logs\\Отчеты\\Лист XLSX 22-24.04.2021.xlsx'
#read_xml(pathXls, 'Спутник', begin_time, end_time)
read_xml(pathXls, 'GPRS', begin_time, end_time)
# read_xml(pathXls, 'Лист2')

# begin_time = datetime.strptime('19.04.2021 10:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('20.04.2021 09:35:16:000', '%d.%m.%Y %H:%M:%S:%f')
# begin_time = datetime.strptime('20.04.2021 10:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('21.04.2021 04:30:00:000', '%d.%m.%Y %H:%M:%S:%f')
