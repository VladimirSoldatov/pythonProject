from datetime import datetime
from datetime import timedelta

import pandas as pd


def max_min(path, sheet):
    a = pd.read_excel(path, sheet_name=sheet, header=None)
    if len(a.columns) == 2:
        a.columns = ['Name', 'Value']
        a.drop(axis=1, columns='Value', inplace=True)
    else:
        a.columns = ['Name']
    for i in range(a.shape[0]):
        if type(a.Name[i]) == str:
            for y in range(10):
                if ' ' + str(y) + ':' in a.Name[i]:
                    a.Name[i].replace(' ' + str(y) + ':', ' 0' + str(y) + ':')
            if '0000' in a.Name[i]:
                a.Name[i] = pd.NA
            elif '[' in a.Name[i]:
                a.Name[i] = datetime.strptime(a.Name[i].replace('[', '').replace(']', ''), '%d.%m.%Y %H:%M:%S:%f')
            else:
                a.Name[i] = pd.NA
        else:
            a.Name[i] = pd.NA
    a = a.dropna()
    print(min(a.Name))
    print(max(a.Name))


def read_xml(path, sheet, begin, end):
    count = 0
    a = pd.read_excel(path, sheet_name=sheet, header=None)
    count_good = 0
    count_bad = 0
    count_fail = 0
    if len(a.columns) == 2:
        a.columns = ['Name', 'Value']
        a.drop(axis=1, columns='Value', inplace=True)
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
                    a.Name[i - 1] += b[1] + b[2]
                    a.Name[i] = pd.NA
                elif int(b[0]) == 20:
                    a.Name[i - 2] += b[1]
                    a.Name[i] = pd.NA
                else:
                    a.Name[i] = b[2]

            elif '[' in a.Name[i]:
                a.Name[i] = datetime.strptime(a.Name[i].replace('[', '').replace(']', ''), '%d.%m.%Y %H:%M:%S:%f')
            else:
                a.Name[i] = pd.NA
        else:
            a.Name[i] = pd.NA
    print('Первый этап закончен')
    x = 0
    if sheet == 'GPRS':
        delta = timedelta(hours=1)
        begin -= delta
        end -= delta

    datetime_count = 0
    for i in range(a.shape[0]):
        if type(a.Name[i]) == datetime:
            if a.Name[i] < begin or a.Name[i] > end:
                a.Name[i] = pd.NA
            datetime_count += 1
            count += 1
        if type(a.Name[i]) == str:
            a.Name[i] = len(str(a.Name[i]).split(' ')) * 2
            if a.Name[i] % 8 == 0 or a.Name[i] % 6 == 0 or a.Name[i] % 10 == 0:
                count_good += 1
            elif 10 < a.Name[i] <= 52:
                count_bad += 1
            else:
                count_fail += 1
            x += a.Name[i]

    a = a.dropna(how='any')
    a.reindex()

    print(sheet)
    print(str(x) + ' байт')
    print(str(count) + ' пакетов')
    print(str(datetime_count) + ' дат')
    print(begin)
    print(end)
    a.to_excel(sheet + ".xlsx")


my_sheet = 'GPRS'
# begin_time = datetime.strptime('24.04.2021 09:47:57:248', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('24.04.2021 21:37:09:429', '%d.%m.%Y %H:%M:%S:%f')
# begin_time: datetime = datetime.strptime('20.04.2021 09:42:21.741', '%d.%m.%Y %H:%M:%S.%f')
# end_time = datetime.strptime('21.04.2021  11:27:16.293', '%d.%m.%Y %H:%M:%S.%f')
pathXls = 'C:\\Users\\СолдатовВВ\\Desktop\\Logs\\Отчеты\\Лист XLSX 19-20.04.2021.xlsx'
# pathXls = 'C:\\Users\\СолдатовВВ\\Desktop\\Logs\\Отчеты\\Лист XLSX 2021-04-25.xlsx'
# max_min(pathXls, my_sheet)

# read_xml(pathXls, 'GPRS', begin_time, end_time)
# read_xml(pathXls, 'Лист2')

begin_time = datetime.strptime('19.04.2021 10:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
end_time = datetime.strptime('20.04.2021 09:35:16:000', '%d.%m.%Y %H:%M:%S:%f')
# begin_time = datetime.strptime('20.04.2021 11:31:30:000', '%d.%m.%Y %H:%M:%S:%f')
# end_time = datetime.strptime('21.04.2021 05:30:00:000', '%d.%m.%Y %H:%M:%S:%f')
# read_xml(pathXls, my_sheet, begin_time, end_time)
print((end_time - begin_time))
