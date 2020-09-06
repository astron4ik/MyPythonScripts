
__version__ = "$Version: 1.1 $"
# $Source$

'''Для работы скрипта необходимо изменить переменную path - путь до эксельвского файла.

Переменную z - ip адрес сервера, логин и пароль от сервера Zabbix.
Переменную path_logfile.
Путь для копирования в shutil.copy.'''

from pyzabbix import ZabbixAPI
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string
import datetime
import os
import shutil

def log(text):
    '''ЗАписывает ошибку в лог файл.

    '''
    file = open(path_logfile, "a")
    file.write(text)
    file.close()


def data():
    '''Определение координат последней записи.

    Цикл проверяет столбец A, находит номер следующей пустой строки и записывает туда текущую дату.'''
    for i in range(1, 1000):
        y = sheet_ranges.cell(row=i, column=1).value
        if i > 7 and y == None:
            coord = sheet_ranges.cell(row=i, column=1).row
            coord = str(coord)
            coord_last = "A" +  coord
            sheet_ranges[coord_last] = date # Определили номер строки
            break
    return coord

def total_pages_and_serial(d, number):
    '''Определяет серийный номер и число напечатанных страниц.

    Получает и обрабатывает данные из Zabbix.'''
    items = z.item.get(hostids=d, output=['itemid','name']) # Это id принтера
    d = items[number]
    d = d.get('itemid')
    a = z.item.get(itemids=d, output=['lastvalue']) #d это id total pages
    b = a[0]
    b = b.get('lastvalue')
    return b

def printers(group):
    '''Получает и записывает значения serial и total pages для каждого принтера.

    Записывает в excel.'''
    hosts = z.host.get(groupids=group, output=['hostid','name']) #Группа Принтеры Офис
    for host in hosts:
        name = host['name'] # Имя узла принтера
        #print(host['hostid'],host['name'])
        host = host.get('hostid')
        pages = int(total_pages_and_serial(host, 5)) # Узнаем Кол-во Страниц, 5 это номер позиции total pages
        serial = total_pages_and_serial(host, 1) # Узнаем Серийник, 1 это номер позиции serial
        serial_zabbix.append(serial)
        for cellObj in sheet_ranges['A2':'CR2']:
            for cell in cellObj:
                if cell.value == serial:
                    column = cell.column
                    coord_pages = column + coord
                    coord_name = column + "4" #Координаты строки места принтера
                    coord_ip = column + "3"
                    sheet_ranges[coord_name] = name
                    sheet_ranges[coord_pages] = pages
                    sheet_ranges[coord_ip] = ip(name)
                    #print(serial, ' - ', pages)
                    dict_excel[cell.value] = column
                elif cell.value != serial: #Создает словарь из серийников, для сравнения
                    column1 = cell.column
                    dict_excel[cell.value] = column1

def raschet(cell, row_2):
    '''Расчет для столбца прирост.

    Создает переменную с формулой расчета прироста total_pages за неделю'''
    column_1 = column_index_from_string(coord_column1)
    column_1 = int(column_1) - 1
    column_2 = get_column_letter(column_1)
    coordinate_last = column_2 + row_2
    coordinate_now = column_2 + str(row)
    letter = "=" + coordinate_now + "-" + coordinate_last
    sheet_ranges[coordinate] = letter
    summa_prom.append(cell.coordinate)
    return summa

def row_1(row):
    '''На 1 ячейку вверх.

    '''
    row2 = int(row) - 1
    row2 = str(row2)
    return row2

def ip(name):
    '''Определяет ip адрес.

    В имени файла последние 4 цифры ip адреса.'''
    ip = name[-4:-1]
    ip = "192.168.3." + ip
    return ip

path_logfile = "\\\\fs-srv-2\\Public\\SPb\\Отд_ИТ\\SYSADMIN\\Тобольцов_БО\\Принтеры\\Script\\logfile.txt" #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
date = datetime.datetime.today().strftime("%d.%m.%y") #Текущая дата

file = open(path_logfile, "a")
file.write("-----------------------------------------------------------")
file.write("\n\n")
file.write("Начало работы в ")
file.write(date)
file.write("\n\n")
file.close()

serial_zabbix = []
dict_excel = {}
path = "\\\\fs-srv-2\\Public\\SPb\\Отд_ИТ\\SYSADMIN\\Тобольцов_БО\\Принтеры\\Script\\Отчет_принтеры.xlsx" #Путь до эксельвского файла, формата например D:\\1\\2\\Отчет_принтеры.xlsx #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
try:
    wb = load_workbook(path) #Загружаем файл
except FileNotFoundError:
    log("Не возможно найти файл excel\n\n")
    raise
try:
    z = ZabbixAPI('http://192.168.40.5', user='tob', password='6c7w1b') # соединяемся с Zabbix #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
except OSError:
    log("Нет соединения с сервером Zabbix\n\n")
    raise
sheet_ranges = wb['Данные'] # Перешли в книгу "Данные"
group_office = 9
group_sklad = 23
coord = data()

printers(group_sklad)
printers(group_office)

coord3 = "A" + coord
coord4 = "CR" + coord
summa_prom = []
summa = []

for cellObj in sheet_ranges[coord3:coord4]:
    for cell in cellObj:
        coordinate = cell.coordinate
        row = cell.row
        row_2 = row_1(row)
        coord_column1 = cell.column
        coord_column = coord_column1 + "5" # 5 строка с заголовками в excel
        value_5 = sheet_ranges[coord_column].value #Чтение ячейки
        if cell.value == None:
            if value_5 == "Счетчик":
                coord6 = coord_column1 + row_2
                value_6 = sheet_ranges[coord6].value
                sheet_ranges[coordinate] = value_6 #Запись в ячейку
            elif value_5 == "Прирост":
                summa = raschet(cell, row_2)
            elif value_5 == "Сумма":
                for element in summa_prom:
                    summa.append(element)
                summa_string = '+'.join(summa)
                summa_string = "=" + summa_string
                sheet_ranges[coordinate] = summa_string
            elif value_5 == "Разница Сумм":
                raschet(cell, row_2)

for dict_keys in dict_excel.keys():
    '''Выставляет Резерв в ячейках.

    Которых нету в zabbix.'''
    if dict_keys not in serial_zabbix and dict_keys != "PAU4609237": #Этот серийник не определяется в Zabbix
        coord7 = dict_excel[dict_keys] + "3"
        coord8 = dict_excel[dict_keys] + "4"
        sheet_ranges[coord7] = "Резерв!"
        sheet_ranges[coord8] = "Резерв!"
try:
    wb.save(path)
except PermissionError:
    log("Не возможно сохранить файл excel, у кого-то он открыт! Сохранение в промежуточный файл!\n\n")
    raise
try:
    shutil.copy(path, '\\\\fs-srv-2\\Public\\SPb\\Отд_ИТ\\SYSADMIN\\Тобольцов_БО\\Принтеры\\Отчет_принтеры.xlsx') #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
except PermissionError:
    log("ПЗДЦ!!!! Файл не скопировался, видимо потому что он открыт у кого-то! Копирование в общую папку!\n\n")
    raise
