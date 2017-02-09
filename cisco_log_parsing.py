#!/usr/bin/python3.5
# encoding=utf8

import xlwt
import os
from datetime import datetime

logfile='/home/opan/workspace/cisco_log_parsing/src/syslog'
namefile='/home/opan/workspace/cisco_log_parsing/src/names.txt'

#Шрифты
font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 2
font0.bold = True
#Стиль заголовков
style0 = xlwt.XFStyle()
style0.font = font0
#Стиль текста
style1 = xlwt.XFStyle()
style1.num_format_str = 'D-MMM-YY'
#Создание листа
wb = xlwt.Workbook()
ws = wb.add_sheet('По входам')
#Заголовок таблицы
ws.write(0, 0, 'Отчет об использовании VPN', style0)
ws.write(1, 0, 'На', style0)
ws.write(1, 1, datetime.now(), style1)
ws.write(2, 0, 'ФИО', style0)
ws.write(2, 1, 'Месяц', style0)
ws.write(2, 2, 'День', style0)
ws.write(2, 3, 'Вход', style0)
ws.write(2, 4, 'Выход', style0)
ws.write(2, 5, 'Разница', style0)
#Счетчик строк для данных
cnt=2

with open(logfile, 'r') as logfile:                         #Открываем лог
    for line in logfile.readlines():                        #Читаем строки
        if '113008' in line:                                #По коду события заходим в if
            cnt=cnt+1                                       #Добавляем счетчик строки
            string = line.split()                           #Строку переводим в формат списка
            #Цикл для замены логинов на полные имена по списку из файла
            for linelist in open(namefile):                 #Открываем файл со списком
                if string[12] in linelist:                  #Если попался логин входим в if
                    fio = linelist.split('=')               #Переводим в формат списка строку по разделителю =
                                                            #и присваеваем fio
                    ws.write(cnt, 0, fio[1], style1)        #Пишем в нужные колонки таблицы
                    ws.write(cnt, 1, string[0], style1)
                    ws.write(cnt, 2, string[1], style1)
                    ws.write(cnt, 3, string[2], style1)

        if '113019' in line:
            cnt=cnt+1
            string = line.split()

            for linelist in open(namefile):
                usrclear = string[10]                       #Приблуда для удаления запятой -
                usrclear = usrclear[0:-1]                   #в этом событии циска формирует логин с запятой
                if usrclear in linelist:
                    fio = linelist.split('=')
                    ws.write(cnt, 0, fio[1], style1)
                    ws.write(cnt, 1, string[0], style1)
                    ws.write(cnt, 2, string[1], style1)
                    ws.write(cnt, 4, string[2], style1)
                    ws.write(cnt, 5, string[20], style1)

wb.save('example.xls')                                      #Закрываем файл
