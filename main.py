import os
import xlrd
from tkinter import *
from tkinter import ttk


def open_vnc():
    os.startfile('O:\!VNC')


def data_crawler():
    response = []
    request = input_data.get().lower()
    excel_data_file = xlrd.open_workbook('hardData.xlsx')
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(1, row_number):
            if request in str(sheet.row(row)[0]).replace('text:','').replace("'", '').replace('ё', 'е').lower() or \
            request in str(sheet.row(row)[1]).replace('text:','').replace("'", '').replace('ё', 'е').lower() or \
            request in str(sheet.row(row)[3]).replace('number:','').replace(".0", '') or \
            request in str(sheet.row(row)[4]).replace('text:','').replace("'", '').replace('ё', 'е').lower():
                response.append(str(sheet.row(row)[6]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[5]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[0]).replace('text:','').replace("'", ''))
                response.append(str(sheet.row(row)[1]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[2]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[4]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[7]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[3]).replace('number:','').replace(".0", '').replace("'",'').replace('empty:', '<нет данных>'))
                response.append('')
        if response != []:
            output_data_lbl03.config(text='Отделение: ')
            output_data_lbl04.config(text=response[0])
            output_data_lbl13.config(text='Должность: ')
            output_data_lbl14.config(text=response[1])
            output_data_lbl23.config(text='ФИО: ')
            output_data_lbl24.config(text=response[2])
            output_data_lbl33.config(text='Логин: ')
            output_data_lbl34.config(text=response[3])
            output_data_lbl35.config(text='ID: ')
            output_data_lbl36.config(text=response[4])
            output_data_lbl43.config(text='Кабинет: ')
            output_data_lbl44.config(text=response[6])
            output_data_lbl45.config(text='Телефон: ')
            output_data_lbl46.config(text=response[7])
            output_data_lbl53.config(text='Имя ПК: ')
            output_data_lbl54.config(text=response[5])
            vnc_but = ttk.Button(text='VNC', command=open_vnc, style='TButton').grid(row=5, column=5, columnspan=2, sticky=W, padx=7, pady=(5, 20))
        else:
            output_data_lbl23.config(text='Совпадений не найдено')


root = Tk()
root.title('Crawler v3.0')
root.geometry('+300+200')
root.resizable(False, False)

ttk.Label(text='Поиск: ', font=('Lucida Console', 10)).grid(row=0, column=0, rowspan=6, sticky=E, padx=(20, 5), pady=20)
input_data = ttk.Entry(width=15)
input_data.grid(row=0, column=1, rowspan=6, sticky=W, pady=10)

ttk.Label(text='Результат: ', font=('Lucida Console', 10)).grid(row=0, column=2, rowspan=6, sticky=W, padx=(20, 5), pady=20)

output_data_lbl03 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl03.grid(row=0, column=3, sticky=E, padx=5, pady=(20, 5))

output_data_lbl04 = ttk.Label(font=('Lucida Console', 10), anchor=W)
output_data_lbl04.grid(row=0, column=4, sticky=W, padx=5, pady=(20, 5), columnspan=3)

output_data_lbl13 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl13.grid(row=1, column=3, sticky=E, padx=5, pady=5)

output_data_lbl14 = ttk.Label(font=('Lucida Console', 10))
output_data_lbl14.grid(row=1, column=4, sticky=W, padx=5, pady=5, columnspan=3)

output_data_lbl23 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl23.grid(row=2, column=3, sticky=E, padx=5, pady=5)

output_data_lbl24 = ttk.Label(font=('Lucida Console', 10))
output_data_lbl24.grid(row=2, column=4, sticky=W, padx=5, pady=5, columnspan=3)

output_data_lbl33 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl33.grid(row=3, column=3, sticky=E, padx=5, pady=5)

output_data_lbl34 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl34.grid(row=3, column=4, padx=5, pady=5)

output_data_lbl35 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl35.grid(row=3, column=5, sticky=E, padx=5, pady=5)

output_data_lbl36 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl36.grid(row=3, column=6, sticky=E, padx=5, pady=5)

output_data_lbl43 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl43.grid(row=4, column=3, sticky=E, padx=5, pady=5)

output_data_lbl44 = ttk.Label(width=15, font=('Lucida Console', 10))
output_data_lbl44.grid(row=4, column=4, sticky=E, padx=5, pady=5)

output_data_lbl45 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl45.grid(row=4, column=5, sticky=E, padx=5, pady=5)

output_data_lbl46 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl46.grid(row=4, column=6, sticky=E, padx=5, pady=5)

output_data_lbl53 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl53.grid(row=5, column=3, sticky=E, padx=5, pady=(5, 20))

output_data_lbl54 = ttk.Label(font=('Lucida Console', 10))
output_data_lbl54.grid(row=5, column=4, sticky=W, padx=5, pady=(5, 20))

button_style = ttk.Style(root)
button_style.configure('TButton', font=('Lucida Console', 10))

show_but = ttk.Button(text='Вывод', command=data_crawler, style='TButton').grid(row=0, column=7, rowspan=6, ipady=10, ipadx=10, pady=60, padx=20)

input_data.focus()

root.mainloop()
