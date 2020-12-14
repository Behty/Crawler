import os
import xlrd
import fnmatch
import subprocess
from tkinter import *
from tkinter import ttk
from functools import partial

VNC_PATH = 'O:\!VNC'
BD_PATH = 'hardData.xlsx'


def open_vnc(path: str) -> None:
    '''Функция запускает сеанс VNC на компьютер пользователя'''

    target = output_data_lbl54.cget('text')
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, f'{target}*'):
                os.startfile(os.path.join(root, name))


def open_rdp() -> None:
    '''Функция запускает сеанс RDP на компьютер пользователя'''
    
    target = output_data_lbl54.cget('text')
    subprocess.run(['mstsc.exe', f'/v:{target}'])


def data_crawler(path: str) -> None:
    '''Функция отрисовывает и заполняет данными окно с результатами поиска'''
    
    #Отрисовка окна с результатами поиска
    res_lbl.grid(row=0, column=2, rowspan=6, sticky=W, padx=(20, 5), pady=20)
    output_data_lbl03.grid(row=0, column=3, sticky=E, padx=5, pady=(20, 5))
    output_data_lbl04.grid(row=0, column=4, sticky=W, padx=5, pady=(20, 5), columnspan=3)
    output_data_lbl13.grid(row=1, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl14.grid(row=1, column=4, sticky=W, padx=5, pady=5, columnspan=3)
    output_data_lbl23.grid(row=2, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl24.grid(row=2, column=4, sticky=W, padx=5, pady=5, columnspan=3)
    output_data_lbl33.grid(row=3, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl34.grid(row=3, column=4, padx=5, pady=5)
    output_data_lbl35.grid(row=3, column=5, sticky=E, padx=5, pady=5)
    output_data_lbl36.grid(row=3, column=6, sticky=E, padx=5, pady=5)
    output_data_lbl43.grid(row=4, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl44.grid(row=4, column=4, sticky=E, padx=5, pady=5)
    output_data_lbl45.grid(row=4, column=5, sticky=E, padx=5, pady=5)
    output_data_lbl46.grid(row=4, column=6, sticky=E, padx=5, pady=5)
    output_data_lbl53.grid(row=5, column=3, sticky=E, padx=5, pady=(5, 20))
    output_data_lbl54.grid(row=5, column=4, sticky=W, padx=5, pady=(5, 20))
    
    #Поиск информации в базе данных
    response = []
    request = input_data.get().lower()
    try:
        excel_data_file = xlrd.open_workbook(path)
    except Exception:
        output_data_lbl03.grid_remove()
        output_data_lbl04.grid_remove()
        output_data_lbl13.grid_remove()
        output_data_lbl14.grid_remove()
        output_data_lbl23.grid_remove()
        output_data_lbl24.grid_remove()
        output_data_lbl33.grid_remove()
        output_data_lbl34.grid_remove()
        output_data_lbl35.grid_remove()
        output_data_lbl36.grid_remove()
        output_data_lbl43.grid_remove()
        output_data_lbl44.grid_remove()
        output_data_lbl45.grid_remove()
        output_data_lbl46.grid_remove()
        output_data_lbl53.grid_remove()
        output_data_lbl54.grid_remove()
        vnc_but.grid_remove()
        rdp_but.grid_remove()
        nomatches_lbl.grid(row=0, column=3, rowspan=6, columnspan=4, padx=10)
        nomatches_lbl.config(width=34, text='Не могу открыть файл базы данных!')
        return None
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(1, row_number):
            if all([(request != ''), (request != ' '), \
            (request in str(sheet.row(row)[0]).replace('text:','').replace("'", '').replace('ё', 'е').lower() or \
            request in str(sheet.row(row)[1]).replace('text:','').replace("'", '').replace('ё', 'е').lower() or \
            request in str(sheet.row(row)[3]).replace('number:','').replace(".0", '') or \
            request in str(sheet.row(row)[4]).replace('text:','').replace("'", '').replace('ё', 'е').lower())]):
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
            
            #Заполнение окна с результатами поиска данными
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
            vnc_but.grid(row=5, column=5, sticky=W, padx=7, pady=(5, 20))
            rdp_but.grid(row=5, column=6, sticky=W, padx=7, pady=(5, 20))
            nomatches_lbl.grid_remove()
        else:
            output_data_lbl03.grid_remove()
            output_data_lbl04.grid_remove()
            output_data_lbl13.grid_remove()
            output_data_lbl14.grid_remove()
            output_data_lbl23.grid_remove()
            output_data_lbl24.grid_remove()
            output_data_lbl33.grid_remove()
            output_data_lbl34.grid_remove()
            output_data_lbl35.grid_remove()
            output_data_lbl36.grid_remove()
            output_data_lbl43.grid_remove()
            output_data_lbl44.grid_remove()
            output_data_lbl45.grid_remove()
            output_data_lbl46.grid_remove()
            output_data_lbl53.grid_remove()
            output_data_lbl54.grid_remove()
            vnc_but.grid_remove()
            rdp_but.grid_remove()
            nomatches_lbl.grid(row=0, column=3, rowspan=6, padx=10)
            nomatches_lbl.config(text='Совпадений не найдено')


#Инициализация главного окна
root = Tk()
root.title('Crawler v6.0')
root.geometry('+300+200')
root.resizable(False, False)

#Задание стиля кнопкам
button_style = ttk.Style(root)
button_style.configure('TButton', font=('Lucida Console', 10))

#Отрисовка интерфейса
search_lbl = ttk.Label(text='Поиск: ', font=('Lucida Console', 10))
search_lbl.grid(row=0, column=0, rowspan=6, sticky=E, padx=(20, 5), pady=20)
res_lbl = ttk.Label(text='Результат: ', font=('Lucida Console', 10))
output_data_lbl03 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl04 = ttk.Label(font=('Lucida Console', 10), anchor=W)
output_data_lbl13 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl14 = ttk.Label(font=('Lucida Console', 10))
output_data_lbl23 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl24 = ttk.Label(font=('Lucida Console', 10))
output_data_lbl33 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl34 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl35 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl36 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl43 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl44 = ttk.Label(width=15, font=('Lucida Console', 10))
output_data_lbl45 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl46 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='red')
output_data_lbl53 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl54 = ttk.Label(font=('Lucida Console', 10))
nomatches_lbl = ttk.Label(width=22, font=('Lucida Console', 10, 'bold'))
input_data = ttk.Entry(width=15)
input_data.grid(row=0, column=1, rowspan=6, sticky=W, pady=10)
vnc_but = ttk.Button(text='VNC', command=partial(open_vnc, VNC_PATH), style='TButton')
rdp_but = ttk.Button(text='RDP', command=open_rdp, style='TButton')
show_but = ttk.Button(text='Вывод', command=partial(data_crawler, BD_PATH), style='TButton')
show_but.grid(row=0, column=7, rowspan=6, ipady=10, ipadx=10, pady=60, padx=20)

#Поместить курсор в окно поиска
input_data.focus()

root.mainloop()
