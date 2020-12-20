import os
import xlrd
import ctypes
import fnmatch
import pyperclip
import subprocess
from tkinter import *
from tkinter import ttk
from functools import partial

VNC_PATH = './VNC'
BD_PATH = './exampleData.xlsx'
ICON_PATH = './spider.ico'


def press(event) -> None:
    """Функция запускает поиск информации по нажатию клавиши."""

    data_crawler(BD_PATH)


def release(event) -> None:
    """Функция отслеживает изменение языка клавиатуры по отжатию клавиш."""

    get_language()


def click(event) -> None:
    """Функция отслеживает клики по меткам с текстом."""

    copy_to_clipboard(event.widget)


def copy_to_clipboard(lbl) -> None:
    """Функция копирует текст из метки в буфер обмена."""

    pyperclip.copy(lbl.cget('text'))
    lbl_text = lbl.cget('text')
    lbl['text'] = '<cкопировано>'
    root.after(1000, label_animation, lbl, lbl_text)


def label_animation(lbl_id, lbl_value: str) -> None:
    """Функция реализует визуальный эффект анимации при клике на метки."""

    lbl_id['text'] = lbl_value


def get_language() -> None:
    """Функция определяет текущий язык ввода текста."""

    lib = ctypes.windll.LoadLibrary('user32.dll')
    keylay = getattr(lib, 'GetKeyboardLayout')
    if hex(keylay(0)) == '0x4190419':
        language_but.config(text='RU')
    if hex(keylay(0)) == '0x4090409':
        language_but.config(text='EN')


def change_language() -> None:
    """Функция меняет язык ввода текста."""

    pass


def open_vnc(path: str) -> None:
    """Функция запускает сеанс VNC на компьютер пользователя."""

    target = output_data_lbl64.cget('text')
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, f'{target}*'):
                os.startfile(os.path.join(root, name))


def open_vnc_folder() -> None:
    """Функция открывает каталог с файлами сеансов VNC."""

    os.startfile(VNC_PATH)


def open_rdp() -> None:
    """Функция запускает сеанс RDP на компьютер пользователя."""

    target = output_data_lbl64.cget('text')
    subprocess.run(['mstsc.exe', f'/v:{target}'])


def data_crawler(path: str) -> None:
    """Функция отрисовывает и заполняет данными окно с результатами поиска."""
    
    # Отрисовка окна с результатами поиска
    res_lbl.grid(row=1, column=2, rowspan=6, sticky=W, padx=(20, 5), pady=20)
    output_data_lbl13.grid(row=1, column=3, sticky=E, padx=5, pady=(20, 5))
    output_data_lbl14.grid(row=1, column=4, sticky=W, padx=5, pady=(20, 5), columnspan=3)
    output_data_lbl23.grid(row=2, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl24.grid(row=2, column=4, sticky=W, padx=5, pady=5, columnspan=3)
    output_data_lbl33.grid(row=3, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl34.grid(row=3, column=4, sticky=W, padx=5, pady=5, columnspan=3)
    output_data_lbl43.grid(row=4, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl44.grid(row=4, column=4, padx=5, pady=5)
    output_data_lbl45.grid(row=4, column=5, sticky=E, padx=5, pady=5)
    output_data_lbl46.grid(row=4, column=6, sticky=E, padx=5, pady=5)
    output_data_lbl53.grid(row=5, column=3, sticky=E, padx=5, pady=5)
    output_data_lbl54.grid(row=5, column=4, sticky=E, padx=5, pady=5)
    output_data_lbl55.grid(row=5, column=5, sticky=E, padx=5, pady=5)
    output_data_lbl56.grid(row=5, column=6, sticky=E, padx=5, pady=5)
    output_data_lbl63.grid(row=6, column=3, sticky=E, padx=5, pady=(5, 20))
    output_data_lbl64.grid(row=6, column=4, sticky=W, padx=5, pady=(5, 20))

    # Поиск информации в базе данных
    response = []
    request = input_data.get().lower()
    try:
        excel_data_file = xlrd.open_workbook(path)
    except Exception:
        output_data_lbl13.grid_remove()
        output_data_lbl14.grid_remove()
        output_data_lbl23.grid_remove()
        output_data_lbl24.grid_remove()
        output_data_lbl33.grid_remove()
        output_data_lbl34.grid_remove()
        output_data_lbl43.grid_remove()
        output_data_lbl44.grid_remove()
        output_data_lbl45.grid_remove()
        output_data_lbl46.grid_remove()
        output_data_lbl53.grid_remove()
        output_data_lbl54.grid_remove()
        output_data_lbl55.grid_remove()
        output_data_lbl56.grid_remove()
        output_data_lbl63.grid_remove()
        output_data_lbl64.grid_remove()
        vnc_but.grid_remove()
        rdp_but.grid_remove()
        no_matches_lbl.grid(row=1, column=3, rowspan=6, columnspan=4, padx=10)
        no_matches_lbl.config(width=34, text='Не могу открыть файл базы данных!')
        return None
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(1, row_number):
            if all([(request != ''), (request != ' '), \
                    (request in str(sheet.row(row)[0]).replace('text:', '').replace("'", '').replace('ё','е').lower() or \
                     request in str(sheet.row(row)[1]).replace('text:', '').replace("'", '').replace('ё','е').lower() or \
                     request in str(sheet.row(row)[3]).replace('number:', '').replace('text:', '').replace(".0", '') or \
                     request in str(sheet.row(row)[4]).replace('text:', '').replace("'", '').replace('ё','е').lower())]):
                response.append(str(sheet.row(row)[6]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[5]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[0]).replace('text:', '').replace("'", ''))
                response.append(str(sheet.row(row)[1]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[2]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[4]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[7]).replace('text:', '').replace("'", '').replace('empty:', '<нет данных>'))
                response.append(str(sheet.row(row)[3]).replace('number:', '').replace(".0", '').replace('text:', '').replace("'",'').replace('empty:', '<нет данных>'))
                response.append('')
        if response != []:

            # Заполнение окна с результатами поиска данными
            output_data_lbl13.config(text='Отделение: ')
            output_data_lbl14.config(text=response[0])
            output_data_lbl23.config(text='Должность: ')
            output_data_lbl24.config(text=response[1])
            output_data_lbl33.config(text='ФИО: ')
            output_data_lbl34.config(text=response[2])
            output_data_lbl43.config(text='Логин: ')
            output_data_lbl44.config(text=response[3])
            output_data_lbl45.config(text='ID: ')
            output_data_lbl46.config(text=response[4])
            output_data_lbl53.config(text='Кабинет: ')
            output_data_lbl54.config(text=response[6])
            output_data_lbl55.config(text='Телефон: ')
            output_data_lbl56.config(text=response[7])
            output_data_lbl63.config(text='Имя ПК: ')
            output_data_lbl64.config(text=response[5])
            vnc_but.grid(row=6, column=5, sticky=W, padx=7, pady=(5, 20))
            rdp_but.grid(row=6, column=6, sticky=W, padx=7, pady=(5, 20))
            no_matches_lbl.grid_remove()
        else:
            output_data_lbl13.grid_remove()
            output_data_lbl14.grid_remove()
            output_data_lbl23.grid_remove()
            output_data_lbl24.grid_remove()
            output_data_lbl33.grid_remove()
            output_data_lbl34.grid_remove()
            output_data_lbl43.grid_remove()
            output_data_lbl44.grid_remove()
            output_data_lbl45.grid_remove()
            output_data_lbl46.grid_remove()
            output_data_lbl53.grid_remove()
            output_data_lbl54.grid_remove()
            output_data_lbl55.grid_remove()
            output_data_lbl56.grid_remove()
            output_data_lbl63.grid_remove()
            output_data_lbl64.grid_remove()
            vnc_but.grid_remove()
            rdp_but.grid_remove()
            no_matches_lbl.grid(row=1, column=3, rowspan=6, padx=10)
            no_matches_lbl.config(text='Совпадений не найдено')


# Инициализация главного окна
root = Tk()
root.geometry('+300+200')
root.title('Crawler v11.0')
root.resizable(False, False)
root.iconbitmap(ICON_PATH)
root.config(bd=5, relief=RIDGE)

# Задание стиля для кнопок
button_style = ttk.Style(root)
button_style.configure('TButton', font=('Lucida Console', 10), background='#900C3F')

# Отрисовка интерфейса
search_lbl = ttk.Label(text='Поиск: ', font=('Lucida Console', 10))
search_lbl.grid(row=1, column=0, rowspan=6, sticky=E, padx=(20, 5), pady=20)
header_lbl = Label(bg='#900C3F', fg='#FFFFFF', text='IT-Crawler!', font=('Lucida Console', 10))
header_lbl.grid(row=0, columnspan=8, sticky=W + E)
footer_lbl = Label(height=2, bg='#900C3F', fg='#FFFFFF', text='© 2020 «Crawler»', font=('Lucida Console', 10))
footer_lbl.grid(row=8, columnspan=8, sticky=W + E)

res_lbl = ttk.Label(text='Результат: ', font=('Lucida Console', 10))
output_data_lbl13 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl14 = ttk.Label(font=('Lucida Console', 10), anchor=W, cursor='spider')
output_data_lbl23 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl24 = ttk.Label(font=('Lucida Console', 10), cursor='spider')
output_data_lbl33 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl34 = ttk.Label(font=('Lucida Console', 10), cursor='spider')
output_data_lbl43 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl44 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='#900C3F', cursor='spider')
output_data_lbl45 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl46 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='#900C3F', cursor='spider')
output_data_lbl53 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl54 = ttk.Label(width=15, font=('Lucida Console', 10), cursor='spider')
output_data_lbl55 = ttk.Label(width=8, font=('Lucida Console', 10, 'bold'))
output_data_lbl56 = ttk.Label(width=15, font=('Lucida Console', 10), foreground='#900C3F', cursor='spider')
output_data_lbl63 = ttk.Label(width=10, font=('Lucida Console', 10, 'bold'))
output_data_lbl64 = ttk.Label(font=('Lucida Console', 10), cursor='spider')
no_matches_lbl = ttk.Label(width=22, font=('Lucida Console', 10, 'bold'))

input_data = ttk.Entry(width=20)
input_data.grid(row=1, column=1, rowspan=6, sticky=W, pady=10)

vnc_but = ttk.Button(text='VNC', command=partial(open_vnc, VNC_PATH), cursor='dotbox')
rdp_but = ttk.Button(text='RDP', command=open_rdp, cursor='dotbox')
language_but = Button(text='RU', font=('Lucida Console', 9), bg='#900C3F', fg='#FFFFFF', relief='flat',command=change_language)
language_but.grid(row=2, column=0, rowspan=5, sticky=SW, padx=(20, 5), pady=(0, 10))
VNC_folder_but = Button(text='VNC', font=('Lucida Console', 9), bg='#900C3F', fg='#FFFFFF', relief='flat',command=open_vnc_folder)
VNC_folder_but.grid(row=2, column=0, rowspan=5, sticky=SW, padx=(50, 5), pady=(0, 10))
show_but = ttk.Button(text='Вывод', command=partial(data_crawler, BD_PATH))
show_but.grid(row=1, column=7, rowspan=6, ipady=10, ipadx=10, pady=60, padx=20)

# Обработка событий нажатия клавиш
output_data_lbl13.bind('<Button-1>', click)
output_data_lbl14.bind('<Button-1>', click)
output_data_lbl23.bind('<Button-1>', click)
output_data_lbl24.bind('<Button-1>', click)
output_data_lbl33.bind('<Button-1>', click)
output_data_lbl34.bind('<Button-1>', click)
output_data_lbl43.bind('<Button-1>', click)
output_data_lbl44.bind('<Button-1>', click)
output_data_lbl45.bind('<Button-1>', click)
output_data_lbl46.bind('<Button-1>', click)
output_data_lbl53.bind('<Button-1>', click)
output_data_lbl54.bind('<Button-1>', click)
output_data_lbl55.bind('<Button-1>', click)
output_data_lbl56.bind('<Button-1>', click)
output_data_lbl63.bind('<Button-1>', click)
output_data_lbl64.bind('<Button-1>', click)
root.bind('<Return>', press)
root.bind('<KeyRelease>', release)

# Поместить курсор в окно поиска
input_data.focus()

get_language()
root.mainloop()
