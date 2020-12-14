import xlrd
from tkinter import *
from tkinter import ttk


def tel_num_crawler():
    '''Ищет абонента по номеру телефона'''
    abonents = []
    file_tel_num = int_tel_num.get()
    excel_data_file = xlrd.open_workbook('./tellist.xlsx')
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(0, row_number):
            if str(file_tel_num) == str(sheet.row(row)[2]).replace('number:','').replace('.0',''):
                abonents.append(str(sheet.row(row)[1]).replace('text:','').replace("'", ''))
                continue
        if abonents != []:
            tel_num_lbl.config(text='\n'.join(abonents))
        else:
            tel_num_lbl.config(text='Совпадений не найдено')


def ID_crawler():
    '''Ищет ID по фамилии'''
    file_tel_num = lastname.get()
    excel_data_file = xlrd.open_workbook('./IDlist.xlsx')
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(0, row_number):
            if str(file_tel_num) == str(sheet.row(row)[0]).replace('text:','').replace("'",''):
                IDword_lbl.config(text=str(sheet.row(row)[2]).replace('text:','').replace("'",''))
                break
            else:
                IDword_lbl.config(text='Совпадений не найдено')


root = Tk()
root.title('Crawler v1.0')
root.geometry('+300+300')
root.resizable(False, False)

#Поиск абонента по номеру телефона
ttk.Label(text='Внутренний номер: ').grid(row=0, column=0, sticky=E, padx=10, pady=20)
int_tel_num = ttk.Entry(width=15)
int_tel_num.grid(row=0, column=1, sticky=W, pady=10)

ttk.Label(text='Абонент: ').grid(row=0, column=2, sticky=E, padx=10, pady=20)
tel_num_lbl = ttk.Label(width=30)
tel_num_lbl.grid(row=0, column=3, sticky=E, padx=10, pady=20)

show_but = ttk.Button(text='Узнать', command=tel_num_crawler).grid(row=0, column=6, ipadx=10, pady=20, padx=50)

#Поиск ID по фамилии
ttk.Label(text='Фамилия: ').grid(row=1, column=0, sticky=E, padx=10, pady=20)
lastname = ttk.Entry(width=15)
lastname.grid(row=1, column=1, sticky=W, pady=10)

ttk.Label(text='ID: ').grid(row=1, column=2, sticky=E, padx=10, pady=20)
IDword_lbl = ttk.Label(width=30)
IDword_lbl.grid(row=1, column=3, sticky=E, padx=10, pady=20)

show_but = ttk.Button(text='Узнать', command=ID_crawler).grid(row=1, column=6, ipadx=10, pady=20, padx=50)

int_tel_num.focus()
root.mainloop()
