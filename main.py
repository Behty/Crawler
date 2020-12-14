import xlrd
from tkinter import *
from tkinter import ttk


def data_crawler():
    response = []
    request = input_data.get().lower()
    excel_data_file = xlrd.open_workbook('hardData.xlsx')
    sheet = excel_data_file.sheet_by_index(0)
    row_number = sheet.nrows
    if row_number > 0:
        for row in range(1, row_number):
            if request in str(sheet.row(row)[0]).replace('text:','').replace("'", '').lower() or \
            request in str(sheet.row(row)[1]).replace('text:','').replace("'", '').lower() or \
            request in str(sheet.row(row)[3]).replace('number:','').replace(".0", '') or \
            request in str(sheet.row(row)[4]).replace('text:','').replace("'", '').lower():
                response.append('Отделение:  ' + str(sheet.row(row)[6]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('Должность:  ' + str(sheet.row(row)[5]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('ФИО:  ' + str(sheet.row(row)[0]).replace('text:','').replace("'", ''))
                response.append('Логин:  ' + str(sheet.row(row)[1]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('ID:  ' + str(sheet.row(row)[2]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('Имя компьютера:  ' + str(sheet.row(row)[4]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('Кабинет:  ' + str(sheet.row(row)[7]).replace('text:','').replace("'", '').replace('empty:', '<нет данных>'))
                response.append('Телефон:  ' + str(sheet.row(row)[3]).replace('number:','').replace(".0", '').replace('empty:', '<нет данных>'))
                response.append('')
        if response != []:
            output_data_lbl.config(text='\n'.join(response))
        else:
            output_data_lbl.config(text='Совпадений не найдено')


root = Tk()
root.title('Crawler v2.0')
root.geometry('+300+200')
root.resizable(False, False)

ttk.Label(text='Поиск: ', font=('Lucida Console', 11)).grid(row=0, column=0, sticky=E, padx=15, pady=20)
input_data = ttk.Entry(width=15)
input_data.grid(row=0, column=1, sticky=W, pady=10)

ttk.Label(text='Сотрудник: ', font=('Lucida Console', 11)).grid(row=0, column=2, sticky=E, padx=10, pady=20)
output_data_lbl = ttk.Label(width=50, font=('Lucida Console', 10))
output_data_lbl.grid(row=0, column=3, sticky=E, padx=10, pady=20)

style = ttk.Style(root)
style.configure('TButton', font=('Lucida Console', 11))
show_but = ttk.Button(text='Вывод', command=data_crawler, style='TButton').grid(row=0, column=6, ipady=10, ipadx=10, pady=20, padx=20)

input_data.focus()

root.mainloop()
