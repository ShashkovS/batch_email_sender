import sys
import os
import test
import xlrd
from PyQt5.Qt import *


def rtv_template(template_name):
    try:
        with open(template_name, encoding='utf-8') as f:
            template = f.read()
    except FileNotFoundError:
        os.chdir('..')
        try:
            with open(template_name, encoding='utf-8') as f:
                template = f.read()
        except FileNotFoundError:
            print('Файл с шаблоном email_template.txt не найден')
            exit(1)
    # ask(template, 'Это правильный шаблон?')
    # template = template.replace('\n', ' ')
    return template


def rtv_settings(email_settings):
    try:
        with open(email_settings, encoding='utf-8') as f:
            settings_rows = f.readlines()
    except FileNotFoundError:
        print('Файл с настойками email_settings.txt не найден')
        exit(1)

    settings = {'FromMail': None, 'gmail/yandex': None, 'FromName': None, 'email_template': None, 'email_list': None}
    for key in settings:
        for row in settings_rows:
            if key in row:
                settings[key] = row[row.find(':')+1:].strip()
    for key, val in settings.items():
        if not val:
            print(f'В файле {email_settings} не заполнен параметр', key)
            exit(1)
    if settings['gmail/yandex'] not in ['gmail', 'yandex']:
        print('В качестве почты (настройка gmail/yandex) поддерживается пока только yandex и gmail')
        exit(1)
    # ask(settings, 'Настройки для отправки в норме?')
    return settings


def rtv_table(xls_name):
    try:
        xl_workbook = xlrd.open_workbook(xls_name)
    except FileNotFoundError:
        print(f'Файл {xls_name} не найден')
        exit(1)
    xl_sheet = xl_workbook.sheet_by_index(0)
    columns = {}
    for i in range(xl_sheet.ncols):
        title = str(xl_sheet.cell(0,i).value)
        if title:
            columns[title] = i
    result = []
    for j in range(1, xl_sheet.nrows):
        cur = columns.copy()
        for key, col in columns.items():
            cur[key] = str(xl_sheet.cell(j,col).value)
        result.append(cur)
    # ask('\n'.join(map(str, result)), 'Данные для отправки похожи на правду?')
    return result


def check_template_vs_table(template, table):
    if table:
        try:
            dummy = (template + '{email}{subject}').format(**table[0])
        except KeyError as e:
            print('Поля', e, 'из шаблона нет в таблице')
            exit(1)


def openfile():
    options = QFileDialog.Options()
    filename, _ = QFileDialog.getOpenFileName(caption='Выберите конфиг', directory='', filter='All files(*)',
                                                        options=options)

    settings = rtv_settings(filename)
    template = rtv_template(settings['email_template'])
    table = rtv_table(settings['email_list'])
    print(table)
    check_template_vs_table(template, table)


app = QApplication(sys.argv)
w = QMainWindow()
ui = test.Ui_MainWindow()
ui.setupUi(w)
w.show()
ui.pushButton.clicked.connect(openfile)
sys.exit(app.exec_())