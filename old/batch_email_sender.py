import xlrd
import getpass
from envelopes import GMailSMTP
from random import choice
from sys import exit
import os
import tempfile
import yagmail
import time

OKS = ['ok', 'good', 'yes', 'sure']

# os.chdir(r'C:\Dropbox\M2021\Собеседование в 7-й класс, 2017\Письма поступающим')
email_settings = r'res_next_email_settings.txt'

# ONLY_ONE = True
ONLY_ONE = False



def transliterate(string):
    capital_letters = {'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'E', 'Ж': 'Zh', 'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'H', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sch', 'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',}
    lower_case_letters = {'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e', 'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',}
    translit_string = ""
    for index, char in enumerate(string):
        if char in lower_case_letters:
            char = lower_case_letters[char]
        elif char in capital_letters:
            char = capital_letters[char]
            if len(string) > index+1:
                if string[index+1] not in lower_case_letters.keys():
                    char = char.upper()
            else:
                char = char.upper()
        translit_string += char
    return translit_string


def ask(data, question):
    cur_cor_ans = choice(OKS)
    print('\n'*5)
    print('*'*100)
    print(data)
    print('*'*100)
    print(question)
    ask = input('Введите "' + cur_cor_ans + '", если всё ОК:\n')
    if ask != cur_cor_ans:
        print('Отменяем')
        exit(0)


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
    ask(template, 'Это правильный шаблон?')
    template = template.replace('\n', ' ')
    return template


def rtv_settings():
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
    ask(settings, 'Настройки для отправки в норме?')
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
    ask('\n'.join(map(str, result)), 'Данные для отправки похожи на правду?')
    return result


def check_template_vs_table(template, table):
    if table:
        try:
            dummy = (template + '{email}{subject}').format(**table[0])
        except KeyError as e:
            print('Поля', e, 'из шаблона нет в таблице')
            exit(1)


def connect_to_smtp(FromMail, gmail_yandex):
    if gmail_yandex == 'gmail':
        GMailSMTP.GMAIL_SMTP_HOST = 'smtp.googlemail.com'
    elif gmail_yandex == 'yandex':
        GMailSMTP.GMAIL_SMTP_HOST = 'smtp.yandex.ru'
    print('Сейчас нужно будет ввести пароль от почты.')
    password = getpass.getpass('Enter password: ')
    # password = input('password')
    print('Пароль принят, пытаемся подключиться.')
    # gmail = GMailSMTP(FromMail, password)
    gmail = yagmail.SMTP(FromMail, password)
    del password
    print('Удалось подключиться для отправки писем.')
    return gmail


def send_mail(FromMail, FromName, gmail, template, data_row, first_mail=[]):
    mail_text = template.format(**data_row)
    if not first_mail:
        ask_mail_text = mail_text[:]
        ask_mail_text = ask_mail_text.replace('<table', '\n<table')
        ask_mail_text = ask_mail_text.replace('<tr>', '\n<tr>')
        ask_mail_text = ask_mail_text.replace('<p>', '\n<p>')
        ask_mail_text = ask_mail_text.replace('<br>', '\n<br>')
        ask_mail_text = ask_mail_text.replace('<br/>', '\n<br/>')
        ask(ask_mail_text, 'Первое письмо выглядит так, это правильно?')
        first_mail.append('Не в первой')
    with tempfile.TemporaryDirectory() as tmpdirname:
        # envelope = Envelope(
        #     from_addr=(FromMail, FromName),
        #     to_addr=data_row['email'],
        #     subject=data_row['subject'],
        #     html_body=mail_text
        # )
        to_addr = data_row['email']
        subject = data_row['subject']
        contents = [mail_text]
        # os.chdir(r'C:\Dropbox\M2021\Собеседование в 7-й класс, 2017\Сканы работ\По школьникам')
        if 'attach1' in data_row:
            fr_n = data_row['attach1']
            contents.append(fr_n)
            # to_n = os.path.join(tmpdirname, transliterate(os.path.basename(fr_n)))
            # shutil.copyfile(fr_n, to_n)
            # envelope.add_attachment(to_n, mimetype='application/pdf')

        if 'attach2' in data_row:
            fr_n = data_row['attach2']
            contents.append(fr_n)
            # to_n = os.path.join(tmpdirname, transliterate(os.path.basename(fr_n)))
            # shutil.copyfile(fr_n, to_n)
            # envelope.add_attachment(to_n, mimetype='application/pdf')

        try:
            # gmail.send(envelope)
            gmail.send(to=to_addr, subject=subject, contents=contents)
            res = 'Отправили письмо по адресу ' + data_row['email']
            print(res)
        except:
            res = 'ОШИБКА ОТПРАВКИ ПО АДРЕСУ ' + data_row['email']
            print(res)
            print('При отправке через сервера gmail, нужно временно разрешить доступ непроверенным приложениям.')
            print('Это нужно сделать по ссылке: https://www.google.com/settings/security/lesssecureapps?rfn=27&rfnc=1&et=0&asae=2')
        if ONLY_ONE: exit()
        return res


print(f'Будем читать настройки из файла {email_settings}')
settings = rtv_settings()
print(f'Будем читать шаблон из файла {settings["email_template"]}')
template = rtv_template(settings['email_template'])
print(f'Будем читать список адресов из файла {settings["email_list"]}')
table = rtv_table(settings['email_list'])
print('Вычитали таблицу, проверяем корректность шаблона')
check_template_vs_table(template, table)
print('Коннектимся...')
gmail = 1
# gmail = connect_to_smtp(settings['FromMail'], settings['gmail/yandex'])
sent = 0
with open('log.txt', 'w', encoding='utf-8') as f:
    for data_row in table:
        if not data_row['email'].strip():
            continue
        sent += 1
        if sent % 60 == 0:
            gmail.close()
            time.sleep(10)
            gmail = connect_to_smtp(settings['FromMail'], settings['gmail/yandex'])
        res = send_mail(settings['FromMail'], settings['FromName'], gmail, template, data_row)
        f.write(res)
        f.write('\n')

print('Всё завершено')
# input('Введить что-нибудь за завершения. Лог отправки можно найти в файле log.txt')
exit(0)
