import sys
import os
import ui2 as GUI
import log2 as LoginForm
import xlrd
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from PyQt5.Qt import *


CONFIG = ''
TEMPLATE = ''
TABLE = ''


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
            return False
    # ask(template, 'Это правильный шаблон?')
    # template = template.replace('\n', ' ')
    return template


def rtv_settings(email_settings):
    try:
        with open(email_settings, encoding='utf-8') as f:
            settings_rows = f.readlines()
    except FileNotFoundError:
        print('Файл с настройками email_settings.txt не найден')
        return False

    settings = {'FromMail': None, 'gmail/yandex': None, 'FromName': None, 'email_template': None, 'email_list': None}
    for key in settings:
        for row in settings_rows:
            if key in row:
                settings[key] = row[row.find(':')+1:].strip()
    for key, val in settings.items():
        if not val:
            print(f'В файле {email_settings} не заполнен параметр', key)
            return False
    if settings['gmail/yandex'] not in ['gmail', 'yandex']:
        print('В качестве почты (настройка gmail/yandex) поддерживается пока только yandex и gmail')
        return False
    # ask(settings, 'Настройки для отправки в норме?')
    return settings


def rtv_table(xls_name):
    try:
        xl_workbook = xlrd.open_workbook(xls_name)
    except FileNotFoundError:
        print(f'Файл {xls_name} не найден')
        return False
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
            return True
        except KeyError as e:
            print('Поля', e, 'из шаблона нет в таблице')
            return False


def openfile():
    global CONFIG, TABLE, TEMPLATE
    def update_temp(item):
        for i in range(ui.listWidget.count()):
            if ui.listWidget.item(i) == item:
                ui.textBrowser.setText(template.format(**table[i]))

    options = QFileDialog.Options()
    filename, _ = QFileDialog.getOpenFileName(caption='Выберите конфиг', directory='', filter='All files(*)',
                                                        options=options)

    CONFIG = settings = rtv_settings(filename)
    TEMPLATE = template = rtv_template(settings['email_template'])
    TABLE = table = rtv_table(settings['email_list'])
    res = check_template_vs_table(template, table)
    if res and template and table:
        ui.textBrowser.setText(template.format(**table[0]))
        ui.listWidget_2.addItems([table[0]['attach1'], table[0]['attach2']])
        for i in table:
            item = QListWidgetItem(' '.join([i['ID'],i['Фамилия'],i['Имя']]))
            ui.listWidget.addItem(item)
            item.setCheckState(Qt.Checked)
        ui.listWidget.itemClicked.connect(update_temp)
        ui.listWidget_2.itemClicked.connect()


def send_mail(send_from, sender_name, send_to, subject, text, smtp, files=None):
    assert isinstance(send_to, list)
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text, 'html'))
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)
    smtp.sendmail(send_from, send_to, msg.as_string())


def connect_to_server(server, login, password):
    smtp = smtplib.SMTP(server)
    smtp.starttls()
    smtp.login(login, password)
    smtp.ehlo()
    return smtp


def send_msg():
    if CONFIG['gmail/yandex'] == 'yandex':
        mailserver = 'smtp.yandex.ru'
    elif CONFIG['gmail/yandex'] == 'gmail':
        mailserver = 'smtp.googlemail.com'
    else:
        raise NotImplementedError # Пока не включены в список
    frommail = CONFIG['FromMail']
    fromname = CONFIG['FromName']

    loginf = QDialog()
    diagui = LoginForm.Ui_Dialog()
    diagui.setupUi(loginf)
    loginf.exec_()
    login = diagui.lineEdit.text()
    passw = diagui.lineEdit_2.text()

    try:
        smtp = connect_to_server(mailserver, login, passw)
    except smtplib.SMTPAuthenticationError:
        QMessageBox.warning(w, 'Ошибка', 'Неправильный логин/пароль')
        return
    except:
        QMessageBox.warning(w, 'Ошибка', 'Не могу подключиться к серверу')
        return
    del passw
    for i in range(ui.listWidget.count()):
        if ui.listWidget.item(i).checkState():
            send_mail(frommail, fromname, [TABLE[i]['email']], TABLE[i]['subject'], TEMPLATE.format(**TABLE[i]),
                      smtp, ['dummy.py'])


app = QApplication(sys.argv)
w = QMainWindow()
ui = GUI.Ui_MainWindow()
ui.setupUi(w)
w.show()
ui.pushButton.clicked.connect(openfile)
ui.pushButton_2.clicked.connect(send_msg)
sys.exit(app.exec_())