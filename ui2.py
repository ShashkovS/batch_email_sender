# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

import sys
import os
import log2 as LoginForm
import xlrd
import smtplib
import subprocess
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from PyQt5.Qt import *


class Ui_MainWindow(object):
    def __init__(self, MainWindow):
        self.CONFIG = ''
        self.TEMPLATE = ''
        self.TABLE = ''
        self.parent = MainWindow
        self.setupUi(MainWindow)
        self.pushButton.clicked.connect(self.open_config)
        self.pushButton_2.clicked.connect(self.send_msg)


    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(691, 691)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout_4 = QGridLayout()
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.pushButton_2 = QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_4.addWidget(self.pushButton_2, 0, 0, 1, 1)
        self.listWidget = QListWidget(self.centralwidget)
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.listWidget.sizePolicy().hasHeightForWidth())
        self.listWidget.setSizePolicy(sizePolicy)
        self.listWidget.setObjectName("listWidget")
        self.gridLayout_4.addWidget(self.listWidget, 1, 0, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout_4, 0, 0, 1, 1)
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 0, 0, 1, 1)
        self.textBrowser = QTextBrowser(self.centralwidget)
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(99)
        sizePolicy.setHeightForWidth(self.textBrowser.sizePolicy().hasHeightForWidth())
        self.textBrowser.setSizePolicy(sizePolicy)
        self.textBrowser.setObjectName("textBrowser")
        self.gridLayout.addWidget(self.textBrowser, 1, 0, 1, 1)
        self.listWidget_2 = QListWidget(self.centralwidget)
        sizePolicy = QSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(1)
        sizePolicy.setHeightForWidth(self.listWidget_2.sizePolicy().hasHeightForWidth())
        self.listWidget_2.setSizePolicy(sizePolicy)
        self.listWidget_2.setMinimumSize(QSize(0, 75))
        self.listWidget_2.setObjectName("listWidget_2")
        self.gridLayout.addWidget(self.listWidget_2, 2, 0, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout, 0, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setGeometry(QRect(0, 0, 691, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton_2.setText(_translate("MainWindow", "отправить письма всем выделенным людям"))
        self.pushButton.setText(_translate("MainWindow", "открыть конфиг"))

    def rtv_template(self, template_name):
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

    def rtv_settings(self, email_settings):
        try:
            with open(email_settings, encoding='utf-8') as f:
                settings_rows = f.readlines()
        except FileNotFoundError:
            print('Файл с настройками email_settings.txt не найден')
            return False

        settings = {'FromMail': None, 'gmail/yandex': None, 'FromName': None, 'email_template': None,
                    'email_list': None}
        for key in settings:
            for row in settings_rows:
                if key in row:
                    settings[key] = row[row.find(':') + 1:].strip()
        for key, val in settings.items():
            if not val:
                print(f'В файле {email_settings} не заполнен параметр', key)
                return False
        if settings['gmail/yandex'] not in ['gmail', 'yandex']:
            print('В качестве почты (настройка gmail/yandex) поддерживается пока только yandex и gmail')
            return False
        # ask(settings, 'Настройки для отправки в норме?')
        return settings

    def rtv_table(self, xls_name):
        try:
            xl_workbook = xlrd.open_workbook(xls_name)
        except FileNotFoundError:
            print(f'Файл {xls_name} не найден')
            return False
        xl_sheet = xl_workbook.sheet_by_index(0)
        columns = {}
        for i in range(xl_sheet.ncols):
            title = str(xl_sheet.cell(0, i).value)
            if title:
                columns[title] = i
        result = []
        for j in range(1, xl_sheet.nrows):
            cur = columns.copy()
            for key, col in columns.items():
                cur[key] = str(xl_sheet.cell(j, col).value)
            result.append(cur)
        # ask('\n'.join(map(str, result)), 'Данные для отправки похожи на правду?')
        return result

    def check_template_vs_table(self, template, table):
        if table:
            try:
                dummy = (template + '{email}{subject}').format(**table[0])
                return True
            except KeyError as e:
                print('Поля', e, 'из шаблона нет в таблице')
                return False

    def open_attach(self, item):
        for i in range(self.listWidget_2.count()):
            if self.listWidget_2.item(i) == item and self.listWidget_2.item(i).text():
                if sys.platform.startswith('darwin'):
                    subprocess.call(('open', item.text()))
                elif os.name == 'nt':
                    os.startfile(item.text())
                elif os.name == 'posix':
                    subprocess.call(('xdg-open', item.text()))

    def open_config(self):

        def update_temp(item):
            for i in range(self.listWidget.count()):
                if self.listWidget.item(i) == item:
                    self.textBrowser.setText(template.format(**table[i]))
                    self.listWidget_2.clear()
                    self.listWidget_2.addItems([table[i]['attach{}'.format(j)] for j in range(1, 3)])
                    break

        self.listWidget.clear()
        self.listWidget_2.clear()
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(caption='Выберите конфиг', directory='', filter='All files(*)',
                                                  options=options)

        self.CONFIG = settings = self.rtv_settings(filename)
        self.TEMPLATE = template = self.rtv_template(settings['email_template'])
        table = self.rtv_table(settings['email_list'])
        self.TABLE = table.copy()
        res = self.check_template_vs_table(template, table)
        if res and template and table:
            self.textBrowser.setText(template.format(**table[0]))
            print([self.TABLE[2]['attach' + str(i)] for i in range(1, 3)])
            self.listWidget_2.addItems([self.TABLE[0]['attach' + str(i)] for i in range(1, 3)])
            for i in table:
                item = QListWidgetItem(' '.join([i['ID'], i['Фамилия'], i['Имя'], i['Школа']]))
                self.listWidget.addItem(item)
                item.setCheckState(Qt.Checked)
            self.listWidget.itemClicked.connect(update_temp)
            self.listWidget_2.itemClicked.connect(self.open_attach)

    def send_mail(self, send_from, sender_name, send_to, subject, text, smtp, files=None):
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

    def connect_to_server(self, server, login, password):
        smtp = smtplib.SMTP(server)
        smtp.starttls()
        smtp.login(login, password)
        smtp.ehlo()
        return smtp

    def send_msg(self):
        if self.CONFIG['gmail/yandex'] == 'yandex':
            mailserver = 'smtp.yandex.ru'
        elif self.CONFIG['gmail/yandex'] == 'gmail':
            mailserver = 'smtp.googlemail.com'
        else:
            raise NotImplementedError  # Пока не включены в список
        frommail = self.CONFIG['FromMail']
        fromname = self.CONFIG['FromName']

        loginf = QDialog()
        diagui = LoginForm.Ui_Dialog()
        diagui.setupUi(loginf)
        if loginf.exec_() == QDialog.Accepted:
            login = diagui.lineEdit.text()
            passw = diagui.lineEdit_2.text()
            try:
                smtp = self.connect_to_server(mailserver, login, passw)
            except smtplib.SMTPAuthenticationError:
                QMessageBox.warning(self.parent, 'Ошибка', 'Неправильный логин/пароль')
                return
            except:
                QMessageBox.warning(self.parent, 'Ошибка', 'Не могу подключиться к серверу')
                return
            del passw
            for i in range(self.listWidget.count()):
                if self.listWidget.item(i).checkState():
                    self.send_mail(frommail, fromname, [self.TABLE[i]['email']], self.TABLE[i]['subject'],
                                   self.TEMPLATE.format(**self.TABLE[i]),
                                   smtp, [self.TABLE[i]['attach1'], self.TABLE[i]['attach2']])


