import sys
import os
import log2 as LoginForm
import xlrd
import smtplib
import subprocess
import ui2
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from PyQt5.Qt import *
import time


class Worker(QObject):
    sig_step = pyqtSignal(int, str)  # worker id, step description: emitted every step through work() loop
    sig_done = pyqtSignal(int)  # worker id: emitted at end of work()
    sig_msg = pyqtSignal(str)  # message to be shown to user

    def __init__(self, id: int):
        super().__init__()
        self.__id = id
        self.__abort = False

    def set_mail_and_smtp_instance(self, send_from, sender_name, send_to, subject, text, smtp, files=None):
        assert isinstance(send_to, list)
        self.msg = MIMEMultipart()
        self.msg['From'] = send_from
        self.msg['To'] = COMMASPACE.join(send_to)
        self.msg['Date'] = formatdate(localtime=True)
        self.msg['Subject'] = subject
        self.msg.attach(MIMEText(text, 'html'))
        for f in files or []:
            with open(f, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(f)
                )
            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            self.msg.attach(part)
        self.smtp = smtp
        self.send_from = send_from
        self.send_to = send_to
        # smtp.sendmail(send_from, send_to, self.msg.as_string())

    @pyqtSlot()
    def work(self):
        thread_name = QThread.currentThread().objectName()
        thread_id = int(QThread.currentThreadId())  # cast to int() is necessary
        self.sig_msg.emit('Running worker #{} from thread "{}" (#{})'.format(self.__id, thread_name, thread_id))
        self.smtp.sendmail(self.send_from, self.send_to, self.msg.as_string())

        self.sig_done.emit(self.__id)

    def abort(self):
        self.sig_msg.emit('Worker #{} notified to abort'.format(self.__id))
        self.__abort = True


class Extended_GUI(ui2.Ui_MainWindow, QObject):
    NUM_THREADS = 1

    sig_abort_workers = pyqtSignal()

    def __init__(self, mainw):
        super().__init__()
        self.setupUi(mainw)
        self.CONFIG = ''
        self.TEMPLATE = ''
        self.TABLE = ''
        self.parent = mainw
        self.pushButton.clicked.connect(self.open_config)
        self.pushButton_2.clicked.connect(self.send_msg)
        self.pushButton_3.clicked.connect(self.abort_workers)
        self.__workers_done = None
        self.__threads = None

    def on_worker_step(self, worker_id: int, data: str):
        # self.log.append('Worker #{}: {}'.format(worker_id, data))
        # self.progress.append('{}: {}'.format(worker_id, data))
        self.statusbar.showMessage('Worker #{}: {}'.format(worker_id, data))

    def on_worker_done(self, worker_id):
        self.statusbar.showMessage('worker #{} done'.format(worker_id))
        self.__workers_done += 1
        if self.__workers_done == self.NUM_THREADS:
            # self.log.append('No more workers active')
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setDisabled(True)
            # self.__threads = None

    def abort_workers(self):
        self.sig_abort_workers.emit()
        self.statusbar.showMessage('Asking each worker to abort')
        for thread, worker in self.__threads:  # note nice unpacking by Python, avoids indexing
            thread.quit()  # this will quit **as soon as thread event loop unblocks**
            thread.wait()  # <- so you need to wait for it to *actually* quit

        # even though threads have exited, there may still be messages on the main thread's
        # queue (messages that threads emitted before the abort):
        self.statusbar.showMessage('All threads exited')
        self.pushButton_3.setDisabled(True)
        self.pushButton_2.setEnabled(True)

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

    def connect_to_server(self, server, login, password):
        smtp = smtplib.SMTP(server)
        smtp.starttls()
        smtp.login(login, password)
        smtp.ehlo()
        return smtp

    def send_msg(self):
        for idx in range(self.NUM_THREADS):
            if self.CONFIG['gmail/yandex'] == 'yandex':
                mailserver = 'smtp.yandex.ru'
            elif self.CONFIG['gmail/yandex'] == 'gmail':
                mailserver = 'smtp.googlemail.com'
            else:
                raise NotImplementedError('Unknown mailserver.')  # Пока не включены в список
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
                self.statusbar.showMessage('starting {} threads'.format(self.NUM_THREADS))
                self.pushButton_2.setDisabled(True)
                self.pushButton_3.setEnabled(True)
                self.__workers_done = 0
                self.__threads = []
                worker = Worker(idx)
                thread = QThread()
                thread.setObjectName('thread_' + str(idx))
                self.__threads.append((thread, worker))  # need to store worker too otherwise will be gc'd
                worker.moveToThread(thread)

                # get progress messages from worker:
                worker.sig_step.connect(self.on_worker_step)
                worker.sig_done.connect(self.on_worker_done)
                worker.sig_msg.connect(self.statusbar.showMessage)

                # control worker:
                self.sig_abort_workers.connect(worker.abort)

                # get read to start worker:
                # self.sig_start.connect(worker.work)  # needed due to PyCharm debugger bug (!); comment out next line
                thread.started.connect(worker.work)
                # this will emit 'started' and start thread's event loop
                for i in range(self.listWidget.count()):
                    if self.listWidget.item(i).checkState():
                        worker.set_mail_and_smtp_instance(frommail, fromname, [self.TABLE[i]['email']], self.TABLE[i]['subject'],
                                       self.TEMPLATE.format(**self.TABLE[i]),
                                       smtp, [self.TABLE[i]['attach1'], self.TABLE[i]['attach2']])
                        thread.start()