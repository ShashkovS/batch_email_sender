import sys
import os
import log2 as LoginForm
import xlrd
import smtplib
import subprocess
import ui2
import files_parsers
import alerts
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from PyQt5.Qt import *
import keyring
KEYRING_SERVICE = "batch_email_sender"
LAST_FROMMAIL = "pSxx7tJyvgz2tk"
LAST_MAILSERVER = "KjdsEYxeRaCk77"
LAST_FROMNAME = "FY8Btthta4n3ZF"

class Worker(QObject):
    sig_step = pyqtSignal(int, str)  # worker id, step description: emitted every step through work() loop
    sig_done = pyqtSignal(int)  # worker id: emitted at end of work()
    sig_msg = pyqtSignal(str)  # message to be shown to user

    def __init__(self, id: int, table, template, xls_name, frommail, fromname, listwidget, server, login, password):
        super().__init__()
        self.__id = id
        self.__abort = False
        self.TABLE = table
        self.TEMPLATE = template
        self.frommail = frommail
        self.fromname = fromname
        self.listWidget = listwidget
        self.xls_name = xls_name
        try:
            self.smtp = self.connect_to_server(server, login, password)
        except smtplib.SMTPAuthenticationError:
            QMessageBox.warning(self.listWidget.parent(), 'Ошибка', 'Неправильный логин/пароль')
            return
        except:
            QMessageBox.warning(self.listWidget.parent(), 'Ошибка', 'Не могу подключиться к серверу')
            return

    def connect_to_server(self, server, login, password):
        smtp = smtplib.SMTP(server)
        smtp.starttls()
        smtp.login(login, password)
        smtp.ehlo()
        return smtp

    def set_mail(self, send_from, sender_name, send_to, subject, text, files=None):
        assert isinstance(send_to, list)
        self.msg = MIMEMultipart()
        self.msg['From'] = send_from
        self.msg['To'] = COMMASPACE.join(send_to)
        self.msg['Date'] = formatdate(localtime=True)
        self.msg['Subject'] = subject
        self.msg.attach(MIMEText(text, 'html'))
        for f in files or []:
            if not f:
                continue
            with open(f, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(f)
                )
            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            self.msg.attach(part)
        self.send_from = send_from
        self.send_to = send_to
        # smtp.sendmail(send_from, send_to, self.msg.as_string())

    @pyqtSlot()
    def work(self):
        thread_name = QThread.currentThread().objectName()
        thread_id = int(QThread.currentThreadId())  # cast to int() is necessary
        self.sig_msg.emit('Running worker #{} from thread "{}" (#{})'.format(self.__id, thread_name, thread_id))
        for i in range(self.listWidget.count()):
            if self.__abort:
                # note that "step" value will not necessarily be same for every thread
                self.sig_msg.emit('Worker #{} aborting work at step {}'.format(self.__id, i))
                break
            if self.listWidget.item(i).checkState():
                self.set_mail(self.frommail, self.fromname, [self.TABLE[i]['email']],
                                                  self.TABLE[i]['subject'],
                                                  self.TEMPLATE.format(**self.TABLE[i]),
                              [self.TABLE[i]['attach1'], self.TABLE[i]['attach2']])
                self.sig_step.emit(self.__id, 'point ' + str(i))
                self.smtp.sendmail(self.send_from, self.send_to, self.msg.as_string())
                files_parsers.set_ok(self.xls_name, self.TABLE[i][files_parsers.ORIGINAL_ROW_NUM])  #
        self.sig_done.emit(self.__id)

    def abort(self):
        self.sig_msg.emit('Worker #{} notified to abort'.format(self.__id))
        self.__abort = True


class Extended_GUI(ui2.Ui_MainWindow, QObject):

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

    @pyqtSlot(int, str)
    def on_worker_step(self, worker_id: int, data: str):
        # self.log.append('Worker #{}: {}'.format(worker_id, data))
        # self.progress.append('{}: {}'.format(worker_id, data))
        self.statusbar.showMessage('Worker #{}: {}'.format(worker_id, data))

    @pyqtSlot(int)
    def on_worker_done(self, worker_id):
        self.statusbar.showMessage('worker #{} done'.format(worker_id))
        self.__workers_done += 1
        if self.__workers_done == 1:
            # self.log.append('No more workers active')
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setDisabled(True)
            # self.__threads = None

    @pyqtSlot()
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
            nonlocal rows_list
            for i in range(self.listWidget.count()):
                if self.listWidget.item(i) == item:
                    self.textBrowser.setText(template.format(**rows_list[i]))
                    self.listWidget_2.clear()
                    self.listWidget_2.addItems([rows_list[i]['attach{}'.format(j)] for j in range(1, 3)])
                    break

        self.listWidget.clear()
        self.listWidget_2.clear()
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(caption='Выберите список или шаблон', directory='',
                                                  options=options,
                                                  filter="Список или шаблон (*list.xlsx *text.html);;Список (*list.xlsx);;Шаблон (*text.html);;All Files (*)")
        # filename — это либо имя excel'ника, либо имя html-шаблона. По имени определяем, что это, и определяем
        # оставшиеся имена
        filename = filename.lower()
        if filename.endswith('list.xlsx'):
            xls_name = filename
            template_name = filename.replace('list.xlsx', 'text.html')
        elif filename.endswith('text.html'):
            xls_name = filename.replace('text.html', 'list.xlsx')
            template_name = filename
        else:
            alerts.alert(f'Нужно выбрать файл со списком ***list.xlsx или файл с шаблоном письма ***text.html')
            return
        rows_list, bold_columns, template = files_parsers.rtv_table_and_template(xls_name, template_name)
        # Удаляем из списка всё, что уже ОК, и у чего не заполнен email
        rows_list = [row for row in rows_list if row[files_parsers.OKOK] != files_parsers.OKOK and '@' in row['email']]
        self.xls_name =xls_name
        self.template_name = template_name
        self.TEMPLATE = template
        self.TABLE = rows_list
        if template and rows_list:
            self.textBrowser.setText(template.format(**rows_list[0]))
            print([self.TABLE[0]['attach' + str(i)] for i in range(1, 3)])
            self.listWidget_2.addItems([self.TABLE[0]['attach' + str(i)] for i in range(1, 3)])
            for i in rows_list:
                item = QListWidgetItem(' '.join([i[col] for col in bold_columns]))
                self.listWidget.addItem(item)
                item.setCheckState(Qt.Checked)
            self.listWidget.itemClicked.connect(update_temp)
            self.listWidget_2.itemClicked.connect(self.open_attach)

    def send_msg(self):
        # if self.CONFIG['gmail/yandex'] == 'yandex':
        #     mailserver = 'smtp.yandex.ru'
        # elif self.CONFIG['gmail/yandex'] == 'gmail':
        #     mailserver = 'smtp.googlemail.com'
        # else:
        #     raise NotImplementedError('Unknown mailserver.')  # Пока не включены в список
        # frommail = self.CONFIG['FromMail']
        # fromname = self.CONFIG['FromName']

        loginf = QDialog()
        diagui = LoginForm.Ui_Dialog()
        diagui.setupUi(loginf)

        last_frommail = keyring.get_password(KEYRING_SERVICE, LAST_FROMMAIL)
        if last_frommail:
            diagui.line_email.setText(last_frommail)
        last_fromname = keyring.get_password(KEYRING_SERVICE, LAST_FROMNAME)
        if last_fromname:
            diagui.line_sender.setText(last_fromname)
        last_mailserver = keyring.get_password(KEYRING_SERVICE, LAST_MAILSERVER)
        if last_mailserver:
            diagui.line_smtpserver.setText(last_mailserver)
        else:
            diagui.line_smtpserver.setText('smtp.googlemail.com')
        if loginf.exec_() == QDialog.Accepted:
            last_frommail = diagui.line_email.text()
            passw = diagui.line_password.text()
            last_fromname = diagui.line_sender.text()
            last_mailserver = diagui.line_smtpserver.text()
            keyring.set_password(KEYRING_SERVICE, LAST_FROMMAIL, last_frommail)
            keyring.set_password(KEYRING_SERVICE, LAST_FROMNAME, last_fromname)
            keyring.set_password(KEYRING_SERVICE, LAST_MAILSERVER, last_mailserver)
            login = last_frommail
            frommail = last_frommail
            fromname = last_fromname
            mailserver = last_mailserver
            self.statusbar.showMessage('starting send-thread')
            self.pushButton_2.setDisabled(True)
            self.pushButton_3.setEnabled(True)
            self.__workers_done = 0
            self.__threads = []
            worker = Worker(1, self.TABLE, self.TEMPLATE, self.xls_name, frommail, fromname, self.listWidget, mailserver, login, passw)
            del passw
            thread = QThread()
            thread.setObjectName('thread_' + str(1))
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
            thread.start()
            thread.quit()
            thread.wait()
            # this will emit 'started' and start thread's event loop


            QMessageBox.information(self.parent, 'OK', 'Все письма успешно отправлены!')

if __name__ == '__main__':
    print('Не-не-не, нужно запускать main.py')