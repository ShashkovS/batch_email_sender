import sys
import traceback
import os
import subprocess
import re
import keyring
from PySide2.QtCore import QObject, QThread, Qt, Signal, Slot
from PySide2.QtGui import QBrush, QColor
from PySide2.QtWidgets import QMessageBox, QListWidgetItem, QFileDialog, QDialog, QApplication, QMainWindow
# from PyQt5.QtCore import QObject, QThread, Qt
# from PyQt5.QtCore import pyqtSignal as Signal
# from PyQt5.QtCore import pyqtSlot as Slot
# from PyQt5.QtGui import QBrush, QColor
# from PyQt5.QtWidgets import QMessageBox, QListWidgetItem, QFileDialog, QDialog, QApplication, QMainWindow

import files_parsers
import ui_email_and_passw
import ui_main_window
import email_stuff


def excepthook(exc_type, exc_value, traceback_obj):
    traceback.print_tb(traceback_obj, exc_type, exc_value)


sys.excepthook = excepthook

KEYRING_SERVICE = "batch_email_sender"
# Ключи, под которыми будут храниться данные
LAST_FROMMAIL = "pSxx7tJyvgz2tk"
LAST_MAILSERVER = "KjdsEYxeRaCk77"
LAST_FROMNAME = "FY8Btthta4n3ZF"
LAST_COPYLIST = "eqRyLeqKPatefP"
LAST_SAVEFLAG = "8AN43xqzGhZHUa"
LAST_PASSWORD = "uLkTjXd6BWa4tw"

EMAIL_REGEX = r"\s*([a-zA-Z0-9'_][a-zA-Z0-9'._+-]{,63}@[a-zA-Z0-9.-]{,254}[a-zA-Z0-9])\s*"


class Worker(QObject):
    sig_step = Signal(int, str)  # worker id, step description: emitted every step through work() loop
    sig_done = Signal(int)  # worker id: emitted at end of work()
    sig_mail_sent = Signal(int, int)
    sig_mail_error = Signal(int)

    def __init__(self, id: int, envelope):
        super().__init__()
        self.__id = id
        self.__abort = False
        self.envelope = envelope

    @Slot()
    def work(self):
        """
        Pretend this worker method does work that takes a long time. During this time, the thread's
        event loop is blocked, except if the application's processEvents() is called: this gives every
        thread (incl. main) a chance to process events, which in this sample means processing signals
        received from GUI (such as abort).
        """
        thread_name = QThread.currentThread().objectName()
        self.sig_step.emit(self.__id, 'Running worker #{} from thread "{}"'.format(self.__id, thread_name))

        while True:
            batch_sender_app.processEvents()  # this could cause change to self.__abort
            if self.__abort:
                self.sig_step.emit(self.__id, 'Worker #{} aborting work'.format(self.__id))
                break
            qt_mail_id, xls_mail_id = -1, -1
            try:
                mail = self.envelope.send_next()
                qt_mail_id, xls_mail_id, sent_to = mail['qt_id'], mail['xls_id'], mail[
                    'to_addrs']  # TODO здесь что-то грязно
            except StopIteration:
                break  # Это — победа
            except Exception as e:
                self.sig_step.emit(self.__id, 'Worker #{} error: {}'.format(self.__id, e))
            if qt_mail_id >= 0:
                self.sig_step.emit(self.__id, 'Worker #{} sent to {}'.format(self.__id, sent_to))
                self.sig_mail_sent.emit(qt_mail_id, xls_mail_id)
        self.sig_done.emit(self.__id)

    def abort(self):
        self.sig_step.emit(self.__id, 'Worker #{} notified to abort'.format(self.__id))
        self.__abort = True


class Extended_GUI(ui_main_window.Ui_MainWindow, QObject):
    NUM_THREADS = 5
    USE_THREADS = None

    sig_abort_workers = Signal()

    def __init__(self, mainw):
        super().__init__()
        self.setupUi(mainw)
        self.CONFIG = ''
        self.template = ''
        self.xlsx_rows_list = ''
        self.parent = mainw
        self.pushButton_open_list_and_template.clicked.connect(self.open_xls_and_template)
        self.pushButton_ask_and_send.clicked.connect(self.send_msg)
        self.pushButton_cancel_send.clicked.connect(self.abort_workers)
        QThread.currentThread().setObjectName('main')  # threads can be named, useful for log output
        self.__workers_done = None
        self.__threads = None

    @Slot(int, str)
    def on_worker_step(self, worker_id: int, data: str):
        self.statusbar.showMessage('Worker #{}: {}'.format(worker_id, data))

    @Slot(int, int)
    def on_mail_sent(self, mail_widget_row_num: int, xls_row_number_ok: int):
        item = self.listWidget_emails.item(mail_widget_row_num)
        item.setBackground(QBrush(QColor("lightGreen")))  # Вах!
        item.setCheckState(Qt.Unchecked)
        try:
            files_parsers.set_ok(self.xls_name, xls_row_number_ok)
        except Exception as e:
            print(e)

    @Slot(int)
    def on_mail_error(self, mail_widget_row_num: int):
        item = self.listWidget_emails.item(mail_widget_row_num)
        item.setBackground(QBrush(QColor("lightRed")))  # Вах!

    @Slot(int)
    def on_worker_done(self, worker_id):
        self.statusbar.showMessage('worker #{} done'.format(worker_id))
        self.__workers_done += 1
        if self.__workers_done == self.USE_THREADS:
            self.statusbar.showMessage('No more workers active')
            self.pushButton_ask_and_send.setEnabled(True)
            self.pushButton_open_list_and_template.setEnabled(True)
            self.pushButton_cancel_send.setDisabled(True)
            QMessageBox.information(self.parent, 'OK', 'Все письма успешно отправлены!')

    @Slot()
    def abort_workers(self):
        self.sig_abort_workers.emit()
        self.statusbar.showMessage('Asking each worker to abort')
        for thread, worker in self.__threads:
            thread.quit()  # this will quit **as soon as thread event loop unblocks**
            thread.wait()  # <- so you need to wait for it to *actually* quit
        # even though threads have exited, there may still be messages on the main thread's
        # queue (messages that threads emitted before the abort):
        self.statusbar.showMessage('All threads exited')
        self.pushButton_ask_and_send.setEnabled(True)
        self.pushButton_open_list_and_template.setEnabled(True)
        self.pushButton_cancel_send.setDisabled(True)

    def show_email_attach(self, item):
        for i in range(self.listWidget_attachments.count()):
            if self.listWidget_attachments.item(i) == item and self.listWidget_attachments.item(i).text():
                if sys.platform.startswith('darwin'):
                    subprocess.call(('open', item.text()))
                elif os.name == 'nt':
                    os.startfile(item.text())
                elif os.name == 'posix':
                    subprocess.call(('xdg-open', item.text()))

    def read_list_and_template(self, filename):
        # filename — это либо имя excel'ника, либо имя html-шаблона. По имени определяем, что это, и определяем
        # оставшиеся имена
        self.template = None
        self.xlsx_rows_list = None
        self.pushButton_ask_and_send.setDisabled(True)
        # filename = filename.lower()
        if filename.endswith('list.xlsx'):
            xls_name = filename
            template_name = filename.replace('list.xlsx', 'text.html')
        elif filename.endswith('text.html'):
            xls_name = filename.replace('text.html', 'list.xlsx')
            template_name = filename
        else:
            raise Exception('Нужно выбрать файл со списком ***list.xlsx или файл с шаблоном письма ***text.html')
        rows_list, preview_columns, template = files_parsers.rtv_table_and_template(xls_name, template_name)
        self.xls_name = xls_name
        self.template_name = template_name
        self.template = template
        self.xlsx_rows_list = rows_list
        if 'email' not in preview_columns:
            preview_columns.append('email')
        self.info_cols = preview_columns
        self.pushButton_ask_and_send.setEnabled(True)

    def update_preview_and_attaches_list(self, item):
        for i in range(self.listWidget_emails.count()):
            if self.listWidget_emails.item(i) is item:
                xlsx_row = self.xlsx_rows_list[i]
                self.textBrowser.setText('<h2>{}</h2><hr>\n'.format(xlsx_row['subject'])
                                         + self.template.format(**xlsx_row))
                self.listWidget_attachments.clear()
                self.listWidget_attachments.addItems(xlsx_row['attach_list'])
                break

    def fill_widgets_with_emails(self):
        if not self.template or not self.xlsx_rows_list:
            return
        for xlsx_row in self.xlsx_rows_list:
            sample_row_text = ' '.join([xlsx_row[col] if col != 'email' else ', '.join(xlsx_row[col])
                                        for col in self.info_cols])
            item = QListWidgetItem(sample_row_text)
            item.xlsx_row = xlsx_row  # Именно отсюда мы возьмём данные для отправки
            item.setCheckState(Qt.Checked)
            if xlsx_row[files_parsers.OKOK] == files_parsers.OKOK:
                item.setCheckState(Qt.Unchecked)
                item.setBackground(QBrush(QColor("lightGreen")))  # Вах!
            self.listWidget_emails.addItem(item)
        self.listWidget_emails.itemClicked.connect(self.update_preview_and_attaches_list)
        self.listWidget_attachments.itemClicked.connect(self.show_email_attach)
        self.update_preview_and_attaches_list(self.listWidget_emails.item(0))

    def open_xls_and_template(self):
        filename, _ = QFileDialog.getOpenFileName(caption='Выберите список или шаблон', directory='',
                                                  options=QFileDialog.Options(),
                                                  filter="Список или шаблон (*list.xlsx *text.html);;"
                                                         "Список (*list.xlsx);;" 
                                                         "Шаблон (*text.html);;All Files (*)")
        if not filename:
            return
        self.listWidget_emails.clear()
        self.listWidget_attachments.clear()
        self.textBrowser.clear()
        os.chdir(os.path.dirname(filename))
        try:
            self.read_list_and_template(filename)
        except Exception as e:
            QMessageBox.warning(self.parent, 'Error', 'Ошибка: ' + str(e))
        self.fill_widgets_with_emails()

    def ask_login_and_create_connection(self):
        loginf = QDialog()
        diagui = ui_email_and_passw.Ui_Dialog()
        diagui.setupUi(loginf)

        last_frommail = keyring.get_password(KEYRING_SERVICE, LAST_FROMMAIL)
        last_fromname = keyring.get_password(KEYRING_SERVICE, LAST_FROMNAME)
        last_mailserver = keyring.get_password(KEYRING_SERVICE, LAST_MAILSERVER)
        last_copylist = keyring.get_password(KEYRING_SERVICE, LAST_COPYLIST)
        last_saveflag = keyring.get_password(KEYRING_SERVICE, LAST_SAVEFLAG)
        last_password = keyring.get_password(KEYRING_SERVICE, LAST_PASSWORD)
        diagui.line_email.setText(last_frommail or '')
        diagui.line_password.setText(last_password or '')
        diagui.line_sender.setText(last_fromname or '')
        diagui.line_smtpserver.setText(last_mailserver or 'smtp.googlemail.com')
        diagui.line_send_copy.setText(last_copylist or '')
        diagui.save_passw_cb.setCheckState([Qt.Unchecked, Qt.Checked][bool(last_saveflag)])
        if loginf.exec_() == QDialog.Accepted:
            last_frommail = diagui.line_email.text()
            last_password = diagui.line_password.text()
            last_fromname = diagui.line_sender.text()
            last_mailserver = diagui.line_smtpserver.text()
            last_copylist = diagui.line_send_copy.text()
            last_saveflag = '1' if diagui.save_passw_cb.checkState() else ''
            keyring.set_password(KEYRING_SERVICE, LAST_FROMMAIL, last_frommail)
            keyring.set_password(KEYRING_SERVICE, LAST_FROMNAME, last_fromname)
            keyring.set_password(KEYRING_SERVICE, LAST_MAILSERVER, last_mailserver)
            keyring.set_password(KEYRING_SERVICE, LAST_COPYLIST, last_copylist)
            keyring.set_password(KEYRING_SERVICE, LAST_SAVEFLAG, last_saveflag)
            keyring.set_password(KEYRING_SERVICE, LAST_PASSWORD, last_password if last_saveflag else '')
            envelope = email_stuff.EmailEnvelope(smtp_server=last_mailserver,
                                                 login=last_frommail,
                                                 password=last_password,
                                                 sender_addr=last_frommail,
                                                 sender_name=last_fromname,
                                                 copy_addrs=re.findall(EMAIL_REGEX, last_copylist))
            envelope.verify_credentials()
            return envelope
        else:
            raise ConnectionAbortedError('Необходимо ввести логин, пароль и т.д.')

    def start_threads_and_send_mails(self):
        mails_to_send = []
        for i in range(self.listWidget_emails.count()):
            item = self.listWidget_emails.item(i)
            if item.checkState():
                item.setSelected(True)
                xlsx_row = item.xlsx_row
                # Сохраняем номер строки, чтобы потом легко пометить её зелёным
                xlsx_row['QListWidgetIndex_WcCRve89'] = i
                mails_to_send.append(item.xlsx_row)
        if not mails_to_send:
            msg = 'Ни одно письмо для отправки не выбрано'
            QMessageBox.information(self.parent, 'OK', msg)
            raise Exception(msg)
        # Создаём по отправляльщику на каждый worker
        # Равномерно распределяем на них на всех почту
        self.USE_THREADS = min(self.NUM_THREADS, len(mails_to_send))
        envelopes = [self.envelope.copy() for __ in range(self.USE_THREADS)]
        for i, xlsx_row in enumerate(mails_to_send):
            envelopes[i % self.USE_THREADS].add_mail_to_queue(recipients=xlsx_row['email'],
                                                              subject=xlsx_row['subject'],
                                                              html=self.template.format(**xlsx_row),
                                                              files=xlsx_row['attach_list'],
                                                              xls_id=xlsx_row[files_parsers.ORIGINAL_ROW_NUM],
                                                              qt_id=xlsx_row['QListWidgetIndex_WcCRve89'])
        # Лочим кнопки
        self.pushButton_ask_and_send.setDisabled(True)
        self.pushButton_open_list_and_template.setDisabled(True)
        self.pushButton_cancel_send.setEnabled(True)
        # Готовим worker'ов
        self.__workers_done = 0
        self.__threads = []
        for idx in range(self.USE_THREADS):
            worker = Worker(idx, envelopes[idx])
            thread = QThread()
            thread.setObjectName('thread_' + str(idx))
            self.__threads.append((thread, worker))  # need to store worker too otherwise will be gc'd
            worker.moveToThread(thread)

            # get progress messages from worker:
            worker.sig_step.connect(self.on_worker_step)
            worker.sig_done.connect(self.on_worker_done)
            worker.sig_mail_sent.connect(self.on_mail_sent)
            worker.sig_mail_error.connect(self.on_mail_error)

            # control worker:
            self.sig_abort_workers.connect(worker.abort)

            # get read to start worker:
            thread.started.connect(worker.work)
            thread.start()  # this will emit 'started' and start thread's event loop

    def send_msg(self):
        # Перед отправкой должен быть закружен шаблон и список
        if not self.xlsx_rows_list or not self.template:
            QMessageBox.warning(self.listWidget_emails.parent(), 'Ошибка',
                                'Сначала нужно открыть шаблон и список рассылки')
            raise Exception()

        self.envelope = None
        try:
            self.envelope = self.ask_login_and_create_connection()
        except ConnectionAbortedError:
            pass
        except Exception as e:
            msg = 'Не могу подключиться к серверу: ' + str(e)
            QMessageBox.warning(self.listWidget_emails.parent(), 'Ошибка', msg)
            raise Exception(msg)
        else:
            self.start_threads_and_send_mails()


if __name__ == '__main__':
    batch_sender_app = QApplication(sys.argv)
    main_window = QMainWindow()
    ui = Extended_GUI(main_window)
    main_window.show()
    sys.exit(batch_sender_app.exec_())
