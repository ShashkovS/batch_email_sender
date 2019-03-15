import smtplib
import ssl
import queue
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate, formataddr
from email.header import Header
from email import encoders
from os.path import basename
from typing import List
# For guessing MIME type based on file name extension
import mimetypes


class EmailEnvelope:
    def __init__(self, smtp_server, login, password, sender_addr, sender_name='', copy_addrs=None):
        """Сохраняем всем параметры в атрибутах. Больше ничего не делаем"""
        self.smtp_server = smtp_server
        self.login = login
        self.password = password
        self.sender_addr = sender_addr
        self.sender_name = sender_name
        self.copy_addrs = copy_addrs or []
        self.smtp = None
        self.send_queue = queue.Queue()
        self.__abort = False
        self.context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)

    def connect_to_server(self):
        """Подключаемся к серверу"""
        if self.smtp is None:
            self.smtp = smtplib.SMTP_SSL(self.smtp_server, port=465)
            # self.smtp.set_debuglevel(1)
            self.smtp.ehlo_or_helo_if_needed()
            self.smtp.login(self.login, self.password)
        # Проверяем подключение
        try:
            status = self.smtp.noop()[0]
        except smtplib.SMTPServerDisconnected as e:
            status = -1
        # Если не ОК, то переподключаемся
        if status != 250:
            self.smtp.connect(host=self.smtp_server, port=587)
            self.smtp.ehlo_or_helo_if_needed()
            self.smtp.starttls(context=self.context)
            self.smtp.ehlo_or_helo_if_needed()
            self.smtp.login(self.login, self.password)

    def verify_credentials(self):
        self.connect_to_server()
        try:
            status, _ = self.smtp.noop()
        except smtplib.SMTPServerDisconnected as e:
            status = -1
        if status != 250:
            raise Exception('Не удалось подключиться для отправки почты с адреса ' + self.login)

    def add_mail_to_queue(self, recipients: List[str], subject, html, files=None, xls_id=None, qt_id=None):
        msg = MIMEMultipart()
        msg['From'] = formataddr((Header(self.sender_name, 'utf-8').encode(), self.sender_addr))
        msg['To'] = COMMASPACE.join(recipients)
        msg['Cc'] = COMMASPACE.join(self.copy_addrs)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = Header(subject, 'utf-8')
        msg.attach(MIMEText(html.encode('utf-8'), 'html', 'utf-8'))
        for file in files or []:
            if not file:
                continue
            ctype, encoding = mimetypes.guess_type(file)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            part = MIMEBase(maintype, subtype)
            with open(file, "rb") as attachment_file:
                part.set_payload(attachment_file.read())
            part.add_header('Content-Disposition', 'attachment', filename=basename(file))
            encoders.encode_base64(part)
            msg.attach(part)
        mail = dict(from_addr=self.sender_addr,
                    to_addrs=recipients + self.copy_addrs,
                    msg=msg.as_string(),
                    xls_id=xls_id,
                    qt_id=qt_id)
        self.send_queue.put(mail)


    def take_next_mail(self):
        try:
            mail = self.send_queue.get(block=False)
        except queue.Empty as e:
            raise StopIteration
        return mail

    def send_next(self, mail):
        self.connect_to_server()
        if self.__abort:
            raise StopIteration
        senderrs = self.smtp.sendmail(from_addr=mail['from_addr'], to_addrs=mail['to_addrs'], msg=mail['msg'])
        return mail

    def __copy__(self):
        return self.__class__(smtp_server=self.smtp_server, login=self.login, password=self.password,
                              sender_addr=self.sender_addr, sender_name=self.sender_name, copy_addrs=self.copy_addrs)

    def copy(self):
        return self.__copy__()

# password = input('Password?:\n')
# sender = EmailEnvelope(smtp_server='smtp.googlemail.com', login='shashkov@179.ru', password=password,
#                        sender_addr='shashkov@179.ru', sender_name='Сергей Шашков',
#                        copy_addrs=['sh57+copy@ya.ru'])
# sender.add_mail_to_queue(recipients=['asdfasdfasdfasdfasdfasdfsh57+t1@ya.ru'],
#                          subject='Тестовая отправка',
#                          html='<h1>Успех?</h1><p>Определённо</p>', files=['sob1.pdf'])
# sender.add_mail_to_queue(recipients=['asdfasdfasdfasdfasdfasdfasdfsh57+t2@ya.ru'],
#                          subject='Тестовая отправка2',
#                          html='<h1>Успех2?</h1><p>Определённо2</p>', files=None)
# sender.send_all()
