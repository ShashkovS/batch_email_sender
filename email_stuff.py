import smtplib
import queue
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, formataddr
from email.header import Header
from os.path import basename
from typing import List

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

    def connect_to_server(self):
        """Подключаемся к серверу"""
        if self.smtp is None:
            self.smtp = smtplib.SMTP()
            self.smtp.set_debuglevel(1)
        # Проверяем подключение
        try:
            status = self.smtp.noop()[0]
        except smtplib.SMTPServerDisconnected as e:
            status = -1
        if status != 250:
            self.smtp.connect(host=self.smtp_server)
            self.smtp.ehlo_or_helo_if_needed()
            self.smtp.starttls()
            self.smtp.login(self.login, self.password)

    def verify_credentials(self):
        smtp = smtplib.SMTP()
        smtp.connect(host=self.smtp_server)
        smtp.ehlo_or_helo_if_needed()
        smtp.starttls()
        smtp.login(self.login, self.password)
        try:
            status, _ = smtp.noop()
            smtp.quit()
        except smtplib.SMTPServerDisconnected as e:
            status = -1
        if status != 250:
            raise Exception(f'Не удалось подключиться для отправки почты с адреса {self.login}')

    def add_mail_to_queue(self, recipients: List[str], subject, html, files=None, id=None):
        msg = MIMEMultipart()
        msg['From'] = formataddr((Header(self.sender_name, 'utf-8').encode(), self.sender_addr))
        msg['To'] = COMMASPACE.join(recipients)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = Header(subject, 'utf-8')
        msg.attach(MIMEText(html.encode('utf-8'), 'html', 'utf-8'))
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
            msg.attach(part)
        mail = dict(from_addr=self.sender_addr, to_addrs=recipients, msg=msg.as_string(), id=id)
        self.send_queue.put(mail)

    def send_all(self):
        while True:
            try:
                mail = self.send_queue.get(block=False)
            except queue.Empty as e:
                return
            self.connect_to_server()
            self.smtp.sendmail(**mail)

# password = input('Password?:\n')
# sender = EmailEnvelope(smtp_server='smtp.googlemail.com', login='shashkov@179.ru', password=password,
#                        sender_addr='shashkov@179.ru', sender_name='Сергей Шашков')
# sender.add_mail_to_queue(recipients=['sh57+t1@ya.ru'],
#                          subject='Тестовая отправка',
#                          html='<h1>Успех?</h1><p>Определённо</p>', files=['sob1.pdf'])
# sender.add_mail_to_queue(recipients=['sh57+t2@ya.ru'],
#                          subject='Тестовая отправка2',
#                          html='<h1>Успех2?</h1><p>Определённо2</p>', files=None)
# sender.send_all()
