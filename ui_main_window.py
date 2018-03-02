# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def __init__(self):
        super().__init__()
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(691, 662)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout_4 = QtWidgets.QGridLayout()
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.pushButton_ask_and_send = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_ask_and_send.setObjectName("pushButton_ask_and_send")
        self.gridLayout_4.addWidget(self.pushButton_ask_and_send, 0, 0, 1, 1)
        self.pushButton_cancel_send = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_cancel_send.setCheckable(False)
        self.pushButton_cancel_send.setObjectName("pushButton_cancel_send")
        self.gridLayout_4.addWidget(self.pushButton_cancel_send, 0, 1, 1, 1)
        self.listWidget_emails = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget_emails.setObjectName("listWidget_emails")
        self.gridLayout_4.addWidget(self.listWidget_emails, 1, 0, 1, 2)
        self.gridLayout_2.addLayout(self.gridLayout_4, 0, 0, 1, 1)
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton_open_list_and_template = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_open_list_and_template.setObjectName("pushButton_open_list_and_template")
        self.gridLayout.addWidget(self.pushButton_open_list_and_template, 0, 0, 1, 1)
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(99)
        sizePolicy.setHeightForWidth(self.textBrowser.sizePolicy().hasHeightForWidth())
        self.textBrowser.setSizePolicy(sizePolicy)
        self.textBrowser.setObjectName("textBrowser")
        self.gridLayout.addWidget(self.textBrowser, 1, 0, 1, 1)
        self.listWidget_attachments = QtWidgets.QListWidget(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(1)
        sizePolicy.setHeightForWidth(self.listWidget_attachments.sizePolicy().hasHeightForWidth())
        self.listWidget_attachments.setSizePolicy(sizePolicy)
        self.listWidget_attachments.setMinimumSize(QtCore.QSize(0, 75))
        self.listWidget_attachments.setObjectName("listWidget_attachments")
        self.gridLayout.addWidget(self.listWidget_attachments, 2, 0, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout, 0, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 691, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Почтослатель"))
        self.pushButton_ask_and_send.setText(_translate("MainWindow", "(2) Отправить письма всем выделенным людям"))
        self.pushButton_ask_and_send.setDisabled(True)
        self.pushButton_cancel_send.setText(_translate("MainWindow", "Отмена"))
        self.pushButton_cancel_send.setDisabled(True)
        self.pushButton_open_list_and_template.setText(_translate("MainWindow", "(1) Открыть *list.xlsx или ***text.html"))

