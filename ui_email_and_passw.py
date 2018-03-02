# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'login_form.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(400, 147)
        self.gridLayout = QtWidgets.QGridLayout(Dialog)
        self.gridLayout.setObjectName("gridLayout")
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setObjectName("formLayout")
        self.line_email = QtWidgets.QLineEdit(Dialog)
        self.line_email.setObjectName("line_email")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.line_email)
        self.label_1 = QtWidgets.QLabel(Dialog)
        self.label_1.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.label_1.setObjectName("label_1")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_1)
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_2)
        self.line_password = QtWidgets.QLineEdit(Dialog)
        self.line_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.line_password.setObjectName("line_password")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.line_password)
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setObjectName("label_3")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_3)
        self.line_sender = QtWidgets.QLineEdit(Dialog)
        self.line_sender.setObjectName("line_sender")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.line_sender)
        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setObjectName("label_4")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.label_4)
        self.line_smtpserver = QtWidgets.QLineEdit(Dialog)
        self.line_smtpserver.setObjectName("line_smtpserver")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.line_smtpserver)

        self.label_5 = QtWidgets.QLabel(Dialog)
        self.label_5.setObjectName("label_5")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.label_5)
        self.line_send_copy = QtWidgets.QLineEdit(Dialog)
        self.line_send_copy.setObjectName("line_send_copy")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.line_send_copy)


        self.label_6 = QtWidgets.QLabel(Dialog)
        self.label_6.setObjectName("label_6")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.LabelRole, self.label_6)

        self.save_passw_cb = QtWidgets.QCheckBox(Dialog)
        self.save_passw_cb.setObjectName("save_passw_cb")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.FieldRole, self.save_passw_cb)

        self.gridLayout.addLayout(self.formLayout, 0, 0, 1, 1)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setCenterButtons(True)
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayout.addWidget(self.buttonBox, 1, 0, 1, 1)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Введите email, пароль и прочее"))
        self.label_1.setText(_translate("Dialog", "Логин (email)"))
        self.label_2.setText(_translate("Dialog", "Пароль"))
        self.label_3.setText(_translate("Dialog", "Отправитель"))
        self.label_4.setText(_translate("Dialog", "SMTP сервер"))
        self.label_5.setText(_translate("Dialog", "Поставить в копию"))
        self.label_6.setText(_translate("Dialog", "Сохранить пароль"))
