from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from main import run


class Ui_Dialog(object):

    def __init__(self):
        self.pathppt = ''
        self.pathdata = ''

    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.setEnabled(True)
        Dialog.resize(399, 300)
        Dialog.setAutoFillBackground(False)
        Dialog.setStyleSheet("background-color: rgb(255, 255, 255)")
        self.pptButton = QtWidgets.QPushButton(Dialog)
        self.pptButton.setGeometry(QtCore.QRect(250, 110, 93, 28))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        self.pptButton.setFont(font)
        self.pptButton.setObjectName("pptButton")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(0, 10, 401, 51))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(51, 148, 68, 28))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.dataButton = QtWidgets.QPushButton(Dialog)
        self.dataButton.setGeometry(QtCore.QRect(250, 150, 93, 28))
        self.dataButton.setStyleSheet("background-color: rgb(255, 255, 255)")
        self.dataButton.setObjectName("dataButton")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(52, 112, 226, 28))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.executarButton = QtWidgets.QPushButton(Dialog)
        self.executarButton.setGeometry(QtCore.QRect(160, 230, 93, 28))
        self.executarButton.setObjectName("executarButton")
        self.label_ppt = QtWidgets.QLabel(Dialog)
        self.label_ppt.setGeometry(QtCore.QRect(350, 120, 55, 16))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setBold(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label_ppt.setFont(font)
        self.label_ppt.setStyleSheet("font-color: rgb(42, 199, 18)")
        self.label_ppt.setObjectName("label_ppt")
        self.label_data = QtWidgets.QLabel(Dialog)
        self.label_data.setGeometry(QtCore.QRect(350, 160, 55, 16))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setBold(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.label_data.setFont(font)
        self.label_data.setObjectName("label_data")
        self.label_3.raise_()
        self.label_2.raise_()
        self.dataButton.raise_()
        self.label.raise_()
        self.pptButton.raise_()
        self.executarButton.raise_()
        self.label_ppt.raise_()
        self.label_data.raise_()

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Atualizador de Slides"))
        self.pptButton.setText(_translate("Dialog", "Procurar"))
        self.label_3.setText(_translate("Dialog", "Atualizador de Slides"))
        self.label_2.setText(_translate("Dialog", "Dados:"))
        self.dataButton.setText(_translate("Dialog", "Procurar"))
        self.label.setText(_translate("Dialog", "Power Point Template:"))
        self.executarButton.setText(_translate("Dialog", "Executar"))
        self.label_ppt.setText(_translate("Dialog", ""))
        self.label_data.setText(_translate("Dialog", ""))
        self.pptButton.clicked.connect(self.open_dialog_box_ppt)
        self.dataButton.clicked.connect(self.open_dialog_box_data)
        self.executarButton.clicked.connect(self.execute)

    def open_dialog_box_ppt(self):
        filename = QFileDialog.getOpenFileName()
        path = filename[0]
        self.label_ppt.setText("Ok")
        self.pathppt = path

    def open_dialog_box_data(self):
        filename = QFileDialog.getOpenFileName()
        path = filename[0]
        self.label_data.setText("Ok")
        self.pathdata = path

    def execute(self):
        if not self.pathdata == '' or self.pathppt == '':
            run(self.pathppt, self.pathdata)

        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Selecione os arquivos")
            msg.exec_()
