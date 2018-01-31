import sys
import uiext as GUI
import traceback
from PyQt5.Qt import *


def excepthook(excType, excValue, tracebackobj):
    traceback.print_tb(tracebackobj, excType, excValue)


sys.excepthook = excepthook

app = QApplication(sys.argv)
w = QMainWindow()
ui = GUI.Extended_GUI(w)
w.show()
sys.exit(app.exec_())