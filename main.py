import sys
import ui2 as GUI
from PyQt5.Qt import *


app = QApplication(sys.argv)
w = QMainWindow()
ui = GUI.Ui_MainWindow(w)
w.show()
sys.exit(app.exec_())