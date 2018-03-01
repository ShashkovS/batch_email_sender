import sys
import gui_logic as GUI
import traceback
from PyQt5.Qt import QApplication, QMainWindow


def excepthook(excType, excValue, tracebackobj):
    traceback.print_tb(tracebackobj, excType, excValue)

sys.excepthook = excepthook


batch_sender_app = QApplication(sys.argv)
main_window = QMainWindow()
ui = GUI.Extended_GUI(main_window)
main_window.show()
sys.exit(batch_sender_app.exec_())