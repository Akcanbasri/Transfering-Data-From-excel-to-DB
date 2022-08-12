import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication, QWidget, QPushButton, QLineEdit
from PyQt5 import uic
import easygui



class WelcomeScreen(QDialog):
    def __init__(self):
        super(WelcomeScreen, self).__init__()
        uic.loadUi("veri_aktarma.ui", self)
        self.text = ""

        # UI dosyasında Komponetleri Py Çekme
        self.fileSelect = self.findChild(QPushButton, "browse")
        self.ileri = self.findChild(QPushButton, "ileri")
        self.iptal = self.findChild(QPushButton, "iptal")
        self.path_f = self.findChild(QLineEdit, "path_file")

        self.fileSelect.clicked.connect(self.FileBrowser)
        self.ileri.clicked.connect(self.Start)
        self.iptal.clicked.connect(self.Cancel)


    def FileBrowser(self):
        path = easygui.fileopenbox(filetypes='.xls ')
        self.path_f.setText(path)

    def Start(self):
        pass

    def Cancel(self):
        sys.exit()




app = QApplication(sys.argv)
welcome = WelcomeScreen()
widget = QtWidgets.QStackedWidget()
widget.addWidget(welcome)
widget.setFixedHeight(220)
widget.setFixedWidth(410)

widget.show()
try:
    sys.exit(app.exec_())
except:
    print("Exiting")
