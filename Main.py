# Libraryies which is we will need - İhtiyacımız olacak kütüphaneler
import timeit
from logging import exception

import xlrd
import os
import sys
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import String, Column, Integer, schema
import json
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QProgressBar
from PyQt5 import uic
import easygui
from time import process_time_ns, sleep

# Lists and variables which is we will use - Kullanacağımız boş listeler
Lisans_Durumu = list()
Lisans_No = list()
Unvan = list()
Vergi_No = list()
Başlangıç_Tarihi = list()
Bitiş_Tarihi = list()
Tesis_Adresi = list()
İlçe = list()
İl = list()
Dağıtım_Şirketi = list()
Dağıtıcı_Tadil_Tarihi = list()
Alt_Başlığı = list()
Kategorisi = list()
İptal_Tarihi = list()
İptal_Açıklaması = list()

all_lists = [Lisans_Durumu, Lisans_No, Unvan, Vergi_No, Başlangıç_Tarihi, Bitiş_Tarihi, Tesis_Adresi,
             İlçe, İl, Dağıtım_Şirketi, Dağıtıcı_Tadil_Tarihi, Alt_Başlığı, Kategorisi, İptal_Tarihi,
             İptal_Açıklaması]


# codes for interface design - arayüz tasarımı için kodlar
class WelcomeScreen(QWidget):
    def __init__(self):
        super(WelcomeScreen, self).__init__()
        uic.loadUi("veri_aktarma.ui", self)
        self.text = ""

        # Getting data components from .UI to .PY - UI dosyasında Komponetleri Py Çekme
        self.fileSelect = self.findChild(QPushButton, "browse")
        self.ileri = self.findChild(QPushButton, "ileri")
        self.iptal = self.findChild(QPushButton, "iptal")
        self.path_f = self.findChild(QLineEdit, "path_file")
        self.bar = self.findChild(QProgressBar, "progressBar")
        self.bar2 = self.findChild(QProgressBar, "progressBar_2")

        # Going to the fucntion which is we want to go - Gitmek istenilen fonksiyona götürür
        self.fileSelect.clicked.connect(self.FileBrowser)
        self.ileri.clicked.connect(self.Start)
        self.iptal.clicked.connect(self.Cancel)

    # Browse file which is for .xls document to translate - Aktarılmak istenilen .xls dosyasının okunmasını sağlar
    def FileBrowser(self):
        path = easygui.fileopenbox(filetypes='.xls')
        self.path_f.setText(path)
        self.wb = xlrd.open_workbook(path)
        self.sheet = self.wb.sheet_by_index(0)

        Lisans_Durumu.clear()
        Lisans_No.clear()
        Unvan.clear()
        Vergi_No.clear()
        Başlangıç_Tarihi.clear()
        Bitiş_Tarihi.clear()
        Tesis_Adresi.clear()
        İlçe.clear()
        İl.clear()
        Dağıtım_Şirketi.clear()
        Dağıtıcı_Tadil_Tarihi.clear()
        Alt_Başlığı.clear()
        Kategorisi.clear()
        İptal_Tarihi.clear()
        İptal_Açıklaması.clear()

        start_time = process_time_ns()
        for i in range(self.sheet.nrows):
            Lisans_Durumu.append(self.sheet.cell_value(i, 0))
        for i in range(self.sheet.nrows):
            if i in Lisans_No:
                continue
            else:
                Lisans_No.append(self.sheet.cell_value(i, 1))
        for i in range(self.sheet.nrows):
            Unvan.append(self.sheet.cell_value(i, 2))
        for i in range(self.sheet.nrows):
            Vergi_No.append(self.sheet.cell_value(i, 3))
        for i in range(self.sheet.nrows):
            Başlangıç_Tarihi.append(self.sheet.cell_value(i, 4))
        for i in range(self.sheet.nrows):
            Bitiş_Tarihi.append(self.sheet.cell_value(i, 5))
        for i in range(self.sheet.nrows):
            Tesis_Adresi.append(self.sheet.cell_value(i, 6))
        for i in range(self.sheet.nrows):
            İlçe.append(self.sheet.cell_value(i, 7))
        for i in range(self.sheet.nrows):
            İl.append(self.sheet.cell_value(i, 8))
        for i in range(self.sheet.nrows):
            Dağıtım_Şirketi.append(self.sheet.cell_value(i, 9))
        for i in range(self.sheet.nrows):
            Dağıtıcı_Tadil_Tarihi.append(self.sheet.cell_value(i, 10))
        for i in range(self.sheet.nrows):
            Alt_Başlığı.append(self.sheet.cell_value(i, 11))
        for i in range(self.sheet.nrows):
            Kategorisi.append(self.sheet.cell_value(i, 12))
        for i in range(self.sheet.nrows):
            İptal_Tarihi.append(self.sheet.cell_value(i, 13))
        for i in range(self.sheet.nrows):
            İptal_Açıklaması.append(self.sheet.cell_value(i, 14))
        stop_time = process_time_ns()

        total_time = stop_time - start_time
        for i in range(101):
            sleep(total_time % 0.05)

            self.bar2.setValue(i)

    # starting point for transfer- Transferi başlatan fonksiyon
    def Start(self):
        start_time = process_time_ns()
        get_data()
        stop_time = process_time_ns()

        total_time = stop_time - start_time
        for i in range(101):
            sleep(total_time % 0.05)

            self.bar.setValue(i)
        easygui.msgbox(
            "Aktarma İşlemi tamamlandı. {} adet bayi datası aktarıldı".format(max(len(x) for x in all_lists)),
            title="Bayi Listesi Transfer")

    # interrupt point for transfer- Transferi kesen fonksiyon
    def Cancel(self):
        sys.exit()


# Connections which is we will need for Oracle - Oracle için gerekli bağlantılar
if getattr(sys, 'frozen', False):
    Current_Path = os.path.dirname(sys.executable)
else:
    Current_Path = str(os.path.dirname(__file__))

if Current_Path is None:
    print("Database Info File Not Found: config.json")
    easygui.msgbox("Database Info File Not Found: config.json", title="Bayi Listesi Transfer")
else:
    filePath = Current_Path + 'here' # config.json
    if os.path.exists(filePath):
        with open(filePath, 'r') as f:
            if f is not None:
                data = json.load(f)
    else:
        print("Database Config File Not Found: config.json")
        easygui.msgbox("Database Config File Not Found: config.json", title="Bayi Listesi Transfer")
        exit()

if data['user'] is None or data['password'] is None or data['host'] is None or data['database'] is None:
    print("Data eksiktir!")
    easygui.msgbox("Data eksiktir!", title="Bayi Listesi Transfer")
else:
    oracle_connection_string = 'here' % (data['user'], data['password'], data['host'],
                                                                   data['database'])  # Connection link for Oracle

app = Flask(__name__)

oracle_schema = 'here' #key for oracle
foreign_key_prefix = oracle_schema + '.'

oracle_db_metadata = schema.MetaData(schema=oracle_schema)
db = SQLAlchemy(app, metadata=oracle_db_metadata)

app.config['SQLALCHEMY_DATABASE_URI'] = oracle_connection_string
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.debug = False


class Bayi_Listesi(db.Model):
    __tablename__ = "here" # table name
    FNM06_ID = Column(Integer, primary_key=True)
    LISANS_DURUMU = Column(String, nullable=True)
    LISANS_NO = Column(String, nullable=True)
    UNVAN = Column(String, nullable=True)
    VERGI_NO = Column(Integer, nullable=True)
    BASLANGIC_TARIHI = Column(String, nullable=True)
    BITIS_TARIHI = Column(String, nullable=True)
    TESIS_ADRESI = Column(String, nullable=True)
    ILCE = Column(String, nullable=True)
    IL = Column(String, nullable=True)
    DAGITIM_SIRKETI = Column(String, nullable=True)
    DAGITICI_TADIL_TARIHI = Column(String, nullable=True)
    ALT_BASLIGI = Column(String, nullable=True)
    KATEGORI = Column(String, nullable=True)
    IPTAL_TARIHI = Column(String, nullable=True)
    IPTAL_ACIKLAMA = Column(String, nullable=True)


# Finding the longest list amoung lists for get data() - get data() için listeler arasında en uzun listeyi bulma
def FindMaxLength(lst):
    maxLength = max(len(x) for x in lst)
    return maxLength


# Database Get User Data
def databaseGetUserData(lisansno):
    result = Bayi_Listesi.query.filter_by(LISANS_NO=lisansno).first()
    return result


# Reading datas from .xls doccument - .xls dosyasından dataları okuma
def get_data():
    for i in range(1, FindMaxLength(all_lists)):
        rs = databaseGetUserData(str(Lisans_No[i]))
        if rs is None:
            new_rec = Bayi_Listesi(
                LISANS_DURUMU=str(Lisans_Durumu[i]),
                LISANS_NO=str(Lisans_No[i]),
                UNVAN=str(Unvan[i]),
                VERGI_NO=str(Vergi_No[i]),
                BASLANGIC_TARIHI=str(Başlangıç_Tarihi[i]),
                BITIS_TARIHI=str(Bitiş_Tarihi[i]),
                TESIS_ADRESI=str(Tesis_Adresi[i]),
                ILCE=str(İlçe[i]),
                IL=str(İl[i]),
                DAGITIM_SIRKETI=str(Dağıtım_Şirketi[i]),
                DAGITICI_TADIL_TARIHI=str(Dağıtıcı_Tadil_Tarihi[i]),
                ALT_BASLIGI=str(Alt_Başlığı[i]),
                KATEGORI=str(Kategorisi[i]),
                IPTAL_TARIHI=str(İptal_Tarihi[i]),
                IPTAL_ACIKLAMA=str(İptal_Açıklaması[i]))

            db.session.add(new_rec)
            db.session.commit()
        else:
            lisansNo = databaseGetUserData(str(Lisans_No[i])).LISANS_NO
            Bayi_Listesi.query.filter_by(LISANS_NO=lisansNo).update(
                dict(LISANS_DURUMU=str(Lisans_Durumu[i]),
                     LISANS_NO=str(Lisans_No[i]),
                     UNVAN=str(Unvan[i]),
                     VERGI_NO=str(Vergi_No[i]),
                     BASLANGIC_TARIHI=str(Başlangıç_Tarihi[i]),
                     BITIS_TARIHI=str(Bitiş_Tarihi[i]),
                     TESIS_ADRESI=str(Tesis_Adresi[i]),
                     ILCE=str(İlçe[i]),
                     IL=str(İl[i]),
                     DAGITIM_SIRKETI=str(Dağıtım_Şirketi[i]),
                     DAGITICI_TADIL_TARIHI=str(Dağıtıcı_Tadil_Tarihi[i]),
                     ALT_BASLIGI=str(Alt_Başlığı[i]),
                     KATEGORI=str(Kategorisi[i]),
                     IPTAL_TARIHI=str(İptal_Tarihi[i]),
                     IPTAL_ACIKLAMA=str(İptal_Açıklaması[i])))
            db.session.commit()


# Settings for window - Pencere için ayarlar
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
