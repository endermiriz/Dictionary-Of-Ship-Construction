from PyQt5.QtWidgets import*
from PyQt5.QtGui import QIcon
from dict_proje import Ui_MainWindow
import pandas as pd
import ctypes
from PyQt5.QtWidgets import *
from PyQt5.QtCore    import *
from PyQt5.QtGui     import *
import openpyxl

myappid = 'dictionary.of.shipbuilding.0.0.1.version.created.by.Ender.MIRIZ'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

kelimeler = []
anlamlar = []
class dictgui(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("Dictonary")
        self.setWindowIcon(QIcon("icon\logo.png"))
        self.setlist()
        self.updatekelimelist()
        self.settextfont()
        self.ui.pushButton.clicked.connect(self.clearline)
        self.ui.lineEdit.textChanged.connect(self.findkelime)
        self.ui.listewidget.itemClicked.connect(self.itemclicked)
    def settextfont(self):
        font = self.font()
        font.setFamily(u"MS Shell Dlg 2")
        font.setPixelSize(14)
        self.ui.textBrowser_2.setTextColor(QColor('#434343'))
        self.ui.textBrowser_2.setFont(font)

    def setlist(self):
        df = pd.read_excel('Data/dict_data.xlsx')
        df.Kelimeler = df.Kelimeler.str.capitalize()
        df.Anlamlar = df.Anlamlar.str.capitalize()
        df = df.sort_values(by=['Kelimeler', 'Anlamlar'])
        df = df.drop_duplicates(subset=['Kelimeler'], keep='last')
        df['Anlamlar'] = df["Kelimeler"] + "\n\n" + df['Anlamlar']
        df.to_excel('Data/dict_data_setting.xlsx',sheet_name="1", index=False)
    def updatekelimelist(self):
        data = pd.read_excel('Data/dict_data_setting.xlsx')
        kelimeler = data["Kelimeler"].tolist()
        for kelime in kelimeler:
            item = QListWidgetItem(kelime)
            self.ui.listewidget.addItem(item)
    def clearline(self):
        self.ui.lineEdit.clear()

    def findkelime(self):

        search_string = self.ui.lineEdit.text()
        match_items = self.ui.listewidget.findItems(search_string, Qt.MatchContains)
        for i in range(self.ui.listewidget.count()):
            it = self.ui.listewidget.item(i)
            it.setHidden(it not in match_items)
    def itemclicked(self,item):
        df = pd.read_excel('Data/dict_data_setting.xlsx')
        i = df.Kelimeler[df.Kelimeler == item.text()].index.tolist()
        i = str(i)
        i = i.replace("[","")
        irep = i.replace("]","")
        irep = int(irep)
        i2 = df.loc[(df.index == irep) & (df.Kelimeler == item.text()), 'Anlamlar'].tolist()
        try:

            for an in i2:
                self.ui.textBrowser_2.setText(an)
        except:
            self.ui.textBrowser_2.setText("")
            QMessageBox.information(self, "UYARI!", item.text()+" Kelimesinin anlamı kayıt edilmemiş!")

        # QMessageBox.information(self, "ListWidget", "You clicked: " +item.text())



uygulama = QApplication([])
pencere = dictgui()
pencere.show()
uygulama.exec_()