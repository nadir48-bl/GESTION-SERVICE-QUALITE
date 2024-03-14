import subprocess
import time
import docx
from PyQt6 import *
from PyQt6 import QtGui, QtWidgets, QtCore
from PyQt6.sip import wrappertype
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6 import QtPrintSupport
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewWidget, QPrinterInfo, QPrintPreviewDialog
import docx2pdf
from docxtpl import DocxTemplate
from docxtpl import *
import datetime
import sys
import os
from confirmité import Conformité_Window
from agréage import Agréage_Window
from moulin import Moulin_Window
from phytosanitaire import Phyto_Window
from refus import Refus_Window
from lgmsec import Stock_Legumesec
try:
    class Window_Ac(QObject):
        loaded = QtCore.pyqtSignal()
        try:
            def conformite_window(self):
                self.windowc = QtWidgets.QMainWindow()
                self.windowconformite = Conformité_Window()
                self.windowconformite.confi_window(self.windowc)
                self.windowcs.append(self.windowc)
                self.windowc.show()



        except Exception as e:
            print(e)
        def stock_legumesec (self):
            self.windowstock=QtWidgets.QMainWindow()
            self.windowstocklegum=Stock_Legumesec()
            self.windowstocklegum.stock_legumesec(self.windowstock)
            self.windowstock.show()

        def agréageWindow(self):
            self.windowa = QtWidgets.QMainWindow()
            self.windowagr = Agréage_Window()
            self.windowagr.agréage(self.windowa)
            self.windowa.show()



        def moulinWindow(self):
            self.windowb = QtWidgets.QMainWindow()
            self.windowm = Moulin_Window()
            self.windowm.mouli_window(self.windowb)
            self.windowbs.append(self.windowb)
            self.windowb.show()

        def phyto(self):
            self.windowph = QtWidgets.QMainWindow()
            self.windowphs.append(self.windowph)
            self.windowphy = Phyto_Window()
            self.windowphy.phyoto_produit(self.windowph)
            self.windowph.show()

        def refus(self):
            self.windowrf=QtWidgets.QMainWindow()
            self.windowrefus=Refus_Window()
            self.windowrefus.refus_produit(self.windowrf)
            self.windowrf.show()

        def __init__(self, parent=None):
            super().__init__(parent)
            self.windowbs=[]
            self.windowcs=[]
            self.windowphs=[]

            MainWindow.setObjectName("MainWindow")
            MainWindow.resize(1340, 800)
            MainWindow.setWindowTitle("CCLS RELIZANE SERVICE QUALITE")
            MainWindow.setWindowIcon(QIcon("images/Picsart_23-03-13_18-53-05-983.ico"))
            MainWindow.setStyleSheet("""
                QWidget 
                {      background: qlineargradient(x1: 1 y1:1, x2: 2, y2: 1,stop: 0 #ff8c00,stop: 1 #ffffff);
                     font-family: "Roboto 16", sans-serif,bold;
                }       

                QPushButton
                {
                    color:#ff8c00 ;
                    background-color: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646);
                    border-width: 1.3px;
                    border-color: #000000;
                    border-style: solid;
                    border-radius: 10px;
                    padding: 3px;
                    font-size: 11pt;
                    font-weight: bold;
                    padding-left: 5px;
                    padding-right: 5px;
                    min-width: 40px;

                }

                QPushButton:disabled
                {
                    background-color:#ff8c00;
                    border-width: 1px;
                    border-color: #454545;
                    border-style: solid;
                    padding-top: 5px;
                    padding-bottom: 5px;
                    padding-left: 10px;
                    padding-right: 10px;
                    border-radius: 8 px;
                    color: #454545;
                }

                QPushButton:focus {
                    background-color:QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646);
                    color:#ff8c00;
                }

                QPushButton:pressed
                {
                    background-color: qlineargradient(x1: 3, y1:0, x2: 1, y2: 0,stop: 0 #F0F8FF,stop: 1 #31363b);
                     color: #ff8cff;
                    padding-top: -15px;
                    padding-bottom: -17px;
                   


                }
                QPushButton:hover
                {
                    border: 1px solid #ff8cff;
                    color: white;
                }
                """)

            self.centralwidget = QtWidgets.QWidget(MainWindow)
            self.centralwidget.setObjectName("centralwidget")
            self.CCCLSTXT = QtWidgets.QLabel(self.centralwidget)
            self.CCCLSTXT.setGeometry(QtCore.QRect(0, 0, 1400, 100))
            self.CCCLSTXT.setStyleSheet(
                "background-color:QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646) ; color: #ff8c00 ;border: 0px solid #ff8c00 ;border-radius: 0px;padding: 2px")
            font = QtGui.QFont()
            font.setPointSize(12)
            font.setBold(True)
            font.setItalic(False)
            font.setWeight(75)
            self.CCCLSTXT.setFont(font)
            self.CCCLSTXT.setFrameShape(QtWidgets.QFrame.Shape.Box)
            self.CCCLSTXT.setFrameShadow(QtWidgets.QFrame.Shadow.Plain)
            self.CCCLSTXT.setLineWidth(2)
            self.CCCLSTXT.setMidLineWidth(0)
            self.CCCLSTXT.setTextFormat(QtCore.Qt.TextFormat.AutoText)
            self.CCCLSTXT.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.CCCLSTXT.setObjectName("CCCLSTXT")
            fontbtn = QtGui.QFont()
            fontbtn.setBold(True)
            fontbtn.setPointSize(14)

            font2 = QtGui.QFont()
            font2.setPointSize(26)
            font2.setBold(True)
            font2.setItalic(False)
            font2.setWeight(75)
            self.gsqr = QtWidgets.QLabel("<h1>G S Q R<h1/>", self.centralwidget)
            self.gsqr.setGeometry(QtCore.QRect(0, 618, 1400, 100))
            self.gsqr.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.gsqr.setStyleSheet(
                "background-color:QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646) ; color: #ff8c00 ;border: 0px solid #ff8c00 ;border-radius: 0px;padding: 0px")
            self.gsqr.setFont(font2)

            self.widget = QtWidgets.QWidget(self.centralwidget)
            self.widget.setGeometry(QtCore.QRect(430, 125, 550, 100))
            self.widget.setObjectName("widget")

            self.verticalLayout = QtWidgets.QVBoxLayout(self.widget)
            self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SizeConstraint.SetMinAndMaxSize)
            self.verticalLayout.setContentsMargins(0, 0, 0, 2)
            self.verticalLayout.setSpacing(10)
            self.verticalLayout.setObjectName("verticalLayout")




            self.conformitButton = QtWidgets.QPushButton(self.widget, clicked=lambda: self.conformite_window())
            self.conformitButton.setMinimumSize(QtCore.QSize(0, 60))
            self.conformitButton.setObjectName("conformitButton")
            self.conformitButton.setIcon(QIcon("images/confir.png"))
            self.conformitButton.setIconSize(QSize(53, 90))
            self.verticalLayout.addWidget(self.conformitButton)
            self.conformitButton.setFont(fontbtn)

            self.moulinButton = QtWidgets.QPushButton(self.widget, clicked=lambda: self.moulinWindow())
            self.moulinButton.setMinimumSize(QtCore.QSize(0, 60))
            self.verticalLayout.addWidget(self.moulinButton)
            self.moulinButton.setFont(fontbtn)
            self.moulinButton.setIcon(QIcon("images/mouln.png"))
            self.moulinButton.setIconSize(QSize(53, 90))

            self.agrageButton = QtWidgets.QPushButton(self.widget, clicked=lambda: self.agréageWindow())
            self.agrageButton.setMinimumSize(QtCore.QSize(0, 60))
            self.agrageButton.setObjectName("agrageButton")
            self.agrageButton.setIcon(QIcon("images/agreage.png"))
            self.agrageButton.setIconSize(QSize(53, 90))
            self.verticalLayout.addWidget(self.agrageButton)

            self.pushButton = QtWidgets.QPushButton(self.widget)
            self.pushButton.setMinimumSize(QtCore.QSize(0, 60))
            self.pushButton.setObjectName("pushButton")
            self.pushButton.setIcon(QIcon('images/preces.png'))
            self.pushButton.setIconSize(QSize(53, 90))
            self.verticalLayout.addWidget(self.pushButton)

            self.refusButton = QtWidgets.QPushButton(self.widget,clicked=lambda :self.refus())
            self.refusButton.setEnabled(True)
            self.refusButton.setMinimumSize(QtCore.QSize(0, 60))
            self.refusButton.setMaximumSize(QtCore.QSize(16777215, 16777215))
            self.refusButton.setBaseSize(QtCore.QSize(17, 0))
            self.refusButton.setIcon(QIcon("images/refus.png"))
            self.refusButton.setIconSize(QSize(53, 90))
            font = QtGui.QFont()
            font.setBold(True)
            font.setWeight(75)
            self.refusButton.setFont(font)
            self.refusButton.setObjectName("refusButton")
            self.verticalLayout.addWidget(self.refusButton)

            self.phytoButton = QtWidgets.QPushButton(self.widget, clicked=lambda: self.phyto())
            self.phytoButton.setMinimumSize(QtCore.QSize(0, 60))
            self.phytoButton.setObjectName("phytoButton")
            self.phytoButton.setIcon(QIcon("images/phyto (2).png"))
            self.phytoButton.setIconSize(QSize(53, 90))
            self.verticalLayout.addWidget(self.phytoButton)

            self.legumesecbutton = QtWidgets.QPushButton("GESTION STOCk LEGUMES SECS", self.widget,
                                                         clicked=lambda: self.stock_legumesec())
            self.legumesecbutton.setMinimumSize(QtCore.QSize(0, 60))
            self.legumesecbutton.setIcon(QIcon("images/lgmsc.png"))
            self.legumesecbutton.setIconSize(QSize(53, 90))
            self.verticalLayout.addWidget(self.legumesecbutton)
            self.legumesecbutton.setFont(fontbtn)

            MainWindow.setCentralWidget(self.centralwidget)
            self.menubar = QtWidgets.QMenuBar(MainWindow)
            self.menubar.setGeometry(QtCore.QRect(0, 0, 1187, 10))
            self.menubar.setObjectName("menubar")
            MainWindow.setMenuBar(self.menubar)
            self.statusbar = QtWidgets.QStatusBar(MainWindow)
            self.statusbar.setObjectName("statusbar")
            MainWindow.setStatusBar(self.statusbar)

            self.retranslateUi(MainWindow)
            QtCore.QMetaObject.connectSlotsByName(MainWindow)

        def retranslateUi(self, MainWindow):
            _translate = QtCore.QCoreApplication.translate
            self.CCCLSTXT.setText(_translate("MainWindow",
                                             "<h1>COOPERATIVE DES CEREALES ET LEGUMES SECSDE RELIZANE<h1/>\n<h1>SERVICE QUALITE<h1/>"))
            # self.SERVICETXT.setText(_translate("MainWindow", "<h1>SERVICE QUALITE<h1/>"))
            self.conformitButton.setText(_translate("MainWindow", "      BULLETIN DE CONFORMITE  "))
            self.moulinButton.setText(_translate("MainWindow", "      BULLETIN  MOULIN                  "))
            self.agrageButton.setText(_translate("MainWindow", "      BULLETIN D'AGREAGE            "))
            self.pushButton.setText(_translate("MainWindow", "      PROCES VERBAL                       "))
            self.refusButton.setText(_translate("MainWindow", "      LES REFUS                                   "))
            self.phytoButton.setText(_translate("MainWindow", "      PRODUITS PHYTOSANITAIRES"))

            # Emit the loaded signal after the ui is loaded
            self.loaded.emit()


    if __name__ == "__main__":
        app = QtWidgets.QApplication(sys.argv)
        pixmap = QPixmap("images/Red Black Bold Car Logo(3).png")
        splash = QSplashScreen(pixmap)
        splash.setEnabled(False)
        splash.show()
        MainWindow = QtWidgets.QMainWindow()
        windowspl = Window_Ac()
        windowspl.__init__(MainWindow)
        MainWindow.show()
        windowspl.loaded.connect(splash.close)
        MainWindow.hide()
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(MainWindow.show)
        timer.timeout.connect(splash.hide)
        timer.start(2000)
        sys.exit(app.exec())


except Exception as e:
    print(e)