import sys
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QTimer
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QWidget, QMessageBox
from PyQt6.QtGui import QIcon, QPixmap
import socket
import subprocess
import mysql.connector

from PyQt6 import QtWidgets

try:
    class LoginUI(object):
        def __init__(self, MainWindow):
            super().__init__()

            MainWindow.setWindowTitle("Se connecter à GSQR")
            MainWindow.setGeometry(500, 100, 700, 400)
            MainWindow.setFixedSize(800, 500)
            MainWindow.setStyleSheet("""QMainWindow {
    background-color:#ffffff;
}

QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QTimeEdit, QDateEdit, QDateTimeEdit {
    border-width: 2px;
    border-radius: 4px;
    border-style: solid;
    border-top-color: qlineargradient(spread:pad, x1:0.5, y1:1, x2:0.5, y2:0, stop:0 #c1c9cf, stop:1 #d2d8dd);
    border-right-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #c1c9cf, stop:1 #d2d8dd);
    border-bottom-color: qlineargradient(spread:pad, x1:0.5, y1:0, x2:0.5, y2:1, stop:0 #c1c9cf, stop:1 #d2d8dd);
    border-left-color: qlineargradient(spread:pad, x1:1, y1:0, x2:0, y2:0, stop:0 #c1c9cf, stop:1 #d2d8dd);
    background-color: #f4f4f4;
    color: #3d3d3d;
}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QTimeEdit:focus, QDateEdit:focus, QDateTimeEdit:focus {
    border-width: 2px;
    border-radius: 4px;
    border-style: solid;
    border-top-color: qlineargradient(spread:pad, x1:0.5, y1:1, x2:0.5, y2:0, stop:0 #85b7e3, stop:1 #9ec1db);
    border-right-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #85b7e3, stop:1 #9ec1db);
    border-bottom-color: qlineargradient(spread:pad, x1:0.5, y1:0, x2:0.5, y2:1, stop:0 #85b7e3, stop:1 #9ec1db);
    border-left-color: qlineargradient(spread:pad, x1:1, y1:0, x2:0, y2:0, stop:0 #85b7e3, stop:1 #9ec1db);
    background-color: #f4f4f4;
    color: #3d3d3d;
}
QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QTimeEdit:disabled, QDateEdit:disabled, QDateTimeEdit:disabled {
    color: #b9b9b9;
}

QLabel, QCheckBox, QRadioButton {
    color: #272727;
}

QTabWidget {
    color:rgb(0,0,0);
    background-color:#000000;
}
QTabWidget::pane {
    border-color: #050a0e;
    background-color: #e0e0e0;
    border-width: 1px;
    border-radius: 4px;
    position: absolute;
    top: -0.5em;
    padding-top: 0.5em;
}

QTabWidget::tab-bar {
    alignment: center;
}
QPushButton
        {
            color: #ffffff;
            background-color:#00ADFC;
            border-width: 1px;
            border-color: #ffffff;
            border-style: solid;
            border-radius: 6;
            padding: 3px;
            font-size: 12px;
            padding-left: 5px;
            padding-right: 5px;
            min-width: 40px;

        }

        QPushButton:disabled
        {
            background-color:#03ecff;
            border-width: 1px;
            border-color: #454545;
            border-style: solid;
            padding-top: 5px;
            padding-bottom: 5px;
            padding-left: 10px;
            padding-right: 10px;
            border-radius: 2px;
            color: #454545;
        }
        QPushButton:pressed
        {
            background-color: #3daee9;
            padding-top: -15px;
            padding-bottom: -17px;
            color:#ffffff;
            border-color: #454545;
        }

""")
            MainWindow.setWindowIcon(QIcon("images/Picsart_23-03-13_18-53-05-983.ico"))
            

            image_label1 = QLabel(MainWindow)
            image_label1.setGeometry(50, -20, 400, 200)
            # Load the image
            image_path1 = "images/_logogsqr2.png"
            image1 = QPixmap(image_path1)
            image_label1.setPixmap(image1)
            image_label1.setScaledContents(True)

            image_label = QLabel(MainWindow)
            image_label.setGeometry(340, 0, 470, 500)
            #Load the image
            image_path = "images/_ab3.png"
            image = QPixmap(image_path)
            image_label.setPixmap(image)
            image_label.setScaledContents(True)



            username_label = QLabel("Nom d'utilisateur:", MainWindow)
            username_label.setStyleSheet("font-size: 14px; color: #555;")
            username_label.setGeometry(60, 160, 150, 30)

            self.username_input = QLineEdit(MainWindow)
            self.username_input.setStyleSheet(
                "font-size: 16px; padding: 8px; border: 1px solid #ddd; border-radius: 5px;")
            self.username_input.setGeometry(60, 190, 360, 40)
            self.username_input.setClearButtonEnabled(True)  # Enable clear button

            # Add icon to the left of the QLineEdit
            icon_user = QPixmap("path_to_username_icon.png")  # Replace with the path to your username icon
            self.username_input.setStyleSheet(
                f"background-image: url({icon_user}); background-position: left center; background-repeat: no-repeat; padding-left: 40px;")

            password_label = QLabel("Mot de passe:", MainWindow)
            password_label.setStyleSheet("font-size: 14px; color: #555;")
            password_label.setGeometry(60, 240, 150, 30)
            self.failed = QtWidgets.QLabel(MainWindow)
            self.failed.setGeometry(100, 400, 370, 35)
            self.failedconection = QtWidgets.QLabel(MainWindow)
            self.failedconection.setGeometry(140, 420, 370, 35)
            self.password_input = QLineEdit(MainWindow)
            self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.password_input.setStyleSheet("font-size: 16px; padding: 8px; border: 1px solid #ddd; border-radius: 5px;")
            self.password_input.setGeometry(60, 270, 360, 40)
            self.password_input.setClearButtonEnabled(True)  # Enable clear button

            # Add icon to the left of the QLineEdit
            icon_password = QPixmap("path_to_password_icon.png")  # Replace with the path to your password icon
            self.password_input.setStyleSheet(
                f"background-image: url({icon_password}); background-position: left center; background-repeat: no-repeat; padding-left: 40px;")

            login_button = QPushButton("Login",MainWindow)
            login_button.clicked.connect(self.on_login_clicked)
            login_button.setGeometry(130, 350, 230, 45)


            self.check_connection()
            self.is_connected()



        def on_login_clicked(self):
            try:
                self.username = self.username_input.text()
                self.password = self.password_input.text()
                userslist=["nadir","qualite"]
                passworlist=["Nadir206@","qualite48"]
                # Add your login logic here (validate username and password)
                if self.username in userslist and self.password in passworlist:
                    # Connecting to MySQL database
                    databaseuser = mysql.connector.connect(user=self.username, host='localhost', password=self.password)
                    database = mysql.connector.connect(
                        host="localhost",
                        user=self.username,
                        password=self.password
                    )

                    curs = database.cursor()

                    # Create the database if it doesn't exist

                    curs.execute("USE datta_legumesec_entry")
                    # Connection for outtable
                    database1 = mysql.connector.connect(
                        host='localhost',
                        user=self.username,
                        password=self.password
                    )
                    curs1 = database1.cursor()
                    # Create the database if it doesn't exist
                    curs1.execute("USE datta_legumsec_out")
                    # Creating database if not exists
                    database.commit()
                    database.close()
                    MainWindow.close()
                    self.RunWindow()
                elif self.username == "" and self.password == "":
                    self.username_input.setPlaceholderText("Entrez le nom d'utilisateur")
                    self.username_input.setStyleSheet("color:red;")
                    self.password_input.setPlaceholderText("Entrez le mot de passe")
                    self.password_input.setStyleSheet("color:red;")
                    self.timer = QTimer()
                    self.timer.timeout.connect(lambda:self.username_input.setStyleSheet("color:#000000;"))
                    self.timer.start(300)
                    self.timer.timeout.connect(lambda: self.password_input.setStyleSheet("color:#000000;"))
                    self.timer.start(300)
                else:
                    self.failed.setText("le nom d'utilisateur ou le mot de passe est incorrect !")

            except Exception as e:
                print(e)



        def check_connection(self):
            while not self.is_connected():
                self.failedconection.setText("Aucune connexion Internet disponible")
                print("Connection Check", "No internet connection available")
                if  self.is_connected():
                    self.failedconection.setText("")
                    print("Connection Check", "Internet connection is available.")
                break
        def is_connected(self):
            try:
                # Check if there's a valid internet connection by trying to resolve a well-known host (e.g., Google's DNS)
                socket.create_connection(("8.8.8.8", 53), timeout=3)
                return True
            except OSError:
                return False

        def RunWindow(self):
            window_ac = "GSQR.pyw"
            subprocess.run(['python',window_ac])


    if __name__ == "__main__":
        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        window = LoginUI(MainWindow)
        MainWindow.show()
        sys.exit(app.exec())

except Exception as e:
    print(e)
