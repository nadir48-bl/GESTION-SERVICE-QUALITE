import os
import subprocess
import tempfile
from docx2pdf import convert
from PyQt6 import QtCore, QtGui, QtWidgets
import docx
from PyQt6 import *
import concurrent.futures
from PyQt6.sip import wrappertype
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6 import QtPrintSupport
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewWidget, QPrinterInfo, QPrintPreviewDialog
from docxtpl import DocxTemplate
from docxtpl import *
import datetime
import sys


try:
    class Conformité_Window(object):
        ##############################################buttuons&def####################
        def clear_all(self):
            self.vpsin.clear()
            self.vhumiditein.clear()
            self.vtotalpremin.clear()
            self.vgraincasséin.clear()
            self.vtotaldemin.clear()
            self.vgrainetrangéin.clear()
            self.vgrainmouchtéin.clear()
            self.vgrainmaigrein.clear()
            self.vgrainechaudein.clear()
            self.vdébrisvéin.clear()
            self.vmatiéreinrtin.clear()
            self.vdébrisvéin.clear()
            self.vmatiéreinrtin.clear()
            self.vgrainsanvaleurin.clear()
            self.vgrainchauféin.clear()
            self.vgraigerméin.clear()
            self.vgrainpunaiséin.clear()
            self.vgrainpiquéin.clear()
            self.vgrainboutefin.clear()
            self.vgraincarréin.clear()
            self.vautrecéréalin.clear()
            self.vautrecéréalin.setValue(0)
            self.vgraigerméin.setValue(0)
            self.vgrainpunaiséin.setValue(0)
            self.vgrainpiquéin.setValue(0)
            self.vgrainboutefin.setValue(0)
            self.vgraincarréin.setValue(0)
            self.vpsin.setValue(0)
            self.vhumiditein.setValue(0)
            self.vtotalpremin.setValue(0)
            self.vgraincasséin.setValue(0)
            self.vtotaldemin.setValue(0)
            self.vgrainetrangéin.setValue(0)
            self.vgrainmouchtéin.setValue(0)
            self.vgrainmaigrein.setValue(0)
            self.vgrainechaudein.setValue(0)
            self.vdébrisvéin.setValue(0)
            self.vmatiéreinrtin.setValue(0)
            self.vdébrisvéin.setValue(0)
            self.vmatiéreinrtin.setValue(0)
            self.vgrainsanvaleurin.setValue(0)
            self.vgrainchauféin.setValue(0)
            self.vtotalpremin.setValue(0)

        def plus_value(self):
            e = self.vgraincasséin.value()
            f = self.vgrainmaigrein.value()
            g = self.vgrainechaudein.value()
            h = self.vgraigerméin.value()
            i = self.vgrainpunaiséin.value()
            j = self.vgrainpiquéin.value()
            k = self.vgrainboutefin.value()
            m = self.vgrainmouchtéin.value()
            n = self.vgrainetrangéin.value()
            # total 2eme cat blé dur
            vtotaldemm = e + f + g + h + i + j + k + m + n
            self.vtotaldemin.setValue(vtotaldemm)
            # total 1er
            p = self.vdébrisvéin.value()
            q = self.vmatiéreinrtin.value()
            r = self.vgrainnuisiblein.value()
            s = self.vgrainsanvaleurin.value()
            v = self.vgrainchauféin.value()
            w = self.vgraincarréin.value()
            self.vtotalpremin.setValue(p + q + r + s + v + w)

        def save_doc(self):
            self.doc = DocxTemplate("Docxfiles/_bulletin entré/template_BULLETIN_ENTRE.docx")
            # matri=[self.matricullist.item(x).text() for x in range(self.matricullist.count())]
            # matriadd=("\n".join(matri))
            matriadd = self.matricullist.toPlainText()
            ebps = self.vpsin.text()
            provance = self.provancetecombo.currentText()
            lieu = self.lieucombo.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantite.text()
            matierinirt = self.vmatiéreinrtin.text()
            debrisv = self.vdébrisvéin.text()
            grainnuisible = self.vgrainnuisiblein.text()
            grainsansvaleur = self.vgrainsanvaleurin.text()
            agrcm = self.agréeeurcombo.currentText()
            tnvv = self.vhumiditein.text()
            totap = self.vtotalpremin.text()
            datte = self.dateedite.text()
            graincasse = self.vgraincasséin.text()
            grainmaigre = self.vgrainmaigrein.text()
            grainmoushte = self.vgrainmouchtéin.text()
            grainpunaise = self.vgrainpunaiséin.text()
            grainpique = self.vgrainpiquéin.text()
            grainboute = self.vgrainpiquéin.text()
            totad = self.vtotaldemin.text()
            autrecereal = self.vautrecéréalin.text()
            bledur = self.vgrainetrangéin.text()
            # self.doc.render({"mat":self.items,"ps": ebps,"uca": provance,"pntc":lieu,"prd": esps,"qtt": qntt,"mtin": matierinirt,"grsn": grainsansvaleur, "grns": grainnuisible, "dbrv": debrisv,"tne": tnvv, "gr": agrcm,"vttp": tota1er, "ps": ebps,"mmmmmmmmmmmmmmmmm": provance,"dt": datte, "gcse": graincasse,"gmgr": grainmaigre, "gmch": grainmoushte, "grpn": grainpunaise,"grpq": grainpique,"grbt": grainboute,"vttd": tota2eme,"autr": autrecereal,"pbtv": bledur})
            self.doc.render({
                "ps": ebps,
                "uca": provance,
                "pntc": lieu,
                "prd": esps,
                "qtt": qntt,
                "mtin": matierinirt,
                "grsn": grainsansvaleur,
                "grns": grainnuisible,
                "dbrv": debrisv,
                "tne": tnvv,
                "gr": agrcm,
                "vttp": totap,
                "dt": datte,
                "gcse": graincasse,
                "gmgr": grainmaigre,
                "gmch": grainmoushte,
                "grpn": grainpunaise,
                "grpq": grainpique,
                "grbt": grainboute,
                "vttd": totad,
                "autr": autrecereal,
                "pbtv": bledur,
                "mat": matriadd
            })
            self.doc_name = provance + datetime.datetime.now().strftime("%d-%m-%y") + ".docx"
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer Fichiers ", self.doc_name,
                                                  "Fichiers DOCX (*.docx)")
            if path:
                self.doc.save(path)
                msg_box = QMessageBox()
                msg_box.setWindowTitle("Confirmation")
                msg_box.setText("Le fichier a été enregistré avec succès. ")
                msg_box.exec()

        def matriculbtnadd(self):
            a = self.matriculedite.text()
            self.matricullist.append(a)
            self.matriculedite.clear()

        def printerin(self):
            self.doc = DocxTemplate("Docxfiles/_bulletin entré/template_BULLETIN_ENTRE.docx")
            matriadd = self.matricullist.toPlainText()
            # matriadd = ("\n".join(matri))
            # [self.matricullist.item(x).text() for x in range(self.matricullist.count())]
            ebps = self.vpsin.text()
            provance = self.provancetecombo.currentText()
            lieu = self.lieucombo.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantite.text()
            matierinirt = self.vmatiéreinrtin.text()
            debrisv = self.vdébrisvéin.text()
            grainnuisible = self.vgrainnuisiblein.text()
            grainsansvaleur = self.vgrainsanvaleurin.text()
            agrcm = self.agréeeurcombo.currentText()
            tnvv = self.vhumiditein.text()
            tota1er = self.vtotalpremin.text()
            datte = self.dateedite.text()
            graincasse = self.vgraincasséin.text()
            grainmaigre = self.vgrainmaigrein.text()
            grainmoushte = self.vgrainmouchtéin.text()
            grainpunaise = self.vgrainpunaiséin.text()
            grainpique = self.vgrainpiquéin.text()
            grainboute = self.vgrainpiquéin.text()
            tota2eme = self.vtotaldemin.text()
            autrecereal = self.vautrecéréalin.text()
            bledur = self.vgrainetrangéin.text()
            # self.doc.render({"mat":self.items,"ps": ebps,"uca": provance,"pntc":lieu,"prd": esps,"qtt": qntt,"mtin": matierinirt,"grsn": grainsansvaleur, "grns": grainnuisible, "dbrv": debrisv,"tne": tnvv, "gr": agrcm,"vttp": tota1er, "ps": ebps,"mmmmmmmmmmmmmmmmm": provance,"dt": datte, "gcse": graincasse,"gmgr": grainmaigre, "gmch": grainmoushte, "grpn": grainpunaise,"grpq": grainpique,"grbt": grainboute,"vttd": tota2eme,"autr": autrecereal,"pbtv": bledur})
            self.doc.render({
                "ps": ebps,
                "uca": provance,
                "pntc": lieu,
                "prd": esps,
                "qtt": qntt,
                "mtin": matierinirt,
                "grsn": grainsansvaleur,
                "grns": grainnuisible,
                "dbrv": debrisv,
                "tne": tnvv,
                "gr": agrcm,
                "vttp": tota1er,
                "dt": datte,
                "gcse": graincasse,
                "gmgr": grainmaigre,
                "gmch": grainmoushte,
                "grpn": grainpunaise,
                "grpq": grainpique,
                "grbt": grainboute,
                "vttd": tota2eme,
                "autr": autrecereal,
                "pbtv": bledur,
                "mat": matriadd
            })
            doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
            doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
            self.doc.save(doc_names)
            try:
                if doc_names:
                    a = self.progress_bar()
                    sys.stderr = open("consoleoutput.log", "w")
                    convert(doc_names, doc_pdf)
                    # Open the resulting .pdf file using the default associated application
                    os.startfile(doc_pdf, 'open')
                    #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                    #subprocess.Popen([app_path, doc_pdf])
            except Exception as e:
                print(e)

        ####################################def sortie
        def save_doc_sortie(self):
            self.docs = DocxTemplate("Docxfiles/_bulletin sortie/template_BULLETIN_SORTIE.docx")
            matriadd = self.matriculcomboso.text()
            observation = self.txtobservationso.toPlainText()
            ebps = self.vpsso.text()
            distination = self.distinationcomboso.currentText()
            expéditeur = self.pointcollectecomboso.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantiteso.text()
            pinsect = self.vpinsect.text()
            espinsct = self.veinsect.text()
            mitadinetbt = self.vmetadinetblt.text()
            mitadini = self.vmitadin.text()
            bletendredandbd = self.vbltdanbld.text()
            mauvaiseherb = self.vgrainmauvaise.text()
            carie = self.vgraincarréso.text()
            matierinirt = self.vmatiéreinrtso.text()
            debrisv = self.vdébrisvéso.text()
            grainnuisible = self.vgrainnuisibleso.text()
            grainsansvaleur = self.vgrainsanvaleurso.text()
            agrcm = self.agréeeurcomboso.currentText()
            totap = self.vtotalpremso.text()
            datte = self.dateediteso.text()
            graincasse = self.vgraincasséso.text()
            grainmaigre = self.vgrainechaudeso.text()
            grainmoushte = self.vgrainmouchtéso.text()
            grainpunaise = self.vgrainpunaiséso.text()
            grainpique = self.vgrainpiquéso.text()
            grainboute = self.vgrainpiquéso.text()
            totad = self.vtotaldemso.text()
            autrecereal = self.vautrecéréalso.text()
            grainavarie = self.vgravar.text()
            self.docs.render({
                "ps": ebps,
                "pro": expéditeur,
                "dst": distination,
                "pl": esps,
                "qtt": qntt,
                "mtint": matierinirt,
                "grsn": grainsansvaleur,
                "nuis": grainnuisible,
                "dbrv": debrisv,
                "gr": agrcm,
                "ttp": totap,
                "dat": datte,
                "grcs": graincasse,
                "grech": grainmaigre,
                "grnmch": grainmoushte,
                "grpn": grainpunaise,
                "grpq": grainpique,
                "grbt": grainboute,
                "vttd": totad,
                "aut": autrecereal,
                "grav": grainavarie,
                "mat": matriadd,
                "pinsct": pinsect,
                "insct": espinsct,
                "mttt": mitadinetbt,
                "mtdn": mitadini,
                "grmh": mauvaiseherb,
                "grcr": carie,
                "gbtbd": bletendredandbd,
                "observ": observation
            })
            self.doc_names = distination + "-" + matriadd + datetime.datetime.now().strftime("%d-%m-%y") + ".docx"
            pdfnames = distination + "-" + matriadd + datetime.datetime.now().strftime("%d-%m-%y") + ".pdf"
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer Fichiers", self.doc_names, "DOCX Files (*.docx)")
            if path:
                self.docs.save(path)
                msg_box = QMessageBox()
                msg_box.setWindowTitle("Confirmation")
                msg_box.setText("Le fichier a été enregistré avec succès. ")
                msg_box.exec()

        def printer_sortie(self):
            try:
                self.docs = DocxTemplate("Docxfiles/_bulletin sortie/template_BULLETIN_SORTIE.docx")
                matriadd = self.matriculcomboso.text()
                observation = self.txtobservationso.toPlainText()
                ebps = self.vpsso.text()
                distination = self.distinationcomboso.currentText()
                expéditeur = self.pointcollectecomboso.currentText()
                esps = self.éspécecombo.currentText()
                qntt = self.quantiteso.text()
                pinsect = self.vpinsect.text()
                espinsct = self.veinsect.text()
                mitadinetbt = self.vmetadinetblt.text()
                mitadini = self.vmitadin.text()
                bletendredandbd = self.vbltdanbld.text()
                mauvaiseherb = self.vgrainmauvaise.text()
                carie = self.vgraincarréso.text()
                matierinirt = self.vmatiéreinrtso.text()
                debrisv = self.vdébrisvéso.text()
                grainnuisible = self.vgrainnuisibleso.text()
                grainsansvaleur = self.vgrainsanvaleurso.text()
                agrcm = self.agréeeurcomboso.currentText()
                totap = self.vtotalpremso.text()
                datte = self.dateediteso.text()
                graincasse = self.vgraincasséso.text()
                grainmaigre = self.vgrainechaudeso.text()
                grainmoushte = self.vgrainmouchtéso.text()
                grainpunaise = self.vgrainpunaiséso.text()
                grainpique = self.vgrainpiquéso.text()
                grainboute = self.vgrainpiquéso.text()
                totad = self.vtotaldemso.text()
                autrecereal = self.vautrecéréalso.text()
                grainavarie = self.vgravar.text()
                self.docs.render({
                    "ps": ebps,
                    "pro": expéditeur,
                    "dst": distination,
                    "pl": esps,
                    "qtt": qntt,
                    "mtint": matierinirt,
                    "grsn": grainsansvaleur,
                    "nuis": grainnuisible,
                    "dbrv": debrisv,
                    "gr": agrcm,
                    "ttp": totap,
                    "dat": datte,
                    "grcs": graincasse,
                    "grech": grainmaigre,
                    "grnmch": grainmoushte,
                    "grpn": grainpunaise,
                    "grpq": grainpique,
                    "grbt": grainboute,
                    "vttd": totad,
                    "aut": autrecereal,
                    "grav": grainavarie,
                    "mat": matriadd,
                    "pinsct": pinsect,
                    "insct": espinsct,
                    "mttt": mitadinetbt,
                    "mtdn": mitadini,
                    "grmh": mauvaiseherb,
                    "grcr": carie,
                    "gbtbd": bletendredandbd,
                    "observ": observation
                })
                doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
                doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
                self.docs.save(doc_names)
                try:
                    if doc_names:
                        a = self.progress_bar()
                        sys.stderr = open("consoleoutput.log", "w")
                        convert(doc_names, doc_pdf)
                        # Open the resulting .pdf file using the default associated application
                        os.startfile(doc_pdf, 'open')
                        #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                        #subprocess.Popen([app_path, doc_pdf])
                except Exception as e:
                    print(e)
            except Exception as e:
                print(e)

        def plus_sortie(self):
            n = self.vmetadinetblt.value()
            aa = self.vbltdanbld.value()
            ab = self.vmitadin.value()
            ac = self.vgrainmauvaise.value()
            p = self.vdébrisvéso.value()
            q = self.vmatiéreinrtso.value()
            r = self.vgrainnuisibleso.value()
            s = self.vgrainsanvaleurso.value()
            w = self.vgraincarréso.value()
            self.vtotaldemso.setValue(q + p + s + w + r)
            self.vtotalpremso.setValue(+aa + ab + ac + n)

        def clear_sortie(self):
            self.vmitadin.clear()
            self.vgrainnuisibleso.clear()
            self.vpsso.clear()
            self.vtotalpremso.clear()
            self.vgraincasséso.clear()
            self.vtotaldemso.clear()
            self.vgrainmouchtéso.clear()
            self.vgrainechaudeso.clear()
            self.vdébrisvéso.clear()
            self.vmatiéreinrtso.clear()
            self.vdébrisvéso.clear()
            self.vmatiéreinrtso.clear()
            self.vgrainsanvaleurso.clear()
            self.vgrainpunaiséso.clear()
            self.vgrainpiquéso.clear()
            self.vgrainboutéso.clear()
            self.vgraincarréso.clear()
            self.vautrecéréalso.clear()
            self.vmetadinetblt.clear()
            self.vbltdanbld.clear()
            self.vgrainmauvaise.clear()
            self.vautrecéréalso.setValue(0)
            self.vgraigerméso.setValue(0)
            self.vgrainpunaiséso.setValue(0)
            self.vgrainpiquéso.setValue(0)
            self.vgrainboutéso.setValue(0)
            self.vgraincarréso.setValue(0)
            self.vpsso.setValue(0)
            self.vtotalpremso.setValue(0)
            self.vgraincasséso.setValue(0)
            self.vtotaldemso.setValue(0)
            self.vgrainmouchtéso.setValue(0)
            self.vgrainechaudeso.setValue(0)
            self.vdébrisvéso.setValue(0)
            self.vmatiéreinrtso.setValue(0)
            self.vdébrisvéso.setValue(0)
            self.vmatiéreinrtso.setValue(0)
            self.vgrainsanvaleurso.setValue(0)
            self.vtotalpremso.setValue(0)

        def progress_bar(self):
            self.widgetprogress = QtWidgets.QDialog()
            self.widgetprogress.setStyleSheet(""" QWidget
                                {
                                    color: #000000;
                                    background-color: #ffffff;
                                    border-width: 1px;
                                    border-color: #1e1e1e;
                                    border-style: solid;
                                    border-radius: 6;
                                    padding: 0px;
                                    font-size: 18px;
                                    padding-left: 1px;
                                    padding-right: 1px
                                }
                                QWidget:item:hover
                                {
                                    background-color: #3daee9;
                                    color: #eff0f1;
                                }
                                QWidget:item:selected
                                {
                                    background-color: #3daee9;
                                }
                                QWidget:disabled
                                {
                                    color: #454545;
                                    background-color: #31363b;
                                }
                                QPushButton
                                {
                                    color: #000000;
                                    background-color:#ade3e7;
                                    border-width: 1px;
                                    border-color: #1e1e1e;
                                    border-style: solid;
                                    border-radius: 6;
                                    padding: 3px;
                                    font-size: 12px;
                                    padding-left: 5px;
                                    padding-right: 5px;
                                    min-width: 40px
                                }
                                QPushButton:disabled
                                {
                                    background-color: #31363b;
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
                                }
                                QPushButton:hover
                                {
                                    border: 1px solid #ff8c00;
                                    color: #000000;
                                }
                                 QLabel
                                {
                                    font-size: 18px;
                                    border: 0px solid orange;
                                }

                            """)
            self.widgetprogress.setWindowTitle("جاري تحميل الملف يرجى الانتظار ")
            self.widgetprogress.setGeometry(550, 450, 250, 20)
            self.progressBar = QtWidgets.QProgressBar(self.widgetprogress)
            self.progressBar.setGeometry(10, 10, 200, 10)
            self.progressBar.setMinimum(0)
            self.progressBar.setMaximum(100)
            self.progressBar.setStyleSheet("""QProgressBar
        {
        border: solid grey;
        border-radius: 15px;
        color: black;
        }
        QProgressBar::chunk 
        {
        background-color: #05B8CC;
        border-radius :15px;
        }      """)
            self.progressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.vbox = QVBoxLayout(self.widgetprogress)
            self.vbox.addWidget(self.progressBar)
            self.timer = QtCore.QTimer()
            self.timer.timeout.connect(self.update_progress)
            self.timer.start(5)  # Update progress every
            self.widgetprogress.show()

        def update_progress(self):
            # Simulate file download progress
            current_value = self.progressBar.value()
            if current_value < 100:
                new_value = current_value + 10
                self.progressBar.setValue(new_value)
                if current_value == 99:
                    self.timer.stop()
                    self.progressBar.close()
                    self.widgetprogress.close()

        #############################################################################################
        ########################################################################################
        ############################################moulin sortie
        def allcallculbd(self):
            self.plusbd()

        def clear_allbd(self):
            self.vpsbd.clear()
            self.vhumiditebd.clear()
            self.vtotalprembd.clear()
            self.vgraincassébd.clear()
            self.vtotaldembd.clear()
            self.vgrainetrangébd.clear()
            self.vgrainmouchtébd.clear()
            self.vgrainmaigrebd.clear()
            self.vgrainechaudebd.clear()
            self.vdébrisvébd.clear()
            self.vmatiéreinrtbd.clear()
            self.vdébrisvébd.clear()
            self.vmatiéreinrtbd.clear()
            self.vgrainsanvaleurbd.clear()
            self.vgrainchaufébd.clear()
            self.vgrainboutébd.clear()
            self.vblétendreinbledur.clear()
            self.vtotalcomplet.clear()
            self.vindicenotin.clear()
            self.vgrainpiquébd.clear()
            self.vgrainpunaisébd.clear()
            self.vgraincarrébd.clear()
            self.vergotbd.clear()
            self.vergotbd.setValue(0)
            self.vgraigermébd.clear()
            self.vgraigermébd.setValue(0)
            self.vgrainnuisiblebd.clear()
            self.vgrainnuisiblebd.setValue(0)
            self.vgraincarrébd.setValue(0)
            self.vgrainpunaisébd.setValue(0)
            self.vgrainpiquébd.setValue(0)
            self.vindicenotin.setValue(0)
            self.vtotalcomplet.setValue(0)
            self.vpsbd.setValue(0)
            self.vhumiditebd.setValue(0)
            self.vtotalprembd.setValue(0)
            self.vgraincassébd.setValue(0)
            self.vtotaldembd.setValue(0)
            self.vgrainetrangébd.setValue(0)
            self.vgrainmouchtébd.setValue(0)
            self.vgrainmaigrebd.setValue(0)
            self.vgrainechaudebd.setValue(0)
            self.vdébrisvébd.setValue(0)
            self.vmatiéreinrtbd.setValue(0)
            self.vdébrisvébd.setValue(0)
            self.vmatiéreinrtbd.setValue(0)
            self.vgrainsanvaleurbd.setValue(0)
            self.vgrainchaufébd.setValue(0)
            self.vtotalprembd.setValue(0)
            self.vgrainboutébd.setValue(0)
            self.vblétendreinbledur.setValue(0)

        def plusbd(self):
            e = self.vgraincassébd.value()
            f = self.vgrainmaigrebd.value()
            g = self.vgrainechaudebd.value()
            h = self.vgraigermébd.value()
            i = self.vgrainpunaisébd.value()
            j = self.vgrainpiquébd.value()
            k = self.vgrainboutébd.value()
            m = self.vgrainmouchtébd.value()
            n = self.vgrainetrangébd.value()
            q = self.vindicenotin.value()
            r = self.vblétendreinbledur.value()
            # total 2eme cat blé dur
            vtotaldemmbd = f + g + h + i + j + m + n
            # total 1er ble dur
            p = self.vdébrisvébd.value()
            qa = self.vmatiéreinrtbd.value()
            ra = self.vgrainchaufébd.value()
            s = self.vgrainsanvaleurbd.value()
            aa = self.vgrainnuisiblebd.value()
            ab = self.vgraincarrébd.value()
            if self.vgraincassébd.value() >= 4.01 and self.vgrainboutébd.value() >= 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r)
            elif self.vgraincassébd.value() >= 4.01 and self.vgrainboutébd.value() < 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd + k)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + k)
            elif self.vgrainboutébd.value() >= 4.01 and self.vgraincassébd.value() < 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd + e)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + e)
            else:
                self.vtotaldembd.setValue(vtotaldemmbd + e + k)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + e + k)

            self.vtotalprembd.setValue(p + qa + ra + s + aa + ab)

        def docx_bdsave(self):
            self.docbd = DocxTemplate("Docxfiles/_bulletin moulin_sortie/_bulletin moulin_Blé DUR/tempfile_blédur.docx")
            ebpsbd = self.vpsbd.text()
            observation = self.txtobservationbd.toPlainText()
            moulnbd = self.moulincombobd.currentText()
            pntclbd = self.pointcollectecombobd.currentText()
            espsbd = self.éspécecombobd.currentText()
            qnttbd = self.quantitetxtbd.text()
            agrcmbd = self.agréeeurcombobd.currentText()
            tnvvbd = self.vhumiditebd.text()
            ttv1bd = self.vtotalprembd.text()
            grcassévbd = self.vgraincassébd.text()
            grmchbd = self.vgrainmouchtébd.value()
            gretrngbd = self.vgrainetrangébd.value()
            total2vbd = self.vtotaldembd.text()
            grnmgrbd = self.vgrainmaigrebd.value()
            grechdbd = self.vgrainechaudebd.value()
            grgrmbd = self.vgraigermébd.value()
            grpnsbd = self.vgrainpunaisébd.value()
            grpqbd = self.vgrainpiquébd.value()
            grbtbd = self.vgrainboutébd.value()
            dattebd = self.dateeditebd.text()

            debrivébd = self.vdébrisvébd.value()
            matinrtdb = self.vmatiéreinrtbd.value()
            grainmaigrebd = self.vgrainmouchtébd.value()
            grainboute = self.vgrainboutébd.value()
            garinssanvaleur = self.vgrainsanvaleurbd.value()
            grainchaufébd = self.vgrainchaufébd.value()
            indice = self.vindicenotin.value()
            ttcomplet = self.vtotalcomplet.text()
            bletendredbd = self.vblétendreinbledur.value()
            graincarre = self.vgraincarrébd.value()
            ergotbd = self.vergotbd.value()
            grainnuisiblebd = self.vgrainnuisiblebd.value()

            self.docbd.render(
                {"gnsv": grainnuisiblebd, "erg": ergotbd,
                 "grcr": graincarre, "grbtf": grbtbd, "grpq": grpqbd, "grpn": grpnsbd, "grg": grgrmbd, "gehv": grechdbd,
                 "gmv": grnmgrbd, "tt1v": ttv1bd, "tnev": tnvvbd, "psv": ebpsbd, "gr": agrcmbd,
                 "esp": espsbd, "mmmmmmmmmmmmmmmmm": moulnbd, "pntc": pntclbd, "qtt": qnttbd, "dt": dattebd,
                 "gcv": grcassévbd, "tt2v": total2vbd,
                 "getv": gretrngbd, "gmv": grmchbd, "gehv": grechdbd,
                 "gnsv": grainnuisiblebd,
                 "dbv": debrivébd,
                 "mtiv": matinrtdb,
                 "obb": observation,
                 "grch": grainchaufébd,
                 "grsn": garinssanvaleur,
                 "gmx": grainmaigrebd,
                 "grbt": grainboute,
                 "indv": indice,
                 "btdv": bletendredbd,
                 "ttcv": ttcomplet,
                 })
            self.docbd_name = moulnbd + "-" + datetime.datetime.now().strftime("%m-%y") + "-" + ".docx"
            # self.docbd.save("_bulletin moulin/_bulletin moulin_Blé DUR/" + self.docbd_name)
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer la fiche", self.docbd_name,
                                                  "Fichiers DOCX (*.docx)")
            if path:
                self.docbd.save(path)
                msgbox = QtWidgets.QMessageBox()
                msgbox.setWindowTitle('confirmation')
                msgbox.setText('Le fichier a été enregistré avec succès.')
                msgbox.exec()

        def printerbd(self):
            self.docbd = DocxTemplate("Docxfiles/_bulletin moulin_sortie/_bulletin moulin_Blé DUR/tempfile_blédur.docx")
            ebpsbd = self.vpsbd.text()
            observation = self.txtobservationbd.toPlainText()
            moulnbd = self.moulincombobd.currentText()
            pntclbd = self.pointcollectecombobd.currentText()
            espsbd = self.éspécecombobd.currentText()
            qnttbd = self.quantitetxtbd.text()
            agrcmbd = self.agréeeurcombobd.currentText()
            tnvvbd = self.vhumiditebd.text()
            ttv1bd = self.vtotalprembd.text()
            grcassévbd = self.vgraincassébd.text()
            grmchbd = self.vgrainmouchtébd.value()
            gretrngbd = self.vgrainetrangébd.value()
            total2vbd = self.vtotaldembd.text()
            grnmgrbd = self.vgrainmaigrebd.value()
            grechdbd = self.vgrainechaudebd.value()
            grgrmbd = self.vgraigermébd.value()
            grpnsbd = self.vgrainpunaisébd.value()
            grpqbd = self.vgrainpiquébd.value()
            grbtbd = self.vgrainboutébd.value()
            dattebd = self.dateeditebd.text()

            debrivébd = self.vdébrisvébd.value()
            matinrtdb = self.vmatiéreinrtbd.value()
            grainmaigrebd = self.vgrainmouchtébd.value()
            grainboute = self.vgrainboutébd.value()
            garinssanvaleur = self.vgrainsanvaleurbd.value()
            grainchaufébd = self.vgrainchaufébd.value()
            indice = self.vindicenotin.value()
            ttcomplet = self.vtotalcomplet.text()
            bletendredbd = self.vblétendreinbledur.value()
            graincarre = self.vgraincarrébd.value()
            ergotbd = self.vergotbd.value()
            grainnuisiblebd = self.vgrainnuisiblebd.value()

            self.docbd.render(
                {"gnsv": grainnuisiblebd, "erg": ergotbd,
                 "grcr": graincarre, "grbtf": grbtbd, "grpq": grpqbd, "grpn": grpnsbd, "grg": grgrmbd, "gehv": grechdbd,
                 "gmv": grnmgrbd, "tt1v": ttv1bd, "tnev": tnvvbd, "psv": ebpsbd, "gr": agrcmbd,
                 "esp": espsbd, "mmmmmmmmmmmmmmmmm": moulnbd, "pntc": pntclbd, "qtt": qnttbd, "dt": dattebd,
                 "gcv": grcassévbd, "tt2v": total2vbd,
                 "getv": gretrngbd, "gmv": grmchbd, "gehv": grechdbd,
                 "gnsv": grainnuisiblebd,
                 "dbv": debrivébd,
                 "mtiv": matinrtdb,
                 "obb": observation,
                 "grch": grainchaufébd,
                 "grsn": garinssanvaleur,
                 "gmx": grainmaigrebd,
                 "grbt": grainboute,
                 "indv": indice,
                 "btdv": bletendredbd,
                 "ttcv": ttcomplet,
                 })
            doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
            doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
            self.docbd.save(doc_names)
            try:
                if doc_names:
                    a = self.progress_bar()
                    sys.stderr = open("consoleoutput.log", "w")
                    convert(doc_names, doc_pdf)
                    # Open the resulting .pdf file using the default associated application
                    os.startfile(doc_pdf, 'open')
                    #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                    #subprocess.Popen([app_path, doc_pdf])
            except Exception as e:
                print(e)

            #################blé tendre def

        def plusbt(self):
            e = self.vgraincassé.value()
            f = self.vgrainmaigre.value()
            g = self.vgrainechaude.value()
            h = self.vgraigermé.value()
            i = self.vgrainpunaisé.value()
            j = self.vgrainpiqué.value()
            k = self.vgrainbouté.value()
            l = self.vgrainboutef.value()
            m = self.vgrainmouchté.value()
            n = self.vgrainetrangé.value()
            p = self.vdébrisvé.value()
            q = self.vmatiéreinrt.value()
            r = self.vgrainchaufé.value()
            s = self.vgrainsanvaleur.value()
            v = self.vgraincarré.value()
            ac = self.vgrainnuisible.value()
            vtotaldemm = f + g + h + i + j + k + l + m + n
            self.vtotalprem.setValue(p + q + r + s + v + ac)
            self.vtotaldem.setValue(vtotaldemm)
            if self.vgraincassé.value() >= 2.01 and self.vgraincassé.value() <= 15:
                self.vtotaldem.setValue(vtotaldemm)
            else:
                self.vtotaldem.setValue(e + vtotaldemm)

            # vtotalpremm=p+q+r+s

        def clear_allbt(self):
            self.vgrainnuisible.clear()
            self.vgrainnuisible.setValue(0)
            self.vgrainchaufé.clear()
            self.vgrainchaufé.setValue(0)
            self.vps.clear()
            self.vgrainboutef.clear()
            self.vgrainbouté.clear()
            self.vhumiditein.clear()
            self.vgrainsanvaleur.clear()
            self.vgraincarré.clear()
            self.vtotalprem.clear()
            self.vgraincassé.clear()
            self.vtotaldem.clear()
            self.vgrainpunaisé.clear()
            self.vgrainpunaisé.setValue(0)
            self.vgraigermé.clear()
            self.vgraigermé.setValue(0)
            self.vgrainetrangé.clear()
            self.vgrainmouchté.clear()
            self.vgrainmaigre.clear()
            self.vgrainechaude.clear()
            self.vgrainbouté.clear()
            self.vdébrisvé.clear()
            self.vmatiéreinrt.clear()
            self.vgrainpiqué.clear()
            self.vgrainpiqué.setValue(0)
            self.vps.setValue(0)
            self.vhumiditein.setValue(0)
            self.vgrainsanvaleur.setValue(0)
            self.vgraincarré.setValue(0)
            self.vtotalprem.setValue(0)
            self.vgraincassé.setValue(0)
            self.vgrainboutef.setValue(0)
            self.vtotaldem.setValue(0)
            self.vgrainetrangé.setValue(0)
            self.vgrainmouchté.setValue(0)
            self.vgrainmaigre.setValue(0)
            self.vgrainechaude.setValue(0)
            self.vdébrisvé.setValue(0)
            self.vmatiéreinrt.setValue(0)
            self.vgrainbouté.setValue(0)

        def docx_file(self):
            self.doc = DocxTemplate("Docxfiles/_bulletin moulin_sortie/_bulletin moulin_Blé TENDRE/invoice_template.docx")
            ebps = self.vps.text()
            ergot = self.vergot.value()
            observationbt = self.txtobservation.toPlainText()
            nuisible = self.vgrainnuisible.value()
            mouln = self.moulincombo.currentText()
            pntcl = self.pointcollectecombo.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantitetxt.text()
            agrcm = self.agréeeurcombo.currentText()
            tnvv = self.vhumiditein.text()
            debritvegetaux = self.vdébrisvé.value()
            matierinert = self.vmatiéreinrt.value()
            grainchaufe = self.vgrainchaufé.value()
            grainsanvaleur = self.vgrainsanvaleur.value()
            graincarie = self.vgraincarré.value()
            graingerme = self.vgraigermé.value()
            grainpunaisés = self.vgrainpunaisé.value()
            ttv1 = self.vtotalprem.text()
            grcassév = self.vgraincassé.text()
            grmgre = self.vgrainmaigre.text()
            grechd = self.vgrainechaude.text()
            grmch = self.vgrainmouchté.text()
            gretrng = self.vgrainetrangé.text()
            total2v = self.vtotaldem.text()
            grnmgr = self.vgrainmaigre.value()
            grechd = self.vgrainechaude.value()
            grgrm = self.vgraigermé.value()
            grpns = self.vgrainpunaisé.value()
            grpq = self.vgrainpiqué.value()
            grbt = self.vgrainbouté.value()
            grnbtf = self.vgrainboutef.value()
            datte = self.dateedite.text()
            self.doc.render(
                {"verg": ergot, "vnsb": nuisible,
                 "grpn": grainpunaisés, "grch": grainchaufe, "grsn": grainsanvaleur, "vgc": graincarie,
                 "grg": graingerme,
                 "grbt": grnbtf, "grbtf": grbt, "dbv": debritvegetaux, "mtiv": matierinert, "grpq": grpq, "grpn": grpns,
                 "grg": grgrm, "gehv": grechd, "gmx": grnmgr, "tt1v": ttv1, "tnev": tnvv,
                 "psv": ebps, "gr": agrcm, "esp": esps, "mmmmmmmmmmmmmmmmm": mouln, "pntc": pntcl, "qtt": qntt,
                 "dt": datte,
                 "gcv": grcassév, "tt2v": total2v,
                 "getv": gretrng, "gmv": grmch, "gehv": grechd, "obb": observationbt, "gmv": grmgre})
            doc_name = mouln + "-" + datetime.datetime.now().strftime("%m-%y") + "-" + ".docx"
            # self.doc.save("_bulletin moulin/_bulletin moulin_Blé TENDRE/" + self.doc_name)
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer la fiche", doc_name, "Fichiers DOCX (*.docx)")
            if path:
                self.doc.save(path)
                msgbox = QtWidgets.QMessageBox()
                msgbox.setWindowTitle('confirmation')
                msgbox.setText('Le fichier a été enregistré avec succès.')
                msgbox.exec()

        def printer(self):
            try:
                self.doc = DocxTemplate("Docxfiles/_bulletin moulin_sortie/_bulletin moulin_Blé TENDRE/invoice_template.docx")
                ebps = self.vps.text()
                ergot = self.vergot.value()
                observationbt = self.txtobservation.toPlainText()
                nuisible = self.vgrainnuisible.value()
                mouln = self.moulincombo.currentText()
                pntcl = self.pointcollectecombo.currentText()
                esps = self.éspécecombo.currentText()
                qntt = self.quantitetxt.text()
                agrcm = self.agréeeurcombo.currentText()
                tnvv = self.vhumiditein.text()
                debritvegetaux = self.vdébrisvé.value()
                matierinert = self.vmatiéreinrt.value()
                grainchaufe = self.vgrainchaufé.value()
                grainsanvaleur = self.vgrainsanvaleur.value()
                graincarie = self.vgraincarré.value()
                graingerme = self.vgraigermé.value()
                grainpunaisés = self.vgrainpunaisé.value()
                ttv1 = self.vtotalprem.text()
                grcassév = self.vgraincassé.text()
                grmgre = self.vgrainmaigre.text()
                grechd = self.vgrainechaude.text()
                grmch = self.vgrainmouchté.text()
                gretrng = self.vgrainetrangé.text()
                total2v = self.vtotaldem.text()
                grnmgr = self.vgrainmaigre.value()
                grechd = self.vgrainechaude.value()
                grgrm = self.vgraigermé.value()
                grpns = self.vgrainpunaisé.value()
                grpq = self.vgrainpiqué.value()
                grbt = self.vgrainbouté.value()
                grnbtf = self.vgrainboutef.value()
                datte = self.dateedite.text()
                self.doc.render(
                    {"verg": ergot, "vnsb": nuisible,
                     "grpn": grainpunaisés, "grch": grainchaufe, "grsn": grainsanvaleur, "vgc": graincarie,
                     "grg": graingerme,
                     "grbt": grnbtf, "grbtf": grbt, "dbv": debritvegetaux, "mtiv": matierinert, "grpq": grpq,
                     "grpn": grpns,
                     "grg": grgrm, "gehv": grechd, "gmx": grnmgr, "tt1v": ttv1, "tnev": tnvv,
                     "psv": ebps, "gr": agrcm, "esp": esps, "mmmmmmmmmmmmmmmmm": mouln, "pntc": pntcl, "qtt": qntt,
                     "dt": datte,
                     "gcv": grcassév, "tt2v": total2v,
                     "getv": gretrng, "gmv": grmch, "gehv": grechd, "obb": observationbt, "gmv": grmgre})

                doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
                doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
                self.doc.save(doc_names)
                try:
                    if doc_names:
                        a = self.progress_bar()
                        sys.stderr = open("consoleoutput.log", "w")
                        convert(doc_names, doc_pdf)
                        # Open the resulting .pdf file using the default associated application
                        os.startfile(doc_pdf, 'open')
                        #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                        #subprocess.Popen([app_path, doc_pdf])
                except Exception as e:
                    print(e)
            except Exception as e:
                print(e)

        ##########################################################################################################################################
        #########################################################################################
        ###########################################################################################
        def confi_window(self, MainWindow):
            MainWindow.setObjectName("MainWindow")
            MainWindow.resize(1338, 700)
            MainWindow.setWindowIcon(QIcon("images/Picsart_23-03-14_20-11-34-387 (1).ico"))
            MainWindow.setStyleSheet("""QWidget
    {
        color: #eff0f1;
        background-color: #ffffff;
        selection-background-color:#3daee9;
        selection-color: #eff0f1;
        background-clip: border;
        border-image: none;
        border: 0px transparent black;
        outline: 0;
    }""")
            self.centralwidget = QtWidgets.QWidget(MainWindow)
            self.centralwidget.setObjectName("centralwidget")
            self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
            self.horizontalLayout.setObjectName("horizontalLayout")
            self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
            self.confirmitewidget = QtWidgets.QTabWidget(self.centralwidget)
            self.confirmitewidget.setStyleSheet("""QToolTip
            {
                border: 1px solid #76797C;
                background-color: rgb(90, 102, 117);;
                color: white;
                padding: 5px;
                opacity: 200;
            }

            QWidget
            {
                color: #000000;
                background-color: #ffffff;
                selection-background-color:#3daee9;
                selection-color: #3daee9;
                background-clip: border;
                border-image: none;
                border: 0px transparent black;
                outline: 0;
            }

            QWidget:item:hover
            {
                background-color: #3daee9;
                color: #eff0f1;
            }

            QWidget:item:selected
            {
                background-color: #3daee9;
            }



            QWidget:disabled
            {
                color: #454545;
                background-color: #31363b;
            }

            QAbstractItemView
            {
                alternate-background-color: #31363b;
                color: #eff0f1;
                border: 1px solid 3A3939;
                border-radius: 2px;
            }

            QWidget:focus, QMenuBar:focus
            {
                border: 1px solid #3daee9;
            }

            QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
            {
                border: none;
            }

            QLineEdit
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color: #000000;
            }
            QDoubleSpinBox
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color:#000000;
                font-size: 11px;
                font-weight: bold;

            }
            QDoubleSpinBox:focus{
                background-color: #FDFEFE;
                border-style: solid;
                border: 2px solid #76797C;
                border-radius: 4px;
                border-color: #ff8c00;
            }
            QDoubleSpinBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 1px;

                border-left-width: 0px;
                border-left-color: #232629;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }



            QGroupBox {
                border:1px solid #76797C;
                border-radius: 2px;
                margin-top: 5px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding-left: 4px;
                padding-right: 4px;
                padding-top: 4px;
            }

            QAbstractScrollArea
            {
                border-radius: 2px;
                border: 1px solid #76797C;
                background-color: transparent;
            }

            QScrollBar:horizontal
            {
                height: 15px;
                margin: 3px 15px 3px 15px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
                background-color: #2A2929;
            }

            QScrollBar::handle:horizontal
            {
                background-color: #605F5F;
                min-width: 5px;
                border-radius: 4px;
            }

            QScrollBar::add-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
                width: 10px;
                height: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }


            QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
            {
                background: none;
            }


            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
            {
                background: none;
            }

            QScrollBar:vertical
            {
                background-color: #2A2929;
                width: 15px;
                margin: 15px 3px 15px 3px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
            }

            QScrollBar::handle:vertical
            {
                background-color: #605F5F;
                min-height: 5px;
                border-radius: 4px;
            }

            QScrollBar::sub-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
            {

                border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }


            QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
            {
                border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
            {
                background: none;
            }


            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
            {
                background: none;
            }

            QTextEdit
            {
                background-color: #FDFEFE;
                color: #000000;
                border: 1px solid #76797C;
            }

            QPlainTextEdit
            {
                background-color: #232629;;
                color: #000000;
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QHeaderView::section
            {
                background-color: #76797C;
                color: #eff0f1;
                padding: 1px;
                border: 1px solid #76797C;
            }

            QSizeGrip {
                width: 12px;
                height: 12px;
            }


            QMainWindow::separator
            {
                background-color: #31363b;
                color: white;
                padding-left: 4px;
                spacing: 2px;
                border: 1px dashed #76797C;
            }

            QMainWindow::separator:hover
            {

                background-color: #787876;
                color: white;
                padding-left: 4px;
                border: 1px solid #76797C;
                spacing: 2px;
            }


            QMenu::separator
            {
                height: 1px;
                background-color: #76797C;
                color: white;
                padding-left: 4px;
                margin-left: 10px;
                margin-right: 5px;
            }


            QFrame
            {
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QFrame[frameShape="0"]
            {
                border-radius: 2px;
                border: 1px transparent #76797C;
            }

            QStackedWidget
            {
                border: 1px transparent black;
            }


            QPushButton
            {
                color: #00000;
                background-color:#ade3e7;
                border-width: 1px;
                border-color: #1e1e1e;
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
            }

            QComboBox
            {
               background-color: #FDFEFE;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                min-width: 40px;
            }

            QPushButton:checked{
                background-color: #76797C;
                border-color: #6A6969;
            }

            QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
            {
                border: 1px solid #ff8c00;
                color: #000000;
            }

            QComboBox:on
            {
                padding-top: 1px;
                padding-left: 1px;
                selection-background-color: #FDFEFE;
            }

            QComboBox QAbstractItemView
            {
                background-color: #FDFEFE;
                border-radius: 2px;
                border: 1px solid #76797C;
                color:#000000;
                selection-background-color: #000000;
            }

            QComboBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;

                border-left-width: 0px;
                border-left-color: ff8c00;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }


            QLabel
            {
                border: 2px solid black;
            }

            QTabWidget{
                border: 0px transparent black;
            }

            QTabWidget::pane {
                border: 1px solid #76797C;
                padding: 5px;
                margin: 0px;
            }

            QTabBar
            {
                qproperty-drawBase: 0;
                left: 5px; /* move to the right by 5px */
                border-radius: 3px;
            }

            QTabBar:focus
            {
                border: 0px transparent black;
            }

            QTabBar::close-button  {
                image: url(:/qss_icons/Dark_rc/close.png);
                background: transparent;
            }

            QTabBar::close-button:hover
            {
                image: url(:/qss_icons/Dark_rc/close-hover.png);
                background: transparent;
            }

            QTabBar::close-button:pressed {
                image: url(:/qss_icons/Dark_rc/close-pressed.png);
                background: transparent;
            }

            /* TOP TABS */
            QTabBar::tab:top {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                min-width: 10px;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;
            }

            QTabBar::tab:top:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;    
            }

            QTabBar::tab:top:!selected:hover {
                background-color: #3daee9;
            }

            /* BOTTOM TABS */
            QTabBar::tab:bottom {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
                min-width: 10px;
            }

            QTabBar::tab:bottom:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:bottom:!selected:hover {
                background-color: #3daee9;
            }

            /* LEFT TABS */
            QTabBar::tab:left {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:left:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:left:!selected:hover {
                background-color: #3daee9;
            }


            /* RIGHT TABS */
            QTabBar::tab:right {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:right:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
            }





            QSlider::groove:horizontal {
                border: 1px solid #565a5e;
                height: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 2px;
            }

            QSlider::handle:horizontal {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: -8px 0;
                border-radius: 9px;
            }

            QSlider::groove:vertical {
                border: 1px solid #565a5e;
                width: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 3px;
            }

            QSlider::handle:vertical {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: 0 -8px;
                border-radius: 9px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px transparent #76797C;
                border-radius: 2px;
                margin: 3px;
                padding: 5px;
            }

            QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
             padding-right: 20px; /* make way for the popup button */
             border: 1px #76797C;
             border-radius: 5px;
            }

            QToolButton[popupMode="2"] { /* only for InstantPopup */
             padding-right: 10px; /* make way for the popup button */
             border: 1px #76797C;
            }


            QToolButton:hover, QToolButton::menu-button:hover {
                background-color: transparent;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            QToolButton:checked, QToolButton:pressed,
                    QToolButton::menu-button:pressed {
                background-color: #3daee9;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
            QToolButton::menu-indicator {
                background-color:ff8c00;
                top: -7px; left: -2px; /* shift it a bit */
            }

            /* the subcontrols below are used only in the MenuButtonPopup mode */
            QToolButton::menu-button {
                border: 1px transparent #76797C;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
                /* 16px width + 4px for border = 20px allocated above */
                width: 16px;
                outline: none;
            }

            QToolButton::menu-arrow {
               background-color:ff8c00;
            }

            QToolButton::menu-arrow:open {
                border: 1px solid #76797C;
            }

            QPushButton::menu-indicator  {
                subcontrol-origin: padding;
                subcontrol-position: bottom right;
                left: 8px;
            }

            QTableView
            {
                border: 1px solid #76797C;
                gridline-color: #31363b;
                background-color: #FDFEFE;
                color:#000000;
            }


            QTableView, QHeaderView
            {
                background-color: #FDFEFE;
                color:#000000;
                border-radius: 0px;
            }

            QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
                background: #FDFEFE;
                color: #000000;
            }

            QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
                background: #3daee9;
                color: #000000;
            }


            QHeaderView
            {
                background-color: #FDFEFE;
                border: 1px transparent;
                border-radius: 0px;
                margin: 0px;
                padding: 0px;

            }

            QHeaderView::section  {
                background-color:#80f1f9;
                color: #000000;
                padding: 5px;
                border: 1px solid #76797C;
                border-radius: 0px;
                text-align: center;
            }

            QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
            {
                border-top: 1px solid #76797C;
            }

            QHeaderView::section::vertical
            {
                border-top: transparent;
            }

            QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
            {
                border-left: 1px solid #76797C;
            }

            QHeaderView::section::horizontal
            {
                border-left: transparent;
            }


            QHeaderView::section:checked
             {
                color: #000000;
                background-color: #3daee9;
             }

             /* style the sort indicator */
            QHeaderView::down-arrow {

            }

            QHeaderView::up-arrow {

            }


            QTableCornerButton::section {
                background-color: #31363b;
                border: 1px transparent #76797C;
                border-radius: 0px;
            }

            QToolBox  {
                padding: 5px;
                border: 1px transparent black;
            }

            QToolBox::tab {
                color: #eff0f1;
                background-color: #31363b;
                border: 1px solid #76797C;
                border-bottom: 1px transparent #31363b;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }

            QToolBox::tab:selected { /* italicize selected tabs */
                font: italic;
                background-color: #31363b;
                border-color: #3daee9;
             }

            QStatusBar::item {
                border: 0px transparent dark;
             }


            QFrame[height="3"], QFrame[width="3"] {
                background-color: #76797C;
            }




            QDateEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;
            }

            QDateEdit:on
            {
                padding-top: 2px;
                padding-left: 2px;
                selection-background-color: #4a4a4a;
            }

            QDateEdit QAbstractItemView
            {
                background-color: #ff8c00;
                border-radius: 2px;
                border: 1px solid #3375A3;
                selection-background-color:ff8c00;
            }

            QDateEdit::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 0px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }   
            QDateTimeEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;

            }    
            """)
            self.confirmitewidget.setObjectName("confirmitewidget")
            self.entrétab = QtWidgets.QWidget()
            self.entrétab.setObjectName("entrétab")
            self.confirmitewidget.addTab(self.entrétab, "Ble tendre")

            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétre = QtWidgets.QLabel("Paramètre recherchés", self.entrétab)
            self.paramétre.move(30, 155)
            self.paramétre.resize(150, 20)
            self.paramétre.setFont(self.font)

            self.txtpsfont = QtGui.QFont()
            self.txtpsfont.setBold(True)
            self.txtpsfont.setPointSize(9)

            ######################Limites(sans bon ni réf)################
            self.ps = QtWidgets.QLineEdit("Poids spécifique (kg/hl): ", self.entrétab, readOnly=True)
            self.ps.resize(319, 20)
            self.ps.move(30, 187)
            self.ps.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.ps.setFont(self.txtpsfont)
            ###############################humidite#############
            self.humidite = QtWidgets.QLineEdit("Teneur en eau(%):", self.entrétab, readOnly=True)
            self.humidite.resize(319, 20)
            self.humidite.move(30, 208)
            self.humidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.humidite.setFont(self.txtpsfont)

            #########################Graines nuisibles (%)##########
            self.grainnuisible = QtWidgets.QLineEdit("Graines nuisibles(%):", self.entrétab, readOnly=True)
            self.grainnuisible.resize(319, 20)
            self.grainnuisible.move(30, 229)
            self.grainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisible.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvé = QtWidgets.QLineEdit("Débris végétaux(%):", self.entrétab, readOnly=True)
            self.débrisvé.resize(319, 20)
            self.débrisvé.move(30, 250)
            self.débrisvé.setFont(self.txtpsfont)
            self.débrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrt = QtWidgets.QLineEdit("Matière inerte(%):", self.entrétab, readOnly=True)
            self.matiéreinrt.resize(319, 20)
            self.matiéreinrt.move(30, 271)
            self.matiéreinrt.setFont(self.txtpsfont)
            self.matiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.grainchaufé = QtWidgets.QLineEdit("Grains chauffés(%):", self.entrétab, readOnly=True)
            self.grainchaufé.resize(319, 20)
            self.grainchaufé.move(30, 292)
            self.grainchaufé.setFont(self.txtpsfont)
            self.grainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleur = QtWidgets.QLineEdit("Grains sans valeur(%):", self.entrétab, readOnly=True)
            self.grainsanvaleur.resize(319, 20)
            self.grainsanvaleur.move(30, 313)
            self.grainsanvaleur.setFont(self.txtpsfont)
            self.grainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:", self.entrétab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 334)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalprem = QtWidgets.QLineEdit("TOTAL 1ére CAT:", self.entrétab, readOnly=True)
            self.totalprem.resize(319, 20)
            self.totalprem.move(30, 355)
            self.totalprem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalprem.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincassé = QtWidgets.QLineEdit("Grains cassés(%):", self.entrétab, readOnly=True)
            self.graincassé.move(30, 376)
            self.graincassé.resize(319, 20)
            self.graincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincassé.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaude = QtWidgets.QLineEdit("Gains échaudés(%):", self.entrétab, readOnly=True)
            self.grainechaude.move(30, 397)
            self.grainechaude.resize(319, 20)
            self.grainechaude.setFont(self.txtpsfont)
            self.grainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################
            self.grainmaigre = QtWidgets.QLineEdit("Grains maigres(%):", self.entrétab, readOnly=True)
            self.grainmaigre.move(30, 523)
            self.grainmaigre.resize(319, 20)
            self.grainmaigre.setFont(self.txtpsfont)
            self.grainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##########################################################Grains germés (%)###################################################
            self.graigermé = QtWidgets.QLineEdit("Grains germés(%):", self.entrétab, readOnly=True)
            self.graigermé.move(30, 439)
            self.graigermé.resize(319, 20)
            self.graigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigermé.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaisé = QtWidgets.QLineEdit("Grain punaisés(%):", self.entrétab, readOnly=True)
            self.grainpunaisé.move(30, 460)
            self.grainpunaisé.resize(319, 20)
            self.grainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaisé.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiqué = QtWidgets.QLineEdit("Grains piqués(%):  ", self.entrétab, readOnly=True)
            self.grainpiqué.move(30, 481)
            self.grainpiqué.resize(319, 20)
            self.grainpiqué.setFont(self.txtpsfont)
            self.grainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.grainboutef = QtWidgets.QLineEdit("Grains boutés (%):", self.entrétab, readOnly=True)
            self.grainboutef.move(30, 502)
            self.grainboutef.resize(319, 20)
            self.grainboutef.setFont(self.txtpsfont)
            self.grainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            self.autrecéréal = QtWidgets.QLineEdit("Autres céréales (%):", self.entrétab, readOnly=True)
            self.autrecéréal.move(30, 565)
            self.autrecéréal.resize(319, 20)
            self.autrecéréal.setFont(self.txtpsfont)
            self.autrecéréal.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchté = QtWidgets.QLineEdit("Grains mouchetés (%):", self.entrétab, readOnly=True)
            self.grainmouchté.move(30, 418)
            self.grainmouchté.resize(319, 20)
            self.grainmouchté.setFont(self.txtpsfont)
            self.grainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.grainetrangé = QtWidgets.QLineEdit("Présence de blé dur dans le blé tendre(%):  ", self.entrétab,
                                                    readOnly=True)
            self.grainetrangé.move(30, 586)
            self.grainetrangé.resize(319, 20)
            self.grainetrangé.setFont(self.txtpsfont)
            self.grainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldem = QtWidgets.QLineEdit("Impuretés 2eme catégorie   ", self.entrétab, readOnly=True)
            self.totaldem.move(30, 544)
            self.totaldem.resize(319, 20)
            self.totaldem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldem.setFont(self.txtpsfont)
            ###########################################################################################################
            #################label valeure##############
            self.valeur = QtWidgets.QLabel("Résultat", self.entrétab)
            self.valeur.move(350, 155)
            self.valeur.resize(100, 20)
            self.valeur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.vpsin = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vpsin.setRange(69, 81.00)
            self.vpsin.setSpecialValueText(' ')
            self.vpsin.resize(100, 20)
            self.vpsin.move(350, 187)
            self.vpsin.setFont(self.txtpsfont)
            self.vpsin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.vhumiditein = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vhumiditein.setRange(8, 14)
            self.vhumiditein.resize(100, 20)
            self.vhumiditein.setSpecialValueText(' ')
            self.vhumiditein.move(350, 208)
            self.vhumiditein.setFont(self.txtpsfont)
            self.vhumiditein.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Graines nuisibles (%)##########
            self.vgrainnuisiblein = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainnuisiblein.setRange(0, 1)
            self.vgrainnuisiblein.setSpecialValueText(' ')
            self.vgrainnuisiblein.resize(100, 20)
            self.vgrainnuisiblein.move(350, 229)
            self.vgrainnuisiblein.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #############################Débris végétaux (%)########
            self.vdébrisvéin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vdébrisvéin.setRange(0, 5)
            self.vdébrisvéin.setSpecialValueText(' ')
            self.vdébrisvéin.resize(100, 20)
            self.vdébrisvéin.move(350, 250)
            self.vdébrisvéin.setFont(self.txtpsfont)
            self.vdébrisvéin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrtin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vmatiéreinrtin.setRange(0, 5)
            self.vmatiéreinrtin.setSpecialValueText(' ')
            self.vmatiéreinrtin.resize(100, 20)
            self.vmatiéreinrtin.move(350, 271)
            self.vmatiéreinrtin.setFont(self.txtpsfont)
            self.vmatiéreinrtin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.vgrainchauféin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainchauféin.setRange(0, 5)
            self.vgrainchauféin.setSpecialValueText(' ')
            self.vgrainchauféin.resize(100, 20)
            self.vgrainchauféin.move(350, 292)
            self.vgrainchauféin.setFont(self.txtpsfont)
            self.vgrainchauféin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleurin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainsanvaleurin.setSpecialValueText(' ')
            self.vgrainsanvaleurin.setRange(0, 5)
            self.vgrainsanvaleurin.resize(100, 20)
            self.vgrainsanvaleurin.move(350, 313)
            self.vgrainsanvaleurin.setFont(self.txtpsfont)
            self.vgrainsanvaleurin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarréin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgraincarréin.setSpecialValueText(' ')
            self.vgraincarréin.setRange(0, 5)
            self.vgraincarréin.resize(100, 20)
            self.vgraincarréin.move(350, 334)
            self.vgraincarréin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Total(%) 1er#####################################################
            self.vtotalpremin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=True)
            self.vtotalpremin.setSpecialValueText(' ')
            self.vtotalpremin.resize(100, 20)
            self.vtotalpremin.move(350, 355)
            self.vtotalpremin.setFont(self.txtpsfont)
            self.vtotalpremin.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##############################################Grains cassés (%) #########################################################
            self.vgraincasséin = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vgraincasséin.move(350, 376)
            self.vgraincasséin.resize(100, 20)
            self.vgraincasséin.setRange(0, 5)
            self.vgraincasséin.setSpecialValueText(" ")
            self.vgraincasséin.setFont(self.txtpsfont)
            self.vgraincasséin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaudein = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vgrainechaudein.setSpecialValueText(" ")
            self.vgrainechaudein.setRange(0, 5)
            self.vgrainechaudein.move(350, 397)
            self.vgrainechaudein.resize(100, 20)
            self.vgrainechaudein.setFont(self.txtpsfont)
            self.vgrainechaudein.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.vgrainmaigrein = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vgrainmaigrein.setRange(0, 5)
            self.vgrainmaigrein.setSpecialValueText(" ")
            self.vgrainmaigrein.move(350, 523)
            self.vgrainmaigrein.setFont(self.txtpsfont)
            self.vgrainmaigrein.resize(100, 20)
            self.vgrainmaigrein.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.vgraigerméin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgraigerméin.move(350, 439)
            self.vgraigerméin.resize(100, 20)
            self.vgraigerméin.setSpecialValueText('  ')
            self.vgraigerméin.setFont(self.txtpsfont)
            self.vgraigerméin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaiséin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainpunaiséin.move(350, 460)
            self.vgrainpunaiséin.resize(100, 20)
            self.vgrainpunaiséin.setSpecialValueText('   ')
            self.vgrainpunaiséin.setFont(self.txtpsfont)
            self.vgrainpunaiséin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiquéin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainpiquéin.move(350, 481)
            self.vgrainpiquéin.resize(100, 20)
            self.vgrainpiquéin.setSpecialValueText('  ')
            self.vgrainpiquéin.setFont(self.txtpsfont)
            self.vgrainpiquéin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.vgrainboutefin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vgrainboutefin.move(350, 502)
            self.vgrainboutefin.resize(100, 20)
            self.vgrainboutefin.setSpecialValueText('  ')
            self.vgrainboutefin.setFont(self.txtpsfont)
            self.vgrainboutefin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")

            self.vautrecéréalin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=False)
            self.vautrecéréalin.move(350, 565)
            self.vautrecéréalin.resize(100, 20)
            self.vautrecéréalin.setSpecialValueText(' ')
            self.vautrecéréalin.setFont(self.txtpsfont)
            self.vautrecéréalin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchtéin = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vgrainmouchtéin.move(350, 418)
            self.vgrainmouchtéin.resize(100, 20)
            self.vgrainmouchtéin.setRange(0, 5)
            self.vgrainmouchtéin.setSpecialValueText(' ')
            self.vgrainmouchtéin.setFont(self.txtpsfont)
            self.vgrainmouchtéin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.vgrainetrangéin = QtWidgets.QDoubleSpinBox(self.entrétab)
            self.vgrainetrangéin.move(350, 586)
            self.vgrainetrangéin.resize(100, 20)
            self.vgrainetrangéin.setRange(0, 3)
            self.vgrainetrangéin.setSpecialValueText(' ')
            self.vgrainetrangéin.setFont(self.txtpsfont)
            self.vgrainetrangéin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.vtotaldemin = QtWidgets.QDoubleSpinBox(self.entrétab, readOnly=True)
            self.vtotaldemin.move(350, 544)
            self.vtotaldemin.resize(100, 20)
            self.vtotaldemin.setSpecialValueText(' ')
            self.vtotaldemin.setFont(self.txtpsfont)
            self.vtotaldemin.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################observation###############
            self.matricullisttxt = QtWidgets.QLabel("Matricule de camion", self.entrétab)
            self.matricullisttxt.move(640, 155)
            self.matricullisttxt.resize(140, 20)
            self.matricullisttxt.setFont(self.txtpsfont)
            self.matricullisttxt.setFont(self.font)

            self.observation = QtWidgets.QLabel("Observation", self.entrétab)
            self.observation.move(470, 155)
            self.observation.resize(100, 20)
            self.observation.setFont(self.txtpsfont)
            self.observation.setFont(self.font)
            ##################################################txtobservation&matricul#################################
            self.matricullist = QtWidgets.QTextEdit(self.entrétab)
            self.matricullist.move(640, 187)
            self.matricullist.resize(150, 417)

            self.matricullist.setStyleSheet(
                "background-color:#FDFEFE;color:#000000;border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            self.matricullist.setFont(self.font)

            self.txtobservation = QtWidgets.QTextEdit(self.entrétab)
            self.txtobservation.move(470, 187)
            self.txtobservation.resize(150, 417)
            self.txtobservation.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelccls = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.entrétab)
            self.labelccls.move(500, 0)
            self.labelccls.resize(438, 80)
            self.labelccls.setFont(self.font)
            self.labelccls.setStyleSheet(
                "background-color: #ffffff; border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            self.LABELBULLETIN = QtWidgets.QLabel("<H2>BULLETIN D’AGREAGE</H2>", self.entrétab)
            self.LABELBULLETIN.move(650, 35)
            self.LABELBULLETIN.resize(180, 20)
            self.LABELBULLETIN.setStyleSheet("background-color: #ffffff")

            self.bletendretxt = QtWidgets.QLabel("<H2>ENTRE<H2/>", self.entrétab)
            self.bletendretxt.move(690, 55)
            self.bletendretxt.resize(140, 20)
            self.bletendretxt.setStyleSheet("background-color: #ffffff")

            #############################################date edit#############################################
            self.daydate = QDate.currentDate()
            self.dateeditetxt = QtWidgets.QLabel("Date:", self.entrétab)
            self.dateeditetxt.move(30, 10)
            self.dateeditetxt.resize(115, 30)
            self.dateeditetxt.setFont((self.font))
            self.dateedite = QtWidgets.QDateEdit(self.entrétab)
            self.dateedite.move(100, 20)
            self.dateedite.resize(130, 20)
            self.dateedite.setDate(self.daydate)
            self.dateedite.setStyleSheet(
                " background-color: #FDFEFE;font-size: 12px;padding: 1px;border-style: solid;border: 1px solid #76797C;border-radius: 0px;color: #000000;")
            self.dateedite.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################
            self.matricul = QtWidgets.QLabel("Matricule de camion: ", self.entrétab)
            self.matricul.move(30, 110)
            self.matricul.resize(140, 20)
            self.matricul.setFont(self.font)
            self.matriculedite = QtWidgets.QLineEdit(self.entrétab)
            self.matriculedite.move(160, 110)
            self.matriculedite.resize(130, 20)
            self.matriculedite.setInputMask("99999-999-99")

            self.matriculedite.setFont(self.font)
            ##################################################quantite###############################################
            self.quantitetxt = QtWidgets.QLabel("Quantité(QX):", self.entrétab)
            self.quantitetxt.move(860, 100)
            self.quantitetxt.resize(85, 40)
            self.quantitetxt.setFont(self.font)
            self.quantite = QtWidgets.QLineEdit("", self.entrétab)
            self.quantite.move(950, 110)
            self.quantite.resize(130, 20)
            self.quantite.setValidator(QDoubleValidator(0.99, 99.99, 2))
            self.quantite.setFont(self.font)

            ####################################################éspece###########################
            self.éspéce = QtWidgets.QLabel("Espèce :", self.entrétab)
            self.éspéce.move(30, 50)
            self.éspéce.resize(120, 40)
            self.éspéce.setFont(self.font)
            self.éspécecombo = QtWidgets.QComboBox(self.entrétab)
            self.éspécecombo.addItem("")
            self.éspécecombo.addItem("Blé Dur IMP")
            self.éspécecombo.addItem("Blé Tendre IMP")
            self.éspécecombo.move(100, 60)
            self.éspécecombo.resize(130, 20)
            self.éspécecombo.setFont(self.font)

            ##########################################Nom de l’acheteur : confirmite###########################################
            self.provance = QtWidgets.QLabel("Provenance:", self.entrétab)
            self.provance.move(357, 100)
            self.provance.resize(118, 40)
            self.provance.setFont(self.font)
            self.provancetecombo = QtWidgets.QComboBox(self.entrétab, editable=True)
            self.provancetecombo.addItem("")
            self.provancetecombo.addItem("UCA ORAN")
            self.provancetecombo.addItem("UCA ALGER")
            self.provancetecombo.addItem("UCA MOSTAGANEM")
            self.provancetecombo.addItem("UCA SKIKDA")
            self.provancetecombo.addItem("UCA BEJAIA")
            self.provancetecombo.addItem("CCLS TELEMCEN")
            self.provancetecombo.addItem("CCLS TENES")
            self.provancetecombo.addItem("CCLS BEROUAGHIA")
            self.provancetecombo.addItem("CCLS LAMTAR")
            self.provancetecombo.addItem("")
            self.provancetecombo.addItem("")
            self.provancetecombo.move(440, 110)
            self.provancetecombo.resize(150, 20)
            self.provancetecombo.setFont(self.font)

            #####################################################Point de collecte : #######################################################
            self.lieutxt = QtWidgets.QLabel("Lieu de la livraison:", self.entrétab)
            self.lieutxt.move(597, 100)
            self.lieutxt.resize(140, 40)
            self.lieutxt.setFont(self.font)
            self.lieucombo = QtWidgets.QComboBox(self.entrétab, editable=True)
            self.lieucombo.addItem("")
            self.lieucombo.addItem("CCLS RELIZANE")
            self.lieucombo.move(725, 110)
            self.lieucombo.resize(130, 20)
            self.lieucombo.setFont(self.font)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeur = QtWidgets.QLabel("Nom de l’Agréeur:", self.entrétab)
            self.agréeeur.move(1085, 100)
            self.agréeeur.resize(112, 40)
            self.agréeeur.setFont(self.font)
            self.agréeeurcombo = QtWidgets.QComboBox(self.entrétab, editable=True)
            self.agréeeurcombo.addItem("")
            self.agréeeurcombo.addItem("FELOUAH OMAR")
            self.agréeeurcombo.addItem("BEKHEDDA AEK")
            self.agréeeurcombo.addItem("BENAISSA YOUCEF")
            self.agréeeurcombo.addItem("REZZAG SOFIANE ")
            self.agréeeurcombo.addItem("BELBACHA M.NADIR")
            self.agréeeurcombo.move(1200, 110)
            self.agréeeurcombo.resize(130, 20)
            self.agréeeurcombo.setFont(self.font)

            ############################################docx2pdf######################
            ###########################buttons################

            self.btnsavebt = QtWidgets.QPushButton("ENREGISTRER", self.entrétab, clicked=lambda: self.save_doc())
            self.btnsavebt.move(832, 187)
            self.btnsavebt.resize(500, 80)
            self.btnsavebt.setFont(self.font)
            self.btnsavebt.setIcon(QIcon("images/savepis.png"))
            self.btnsavebt.setIconSize(QSize(70, 80))

            self.btnprintbt = QtWidgets.QPushButton("IMPRIMER", self.entrétab, clicked=lambda: self.printerin())
            self.btnprintbt.move(832, 292)
            self.btnprintbt.resize(500, 80)
            self.btnprintbt.setFont(self.font)
            self.btnprintbt.setIcon(QIcon("images/print125.png"))
            self.btnprintbt.setIconSize(QSize(70, 80))

            self.btnefacebt = QtWidgets.QPushButton("EFACER", self.entrétab, clicked=lambda: self.clear_all())
            self.btnefacebt.move(832, 397)
            self.btnefacebt.resize(500, 80)
            self.btnefacebt.setIcon(QIcon("images/eraser45877.png"))
            self.btnefacebt.setIconSize(QSize(70, 80))
            self.btnefacebt.setFont(self.font)

            self.btnmat = QtWidgets.QPushButton("Ajouté", self.entrétab, clicked=lambda: self.matriculbtnadd())
            self.btnmat.move(292, 110)
            self.btnmat.resize(30, 20)
            #############################################BLE DUR
            ########################################################
            #####################################################################
            ###############################################################################
            ##############################################################################################
            self.sortietab = QtWidgets.QWidget()
            self.sortietab.setObjectName("sortietab")
            self.confirmitewidget.addTab(self.sortietab, "")
            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétreso = QtWidgets.QLabel("Paramètre recherchés", self.sortietab)
            self.paramétreso.move(30, 122)
            self.paramétreso.resize(80, 20)
            self.paramétreso.setFont(self.font)
            self.txtpsfontso = QtGui.QFont()
            self.txtpsfontso.setBold(True)
            self.txtpsfontso.setPointSize(9)
            ######################Limites(sans bon ni réf)################
            self.pinsect = QtWidgets.QLineEdit("Présence d’insectes(morts/vivants) : ", self.sortietab, readOnly=True)
            self.pinsect.resize(319, 20)
            self.pinsect.move(30, 145)
            self.pinsect.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.pinsect.setFont(self.txtpsfont)
            self.pinsect.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")

            self.einsect = QtWidgets.QLineEdit("Espèce d’insecte : ", self.sortietab, readOnly=True)
            self.einsect.resize(319, 20)
            self.einsect.move(30, 166)
            self.einsect.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.einsect.setFont(self.txtpsfont)
            self.einsect.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")

            self.psso = QtWidgets.QLineEdit("Poids spécifique (kg/hl): ", self.sortietab, readOnly=True)
            self.psso.resize(319, 20)
            self.psso.move(30, 187)
            self.psso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.psso.setFont(self.txtpsfont)
            self.psso.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.metadinetblt = QtWidgets.QLineEdit("Mitadin et blé tendre(%):", self.sortietab, readOnly=True)
            self.metadinetblt.resize(319, 20)
            self.metadinetblt.move(30, 208)
            self.metadinetblt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.metadinetblt.setFont(self.txtpsfont)
            #########################Graines nuisibles (%)##########
            self.bltdanbld = QtWidgets.QLineEdit("Blé tendre dans le blé dur(%):", self.sortietab, readOnly=True)
            self.bltdanbld.resize(319, 20)
            self.bltdanbld.move(30, 229)
            self.bltdanbld.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.bltdanbld.setFont(self.txtpsfont)

            self.mitadin = QtWidgets.QLineEdit("Mitadin(%):", self.sortietab, readOnly=True)
            self.mitadin.resize(319, 20)
            self.mitadin.move(30, 250)
            self.mitadin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.mitadin.setFont(self.txtpsfont)

            self.gravar = QtWidgets.QLineEdit("Grains avariés (%):", self.sortietab, readOnly=True)
            self.gravar.resize(319, 20)
            self.gravar.move(30, 586)
            self.gravar.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.gravar.setFont(self.txtpsfont)

            self.autrecéréalso = QtWidgets.QLineEdit("Autres céréales(%):", self.sortietab, readOnly=True)
            self.autrecéréalso.resize(319, 20)
            self.autrecéréalso.move(30, 481)
            self.autrecéréalso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.autrecéréalso.setFont(self.txtpsfont)

            self.grainmauvaise = QtWidgets.QLineEdit("Graines de mauvaises herbes:", self.sortietab, readOnly=True)
            self.grainmauvaise.resize(319, 20)
            self.grainmauvaise.move(30, 271)
            self.grainmauvaise.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainmauvaise.setFont(self.txtpsfont)

            self.grainnuisibleso = QtWidgets.QLineEdit("Graines nuisibles(%):", self.sortietab, readOnly=True)
            self.grainnuisibleso.resize(319, 20)
            self.grainnuisibleso.move(30, 397)
            self.grainnuisibleso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisibleso.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvéso = QtWidgets.QLineEdit("Débris végétaux(%):  ", self.sortietab, readOnly=True)
            self.débrisvéso.resize(319, 20)
            self.débrisvéso.move(30, 334)
            self.débrisvéso.setFont(self.txtpsfont)
            self.débrisvéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrtso = QtWidgets.QLineEdit("Matière inerte(%):", self.sortietab, readOnly=True)
            self.matiéreinrtso.resize(319, 20)
            self.matiéreinrtso.move(30, 313)
            self.matiéreinrtso.setFont(self.txtpsfont)
            self.matiéreinrtso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################

            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleurso = QtWidgets.QLineEdit("Grains sans valeur(%):", self.sortietab, readOnly=True)
            self.grainsanvaleurso.resize(319, 20)
            self.grainsanvaleurso.move(30, 355)
            self.grainsanvaleurso.setFont(self.txtpsfont)
            self.grainsanvaleurso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:   ", self.sortietab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 376)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalpremso = QtWidgets.QLineEdit("Impuretés 1ere catégorie: ", self.sortietab, readOnly=True)
            self.totalpremso.resize(319, 20)
            self.totalpremso.move(30, 292)
            self.totalpremso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalpremso.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincasséso = QtWidgets.QLineEdit("Grains cassés(%):   ≤2", self.sortietab, readOnly=True)
            self.graincasséso.move(30, 439)
            self.graincasséso.resize(319, 20)
            self.graincasséso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincasséso.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaudeso = QtWidgets.QLineEdit("Gains échaudés/maigres(%):   ", self.sortietab, readOnly=True)
            self.grainechaudeso.move(30, 460)
            self.grainechaudeso.resize(319, 20)
            self.grainechaudeso.setFont(self.txtpsfont)
            self.grainechaudeso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################

            ##########################################################Grains germés (%)###################################################
            self.graigerméso = QtWidgets.QLineEdit("Grains germés(%):", self.sortietab, readOnly=True)
            self.graigerméso.move(30, 418)
            self.graigerméso.resize(319, 20)
            self.graigerméso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigerméso.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaiséso = QtWidgets.QLineEdit("Grain punaisés(%): ", self.sortietab, readOnly=True)
            self.grainpunaiséso.move(30, 544)
            self.grainpunaiséso.resize(319, 20)
            self.grainpunaiséso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaiséso.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiquéso = QtWidgets.QLineEdit("Grains piqués(%):  ", self.sortietab, readOnly=True)
            self.grainpiquéso.move(30, 565)
            self.grainpiquéso.resize(319, 20)
            self.grainpiquéso.setFont(self.txtpsfont)
            self.grainpiquéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.grainboutéso = QtWidgets.QLineEdit("Grains boutés (%):", self.sortietab, readOnly=True)
            self.grainboutéso.move(30, 523)
            self.grainboutéso.resize(319, 20)
            self.grainboutéso.setFont(self.txtpsfont)
            self.grainboutéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchtéso = QtWidgets.QLineEdit("Grains fortement mouchetés (%):", self.sortietab, readOnly=True)
            self.grainmouchtéso.move(30, 502)
            self.grainmouchtéso.resize(319, 20)
            self.grainmouchtéso.setFont(self.txtpsfont)
            self.grainmouchtéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldemso = QtWidgets.QLineEdit("Impuretés 2eme catégorie", self.sortietab, readOnly=True)
            self.totaldemso.move(30, 418)
            self.totaldemso.resize(319, 20)
            self.totaldemso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldemso.setFont(self.txtpsfont)
            #########################################################indice notin##################################################

            #######################blétendre dans blé dur########

            #################label valeure##############
            self.valeurso = QtWidgets.QLabel("Résultat", self.sortietab)
            self.valeurso.move(350, 122)
            self.valeurso.resize(100, 20)
            self.valeurso.setFont(self.font)
            ######################Limites(sans bon ni réf)################

            self.vpinsect = QtWidgets.QLineEdit(self.sortietab)
            self.vpinsect.resize(100, 20)
            self.vpinsect.move(350, 145)
            self.vpinsect.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vpinsect.setFont(self.txtpsfont)
            self.vpinsect.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")

            self.veinsect = QtWidgets.QLineEdit(self.sortietab)
            self.veinsect.resize(100, 20)
            self.veinsect.move(350, 166)
            self.veinsect.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.veinsect.setFont(self.txtpsfont)
            self.veinsect.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")

            self.vpsso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vpsso.setSpecialValueText(" ")
            self.vpsso.resize(100, 20)
            self.vpsso.move(350, 187)
            self.vpsso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vpsso.setFont(self.txtpsfont)
            self.vpsso.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.vmetadinetblt = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vmetadinetblt.setSpecialValueText(" ")
            self.vmetadinetblt.resize(100, 20)
            self.vmetadinetblt.move(350, 208)
            self.vmetadinetblt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vmetadinetblt.setFont(self.txtpsfont)
            #########################Graines nuisibles (%)##########
            self.vbltdanbld = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vbltdanbld.setSpecialValueText(" ")
            self.vbltdanbld.resize(100, 20)
            self.vbltdanbld.move(350, 229)
            self.vbltdanbld.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vbltdanbld.setFont(self.txtpsfont)

            self.vmitadin = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vmitadin.setSpecialValueText(" ")
            self.vmitadin.resize(100, 20)
            self.vmitadin.move(350, 250)
            self.vmitadin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vmitadin.setFont(self.txtpsfont)

            self.vgravar = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgravar.setSpecialValueText(" ")
            self.vgravar.resize(100, 20)
            self.vgravar.move(350, 586)
            self.vgravar.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgravar.setFont(self.txtpsfont)

            self.vautrecéréalso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vautrecéréalso.setSpecialValueText(" ")
            self.vautrecéréalso.resize(100, 20)
            self.vautrecéréalso.move(350, 481)
            self.vautrecéréalso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vautrecéréalso.setFont(self.txtpsfont)

            self.vgrainmauvaise = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainmauvaise.setSpecialValueText(" ")
            self.vgrainmauvaise.resize(100, 20)
            self.vgrainmauvaise.move(350, 271)
            self.vgrainmauvaise.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgrainmauvaise.setFont(self.txtpsfont)

            self.vgrainnuisibleso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainnuisibleso.setSpecialValueText(" ")
            self.vgrainnuisibleso.resize(100, 20)
            self.vgrainnuisibleso.move(350, 397)
            self.vgrainnuisibleso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgrainnuisibleso.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.vdébrisvéso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vdébrisvéso.setSpecialValueText(" ")
            self.vdébrisvéso.resize(100, 20)
            self.vdébrisvéso.move(350, 334)
            self.vdébrisvéso.setFont(self.txtpsfont)
            self.vdébrisvéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrtso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vmatiéreinrtso.setSpecialValueText(" ")
            self.vmatiéreinrtso.resize(100, 20)
            self.vmatiéreinrtso.move(350, 313)
            self.vmatiéreinrtso.setFont(self.txtpsfont)
            self.vmatiéreinrtso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################

            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleurso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainsanvaleurso.setSpecialValueText(" ")
            self.vgrainsanvaleurso.resize(100, 20)
            self.vgrainsanvaleurso.move(350, 355)
            self.vgrainsanvaleurso.setFont(self.txtpsfont)
            self.vgrainsanvaleurso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarréso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgraincarréso.setSpecialValueText(" ")
            self.vgraincarréso.resize(100, 20)
            self.vgraincarréso.move(350, 376)
            self.vgraincarréso.setFont(self.txtpsfont)
            self.vgraincarréso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.vtotalpremso = QtWidgets.QDoubleSpinBox(self.sortietab, readOnly=True)
            self.vtotalpremso.setSpecialValueText(" ")
            self.vtotalpremso.resize(100, 20)
            self.vtotalpremso.move(350, 292)
            self.vtotalpremso.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vtotalpremso.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.vgraincasséso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgraincasséso.setSpecialValueText(" ")
            self.vgraincasséso.move(350, 439)
            self.vgraincasséso.resize(100, 20)
            self.vgraincasséso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgraincasséso.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaudeso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainechaudeso.setSpecialValueText(" ")
            self.vgrainechaudeso.move(350, 460)
            self.vgrainechaudeso.resize(100, 20)
            self.vgrainechaudeso.setFont(self.txtpsfont)
            self.vgrainechaudeso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################

            ##########################################################Grains germés (%)###################################################
            self.vgraigerméso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgraigerméso.setSpecialValueText(" ")
            self.vgraigerméso.move(350, 418)
            self.vgraigerméso.resize(100, 20)
            self.vgraigerméso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgraigerméso.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaiséso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainpunaiséso.setSpecialValueText(" ")
            self.vgrainpunaiséso.move(350, 544)
            self.vgrainpunaiséso.resize(100, 20)
            self.vgrainpunaiséso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vgrainpunaiséso.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiquéso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainpiquéso.setSpecialValueText(" ")
            self.vgrainpiquéso.move(350, 565)
            self.vgrainpiquéso.resize(100, 20)
            self.vgrainpiquéso.setFont(self.txtpsfont)
            self.vgrainpiquéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.vgrainboutéso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainboutéso.setSpecialValueText(" ")
            self.vgrainboutéso.move(350, 523)
            self.vgrainboutéso.resize(100, 20)
            self.vgrainboutéso.setFont(self.txtpsfont)
            self.vgrainboutéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchtéso = QtWidgets.QDoubleSpinBox(self.sortietab)
            self.vgrainmouchtéso.setSpecialValueText(" ")
            self.vgrainmouchtéso.move(350, 502)
            self.vgrainmouchtéso.resize(100, 20)
            self.vgrainmouchtéso.setFont(self.txtpsfont)
            self.vgrainmouchtéso.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.vtotaldemso = QtWidgets.QDoubleSpinBox(self.sortietab, readOnly=True)
            self.vtotaldemso.setSpecialValueText("  ")
            self.vtotaldemso.move(350, 418)
            self.vtotaldemso.resize(100, 20)
            self.vtotaldemso.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.vtotaldemso.setFont(self.txtpsfont)
            #########################################################observation###############
            self.observationso = QtWidgets.QLabel("Observation", self.sortietab)
            self.observationso.move(470, 122)
            self.observationso.resize(100, 20)
            self.observationso.setFont(self.txtpsfont)
            self.observationso.setFont(self.font)
            ##################################################txtobservation##################################
            self.txtobservationso = QtWidgets.QTextEdit("<h2><h2/>  <h2><h2/>  <h2><h2/> <h2><h2/>   <h3><h3/>",
                                                        self.sortietab)
            self.txtobservationso.move(470, 145)
            self.txtobservationso.resize(300, 460)
            self.txtobservationso.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelcclsso = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.sortietab)
            self.labelcclsso.move(500, 0)
            self.labelcclsso.resize(438, 80)
            self.labelcclsso.setFont(self.font)
            self.labelcclsso.setStyleSheet(
                "background-color: #ffffff; border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            self.LABELBULLETINso = QtWidgets.QLabel("<H2>BULLETIN CONFORMITE</H2>", self.sortietab)
            self.LABELBULLETINso.move(650, 35)
            self.LABELBULLETINso.resize(180, 20)
            self.LABELBULLETINso.setStyleSheet("background-color: #ffffff")
            self.bledurtxt = QtWidgets.QLabel("<H2>SORTIE<H2/>", self.sortietab)
            self.bledurtxt.move(698, 55)
            self.bledurtxt.resize(120, 20)
            self.bledurtxt.setStyleSheet("background-color: #ffffff")

            #############################################date edit#############################################
            self.dateediteso = QtWidgets.QDateEdit(self.sortietab)
            self.dateediteso.setDate(self.daydate)
            self.dateediteso.move(100, 10)
            self.dateediteso.setStyleSheet(
                " background-color: #FDFEFE;padding: 1px;font-size: 12px;border-style: solid;border: 1px solid #76797C;border-radius: 0px;color: #000000;")
            self.dateediteso.resize(130, 20)
            self.dateediteso.setFont(self.font)

            self.dateeditesotxt = QtWidgets.QLabel("Date:", self.sortietab)
            self.dateeditesotxt.move(30, 10)
            self.dateeditesotxt.resize(60, 20)
            self.dateeditesotxt.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################
            self.matriculso = QtWidgets.QLabel("Matricul camion:", self.sortietab)
            self.matriculso.move(30, 90)
            self.matriculso.resize(130, 20)
            self.matriculso.setFont(self.font)
            self.matriculcomboso = QtWidgets.QLineEdit(self.sortietab)
            self.matriculcomboso.setInputMask("99999-999-99")
            self.matriculcomboso.move(140, 90)
            self.matriculcomboso.resize(130, 20)

            ##################################################quantite###############################################
            self.quantiteso = QtWidgets.QLabel("Quantité(QX):", self.sortietab)
            self.quantiteso.move(720, 80)
            self.quantiteso.resize(85, 40)
            self.quantiteso.setFont(self.font)
            self.quantitetxtso = QtWidgets.QLineEdit("", self.sortietab)
            self.quantitetxtso.move(820, 90)
            self.quantitetxtso.resize(130, 20)
            self.quantitetxtso.setValidator(QDoubleValidator(0.99, 99.99, 2))

            ####################################################éspece###########################
            self.éspéceso = QtWidgets.QLabel("Espèce :", self.sortietab)
            self.éspéceso.move(30, 35)
            self.éspéceso.resize(60, 40)
            self.éspéceso.setFont(self.font)
            self.éspécecomboso = QtWidgets.QComboBox(self.sortietab)
            self.éspécecomboso.addItem("Blé Dur")
            self.éspécecomboso.move(100, 45)
            self.éspécecomboso.resize(130, 20)

            ##########################################Nom de l’acheteur : confirmite###########################################
            self.distination = QtWidgets.QLabel("Distination:", self.sortietab)
            self.distination.move(280, 80)
            self.distination.resize(118, 40)
            self.distination.setFont(self.font)
            self.distinationcomboso = QtWidgets.QComboBox(self.sortietab, editable=True)
            self.distinationcomboso.addItem("")
            self.distinationcomboso.addItem("CCLS BLIDA ")
            self.distinationcomboso.addItem("UCA MOSTAGANEM")
            self.distinationcomboso.addItem("CCLS B.BOU ARRERIDJ")
            self.distinationcomboso.addItem("")
            self.distinationcomboso.addItem("")
            self.distinationcomboso.addItem("")
            self.distinationcomboso.addItem("")
            self.distinationcomboso.addItem("")
            self.distinationcomboso.move(360, 90)
            self.distinationcomboso.resize(138, 20)

            #####################################################Point de collecte : #######################################################
            self.pointcollecteso = QtWidgets.QLabel("Expéditeur:", self.sortietab)
            self.pointcollecteso.move(500, 80)
            self.pointcollecteso.resize(106, 40)
            self.pointcollecteso.setFont(self.font)
            self.pointcollectecomboso = QtWidgets.QComboBox(self.sortietab)
            self.pointcollectecomboso.addItem("")
            self.pointcollectecomboso.addItem("CCLS RELIZANE")
            self.pointcollectecomboso.move(580, 90)
            self.pointcollectecomboso.resize(130, 20)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeurso = QtWidgets.QLabel("Nom de l’Agréeur:", self.sortietab)
            self.agréeeurso.move(960, 80)
            self.agréeeurso.resize(112, 40)
            self.agréeeurso.setFont(self.font)
            self.agréeeurcomboso = QtWidgets.QComboBox(self.sortietab, editable=False)
            self.agréeeurcomboso.addItem("")
            self.agréeeurcomboso.addItem("FELOUAH OMAR")
            self.agréeeurcomboso.addItem("BEKHEDDA AEK")
            self.agréeeurcomboso.addItem("BENAISSA YOUCEF")
            self.agréeeurcomboso.addItem("REZZAG SOFIANE ")
            self.agréeeurcomboso.addItem("BELBACHA M.NADIR")
            self.agréeeurcomboso.move(1090, 90)
            self.agréeeurcomboso.resize(130, 20)

            ############################################docx2pdf######################
            self.aab = QTimer()
            self.aab.timeout.connect(self.plus_value)
            self.aab.setInterval(1000)
            self.aab.start()

            self.aa = QTimer()
            self.aa.timeout.connect(self.plus_sortie)
            self.aa.setInterval(1000)
            self.aa.start()

            ###########################buttons################

            self.btnsaveso = QtWidgets.QPushButton("ENREGISTRER", self.sortietab,
                                                   clicked=lambda: self.save_doc_sortie())
            self.btnsaveso.move(790, 145)
            self.btnsaveso.resize(430, 80)
            self.btnsaveso.setFont(self.font)
            self.btnsaveso.setIcon(QIcon("images/savepis.png"))
            self.btnsaveso.setIconSize(QSize(70, 80))

            self.btnprintso = QtWidgets.QPushButton("IMPRIMER", self.sortietab, clicked=lambda: self.printer_sortie())
            self.btnprintso.move(790, 265)
            self.btnprintso.resize(430, 80)
            self.btnprintso.setFont(self.font)
            self.btnprintso.setIcon(QIcon("images/print125.png"))
            self.btnprintso.setIconSize(QSize(70, 80))

            self.btnefaceso = QtWidgets.QPushButton("EFACER", self.sortietab, clicked=lambda: self.clear_sortie())
            self.btnefaceso.move(790, 385)
            self.btnefaceso.resize(430, 80)
            self.btnefaceso.setIcon(QIcon("images/eraser45877.png"))
            self.btnefaceso.setIconSize(QSize(70, 80))
            self.btnefaceso.setFont(self.font)

            ###############################################################################################################
            ###########################################################################################################################
            ########################################################sortie poure moulin
            self.bltendretab = QtWidgets.QWidget()
            self.bltendretab.setStyleSheet("""QToolTip
            {
                border: 1px solid #76797C;
                background-color: rgb(90, 102, 117);;
                color: white;
                padding: 5px;
                opacity: 200;
            }

            QWidget
            {
                color: #000000;
                background-color: #ffffff;
                selection-background-color:#3daee9;
                selection-color: #3daee9;
                background-clip: border;
                border-image: none;
                border: 0px transparent black;
                outline: 0;
            }

            QWidget:item:hover
            {
                background-color: #3daee9;
                color: #eff0f1;
            }

            QWidget:item:selected
            {
                background-color: #3daee9;
            }



            QWidget:disabled
            {
                color: #454545;
                background-color: #31363b;
            }

            QAbstractItemView
            {
                alternate-background-color: #31363b;
                color: #eff0f1;
                border: 1px solid 3A3939;
                border-radius: 2px;
            }

            QWidget:focus, QMenuBar:focus
            {
                border: 1px solid #3daee9;
            }

            QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
            {
                border: none;
            }

            QLineEdit
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color: #000000;
            }
            QDoubleSpinBox
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color:#000000;
                font-size: 11px;
                font-weight: bold;

            }
            QDoubleSpinBox:focus{
                background-color: #FDFEFE;
                border-style: solid;
                border: 2px solid #76797C;
                border-radius: 4px;
                border-color: #ff8c00;
            }
            QDoubleSpinBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 1px;

                border-left-width: 0px;
                border-left-color: #232629;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }



            QGroupBox {
                border:1px solid #76797C;
                border-radius: 2px;
                margin-top: 5px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding-left: 4px;
                padding-right: 4px;
                padding-top: 4px;
            }

            QAbstractScrollArea
            {
                border-radius: 2px;
                border: 1px solid #76797C;
                background-color: transparent;
            }

            QScrollBar:horizontal
            {
                height: 15px;
                margin: 3px 15px 3px 15px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
                background-color: #2A2929;
            }

            QScrollBar::handle:horizontal
            {
                background-color: #605F5F;
                min-width: 5px;
                border-radius: 4px;
            }

            QScrollBar::add-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
                width: 10px;
                height: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }


            QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
            {
                background: none;
            }


            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
            {
                background: none;
            }

            QScrollBar:vertical
            {
                background-color: #2A2929;
                width: 15px;
                margin: 15px 3px 15px 3px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
            }

            QScrollBar::handle:vertical
            {
                background-color: #605F5F;
                min-height: 5px;
                border-radius: 4px;
            }

            QScrollBar::sub-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
            {

                border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }


            QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
            {
                border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
            {
                background: none;
            }


            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
            {
                background: none;
            }

            QTextEdit
            {
                background-color: #FDFEFE;
                color: #000000;
                border: 1px solid #76797C;
            }

            QPlainTextEdit
            {
                background-color: #232629;;
                color: #000000;
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QHeaderView::section
            {
                background-color: #76797C;
                color: #eff0f1;
                padding: 1px;
                border: 1px solid #76797C;
            }

            QSizeGrip {
                width: 12px;
                height: 12px;
            }


            QMainWindow::separator
            {
                background-color: #31363b;
                color: white;
                padding-left: 4px;
                spacing: 2px;
                border: 1px dashed #76797C;
            }

            QMainWindow::separator:hover
            {

                background-color: #787876;
                color: white;
                padding-left: 4px;
                border: 1px solid #76797C;
                spacing: 2px;
            }


            QMenu::separator
            {
                height: 1px;
                background-color: #76797C;
                color: white;
                padding-left: 4px;
                margin-left: 10px;
                margin-right: 5px;
            }


            QFrame
            {
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QFrame[frameShape="0"]
            {
                border-radius: 2px;
                border: 1px transparent #76797C;
            }

            QStackedWidget
            {
                border: 1px transparent black;
            }


            QPushButton
            {
                color: #00000;
                background-color:#ade3e7;
                border-width: 1px;
                border-color: #1e1e1e;
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
            }

            QComboBox
            {
               background-color: #FDFEFE;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                min-width: 40px;
            }

            QPushButton:checked{
                background-color: #76797C;
                border-color: #6A6969;
            }

            QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
            {
                border: 1px solid #ff8c00;
                color: #000000;
            }

            QComboBox:on
            {
                padding-top: 1px;
                padding-left: 1px;
                selection-background-color: #FDFEFE;
            }

            QComboBox QAbstractItemView
            {
                background-color: #FDFEFE;
                border-radius: 2px;
                border: 1px solid #76797C;
                color:#000000;
                selection-background-color: #000000;
            }

            QComboBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;

                border-left-width: 0px;
                border-left-color: ff8c00;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }


            QLabel
            {
                border: 2px solid black;
            }

            QTabWidget{
                border: 0px transparent black;
            }

            QTabWidget::pane {
                border: 1px solid #76797C;
                padding: 5px;
                margin: 0px;
            }

            QTabBar
            {
                qproperty-drawBase: 0;
                left: 5px; /* move to the right by 5px */
                border-radius: 3px;
            }

            QTabBar:focus
            {
                border: 0px transparent black;
            }

            QTabBar::close-button  {
                image: url(:/qss_icons/Dark_rc/close.png);
                background: transparent;
            }

            QTabBar::close-button:hover
            {
                image: url(:/qss_icons/Dark_rc/close-hover.png);
                background: transparent;
            }

            QTabBar::close-button:pressed {
                image: url(:/qss_icons/Dark_rc/close-pressed.png);
                background: transparent;
            }

            /* TOP TABS */
            QTabBar::tab:top {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                min-width: 10px;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;
            }

            QTabBar::tab:top:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;    
            }

            QTabBar::tab:top:!selected:hover {
                background-color: #3daee9;
            }

            /* BOTTOM TABS */
            QTabBar::tab:bottom {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
                min-width: 10px;
            }

            QTabBar::tab:bottom:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:bottom:!selected:hover {
                background-color: #3daee9;
            }

            /* LEFT TABS */
            QTabBar::tab:left {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:left:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:left:!selected:hover {
                background-color: #3daee9;
            }


            /* RIGHT TABS */
            QTabBar::tab:right {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:right:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
            }





            QSlider::groove:horizontal {
                border: 1px solid #565a5e;
                height: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 2px;
            }

            QSlider::handle:horizontal {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: -8px 0;
                border-radius: 9px;
            }

            QSlider::groove:vertical {
                border: 1px solid #565a5e;
                width: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 3px;
            }

            QSlider::handle:vertical {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: 0 -8px;
                border-radius: 9px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px transparent #76797C;
                border-radius: 2px;
                margin: 3px;
                padding: 5px;
            }

            QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
             padding-right: 20px; /* make way for the popup button */
             border: 1px #76797C;
             border-radius: 5px;
            }

            QToolButton[popupMode="2"] { /* only for InstantPopup */
             padding-right: 10px; /* make way for the popup button */
             border: 1px #76797C;
            }


            QToolButton:hover, QToolButton::menu-button:hover {
                background-color: transparent;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            QToolButton:checked, QToolButton:pressed,
                    QToolButton::menu-button:pressed {
                background-color: #3daee9;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
            QToolButton::menu-indicator {
                background-color:ff8c00;
                top: -7px; left: -2px; /* shift it a bit */
            }

            /* the subcontrols below are used only in the MenuButtonPopup mode */
            QToolButton::menu-button {
                border: 1px transparent #76797C;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
                /* 16px width + 4px for border = 20px allocated above */
                width: 16px;
                outline: none;
            }

            QToolButton::menu-arrow {
               background-color:ff8c00;
            }

            QToolButton::menu-arrow:open {
                border: 1px solid #76797C;
            }

            QPushButton::menu-indicator  {
                subcontrol-origin: padding;
                subcontrol-position: bottom right;
                left: 8px;
            }

            QTableView
            {
                border: 1px solid #76797C;
                gridline-color: #31363b;
                background-color: #FDFEFE;
                color:#000000;
            }


            QTableView, QHeaderView
            {
                background-color: #FDFEFE;
                color:#000000;
                border-radius: 0px;
            }

            QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
                background: #FDFEFE;
                color: #000000;
            }

            QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
                background: #3daee9;
                color: #000000;
            }


            QHeaderView
            {
                background-color: #FDFEFE;
                border: 1px transparent;
                border-radius: 0px;
                margin: 0px;
                padding: 0px;

            }

            QHeaderView::section  {
                background-color:#80f1f9;
                color: #000000;
                padding: 5px;
                border: 1px solid #76797C;
                border-radius: 0px;
                text-align: center;
            }

            QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
            {
                border-top: 1px solid #76797C;
            }

            QHeaderView::section::vertical
            {
                border-top: transparent;
            }

            QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
            {
                border-left: 1px solid #76797C;
            }

            QHeaderView::section::horizontal
            {
                border-left: transparent;
            }


            QHeaderView::section:checked
             {
                color: #000000;
                background-color: #3daee9;
             }

             /* style the sort indicator */
            QHeaderView::down-arrow {

            }

            QHeaderView::up-arrow {

            }


            QTableCornerButton::section {
                background-color: #31363b;
                border: 1px transparent #76797C;
                border-radius: 0px;
            }

            QToolBox  {
                padding: 5px;
                border: 1px transparent black;
            }

            QToolBox::tab {
                color: #eff0f1;
                background-color: #31363b;
                border: 1px solid #76797C;
                border-bottom: 1px transparent #31363b;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }

            QToolBox::tab:selected { /* italicize selected tabs */
                font: italic;
                background-color: #31363b;
                border-color: #3daee9;
             }

            QStatusBar::item {
                border: 0px transparent dark;
             }


            QFrame[height="3"], QFrame[width="3"] {
                background-color: #76797C;
            }




            QDateEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;
            }

            QDateEdit:on
            {
                padding-top: 2px;
                padding-left: 2px;
                selection-background-color: #4a4a4a;
            }

            QDateEdit QAbstractItemView
            {
                background-color: #ff8c00;
                border-radius: 2px;
                border: 1px solid #3375A3;
                selection-background-color:ff8c00;
            }

            QDateEdit::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 0px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }   
            QDateTimeEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;

            }    
            """)
            self.bltendretab.setObjectName("bltendretab")
            self.confirmitewidget.addTab(self.bltendretab, "sortie moulin BT")

            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétre = QtWidgets.QLabel("Paramètre", self.bltendretab)
            self.paramétre.move(30, 145)
            self.paramétre.resize(80, 20)
            self.paramétre.setFont(self.font)

            self.txtpsfont = QtGui.QFont()
            self.txtpsfont.setBold(True)
            self.txtpsfont.setPointSize(9)
            ################Limites(sans bon ni réf)###############
            self.valeur = QtWidgets.QLabel("""Limite ssans bon ni réf)""", self.bltendretab)
            self.valeur.move(165, 127)
            self.valeur.resize(200, 55)
            self.valeur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.ps = QtWidgets.QLineEdit("Poids spécifique (kg/hl):   (75.500-75.899)", self.bltendretab,
                                          readOnly=True)
            self.ps.resize(319, 20)
            self.ps.move(30, 167)
            self.ps.setFont(self.txtpsfont)
            self.ps.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.humidite = QtWidgets.QLineEdit("Teneur en eau(%): (13.5-15)", self.bltendretab, readOnly=True)
            self.humidite.resize(319, 20)
            self.humidite.move(30, 188)
            self.humidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.humidite.setFont(self.txtpsfont)

            #######################ergot#########################
            self.ergot = QtWidgets.QLineEdit("Ergo(% :\t<0.001 ", self.bltendretab, readOnly=True)
            self.ergot.resize(319, 20)
            self.ergot.move(30, 209)
            self.ergot.setStyleSheet("background-color: #232629")
            self.ergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.ergot.setFont(self.txtpsfont)

            #########################Graines nuisibles (%)##########
            self.grainnuisible = QtWidgets.QLineEdit("Graines nuisibles(%):\t<0.001", self.bltendretab, readOnly=True)
            self.grainnuisible.resize(319, 20)
            self.grainnuisible.move(30, 230)
            self.grainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisible.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvé = QtWidgets.QLineEdit("Débris végétaux(%):     ", self.bltendretab, readOnly=True)
            self.débrisvé.resize(319, 20)
            self.débrisvé.move(30, 251)
            self.débrisvé.setFont(self.txtpsfont)
            self.débrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrt = QtWidgets.QLineEdit("Matière inerte(%):", self.bltendretab, readOnly=True)
            self.matiéreinrt.resize(319, 20)
            self.matiéreinrt.move(30, 272)
            self.matiéreinrt.setFont(self.txtpsfont)
            self.matiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.grainchaufé = QtWidgets.QLineEdit("Grains chauffés(%):    ", self.bltendretab, readOnly=True)
            self.grainchaufé.resize(319, 20)
            self.grainchaufé.move(30, 293)
            self.grainchaufé.setFont(self.txtpsfont)
            self.grainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleur = QtWidgets.QLineEdit("Grains sans valeur(%):", self.bltendretab, readOnly=True)
            self.grainsanvaleur.resize(319, 20)
            self.grainsanvaleur.move(30, 314)
            self.grainsanvaleur.setFont(self.txtpsfont)
            self.grainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:   ", self.bltendretab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 335)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalprem = QtWidgets.QLineEdit("TOTAL 1ére CAT:     ≤1", self.bltendretab, readOnly=True)
            self.totalprem.resize(319, 20)
            self.totalprem.move(30, 356)
            self.totalprem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalprem.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincassé = QtWidgets.QLineEdit("Grains cassés(%):   ≤2", self.bltendretab, readOnly=True)
            self.graincassé.move(30, 377)
            self.graincassé.resize(319, 20)
            self.graincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincassé.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaude = QtWidgets.QLineEdit("Gains échaudés(%):   ", self.bltendretab, readOnly=True)
            self.grainechaude.move(30, 419)
            self.grainechaude.resize(319, 20)
            self.grainechaude.setFont(self.txtpsfont)
            self.grainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################
            self.grainmaigre = QtWidgets.QLineEdit("Grains maigres(%):", self.bltendretab, readOnly=True)
            self.grainmaigre.move(30, 398)
            self.grainmaigre.resize(319, 20)
            self.grainmaigre.setFont(self.txtpsfont)
            self.grainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##########################################################Grains germés (%)###################################################
            self.graigermé = QtWidgets.QLineEdit("Grains germés(%): ≤2", self.bltendretab, readOnly=True)
            self.graigermé.move(30, 440)
            self.graigermé.resize(319, 20)
            self.graigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigermé.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaisé = QtWidgets.QLineEdit("Grain punaisés(%): ≤1", self.bltendretab, readOnly=True)
            self.grainpunaisé.move(30, 461)
            self.grainpunaisé.resize(319, 20)
            self.grainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaisé.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiqué = QtWidgets.QLineEdit("Grains piqués(%):  ", self.bltendretab, readOnly=True)
            self.grainpiqué.move(30, 482)
            self.grainpiqué.resize(319, 20)
            self.grainpiqué.setFont(self.txtpsfont)
            self.grainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.grainboutef = QtWidgets.QLineEdit("Grains boutés « faible » (%):", self.bltendretab, readOnly=True)
            self.grainboutef.move(30, 503)
            self.grainboutef.resize(319, 20)
            self.grainboutef.setFont(self.txtpsfont)
            self.grainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.grainbouté = QtWidgets.QLineEdit("Grains boutés  « forte » (%):", self.bltendretab, readOnly=True)
            self.grainbouté.move(30, 524)
            self.grainbouté.resize(319, 20)
            self.grainbouté.setFont(self.txtpsfont)
            self.grainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchté = QtWidgets.QLineEdit("Grains mouchetés (%):", self.bltendretab, readOnly=True)
            self.grainmouchté.move(30, 545)
            self.grainmouchté.resize(319, 20)
            self.grainmouchté.setFont(self.txtpsfont)
            self.grainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.grainetrangé = QtWidgets.QLineEdit("Grain étrangers Utilisables pour le bétail (%):  ",
                                                    self.bltendretab, readOnly=True)
            self.grainetrangé.move(30, 566)
            self.grainetrangé.resize(319, 20)
            self.grainetrangé.setFont(self.txtpsfont)
            self.grainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldem = QtWidgets.QLineEdit("Total(%)  Imp2eme cat   ≤5", self.bltendretab, readOnly=True)
            self.totaldem.move(30, 587)
            self.totaldem.resize(319, 20)
            self.totaldem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldem.setFont(self.txtpsfont)
            ###########################################################################################################

            #################label valeure##############
            self.valeur = QtWidgets.QLabel("valeur", self.bltendretab)
            self.valeur.move(350, 144)
            self.valeur.resize(100, 20)
            self.valeur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.vps = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vps.setRange(60, 81.00)
            self.vps.setSpecialValueText(' ')
            self.vps.resize(100, 20)
            self.vps.move(350, 167)
            self.vps.setFont(self.txtpsfont)
            self.vps.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.vhumidite = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vhumidite.setRange(4, 18)
            self.vhumidite.resize(100, 20)
            self.vhumidite.setSpecialValueText(' ')
            self.vhumidite.move(350, 188)
            self.vhumidite.setFont(self.txtpsfont)
            self.vhumidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################ergot#########################
            self.vergot = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vergot.setRange(0, 1)
            self.vergot.setSpecialValueText(' ')
            self.vergot.resize(100, 20)
            self.vergot.move(350, 209)
            self.vergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Graines nuisibles (%)##########
            self.vgrainnuisible = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainnuisible.setRange(0, 1)
            self.vgrainnuisible.setSpecialValueText(' ')
            self.vgrainnuisible.resize(100, 20)
            self.vgrainnuisible.move(350, 230)
            self.vgrainnuisible.setFont(self.txtpsfont)
            self.vgrainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #############################Débris végétaux (%)########
            self.vdébrisvé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vdébrisvé.setRange(0, 10)
            self.vdébrisvé.setSpecialValueText(' ')
            self.vdébrisvé.resize(100, 20)
            self.vdébrisvé.move(350, 251)
            self.vdébrisvé.setFont(self.txtpsfont)
            self.vdébrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrt = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vmatiéreinrt.setRange(0, 10)
            self.vmatiéreinrt.setSpecialValueText(' ')
            self.vmatiéreinrt.resize(100, 20)
            self.vmatiéreinrt.move(350, 272)
            self.vmatiéreinrt.setFont(self.txtpsfont)
            self.vmatiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.vgrainchaufé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainchaufé.setRange(0, 7)
            self.vgrainchaufé.setSpecialValueText(' ')
            self.vgrainchaufé.resize(100, 20)
            self.vgrainchaufé.move(350, 293)
            self.vgrainchaufé.setFont(self.txtpsfont)
            self.vgrainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleur = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainsanvaleur.setSpecialValueText(' ')
            self.vgrainsanvaleur.setRange(0, 10)
            self.vgrainsanvaleur.resize(100, 20)
            self.vgrainsanvaleur.move(350, 314)
            self.vgrainsanvaleur.setFont(self.txtpsfont)
            self.vgrainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarré = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgraincarré.setSpecialValueText(' ')
            self.vgraincarré.setFont(self.txtpsfont)
            self.vgraincarré.setRange(0, 10)
            self.vgraincarré.resize(100, 20)
            self.vgraincarré.move(350, 335)
            self.vgraincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Total(%) 1er#####################################################
            self.vtotalprem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.vtotalprem.setSpecialValueText(' ')
            self.vtotalprem.setRange(0, 20)
            self.vtotalprem.resize(100, 20)
            self.vtotalprem.move(350, 356)
            self.vtotalprem.setFont(self.txtpsfont)
            self.vtotalprem.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##############################################Grains cassés (%) #########################################################
            self.vgraincassé = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgraincassé.move(350, 377)
            self.vgraincassé.resize(100, 20)
            self.vgraincassé.setRange(0, 20)
            self.vgraincassé.setSpecialValueText(" ")
            self.vgraincassé.setFont(self.txtpsfont)
            self.vgraincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaude = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainechaude.setSpecialValueText(" ")
            self.vgrainechaude.setRange(0, 10)
            self.vgrainechaude.move(350, 419)
            self.vgrainechaude.resize(100, 20)
            self.vgrainechaude.setFont(self.txtpsfont)
            self.vgrainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.vgrainmaigre = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainmaigre.setRange(0, 10)
            self.vgrainmaigre.setSpecialValueText(" ")
            self.vgrainmaigre.move(350, 398)
            self.vgrainmaigre.setFont(self.txtpsfont)
            self.vgrainmaigre.resize(100, 20)
            self.vgrainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.vgraigermé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgraigermé.move(350, 440)
            self.vgraigermé.resize(100, 20)
            self.vgraigermé.setSpecialValueText('  ')
            self.vgraigermé.setFont(self.txtpsfont)
            self.vgraigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaisé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainpunaisé.move(350, 461)
            self.vgrainpunaisé.resize(100, 20)
            self.vgrainpunaisé.setSpecialValueText('   ')
            self.vgrainpunaisé.setFont(self.txtpsfont)
            self.vgrainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiqué = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainpiqué.move(350, 482)
            self.vgrainpiqué.resize(100, 20)
            self.vgrainpiqué.setSpecialValueText('  ')
            self.vgrainpiqué.setFont(self.txtpsfont)
            self.vgrainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.vgrainboutef = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainboutef.move(350, 503)
            self.vgrainboutef.resize(100, 20)
            self.vgrainboutef.setRange(0, 10)
            self.vgrainboutef.setSpecialValueText('  ')
            self.vgrainboutef.setFont(self.txtpsfont)
            self.vgrainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.vgrainbouté = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainbouté.move(350, 524)
            self.vgrainbouté.resize(100, 20)
            self.vgrainbouté.setRange(0, 10)
            self.vgrainbouté.setSpecialValueText('  ')
            self.vgrainboutef.setFont(self.txtpsfont)
            self.vgrainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchté = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainmouchté.move(350, 545)
            self.vgrainmouchté.resize(100, 20)
            self.vgrainmouchté.setRange(0, 5)
            self.vgrainmouchté.setSpecialValueText(' ')
            self.vgrainmouchté.setFont(self.txtpsfont)
            self.vgrainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.vgrainetrangé = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainetrangé.move(350, 566)
            self.vgrainetrangé.resize(100, 20)
            self.vgrainetrangé.setRange(0, 10)
            self.vgrainetrangé.setSpecialValueText(' ')
            self.vgrainetrangé.setFont(self.txtpsfont)
            self.vgrainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.vtotaldem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.vtotaldem.setRange(1, 30)
            self.vtotaldem.move(350, 587)
            self.vtotaldem.resize(100, 20)
            self.vtotaldem.setSpecialValueText(' ')
            self.vtotaldem.setFont(self.txtpsfont)
            self.vtotaldem.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #########################################################observation###############
            self.observation = QtWidgets.QLabel("Observation", self.bltendretab)
            self.observation.move(490, 145)
            self.observation.resize(100, 20)
            self.observation.setFont(self.txtpsfont)
            self.observation.setFont(self.font)
            ##################################################txtobservation##################################
            self.txtobservation = QtWidgets.QTextEdit("<h2><h2/>  <h2><h2/>  <h2><h2/> <h2><h2/>   <h3><h3/>",
                                                      self.bltendretab)
            self.txtobservation.move(490, 167)
            self.txtobservation.resize(250, 438)
            self.txtobservation.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelccls = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.bltendretab)
            self.labelccls.move(500, 0)
            self.labelccls.resize(438, 90)
            self.labelccls.setFont(self.font)
            self.labelccls.setStyleSheet(
                "background-color: #ffffff; border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            self.LABELBULLETIN = QtWidgets.QLabel("<H2>BULLETIN SORTIE MOULIN</H2>", self.bltendretab)
            self.LABELBULLETIN.move(600, 35)
            self.LABELBULLETIN.resize(240, 30)
            self.LABELBULLETIN.setStyleSheet("background-color: #ffffff")

            self.bletendretxt = QtWidgets.QLabel("<H2>Blé Tendre<H2/>", self.bltendretab)
            self.bletendretxt.move(690, 60)
            self.bletendretxt.resize(140, 23)
            self.bletendretxt.setStyleSheet("background-color: #fffff")

            #############################################date edit#############################################
            self.dateeditetxt = QtWidgets.QLabel("Date:", self.bltendretab)
            self.dateeditetxt.setGeometry(QtCore.QRect(20, 20, 100, 23))
            self.dateeditetxt.setFont(self.font)
            self.dateedite = QtWidgets.QDateEdit(self.bltendretab)
            self.dateedite.setDate(self.daydate)
            self.dateedite.setStyleSheet(
                " background-color: #FDFEFE;padding: 1px;font-size: 12px;border-style: solid;border: 1px solid #76797C;border-radius: 0px;color: #000000;")
            self.dateedite.move(84, 20)
            self.dateedite.resize(130, 20)
            self.dateedite.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################

            ##################################################quantite###############################################
            self.quantite = QtWidgets.QLabel("Quantité(QX):", self.bltendretab)
            self.quantite.move(840, 100)
            self.quantite.resize(85, 40)
            self.quantite.setFont(self.font)
            self.quantitetxt = QtWidgets.QLineEdit("", self.bltendretab)

            self.quantitetxt.move(928, 110)
            self.quantitetxt.resize(80, 20)
            self.quantitetxt.setValidator(QDoubleValidator(0.99, 99.99, 2))
            # self.quantitetxt.setStyleSheet("background-color: #31363b")

            ####################################################éspece###########################
            self.éspéce = QtWidgets.QLabel("Espèce :", self.bltendretab)
            self.éspéce.move(20, 100)
            self.éspéce.resize(120, 40)
            self.éspéce.setFont(self.font)
            self.éspécecombo = QtWidgets.QComboBox(self.bltendretab)
            self.éspécecombo.addItem("Blé Tendre")
            self.éspécecombo.move(84, 110)
            self.éspécecombo.resize(130, 20)
            ##########################################Nom de l’acheteur : moulin###########################################
            self.moulin = QtWidgets.QLabel("Nom de l’acheteur:", self.bltendretab)
            self.moulin.move(220, 100)
            self.moulin.resize(118, 40)
            self.moulin.setFont(self.font)
            self.moulincombo = QtWidgets.QComboBox(self.bltendretab)
            self.moulincombo.addItem("")
            self.moulincombo.addItem("EURL DJERBIR INDUSTRIELE")
            self.moulincombo.addItem("SARL MOULIN O_ABBES")
            self.moulincombo.addItem("EURL MOULIN BELACEL")
            self.moulincombo.addItem("MOULIN TAHAR MESSAOUD")
            self.moulincombo.addItem("SARL MOULIN BENABDELLAH")
            self.moulincombo.addItem("SARL MATAHIN EL HARAMAIN")
            self.moulincombo.addItem("MINOTERIE NOUR EL HAYAT")
            self.moulincombo.addItem("MOULIN MERINE SASSI")
            self.moulincombo.addItem("SARL DJENDLI")
            self.moulincombo.addItem("SARL TRX HYDRO BENHADJAR")
            self.moulincombo.addItem("EURL MOULIN AIN RAHMA")
            self.moulincombo.addItem("MOULIN FARINE BLANCHE")
            self.moulincombo.addItem("MOULIN MAAMAR BENHADJAR")
            self.moulincombo.addItem("MOULIN OULD BENAICHOUCHE")
            self.moulincombo.addItem("SARL FARINIERE DE L’OUEST")
            self.moulincombo.addItem("SARL MATAHINE ADJINE")
            self.moulincombo.addItem("MOULIN CHOUIKH YOUCEF")
            self.moulincombo.addItem("EURL MOULIN DAMAKO")
            self.moulincombo.addItem("SARL MATAHINE SIDI ABDELHADI")
            self.moulincombo.addItem("SARL MATAHINE MINA")
            self.moulincombo.addItem("EURL ELFORSANE PRODUCTION")
            self.moulincombo.addItem("SARL MATAHINE TOUFIK")
            self.moulincombo.move(340, 110)
            self.moulincombo.resize(220, 20)

            #####################################################Point de collecte : #######################################################
            self.pointcollecte = QtWidgets.QLabel("Point de collecte:", self.bltendretab)
            self.pointcollecte.move(570, 100)
            self.pointcollecte.resize(106, 40)
            self.pointcollecte.setFont(self.font)
            self.pointcollectecombo = QtWidgets.QComboBox(self.bltendretab)
            self.pointcollectecombo.addItem("")
            self.pointcollectecombo.addItem("DOCK SILO CENTRAL")
            self.pointcollectecombo.move(680, 110)
            self.pointcollectecombo.resize(150, 20)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeur = QtWidgets.QLabel("Nom de l’Agréeur:", self.bltendretab)
            self.agréeeur.move(1015, 100)
            self.agréeeur.resize(112, 40)
            self.agréeeur.setFont(self.font)
            self.agréeeurcombobt = QtWidgets.QComboBox(self.bltendretab, editable=False)
            self.agréeeurcombobt.addItem("")
            self.agréeeurcombobt.addItem("FELOUAH OMAR")
            self.agréeeurcombobt.addItem("BEKHEDDA AEK")
            self.agréeeurcombobt.addItem("BENAISSA YOUCEF")
            self.agréeeurcombobt.addItem("REZZAG SOFIANE ")
            self.agréeeurcombobt.addItem("BELBACHA M.NADIR")
            self.agréeeurcombobt.move(1133, 110)
            self.agréeeurcombobt.resize(147, 20)

            ############################################docx2pdf######################

            self.tamerbt = QTimer()
            self.tamerbt.timeout.connect(self.plusbt)
            self.tamerbt.setInterval(1000)
            self.tamerbt.start()

            ###########################buttons################

            self.btnsavebt = QtWidgets.QPushButton("ENREGISTRER", self.bltendretab, clicked=lambda: self.docx_file())
            self.btnsavebt.move(800, 167)
            self.btnsavebt.resize(480, 80)
            self.btnsavebt.setFont(self.font)
            self.btnsavebt.setIcon(QIcon("images/savepis.png"))
            self.btnsavebt.setIconSize(QSize(70, 80))
            # self.btnsavebt.clicked.connect(self.docx_file)

            self.btnprintbt = QtWidgets.QPushButton("IMPRIMER", self.bltendretab, clicked=lambda: self.printer())
            self.btnprintbt.move(800, 272)
            self.btnprintbt.resize(480, 80)
            self.btnprintbt.setFont(self.font)
            self.btnprintbt.setIcon(QIcon("images/print125.png"))
            self.btnprintbt.setIconSize(QSize(70, 80))
            # btnprint.clicked.connect(printer)

            self.btnefacebt = QtWidgets.QPushButton("EFACER", self.bltendretab, clicked=lambda: self.clear_allbt())
            self.btnefacebt.move(800, 377)
            self.btnefacebt.resize(480, 80)
            self.btnefacebt.setIcon(QIcon("images/eraser45877.png"))
            self.btnefacebt.setIconSize(QSize(70, 80))
            self.btnefacebt.setFont(self.font)
            # self.btnefacebt.clicked.connect(self.clear_all)

            #############################################BLE DUR
            ########################################################
            #####################################################################
            ###############################################################################
            ##############################################################################################
            self.bldurtab = QtWidgets.QWidget()
            self.bldurtab.setObjectName("bldurtab")
            self.confirmitewidget.addTab(self.bldurtab, "sortie moulin BD")
            self.bldurtab.setStyleSheet("""QToolTip
            {
                border: 1px solid #76797C;
                background-color: rgb(90, 102, 117);;
                color: white;
                padding: 5px;
                opacity: 200;
            }

            QWidget
            {
                color: #000000;
                background-color: #ffffff;
                selection-background-color:#3daee9;
                selection-color: #3daee9;
                background-clip: border;
                border-image: none;
                border: 0px transparent black;
                outline: 0;
            }

            QWidget:item:hover
            {
                background-color: #3daee9;
                color: #eff0f1;
            }

            QWidget:item:selected
            {
                background-color: #3daee9;
            }



            QWidget:disabled
            {
                color: #454545;
                background-color: #31363b;
            }

            QAbstractItemView
            {
                alternate-background-color: #31363b;
                color: #eff0f1;
                border: 1px solid 3A3939;
                border-radius: 2px;
            }

            QWidget:focus, QMenuBar:focus
            {
                border: 1px solid #3daee9;
            }

            QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
            {
                border: none;
            }

            QLineEdit
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color: #000000;
            }
            QDoubleSpinBox
            {
                background-color: #FDFEFE;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 0px;
                color:#000000;
                font-size: 11px;
                font-weight: bold;

            }
            QDoubleSpinBox:focus{
                background-color: #FDFEFE;
                border-style: solid;
                border: 2px solid #76797C;
                border-radius: 4px;
                border-color: #ff8c00;
            }
            QDoubleSpinBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 1px;

                border-left-width: 0px;
                border-left-color: #232629;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }



            QGroupBox {
                border:1px solid #76797C;
                border-radius: 2px;
                margin-top: 5px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding-left: 4px;
                padding-right: 4px;
                padding-top: 4px;
            }

            QAbstractScrollArea
            {
                border-radius: 2px;
                border: 1px solid #76797C;
                background-color: transparent;
            }

            QScrollBar:horizontal
            {
                height: 15px;
                margin: 3px 15px 3px 15px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
                background-color: #2A2929;
            }

            QScrollBar::handle:horizontal
            {
                background-color: #605F5F;
                min-width: 5px;
                border-radius: 4px;
            }

            QScrollBar::add-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
                width: 10px;
                height: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }


            QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
            {
                background: none;
            }


            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
            {
                background: none;
            }

            QScrollBar:vertical
            {
                background-color: #2A2929;
                width: 15px;
                margin: 15px 3px 15px 3px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
            }

            QScrollBar::handle:vertical
            {
                background-color: #605F5F;
                min-height: 5px;
                border-radius: 4px;
            }

            QScrollBar::sub-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
            {

                border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }


            QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
            {
                border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
            {
                background: none;
            }


            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
            {
                background: none;
            }

            QTextEdit
            {
                background-color: #FDFEFE;
                color: #000000;
                border: 1px solid #76797C;
            }

            QPlainTextEdit
            {
                background-color: #232629;;
                color: #000000;
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QHeaderView::section
            {
                background-color: #76797C;
                color: #eff0f1;
                padding: 1px;
                border: 1px solid #76797C;
            }

            QSizeGrip {
                width: 12px;
                height: 12px;
            }


            QMainWindow::separator
            {
                background-color: #31363b;
                color: white;
                padding-left: 4px;
                spacing: 2px;
                border: 1px dashed #76797C;
            }

            QMainWindow::separator:hover
            {

                background-color: #787876;
                color: white;
                padding-left: 4px;
                border: 1px solid #76797C;
                spacing: 2px;
            }


            QMenu::separator
            {
                height: 1px;
                background-color: #76797C;
                color: white;
                padding-left: 4px;
                margin-left: 10px;
                margin-right: 5px;
            }


            QFrame
            {
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QFrame[frameShape="0"]
            {
                border-radius: 2px;
                border: 1px transparent #76797C;
            }

            QStackedWidget
            {
                border: 1px transparent black;
            }


            QPushButton
            {
                color: #00000;
                background-color:#ade3e7;
                border-width: 1px;
                border-color: #1e1e1e;
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
            }

            QComboBox
            {
               background-color: #FDFEFE;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                min-width: 40px;
            }

            QPushButton:checked{
                background-color: #76797C;
                border-color: #6A6969;
            }

            QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
            {
                border: 1px solid #ff8c00;
                color: #000000;
            }

            QComboBox:on
            {
                padding-top: 1px;
                padding-left: 1px;
                selection-background-color: #FDFEFE;
            }

            QComboBox QAbstractItemView
            {
                background-color: #FDFEFE;
                border-radius: 2px;
                border: 1px solid #76797C;
                color:#000000;
                selection-background-color: #000000;
            }

            QComboBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;

                border-left-width: 0px;
                border-left-color: ff8c00;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }


            QLabel
            {
                border: 2px solid black;
            }

            QTabWidget{
                border: 0px transparent black;
            }

            QTabWidget::pane {
                border: 1px solid #76797C;
                padding: 5px;
                margin: 0px;
            }

            QTabBar
            {
                qproperty-drawBase: 0;
                left: 5px; /* move to the right by 5px */
                border-radius: 3px;
            }

            QTabBar:focus
            {
                border: 0px transparent black;
            }

            QTabBar::close-button  {
                image: url(:/qss_icons/Dark_rc/close.png);
                background: transparent;
            }

            QTabBar::close-button:hover
            {
                image: url(:/qss_icons/Dark_rc/close-hover.png);
                background: transparent;
            }

            QTabBar::close-button:pressed {
                image: url(:/qss_icons/Dark_rc/close-pressed.png);
                background: transparent;
            }

            /* TOP TABS */
            QTabBar::tab:top {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                min-width: 10px;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;
            }

            QTabBar::tab:top:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;    
            }

            QTabBar::tab:top:!selected:hover {
                background-color: #3daee9;
            }

            /* BOTTOM TABS */
            QTabBar::tab:bottom {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
                min-width: 10px;
            }

            QTabBar::tab:bottom:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:bottom:!selected:hover {
                background-color: #3daee9;
            }

            /* LEFT TABS */
            QTabBar::tab:left {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:left:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:left:!selected:hover {
                background-color: #3daee9;
            }


            /* RIGHT TABS */
            QTabBar::tab:right {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:right:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
            }





            QSlider::groove:horizontal {
                border: 1px solid #565a5e;
                height: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 2px;
            }

            QSlider::handle:horizontal {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: -8px 0;
                border-radius: 9px;
            }

            QSlider::groove:vertical {
                border: 1px solid #565a5e;
                width: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 3px;
            }

            QSlider::handle:vertical {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: 0 -8px;
                border-radius: 9px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px transparent #76797C;
                border-radius: 2px;
                margin: 3px;
                padding: 5px;
            }

            QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
             padding-right: 20px; /* make way for the popup button */
             border: 1px #76797C;
             border-radius: 5px;
            }

            QToolButton[popupMode="2"] { /* only for InstantPopup */
             padding-right: 10px; /* make way for the popup button */
             border: 1px #76797C;
            }


            QToolButton:hover, QToolButton::menu-button:hover {
                background-color: transparent;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            QToolButton:checked, QToolButton:pressed,
                    QToolButton::menu-button:pressed {
                background-color: #3daee9;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
            QToolButton::menu-indicator {
                background-color:ff8c00;
                top: -7px; left: -2px; /* shift it a bit */
            }

            /* the subcontrols below are used only in the MenuButtonPopup mode */
            QToolButton::menu-button {
                border: 1px transparent #76797C;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
                /* 16px width + 4px for border = 20px allocated above */
                width: 16px;
                outline: none;
            }

            QToolButton::menu-arrow {
               background-color:ff8c00;
            }

            QToolButton::menu-arrow:open {
                border: 1px solid #76797C;
            }

            QPushButton::menu-indicator  {
                subcontrol-origin: padding;
                subcontrol-position: bottom right;
                left: 8px;
            }

            QTableView
            {
                border: 1px solid #76797C;
                gridline-color: #31363b;
                background-color: #FDFEFE;
                color:#000000;
            }


            QTableView, QHeaderView
            {
                background-color: #FDFEFE;
                color:#000000;
                border-radius: 0px;
            }

            QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
                background: #FDFEFE;
                color: #000000;
            }

            QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
                background: #3daee9;
                color: #000000;
            }


            QHeaderView
            {
                background-color: #FDFEFE;
                border: 1px transparent;
                border-radius: 0px;
                margin: 0px;
                padding: 0px;

            }

            QHeaderView::section  {
                background-color:#80f1f9;
                color: #000000;
                padding: 5px;
                border: 1px solid #76797C;
                border-radius: 0px;
                text-align: center;
            }

            QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
            {
                border-top: 1px solid #76797C;
            }

            QHeaderView::section::vertical
            {
                border-top: transparent;
            }

            QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
            {
                border-left: 1px solid #76797C;
            }

            QHeaderView::section::horizontal
            {
                border-left: transparent;
            }


            QHeaderView::section:checked
             {
                color: #000000;
                background-color: #3daee9;
             }

             /* style the sort indicator */
            QHeaderView::down-arrow {

            }

            QHeaderView::up-arrow {

            }


            QTableCornerButton::section {
                background-color: #31363b;
                border: 1px transparent #76797C;
                border-radius: 0px;
            }

            QToolBox  {
                padding: 5px;
                border: 1px transparent black;
            }

            QToolBox::tab {
                color: #eff0f1;
                background-color: #31363b;
                border: 1px solid #76797C;
                border-bottom: 1px transparent #31363b;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }

            QToolBox::tab:selected { /* italicize selected tabs */
                font: italic;
                background-color: #31363b;
                border-color: #3daee9;
             }

            QStatusBar::item {
                border: 0px transparent dark;
             }


            QFrame[height="3"], QFrame[width="3"] {
                background-color: #76797C;
            }




            QDateEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;
            }

            QDateEdit:on
            {
                padding-top: 2px;
                padding-left: 2px;
                selection-background-color: #4a4a4a;
            }

            QDateEdit QAbstractItemView
            {
                background-color: #ff8c00;
                border-radius: 2px;
                border: 1px solid #3375A3;
                selection-background-color:ff8c00;
            }

            QDateEdit::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 0px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }   
            QDateTimeEdit
            {
                background-color: #232629;;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;

            }    
            """)
            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétrebd = QtWidgets.QLabel("Paramètre", self.bldurtab)
            self.paramétrebd.move(30, 111)
            self.paramétrebd.resize(80, 20)
            self.paramétrebd.setFont(self.font)
            self.txtpsfontbd = QtGui.QFont()
            self.txtpsfontbd.setBold(True)
            self.txtpsfontbd.setPointSize(9)
            ################Limites(sans bon ni réf)###############
            self.valeurbd = QtWidgets.QLabel("""Limite-ssans-bon-ni-réf)""", self.bldurtab)
            self.valeurbd.move(170, 95)
            self.valeurbd.resize(145, 55)
            self.valeurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.psbd = QtWidgets.QLineEdit("Poids spécifique (kg/hl):   (75.500-75.899)", self.bldurtab, readOnly=True)
            self.psbd.resize(319, 20)
            self.psbd.move(30, 131)
            self.psbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.psbd.setFont(self.txtpsfont)
            self.psbd.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.humiditebd = QtWidgets.QLineEdit("Teneur en eau(%):  (13.5-15)", self.bldurtab, readOnly=True)
            self.humiditebd.resize(319, 20)
            self.humiditebd.move(30, 152)
            self.humiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.humiditebd.setFont(self.txtpsfont)

            #######################ergot#########################
            self.ergotbd = QtWidgets.QLineEdit("Ergo(% :<0.001 ", self.bldurtab, readOnly=True)
            self.ergotbd.resize(319, 20)
            self.ergotbd.move(30, 173)
            self.ergotbd.setStyleSheet("background-color: #232629")
            self.ergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.ergotbd.setFont(self.txtpsfont)

            #########################Graines nuisibles (%)##########
            self.grainnuisiblebd = QtWidgets.QLineEdit("Graines nuisibles(%): <0.001", self.bldurtab, readOnly=True)
            self.grainnuisiblebd.resize(319, 20)
            self.grainnuisiblebd.move(30, 194)
            self.grainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisiblebd.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvébd = QtWidgets.QLineEdit("Débris végétaux(%):  ", self.bldurtab, readOnly=True)
            self.débrisvébd.resize(319, 20)
            self.débrisvébd.move(30, 215)
            self.débrisvébd.setFont(self.txtpsfont)
            self.débrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrtbd = QtWidgets.QLineEdit("Matière inerte(%):", self.bldurtab, readOnly=True)
            self.matiéreinrtbd.resize(319, 20)
            self.matiéreinrtbd.move(30, 236)
            self.matiéreinrtbd.setFont(self.txtpsfont)
            self.matiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.grainchaufébd = QtWidgets.QLineEdit("Grains chauffés(%): ", self.bldurtab, readOnly=True)
            self.grainchaufébd.resize(319, 20)
            self.grainchaufébd.move(30, 257)
            self.grainchaufébd.setFont(self.txtpsfont)
            self.grainchaufébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleurbd = QtWidgets.QLineEdit("Grains sans valeur(%):", self.bldurtab, readOnly=True)
            self.grainsanvaleurbd.resize(319, 20)
            self.grainsanvaleurbd.move(30, 278)
            self.grainsanvaleurbd.setFont(self.txtpsfont)
            self.grainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:   ", self.bldurtab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 299)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalprembd = QtWidgets.QLineEdit("TOTAL 1ére CAT:   ≤1", self.bldurtab, readOnly=True)
            self.totalprembd.resize(319, 20)
            self.totalprembd.move(30, 320)
            self.totalprembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalprembd.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincassébd = QtWidgets.QLineEdit("Grains cassés(%):   ≤2", self.bldurtab, readOnly=True)
            self.graincassébd.move(30, 341)
            self.graincassébd.resize(319, 20)
            self.graincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincassébd.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaudebd = QtWidgets.QLineEdit("Gains échaudés(%):   ", self.bldurtab, readOnly=True)
            self.grainechaudebd.move(30, 362)
            self.grainechaudebd.resize(319, 20)
            self.grainechaudebd.setFont(self.txtpsfont)
            self.grainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################
            self.grainmaigrebd = QtWidgets.QLineEdit("Grains maigres(%):", self.bldurtab, readOnly=True)
            self.grainmaigrebd.move(30, 383)
            self.grainmaigrebd.resize(319, 20)
            self.grainmaigrebd.setFont(self.txtpsfont)
            self.grainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##########################################################Grains germés (%)###################################################
            self.graigermébd = QtWidgets.QLineEdit("Grains germés(%): ≤2", self.bldurtab, readOnly=True)
            self.graigermébd.move(30, 404)
            self.graigermébd.resize(319, 20)
            self.graigermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigermébd.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaisébd = QtWidgets.QLineEdit("Grain punaisés(%): ≤1", self.bldurtab, readOnly=True)
            self.grainpunaisébd.move(30, 425)
            self.grainpunaisébd.resize(319, 20)
            self.grainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaisébd.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiquébd = QtWidgets.QLineEdit("Grains piqués(%):  ", self.bldurtab, readOnly=True)
            self.grainpiquébd.move(30, 446)
            self.grainpiquébd.resize(319, 20)
            self.grainpiquébd.setFont(self.txtpsfont)
            self.grainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.grainboutébd = QtWidgets.QLineEdit("Grains boutés  « forte » (%):", self.bldurtab, readOnly=True)
            self.grainboutébd.move(30, 467)
            self.grainboutébd.resize(319, 20)
            self.grainboutébd.setFont(self.txtpsfont)
            self.grainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchtébd = QtWidgets.QLineEdit("Grains mouchetés (%):", self.bldurtab, readOnly=True)
            self.grainmouchtébd.move(30, 488)
            self.grainmouchtébd.resize(319, 20)
            self.grainmouchtébd.setFont(self.txtpsfont)
            self.grainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.grainetrangébd = QtWidgets.QLineEdit("Grain étrangers Utilisables pour le bétail (%):  ",
                                                      self.bldurtab,
                                                      readOnly=True)
            self.grainetrangébd.move(30, 509)
            self.grainetrangébd.resize(319, 20)
            self.grainetrangébd.setFont(self.txtpsfont)
            self.grainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldembd = QtWidgets.QLineEdit("Total(%)  Imp2eme cat   ≤5", self.bldurtab, readOnly=True)
            self.totaldembd.move(30, 530)
            self.totaldembd.resize(319, 20)
            self.totaldembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldembd.setFont(self.txtpsfont)
            #########################################################indice notin##################################################
            self.indicebd = QtWidgets.QLineEdit('Indice Notin ', self.bldurtab, readOnly=True)
            self.indicebd.move(30, 551)
            self.indicebd.resize(319, 20)
            self.indicebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################blétendre dans blé dur########
            self.bletendreinbledur = QtWidgets.QLineEdit("Ble tendre dans ble dur(%)", self.bldurtab, readOnly=True)
            self.bletendreinbledur.move(30, 572)
            self.bletendreinbledur.resize(319, 20)
            self.bletendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################total complet######
            self.totalcomplet = QtWidgets.QLineEdit("TOTAL", self.bldurtab, readOnly=True)
            self.totalcomplet.move(30, 593)
            self.totalcomplet.resize(319, 20)
            self.totalcomplet.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #################label valeure##############
            self.valeurbd = QtWidgets.QLabel("valeur", self.bldurtab)
            self.valeurbd.move(350, 112)
            self.valeurbd.resize(100, 20)
            self.valeurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.vpsbd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vpsbd.setRange(71, 84.00)
            self.vpsbd.setSpecialValueText(' ')
            self.vpsbd.resize(100, 20)
            self.vpsbd.move(350, 131)
            self.vpsbd.setFont(self.txtpsfont)
            self.vpsbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.vhumiditebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vhumiditebd.setRange(8, 14)
            self.vhumiditebd.resize(100, 20)
            self.vhumiditebd.setSpecialValueText(' ')
            self.vhumiditebd.move(350, 152)
            self.vhumiditebd.setFont(self.txtpsfont)
            self.vhumiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################ergot#########################
            self.vergotbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vergotbd.setRange(0, 10)
            self.vergotbd.setSpecialValueText(' ')
            self.vergotbd.resize(100, 20)
            self.vergotbd.move(350, 173)
            self.vergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Graines nuisibles (%)##########
            self.vgrainnuisiblebd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainnuisiblebd.setRange(0, 10)
            self.vgrainnuisiblebd.setSpecialValueText(' ')
            self.vgrainnuisiblebd.resize(100, 20)
            self.vgrainnuisiblebd.move(350, 194)
            self.vgrainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #############################Débris végétaux (%)########
            self.vdébrisvébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vdébrisvébd.setRange(0, 10)
            self.vdébrisvébd.setSpecialValueText(' ')
            self.vdébrisvébd.resize(100, 20)
            self.vdébrisvébd.move(350, 215)
            self.vdébrisvébd.setFont(self.txtpsfont)
            self.vdébrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrtbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vmatiéreinrtbd.setRange(0, 10)
            self.vmatiéreinrtbd.setSpecialValueText(' ')
            self.vmatiéreinrtbd.resize(100, 20)
            self.vmatiéreinrtbd.move(350, 236)
            self.vmatiéreinrtbd.setFont(self.txtpsfont)
            self.vmatiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.vgrainchaufébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainchaufébd.setRange(0, 10)
            self.vgrainchaufébd.setSpecialValueText(' ')
            self.vgrainchaufébd.resize(100, 20)
            self.vgrainchaufébd.move(350, 257)
            self.vgrainchaufébd.setFont(self.txtpsfont)
            self.vgrainchaufébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleurbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainsanvaleurbd.setSpecialValueText(' ')
            self.vgrainsanvaleurbd.setRange(0, 10)
            self.vgrainsanvaleurbd.resize(100, 20)
            self.vgrainsanvaleurbd.move(350, 278)
            self.vgrainsanvaleurbd.setFont(self.txtpsfont)
            self.vgrainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarrébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgraincarrébd.setSpecialValueText(' ')
            self.vgraincarrébd.setRange(0, 10)
            self.vgraincarrébd.resize(100, 20)
            self.vgraincarrébd.move(350, 299)
            self.vgraincarrébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Total(%) 1er#####################################################
            self.vtotalprembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotalprembd.setSpecialValueText(' ')
            self.vtotalprembd.setRange(0, 10)
            self.vtotalprembd.resize(100, 20)
            self.vtotalprembd.move(350, 320)
            self.vtotalprembd.setFont(self.txtpsfont)
            self.vtotalprembd.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##############################################Grains cassés (%) #########################################################
            self.vgraincassébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgraincassébd.move(350, 341)
            self.vgraincassébd.resize(100, 20)
            self.vgraincassébd.setRange(0, 10)
            self.vgraincassébd.setSpecialValueText(" ")
            self.vgraincassébd.setFont(self.txtpsfont)
            self.vgraincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaudebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainechaudebd.setSpecialValueText(" ")
            self.vgrainechaudebd.setRange(0, 10)
            self.vgrainechaudebd.move(350, 362)
            self.vgrainechaudebd.resize(100, 20)
            self.vgrainechaudebd.setFont(self.txtpsfont)
            self.vgrainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.vgrainmaigrebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainmaigrebd.setRange(0, 10)
            self.vgrainmaigrebd.setSpecialValueText(" ")
            self.vgrainmaigrebd.move(350, 383)
            self.vgrainmaigrebd.setFont(self.txtpsfont)
            self.vgrainmaigrebd.resize(100, 20)
            self.vgrainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.vgraigermébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgraigermébd.move(350, 404)
            self.vgraigermébd.resize(100, 20)
            self.vgraigermébd.setSpecialValueText('  ')
            self.vgraigermébd.setFont(self.txtpsfont)
            self.vgraigermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaisébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainpunaisébd.move(350, 425)
            self.vgrainpunaisébd.resize(100, 20)
            self.vgrainpunaisébd.setSpecialValueText('   ')
            self.vgrainpunaisébd.setFont(self.txtpsfont)
            self.vgrainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiquébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainpiquébd.move(350, 446)
            self.vgrainpiquébd.resize(100, 20)
            self.vgrainpiquébd.setSpecialValueText('  ')
            self.vgrainpiquébd.setFont(self.txtpsfont)
            self.vgrainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.vgrainboutébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainboutébd.move(350, 467)
            self.vgrainboutébd.resize(100, 20)
            self.vgrainboutébd.setSpecialValueText('  ')
            self.vgrainboutébd.setFont(self.txtpsfont)
            self.vgrainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchtébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainmouchtébd.move(350, 488)
            self.vgrainmouchtébd.resize(100, 20)
            self.vgrainmouchtébd.setRange(0, 10)
            self.vgrainmouchtébd.setSpecialValueText(' ')
            self.vgrainmouchtébd.setFont(self.txtpsfont)
            self.vgrainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.vgrainetrangébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainetrangébd.move(350, 509)
            self.vgrainetrangébd.resize(100, 20)
            self.vgrainetrangébd.setRange(0, 10)
            self.vgrainetrangébd.setSpecialValueText(' ')
            self.vgrainetrangébd.setFont(self.txtpsfont)
            self.vgrainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.vtotaldembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotaldembd.setRange(1, 20)
            self.vtotaldembd.move(350, 530)
            self.vtotaldembd.resize(100, 20)
            self.vtotaldembd.setSpecialValueText(' ')
            self.vtotaldembd.setFont(self.txtpsfont)
            self.vtotaldembd.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################indicenotin #################
            self.vindicenotin = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vindicenotin.setSpecialValueText(' ')
            self.vindicenotin.setFont(self.txtpsfont)
            self.vindicenotin.move(350, 551)
            self.vindicenotin.resize(100, 20)
            self.vindicenotin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################ble tendre dand blé dur############
            self.vblétendreinbledur = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vblétendreinbledur.setSpecialValueText(" ")
            self.vblétendreinbledur.move(350, 572)
            self.vblétendreinbledur.resize(100, 20)
            self.vblétendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################bonification total complet##############
            self.vtotalcomplet = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotalcomplet.setSpecialValueText(' ')
            self.vtotalcomplet.move(350, 593)
            self.vtotalcomplet.resize(100, 20)
            self.vtotalcomplet.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################réfaction##############################################

            #########################################################observation###############
            self.observationbd = QtWidgets.QLabel("Observation", self.bldurtab)
            self.observationbd.move(490, 111)
            self.observationbd.resize(100, 20)
            self.observationbd.setFont(self.txtpsfont)
            self.observationbd.setFont(self.font)
            ##################################################txtobservation##################################
            self.txtobservationbd = QtWidgets.QTextEdit("<h2><h2/>  <h2><h2/>  <h2><h2/> <h2><h2/>   <h3><h3/>",
                                                        self.bldurtab)
            self.txtobservationbd.move(490, 131)
            self.txtobservationbd.resize(250, 482)
            self.txtobservationbd.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelcclsbd = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.bldurtab)
            self.labelcclsbd.move(500, 0)
            self.labelcclsbd.resize(438, 80)
            self.labelcclsbd.setFont(self.font)
            self.labelcclsbd.setStyleSheet(
                "background-color: #ffffff; border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            self.LABELBULLETINbd = QtWidgets.QLabel("<H2>BULLETIN SORTIE MOULIN</H2>", self.bldurtab)
            self.LABELBULLETINbd.move(600, 30)
            self.LABELBULLETINbd.resize(240, 23)
            self.LABELBULLETINbd.setStyleSheet("background-color: #ffffff")
            self.bledurtxt = QtWidgets.QLabel("<H2>Blé DUR<H2/>", self.bldurtab)
            self.bledurtxt.move(698, 53)
            self.bledurtxt.resize(120, 23)
            self.bledurtxt.setStyleSheet("background-color: #ffffff")

            #############################################date edit#############################################
            self.dateeditetxtbd = QtWidgets.QLabel("Date:", self.bldurtab)
            self.dateeditetxtbd.setGeometry(QtCore.QRect(30, 5, 100, 20))
            self.dateeditetxtbd.setFont(self.font)
            self.dateeditebd = QtWidgets.QDateEdit(self.bldurtab)
            self.dateeditebd.setDate(self.daydate)
            self.dateeditebd.setStyleSheet(
                " background-color: #FDFEFE;padding: 1px;border-style: solid;border: 1px solid #76797C;font-size: 12px;border-radius: 0px;color: #000000;")
            self.dateeditebd.move(100, 5)
            self.dateeditebd.resize(112, 30)
            self.dateeditebd.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################

            ##################################################quantite###############################################
            self.quantitebd = QtWidgets.QLabel("Quantité(QX):", self.bldurtab)
            self.quantitebd.move(840, 88)
            self.quantitebd.resize(85, 25)
            self.quantitebd.setFont(self.font)
            self.quantitetxtbd = QtWidgets.QLineEdit("", self.bldurtab)
            self.quantitetxtbd.move(928, 85)
            self.quantitetxtbd.resize(80, 30)
            self.quantitetxtbd.setValidator(QDoubleValidator(0.99, 99.99, 2))

            ####################################################éspece###########################
            self.éspécebd = QtWidgets.QLabel("Espèce :", self.bldurtab)
            self.éspécebd.move(30, 88)
            self.éspécebd.resize(60, 20)
            self.éspécebd.setFont(self.font)
            self.éspécecombobd = QtWidgets.QComboBox(self.bldurtab)
            self.éspécecombobd.addItem("Blé Dur")
            self.éspécecombobd.move(100, 85)
            self.éspécecombobd.resize(112, 30)

            ##########################################Nom de l’acheteur : moulin###########################################
            self.moulinbd = QtWidgets.QLabel("Nom de l’acheteur:", self.bldurtab)
            self.moulinbd.move(220, 88)
            self.moulinbd.resize(118, 20)
            self.moulinbd.setFont(self.font)
            self.moulincombobd = QtWidgets.QComboBox(self.bldurtab)
            self.moulincombobd.addItem("")
            self.moulincombobd.addItem("SARL MOULIN BENABDELLAH")
            self.moulincombobd.move(340, 85)
            self.moulincombobd.resize(220, 30)

            #####################################################Point de collecte : #######################################################
            self.pointcollectebd = QtWidgets.QLabel("Point de collecte:", self.bldurtab)
            self.pointcollectebd.move(570, 80)
            self.pointcollectebd.resize(106, 40)
            self.pointcollectebd.setFont(self.font)
            self.pointcollectecombobd = QtWidgets.QComboBox(self.bldurtab)
            self.pointcollectecombobd.addItem("")
            self.pointcollectecombobd.addItem("DOCK SILO CENTRAL")
            self.pointcollectecombobd.move(680, 85)
            self.pointcollectecombobd.resize(150, 30)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeurbd = QtWidgets.QLabel("Nom de l’Agréeur:", self.bldurtab)
            self.agréeeurbd.move(1015, 80)
            self.agréeeurbd.resize(112, 40)
            self.agréeeurbd.setFont(self.font)
            self.agréeeurcombobd = QtWidgets.QComboBox(self.bldurtab, editable=False)
            self.agréeeurcombobd.addItem("")
            self.agréeeurcombobd.addItem("FELOUAH OMAR")
            self.agréeeurcombobd.addItem("BEKHEDDA AEK")
            self.agréeeurcombobd.addItem("BENAISSA YOUCEF")
            self.agréeeurcombobd.addItem("REZZAG SOFIANE ")
            self.agréeeurcombobd.addItem("BELBACHA M.NADIR")
            self.agréeeurcombobd.move(1133, 85)
            self.agréeeurcombobd.resize(147, 30)

            ############################################docx2pdf######################
            self.timerbd = QTimer()
            self.timerbd.timeout.connect(self.allcallculbd)
            self.timerbd.setInterval(1000)
            self.timerbd.start()

            ###########################buttons################

            self.btnsavebd = QtWidgets.QPushButton("ENREGISTRER", self.bldurtab, clicked=lambda: self.docx_bdsave())
            self.btnsavebd.move(800, 131)
            self.btnsavebd.resize(480, 80)
            self.btnsavebd.setFont(self.font)
            self.btnsavebd.setIcon(QIcon("images/savepis.png"))
            self.btnsavebd.setIconSize(QSize(70, 80))

            self.btnprintbd = QtWidgets.QPushButton("IMPRIMER", self.bldurtab, clicked=lambda: self.printerbd())
            self.btnprintbd.move(800, 236)
            self.btnprintbd.resize(480, 80)
            self.btnprintbd.setFont(self.font)
            self.btnprintbd.setIcon(QIcon("images/print125.png"))
            self.btnprintbd.setIconSize(QSize(70, 80))

            self.btnefacebd = QtWidgets.QPushButton("EFACER", self.bldurtab, clicked=lambda: self.clear_allbd())
            self.btnefacebd.move(800, 341)
            self.btnefacebd.resize(480, 80)
            self.btnefacebd.setIcon(QIcon("images/eraser45877.png"))
            self.btnefacebd.setIconSize(QSize(70, 80))
            self.btnefacebd.setFont(self.font)

            # self.btncalculbd = QtWidgets.QPushButton("CALCULER", self.bldurtab,clicked=lambda :self.allcallculbd())
            # self.btncalculbd.move(1120, 525)
            # self.btncalculbd.resize(200, 80)
            # self.btncalculbd.setFont(self.font)
            # self.btncalculbd.setIcon((QIcon("images/calcul12544.png")))
            # self.btncalculbd.setIconSize(QSize(70, 80))

            self.horizontalLayout.addWidget(self.confirmitewidget)
            MainWindow.setCentralWidget(self.centralwidget)
            self.statusbar = QtWidgets.QStatusBar(MainWindow)
            self.statusbar.setObjectName("statusbar")
            MainWindow.setStatusBar(self.statusbar)

            self.retranslateUi(MainWindow)
            self.confirmitewidget.setCurrentIndex(1)
            QtCore.QMetaObject.connectSlotsByName(MainWindow)



        def retranslateUi(self, MainWindow):
            _translate = QtCore.QCoreApplication.translate
            MainWindow.setWindowTitle(_translate("MainWindow", "ccls relizane service qualité"))
            self.confirmitewidget.setTabText(self.confirmitewidget.indexOf(self.entrétab),
                                             _translate("MainWindow", "ENTRE"))
            self.confirmitewidget.setTabText(self.confirmitewidget.indexOf(self.sortietab),
                                             _translate("MainWindow", "SORTIE"))


    if __name__ == "__main__":
        import sys
        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        ui = Conformité_Window()
        ui.confi_window(MainWindow)
        MainWindow.show()
        sys.exit(app.exec())
except Exception as e:
    print(e)
