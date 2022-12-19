##### Install Instructions
# use py2app
# create a setup.py file with py2applet --make-setup src/main.py
# delete the any old dist and build directories
# Then run python setup.py py2app -A

import PyQt6
from PyQt6 import QtCore, QtGui, QtWidgets
from PyPDF4 import PdfFileMerger
import subprocess
import time
import pdftotext
import os
import ocrmypdf
import docx2pdf



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1094, 771)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.garbage_bin = []

        self.file_list = ListDragWidget(self.centralwidget)
        self.file_list.setGeometry(QtCore.QRect(40, 90, 511, 271))
        self.file_list.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)
        self.file_list.setObjectName("file_list")

        self.file_list_label = QtWidgets.QLabel(self.centralwidget)
        self.file_list_label.setGeometry(QtCore.QRect(270, 60, 31, 16))
        self.file_list_label.setObjectName("file_list_label")

        self.ocr_languages_list = QtWidgets.QListWidget(self.centralwidget)
        self.ocr_languages_list.setGeometry(QtCore.QRect(580, 90, 241, 271))
        self.ocr_languages_list.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)
        self.ocr_languages_list.setObjectName("ocr_languages_list")

        self.available_ocr_languages_label = QtWidgets.QLabel(self.centralwidget)
        self.available_ocr_languages_label.setGeometry(QtCore.QRect(630, 60, 171, 16))
        self.available_ocr_languages_label.setObjectName("available_ocr_languages_label")

        self.app_title_label = QtWidgets.QLabel(self.centralwidget)
        self.app_title_label.setGeometry(QtCore.QRect(480, 0, 151, 31))
        font = QtGui.QFont()
        font.setPointSize(24)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        self.app_title_label.setFont(font)
        self.app_title_label.setTextFormat(QtCore.Qt.TextFormat.AutoText)
        self.app_title_label.setObjectName("app_title_label")

        self.file_browse_button = QtWidgets.QPushButton(self.centralwidget)
        self.file_browse_button.setGeometry(QtCore.QRect(40, 380, 141, 32))
        self.file_browse_button.setObjectName("file_browse_button")

        self.remove_selected_items_button = QtWidgets.QPushButton(self.centralwidget)
        self.remove_selected_items_button.setGeometry(QtCore.QRect(180, 380, 181, 32))
        self.remove_selected_items_button.setObjectName("remove_selected_items_button")

        self.clear_all_items_button = QtWidgets.QPushButton(self.centralwidget)
        self.clear_all_items_button.setGeometry(QtCore.QRect(360, 380, 121, 32))
        self.clear_all_items_button.setObjectName("clear_all_items_button")

        self.actions_label = QtWidgets.QLabel(self.centralwidget)
        self.actions_label.setGeometry(QtCore.QRect(940, 60, 60, 16))
        self.actions_label.setObjectName("actions_label")

        self.merge_pdfs_button = QtWidgets.QPushButton(self.centralwidget)
        self.merge_pdfs_button.setGeometry(QtCore.QRect(890, 90, 151, 32))
        self.merge_pdfs_button.setObjectName("merge_pdfs_button")

        self.ocr_pdfs_button = QtWidgets.QPushButton(self.centralwidget)
        self.ocr_pdfs_button.setGeometry(QtCore.QRect(890, 150, 151, 32))
        self.ocr_pdfs_button.setObjectName("ocr_pdfs_button")

        self.extract_text_button = QtWidgets.QPushButton(self.centralwidget)
        self.extract_text_button.setGeometry(QtCore.QRect(890, 210, 151, 32))
        self.extract_text_button.setObjectName("extract_text_button")

        self.extract_annotations_button = QtWidgets.QPushButton(self.centralwidget)
        self.extract_annotations_button.setGeometry(QtCore.QRect(890, 270, 151, 32))
        self.extract_annotations_button.setObjectName("extract_annotations_button")

        self.convert_files_button = QtWidgets.QPushButton(self.centralwidget)
        self.convert_files_button.setGeometry(QtCore.QRect(890, 330, 151, 32))
        self.convert_files_button.setObjectName("convert_files_button")

        self.clear_language_selections_button = QtWidgets.QPushButton(self.centralwidget)
        self.clear_language_selections_button.setGeometry(QtCore.QRect(600, 380, 191, 32))
        self.clear_language_selections_button.setObjectName("clear_language_selections_button")

        self.console_text_box = QtWidgets.QTextEdit(self.centralwidget)
        self.console_text_box.setGeometry(QtCore.QRect(40, 470, 781, 231))
        self.console_text_box.setReadOnly(True)
        self.console_text_box.setObjectName("console_text_box")

        self.status_of_operations_label = QtWidgets.QLabel(self.centralwidget)
        self.status_of_operations_label.setGeometry(QtCore.QRect(360, 440, 131, 16))
        self.status_of_operations_label.setObjectName("status_of_operations_label")

        self.exit_button = QtWidgets.QPushButton(self.centralwidget)
        self.exit_button.setGeometry(QtCore.QRect(890, 670, 151, 32))
        self.exit_button.setObjectName("exit_button")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1094, 24))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.get_langs()
        self.username = os.getenv('USER')

        self.retranslateUi(MainWindow)
        self.file_browse_button.clicked.connect(self.browseFiles)
        self.remove_selected_items_button.clicked.connect(self.removeFiles)
        self.clear_all_items_button.clicked.connect(self.clearFiles)
        self.merge_pdfs_button.clicked.connect(self.mergePDFs)
        self.ocr_pdfs_button.clicked.connect(self.ocrPDFs)
        self.extract_text_button.clicked.connect(self.extractText)
        self.extract_annotations_button.clicked.connect(self.extractAnnotations)
        self.convert_files_button.clicked.connect(self.convertFiles)
        self.exit_button.clicked.connect(self.exitApp)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.file_list_label.setText(_translate("MainWindow", "Files"))
        self.available_ocr_languages_label.setText(_translate("MainWindow", "Available OCR Languages"))
        self.app_title_label.setText(_translate("MainWindow", "PDF Tool Box"))
        self.file_browse_button.setText(_translate("MainWindow", "Browse for Files"))
        self.remove_selected_items_button.setText(_translate("MainWindow", "Remove Selected Items"))
        self.clear_all_items_button.setText(_translate("MainWindow", "Clear All Items"))
        self.actions_label.setText(_translate("MainWindow", "Actions"))
        self.merge_pdfs_button.setText(_translate("MainWindow", "Merge PDFs"))
        self.ocr_pdfs_button.setText(_translate("MainWindow", "OCR PDFs"))
        self.extract_text_button.setText(_translate("MainWindow", "Extract Text"))
        self.extract_annotations_button.setText(_translate("MainWindow", "Extract Annotations"))
        self.convert_files_button.setText(_translate("MainWindow", "Convert to PDF"))
        self.clear_language_selections_button.setText(_translate("MainWindow", "Clear Language Selections"))
        self.status_of_operations_label.setText(_translate("MainWindow", "Status of Operations"))
        self.exit_button.setText(_translate("MainWindow", "Exit"))

    def get_langs(self):
        """"
        #This is the way to get all installed languages available on the user's system
        #The issue is that Pyinstaller cannnot run subprocess commands
        #So the user can never create a stand-alone application

        self.ocr_languages_list.addItems(subprocess.run(['tesseract', '--list-langs'],
                                                        capture_output=True, text=True).stdout.splitlines())
        self.ocr_languages_list.takeItem(
            0)  # removes first output line that is just informational output for command above
        """
        self.ocr_languages_list.addItem("eng")
        self.ocr_languages_list.addItem("deu")
        self.ocr_languages_list.addItem('script/Fraktur')
        self.ocr_languages_list.addItem("fra")

    def clearFiles(self):
        self.file_list.clear()

    def removeFiles(self):
        if self.file_list.count() > 0:  # prevents crash if nothing in list
            indices = [i.row() for i in self.file_list.selectedIndexes()]

            for idx in indices:
                self.file_list.takeItem(idx)

    def browseFiles(self):

        new_files, _ = QtWidgets.QFileDialog.getOpenFileNames(None, "Select Files to Merge",
                                                              f"/Users/{self.username}/Desktop/",
                                                              "PDF Files(*.pdf)")  # *.pdf limits selection to pdf files only
        if new_files:  # check to make sure files were selected
            # new_files is separate from file_names in case user browses multiple times before merging
            for file_name in new_files:
                self.file_list.addItem(file_name)

    def mergePDFs(self):
        output_file_name = 'merged.pdf'  # default name for file output

        #make sure there are no extraneous files hanging around if user previously converted files
        # i.e. if they convert some first, then run merge, those files would get deleted too
        self.garbageCollection(delete_files=False)

        if self.file_list.count() > 1:  # no merging unless there are enough documents to merge

            output_file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
                None, "Save File", f"/Users/{self.username}/Desktop/merged", "PDF File (*.pdf)")

            for i in range(self.file_list.count()):
                if output_file_name + ".pdf" == self.file_list.item(i).text():
                    error_message = QtWidgets.QMessageBox.critical(None, "Error!",
                                                                   "Error! Your file name is already in use!")
                    return

            self.console_text_box.append(f"Merging PDFs into {output_file_name}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

            if output_file_name:  # check to make sure there is a name
                # user's file name won't include .pdf unless they type it in
                output_file_name = output_file_name + '.pdf'

                # create PDF merger object
                pdf_merger = PdfFileMerger(open(output_file_name, "wb"))

                # Ensure that all files are PDFs
                self.convertFiles()

                for i in range(self.file_list.count()):
                    # get everything from the file list
                    pdf_merger.append(self.file_list.item(i).text())

                pdf_merger.write(output_file_name)
                pdf_merger.close()

                self.garbageCollection(delete_files=True)

                self.console_text_box.append(f"COMPLETED MERGING PROCESS OF ALL FILES INTO {output_file_name}")
                self.console_text_box.repaint()
                QtCore.QCoreApplication.processEvents()

    def ocrPDFs(self):

        languages = [item.text() for item in self.ocr_languages_list.selectedItems()]

        total_time_start = time.time()

        for i in range(self.file_list.count()):
            self.console_text_box.append(f"Currently OCRing File {self.file_list.item(i).text()}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

            input_pdf = self.file_list.item(i).text()
            output_pdf = input_pdf[0:-4] + "-OCR.pdf"

            start_time = time.time()

            #This subprocess version works, but not in a Pyinstaller app
            ##cmd = ["ocrmypdf", "--output-type", "pdf", "--redo-ocr", input_pdf, output_pdf]
            ##subprocess.run(cmd, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)

            ocrmypdf.ocr(input_pdf, output_pdf, output_type="pdf", redo_ocr=True, language=languages, clean=True)

            end_time = time.time()

            self.console_text_box.append(
                f"Finished OCRing File {self.file_list.item(i).text()} in {end_time - start_time} seconds.")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

        total_time_end = time.time()

        self.console_text_box.append(
            f"COMPLETED OCR PROCESS FOR ALL FILES IN {total_time_end - total_time_start} SECONDS")
        self.console_text_box.repaint()
        QtCore.QCoreApplication.processEvents()

    def extractText(self):
        for i in range(self.file_list.count()):
            self.console_text_box.append(f"Extracting Text for File {self.file_list.item(i).text()}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

            input_pdf = self.file_list.item(i).text()
            output_txt = input_pdf[0:-4] + "-text.txt"

            with open(input_pdf, "rb") as pdf_file:
                pdf = pdftotext.PDF(pdf_file)

            with open(output_txt, "w") as text_file:
                for page in pdf:
                    text_file.write(page)

            pdf_file.close()
            text_file.close()

            self.console_text_box.append(f"Completed Extracting Text for File {self.file_list.item(i).text()}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

        self.console_text_box.append(f"COMPLETED EXTRACTING TEXT FOR ALL FILES")
        self.console_text_box.repaint()
        QtCore.QCoreApplication.processEvents()

    def extractAnnotations(self):

        for i in range(self.file_list.count()):
            self.console_text_box.append(f"Extracting Annotations for File {self.file_list.item(i).text()}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

            input_file = self.file_list.item(i).text()
            output_file = input_file[0:-4] + "-annotations.txt"

            cmd = ["pdfannots", input_file, "-o", output_file]

            subprocess.run(cmd, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)

            self.console_text_box.append(f"Completed Extracting Annotations for File {self.file_list.item(i).text()}")
            self.console_text_box.repaint()
            QtCore.QCoreApplication.processEvents()

        self.console_text_box.append(f"COMPLETED EXTRACTING ANNOTATIONS FOR ALL FILES")
        self.console_text_box.repaint()
        QtCore.QCoreApplication.processEvents()

    def convertFiles(self, from_merge=False):

        for i in range(self.file_list.count()):
            file = self.file_list.item(i).text()
            if file.endswith("docx") or file.endswith("doc"):
                self.convert_word_to_pdf(i)

        self.console_text_box.append(f"COMPLETED CONVERTING ALL FILES TO PDF")
        self.console_text_box.repaint()
        QtCore.QCoreApplication.processEvents()

    def convert_word_to_pdf(self, i):
        original_file = self.file_list.item(i).text()
        output_file = ""

        if original_file.endswith("docx"):
            output_file = original_file.replace(".docx", ".pdf")
        elif original_file.endswith("doc"):
            output_file = original_file.replace(".doc", ".pdf")

        docx2pdf.convert(original_file, output_file)

        self.file_list.item(i).setText(output_file)

        # add to garbage bin so if it was just merging, the extraneous files will be deleted later
        self.garbage_bin.append(output_file)

        self.console_text_box.append(
            f"Finished converting Word File {original_file} to PDF file {output_file}.")
        self.console_text_box.repaint()
        QtCore.QCoreApplication.processEvents()

    def garbageCollection(self, delete_files=False):
        #delete any extra files created that user will not need

        if delete_files:
            for item in self.garbage_bin:
                os.remove(item)

        self.garbage_bin.clear()


    def exitApp(self):
        QtCore.QCoreApplication.instance().quit()


class ListDragWidget(QtWidgets.QListWidget):
    """Creates a list widget that allows user to drag and drop PDF
    files into the widget area to add these files."""

    def __init__(self, parent=None):
        super(ListDragWidget, self).__init__(parent)
        self.setDragDropMode(QtWidgets.QAbstractItemView.DragDropMode.DragDrop)
        self.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super(ListDragWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super(ListDragWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            for file in event.mimeData().urls():
                if file.path().endswith('.pdf') or file.path().endswith('.docx') or file.path().endswith('.doc'):  # make sure it is a PDF file
                    self.addItem(file.toLocalFile())
        else:
            super(ListDragWidget, self).dropEvent(event)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
