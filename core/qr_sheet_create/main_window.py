import os
from PyQt6.QtCore import pyqtSlot,Qt
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QMainWindow

from frontend.gui_main_window import GuiMainWindow

class MainWindow(QMainWindow, GuiMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_gui(self)

    # ui를 제외한 로직은 여기에 빼서 구현하려했는데, 시간상 gui_main_window.py에 구현함
    # def openFile(self):
    #     options = QFileDialog.Options()
    #     options |= QFileDialog.Option.ReadOnly
    #     fileName, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsx)", options=options)
    #     if fileName:
    #         self.lineEdit.setText(fileName)

    # def saveFile(self):
    #     options = QFileDialog.Options()
    #     options |= QFileDialog.Option.ReadOnly
    #     fileName, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)", options=options)
    #     if fileName:
    #         self.lineEdit_2.setText(fileName)

    # def createQR(self):
    #     input_file = self.lineEdit.text()
    #     output_file = self.lineEdit_2.text()

    #     if not input_file or not output_file:
    #         QMessageBox.critical(self, "Error", "Please select input and output files.")
    #         return

    #     if not os.path.exists(input_file):
    #         QMessageBox.critical(self, "Error", "Input file does not exist.")
    #         return