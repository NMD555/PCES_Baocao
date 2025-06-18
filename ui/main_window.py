from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QCompleter, QComboBox, QMessageBox
from PyQt5 import QtWidgets
from PyQt5.QtCore import QStringListModel, Qt, QSortFilterProxyModel

from ui.start_task_window import StartTaskWindow
from ui.end_task_window import EndTaskWindow

import openpyxl
from typing import List

from PyQt5 import uic
import os

from logic.logic_handle import AutoCompleterComboBox

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        uiFilePath = os.path.join(os.path.dirname(__file__), 'ui.ui')
        uic.loadUi(uiFilePath, self)
        self.setWindowTitle("Report Tool")
        self.worker_name_box = self.findChild(QtWidgets.QComboBox, "comboBox")
        excel_path = os.path.join(os.path.dirname(__file__),'..', 'inputdata.xlsx')
        names = get_names_from_excel(excel_path)

        # Gan completer vao o name
        AutoCompleterComboBox(self.worker_name_box, names)

        self.start_button = self.findChild(QtWidgets.QPushButton, "pushButton")  # nút Start Task
        self.start_button.clicked.connect(self.start_task)
        self.end_button = self.findChild(QtWidgets.QPushButton, "pushButton_2")
        self.end_button.clicked.connect(self.end_task)

        self.second_window = None  # Biến giữ tham chiếu cửa sổ phụ

    def start_task(self):
        if not self.worker_name_box.currentText():
            QMessageBox.warning(self, "Thông báo", "Điền tên nhân viên!")
            return
        worker_name = self.worker_name_box.currentText()
        self.startTask = StartTaskWindow(worker_name)
        self.startTask.resize(500, 200)
        self.startTask.show()
    def end_task(self):
        if not self.worker_name_box.currentText():
            QMessageBox.warning(self, "Thông báo", "Điền tên nhân viên!")
            return
        worker_name = self.worker_name_box.currentText()
        self.endTask = EndTaskWindow(worker_name)
        self.endTask.resize(500, 200)
        self.endTask.show()

def get_names_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook["Name"]
    names = []

    for row in sheet.iter_rows(min_row=2, max_col=1):  # Bỏ qua tiêu đề
        cell = row[0]
        if cell.value:
            names.append(str(cell.value))
    return names


        