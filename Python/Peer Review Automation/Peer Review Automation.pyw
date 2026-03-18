import sys
import pandas as pd
import csv
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QFileDialog, QCheckBox, QMessageBox
from PyQt5.QtGui import *
from PyQt5 import uic
from PRA import *
import traceback


class Widget(QMainWindow):
    def __init__(self):
        """_summary_"""
        super().__init__()
        self.file_path = ""
        self.file_path_out = ""
        # Load the ui file
        uic.loadUi("PRAwindow.ui", self)

        # Define our widgets
        # Select SDS file button
        self.button = self.findChild(QPushButton, "pushButton")
        # Select Save Location file button
        self.button2 = self.findChild(QPushButton, "pushButton_2")
        # Select validation file button
        self.button3 = self.findChild(QPushButton, "pushButton_3")
        # Run Peer Review Automation button
        self.button4 = self.findChild(QPushButton, "pushButton_4")

        # Label that shows the file path
        self.label = self.findChild(QLabel, "label")
        # Label that shwos the Validation
        self.label_2 = self.findChild(QLabel, "label_2")
        # Label that shwos the output file path
        self.label_3 = self.findChild(QLabel, "label_3")
        # Label that shows the confirmation if it works
        self.label_confirm1 = self.findChild(QLabel, "label_confirm1")
        # Check box that check naming
        self.checkBox = self.findChild(QCheckBox, "checkBox")
        # Click and select the SDS file
        self.button.clicked.connect(self.clicker)
        # Click and select the Validation file
        self.button2.clicked.connect(self.clicker2)
        # Click and run the Peer Review Automation
        self.button3.clicked.connect(self.clicker3)
        # Click and run the Peer Review Automation
        self.button4.clicked.connect(self.runPRA)
        # Check box is checked
        self.checkBox.stateChanged.connect(self.checkedBox)
        # Show the app
        self.show()
        self.SDS_file_path = None
        self.val_file_path = None

    def checkedBox(self):
        self.checkName = self.checkBox.isChecked()

    def clicker(self):
        # self.label.setText("You Clicked")
        fname = QFileDialog.getOpenFileName(self, "Open SDS File", "", " Excel Files (*.xlsx);; CSV Files (*.csv)")

        # output filename to screen
        if fname:
            self.label.setText("SDS input: " + fname[0])
            self.SDS_file_path = fname[0]

    def clicker2(self):
        # self.label.setText("You Clicked")
        fname = QFileDialog.getOpenFileName(
            self, "Open Validation File", "", " CSV Files (*.csv);; Excel Files (*.xlsx)"
        )

        # output filename to screen
        if fname:
            self.label_2.setText("Validation input: " + fname[0])
            self.val_file_path = fname[0]

    def clicker3(self):
        # self.label.setText("You Clicked")
        fname_out = QFileDialog.getExistingDirectory(self, "Select Directory for Output")

        # output filename to screen
        if fname_out:
            self.label_3.setText("Output: " + fname_out)
            self.file_path_out = fname_out

    def runPRA(self):
        try:
            self.checkedBox()
            if self.SDS_file_path == "":
                QMessageBox.warning(self, "Warning", "Please select a SDS file")
            SDS = SDSData(self.checkName, self.SDS_file_path, self.val_file_path)
            SDS.output(self.file_path_out)
            self.label_confirm1.setText("Confirmed: it works")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e) + traceback.format_exc())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Widget()
    widget.show()
    sys.exit(app.exec())
