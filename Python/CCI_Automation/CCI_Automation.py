# This Python file uses the following encoding: utf-8
from util import *
from pathlib import Path
import sys
from datetime import date
from EnrollmentLog.EnrollmentLog import EnrollmentLog
from Clockify.Clockify import ClockifyDashboard
from DSMB.DSMB import DSMB
import traceback
from ui_form import Ui_Widget

from PySide6.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox
import warnings

class Widget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Widget()
        self.ui.setupUi(self)
        self.load_ui()


    def load_ui(self):
        self.setWindowTitle("CCI Automation")
        # Combo Box
        comboBox_options = ["","Enrollment Log", "DSMB Report",  "Clockify Dashboard", "Add templated tasks"]
        self.ui.comboBox.addItems(comboBox_options)
        self.ui.comboBox.currentIndexChanged.connect(self.update_comboBox2)

        # Combo Box2
        comboBox2_options = [""]
        self.ui.comboBox_2.addItems(comboBox2_options)
        
        #Date edit check box
        self.ui.checkBox.stateChanged.connect(self.toggle_dateEdit)
        self.ui.checkBox.setVisible(False)
        
        # QDateEdit creation
        self.ui.dateEdit.setVisible(False)  # Date picker hidden by default

        # Choose input Directory
        self.ui.pushButton.clicked.connect(self.clicker)

        # Choose output Directory
        self.ui.pushButton_2.clicked.connect(self.clicker_output_folder)

        # Input text

        # Run
        self.ui.pushButton_3.clicked.connect(self.run)

        # Input File Path Label: label_2
        # Destination File Path Label: label_3
        # Run Confirmation Label: label_5

        ###link interaction afterdeclare variables
        #Combo box 1
        self.onComboBoxChanged()
        self.ui.comboBox.currentIndexChanged.connect(self.onComboBoxChanged)
        #Combo box 2
        self.onComboBoxChanged2()
        self.ui.comboBox_2.currentIndexChanged.connect(self.onComboBoxChanged2)
        #Line edit for Input
        self.on_lineEdit_changed()
        self.ui.lineEdit.textChanged.connect(self.on_lineEdit_changed)
        self.input_file_path = None

    def clicker(self):
        """ Open a file dialog when the user clicks the button.
            Options: DSMB Report, Enrollment Log, Clockify Dashboard
        """
        if self.selected_option == "DSMB Report" or self.selected_option == "Enrollment Log":
            default_dir = "C:/Users/hmn39/Downloads"
            file_path, _ = QFileDialog.getOpenFileName(self, "Select the raw file of Clockify Dashboard", default_dir, "Zip files (*.zip)")
            # output filename to screen
            if file_path:
                self.ui.label_2.setText(file_path)
                self.input_file_path = file_path

        elif self.selected_option == "Clockify Dashboard":
            default_dir = "C:/Users/hmn39/Downloads"
            file_path, _ = QFileDialog.getOpenFileName(self, "Select the raw file of Clockify Dashboard", default_dir, "CSV files (*.csv)")
            # output filename to screen
            if file_path:
                self.ui.label_2.setText(file_path)
                self.input_file_path = file_path
                
    def update_comboBox2(self, index):
        self.ui.comboBox_2.clear()
        if index == 1:
            self.ui.comboBox_2.addItems(["03821", "11823", "12423", "15122", "15420", "16321"])
        elif index == 2:
            self.ui.comboBox_2.addItems(["15420", "12423"])
        elif index == 3:
            self.ui.comboBox_2.addItems(clockify_get_list_projects(clockify_get_api_key(),clockify_get_workplace_id()))
        elif index == 4:
            self.ui.comboBox_2.addItems(clockify_get_list_projects(clockify_get_api_key(),clockify_get_workplace_id()))
            
        if index == 4:
            self.ui.pushButton.setEnabled(False)
            self.ui.pushButton_2.setEnabled(False)
        else:
            self.ui.pushButton.setEnabled(True)
            self.ui.pushButton_2.setEnabled(True)

    def toggle_dateEdit(self, state):
        """ Change the state of the date picker based on the checkbox state.
            If checkbox is checked (state == 2), show the date picker. Otherwise, hide it.
        Args:
            state (_type_): _description_
        """
        
        self.ui.dateEdit.setVisible(state == 2)

    def clicker_output_folder(self):
        #self.ui.label.setText("You Clicked")
        default_dir = "C:/Users/hmn39/Downloads"
        output_folder_name = QFileDialog.getExistingDirectory(self, "Select Output Folder", default_dir)
        # output filename to screen
        if output_folder_name:
            self.ui.label_3.setText(output_folder_name)
            self.output_folder_name = output_folder_name

    def onComboBoxChanged(self):
        """_summary_
        """
        # Get the selected item text
        self.selected_option = self.ui.comboBox.currentText()
        if self.selected_option == "Enrollment Log":
            self.ui.checkBox.setVisible(True)
            self.ui.lineEdit.setText(datetime.now().strftime("%y%m%d") + "-" + self.selected_option2 + " Enrollment Log")
            self.ui.lineEdit.setEnabled(True)
            self.ui.lineEdit.show()
        elif self.selected_option == "DSMB Report":
            self.ui.checkBox.setVisible(True)
            self.ui.lineEdit.setText(datetime.now().strftime("%y%m%d") + "-" + self.selected_option2 + " DSMB Report")
            self.ui.lineEdit.setEnabled(True)
        elif self.selected_option == "Clockify Dashboard":
            self.ui.checkBox.setVisible(False)
            self.ui.dateEdit.setVisible(False)
            self.ui.lineEdit.setText(date.today().strftime("%y%m%d") + "-" + self.selected_option2 + "-Clockify Dashboard")
            self.ui.lineEdit.setEnabled(False)
        elif self.selected_option == "Add templated tasks":
            self.ui.checkBox.setVisible(False)
            self.ui.dateEdit.setVisible(False)
            self.ui.lineEdit.setText("")
            self.ui.lineEdit.setEnabled(False)

    def onComboBoxChanged2(self):
        """_summary_
        """
        # Get the selected item text
        self.selected_option2 = self.ui.comboBox_2.currentText()
        if self.selected_option == "Enrollment Log":
            self.ui.lineEdit.setText(date.today().strftime("%y%m%d") + "-" + self.selected_option2 + " Enrollment Log")
        elif self.selected_option == "Clockify Dashboard":
            self.ui.lineEdit.setText(date.today().strftime("%y%m%d") + "-" + self.selected_option2 + "-Clockify Dashboard")
        elif self.selected_option == "Add templated tasks":
            self.ui.lineEdit.setText("")
        elif self.selected_option == "DSMB Report":
            self.ui.lineEdit.setText(datetime.now().strftime("%y%m%d") + "-" + self.selected_option2 + " DSMB Report")
            self.ui.lineEdit.setEnabled(True)

    def on_lineEdit_changed(self, text = None):
        #update the changed line edit
        self.output_file_name = text

    def run(self):

        if self.selected_option == "Enrollment Log":
            #option for enrollment log
            if not self.selected_option2 or not self.input_file_path or not self.output_folder_name or not self.output_file_name:
                QMessageBox.warning(self, 'Warning', 'Please make sure all those selections are entered correctly.')
            else:
                try:
                    if self.ui.checkBox.isChecked():
                        cut_off_date = self.ui.dateEdit.date().toPython()
                        EnrollmentLog(self.selected_option2, self.input_file_path, self.output_folder_name, self.output_file_name, cut_off_date)
                    else:
                        EnrollmentLog(self.selected_option2, self.input_file_path, self.output_folder_name, self.output_file_name)
                    self.ui.label_5.setText("Confirmed: Enrollment Log " + self.selected_option2 + " study selected. Output file name is " +  self.output_file_name + ".xlsx")
                except Exception as e:
                    self.ui.label_5.setText("Error encountered.")
                    QMessageBox.warning(self, 'Error', str(e) + traceback.format_exc())
        elif self.selected_option == "DSMB Report":
            #option for enrollment log
            if not self.selected_option2 or not self.input_file_path or not self.output_folder_name:
                QMessageBox.warning(self, 'Warning', 'Please make sure all those selections are entered correctly.')
            else:
                try:
                    if self.ui.checkBox.isChecked():
                        cut_off_date = self.ui.dateEdit.date().toPython()
                        DSMB(self.selected_option2, self.input_file_path, self.output_folder_name, self.ui.lineEdit.text(), cut_off_date)
                    else:
                        DSMB(self.selected_option2, self.input_file_path, self.output_folder_name, self.ui.lineEdit.text())
                    self.ui.label_5.setText("Confirmed: DSMB Report " + self.selected_option2 + " study selected. Output file name is "  +  self.ui.lineEdit.text() + ".xlsx")
                except Exception as e:
                    self.ui.label_5.setText("Error encountered.")
                    QMessageBox.warning(self, 'Error', str(e) + traceback.format_exc())
        elif self.selected_option == "Clockify Dashboard":
            #option for Clockify Dashboard
            if not self.selected_option2 or not self.output_folder_name:
                QMessageBox.warning(self, 'Warning', 'Please make sure all those selections are entered correctly.')
            else:
                try:
                    ClockifyDashboard(self.selected_option2, self.output_folder_name, self.input_file_path)
                    self.ui.label_5.setText("Confirmed: ClockifyDashboard " + self.selected_option2 + " project selected. Output file name is " +  date.today().strftime("%y%m%d") + "-" + self.selected_option2 + "-Clockify Dashboard" + ".xlsx")
                except Exception as e:
                    self.ui.label_5.setText("Error encountered.")
                    QMessageBox.warning(self, 'Error', str(e) + traceback.format_exc())
        elif self.selected_option == "Add templated tasks":
            if not self.selected_option2:
                QMessageBox.warning(self, 'Warning', 'Please make sure all those selections are entered correctly.')
            else:
                try:
                    clockify_create_tasks(clockify_get_api_key(), clockify_get_workplace_id(), self.selected_option2)
                    self.ui.label_5.setText("Confirmed: template task has been added to " + self.selected_option2)
                except Exception as e:
                    self.ui.label_5.setText("Error encountered.")
                    QMessageBox.warning(self, 'Error', str(e) + traceback.format_exc())


if __name__ == "__main__":
    warnings.filterwarnings("ignore", category=RuntimeWarning)
    app = QApplication(sys.argv)

    widget = Widget()
    widget.show()
    sys.exit(app.exec())
