#!/usr/bin/env python3
#!/usr/bin/env python3
from autoroster import core
from PyQt5 import QtGui, QtWidgets
import sys
import time


class MainWindow(QtWidgets.QWidget):

    def __init__(self):
        super(MainWindow, self).__init__()

        self.initUI()

    def initUI(self):
        grid = QtWidgets.QGridLayout()

        self.roster_name = QtWidgets.QLabel('No Roster Selected')
        self.roster_list = QtWidgets.QListWidget()
        self.roster_list.setSelectionMode(
            QtWidgets.QAbstractItemView.ExtendedSelection)

        report_type_label = QtWidgets.QLabel("Report Type: ")
        self.report_type = QtWidgets.QComboBox()
        self.report_type.addItem("Daily")
        self.report_type.addItem("Weekly")

        open_button = QtWidgets.QPushButton("Open")
        open_button.clicked.connect(self.open_roster)

        generate_button = QtWidgets.QPushButton("Generate")
        generate_button.clicked.connect(self.generate_report)

        # Set grid layout
        grid.addWidget(self.roster_name, 0, 0)
        grid.addWidget(open_button, 0, 1)
        grid.addWidget(report_type_label, 1, 0)
        grid.addWidget(self.report_type, 1, 1)
        grid.addWidget(self.roster_list, 2, 0, 2, 1)
        grid.addWidget(generate_button, 4, 0, 2, 1)

        self.setLayout(grid)
        self.setWindowTitle('AutoRoster')
        self.show()

    def open_roster(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Open File', "", "Excel Files (*.xlsx)")[0]
        if filename:
            self.roster_list.clear()
            shortend = filename.split('/')[-1]
            self.roster_name.setText(shortend)
            self.report_dataframe = core.read_in_roster(filename)
            exam_list = core.get_exam_list(self.report_dataframe)
            for exam in exam_list:
                self.roster_list.addItem(exam)
        else:
            self.roster_name.setText('No Roster Selected')

    def generate_report(self):
        selection = self.roster_list.selectedItems()
        if selection:
            # selection is a list of object, change into strings
            selection = [i.text() for i in selection]
            wb = core.make_workbook()

            if str(self.report_type.currentText()) == 'Daily':
                default = time.strftime(
                    "Daily %m-%d-%y.xlsx", time.localtime())
                date = time.strftime("%x", time.localtime())
                for exam in selection:
                    student_info = core.get_student_info(
                        exam, self.report_dataframe)
                    asterisk_free = core.ignore_asterisk(student_info)
                    core.make_daily_report(wb, exam, date, asterisk_free)

            elif str(self.report_type.currentText()) == 'Weekly':
                default = time.strftime(
                    "Weekly %m-%d-%y.xlsx", time.localtime())
                year = time.strftime("%Y", time.localtime())
                for exam in selection:
                    student_info = core.get_student_info(
                        exam, self.report_dataframe)
                    asterisk_free = core.ignore_asterisk(student_info)
                    core.make_weekly_report(wb, exam, year, asterisk_free)

            else:
                QtWidgets.QMessageBox.critical(
                    self, "Error", "Error Unexpected report type selected, please contact Daniel at u0346076@utah.edu")

            core.delete_blank_sheets(wb)
            filename = QtWidgets.QFileDialog.getSaveFileName(
                self, 'Save File', default, "Excel Files (*.xlsx)")[0]
            core.save_workbook(wb, filename)
            QtWidgets.QMessageBox.information(
                self, "Message", "File successfully created")

        else:
            QtWidgets.QMessageBox.critical(self, "Error", "No Exam Selected")


def main():

    app = QtWidgets.QApplication([])
    ex = MainWindow()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
    sys.exit()
