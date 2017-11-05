import sys
from Compound_Finder_Functions import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton, QAction, QFileDialog, QLineEdit, QMessageBox, QLabel, QCheckBox
from CofA_Functions import *

library_imported = False

class window(QMainWindow):


    def __init__(self):

        super(window, self).__init__()
        self.setGeometry(150, 150, 440, 280)
        self.setWindowTitle("Compound Finder")



        extractAction = QAction("Quit", self)
        extractAction.setShortcut("Ctrl+X")
        extractAction.setStatusTip("Closes the program")
        extractAction.triggered.connect(self.close_application)

        self.statusBar()

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu("&File")
        fileMenu.addAction(extractAction)

        self.home()

    def home(self):

        library_button = QPushButton("Library", self)
        library_button.resize(75, 25)
        library_button.move(355, 75)
        library_button = library_button.clicked.connect(self.library_text)

        file_button = QPushButton("Input File", self)
        file_button.resize(75, 25)
        file_button.move(355, 105)
        file_button = file_button.clicked.connect(self.file_text)

        run_button = QPushButton("Run", self)
        run_button.resize(75, 25)
        run_button.move(355, 195)
        run_button = run_button.clicked.connect(self.file_open)

        self.textbox = QLineEdit(self)
        self.textbox.move(75, 75)
        self.textbox.resize(280, 25)

        self.textbox2 = QLineEdit(self)
        self.textbox2.move(75, 105)
        self.textbox2.resize(280, 25)

        self.textbox3 = QLineEdit(self)
        self.textbox3.move(75, 135)
        self.textbox3.resize(280, 25)

        self.textbox4 = QLineEdit(self)
        self.textbox4.move(75, 195)
        self.textbox4.resize(280, 25)

        self.library_label = QLabel(self)
        self.library_label.setText("Ret. Library")
        self.library_label.setAlignment(Qt.AlignVCenter)
        self.library_label.move(16, 72)

        self.file_label = QLabel(self)
        self.file_label.setText("GCMS File")
        self.file_label.setAlignment(Qt.AlignVCenter)
        self.file_label.move(26, 102)

        self.output_label = QLabel(self)
        self.output_label.setText("Output name")
        self.output_label.setAlignment(Qt.AlignVCenter)
        self.output_label.move(10, 132)

        self.lot_label = QLabel(self)
        self.lot_label.setText("Lot #")
        self.lot_label.setAlignment(Qt.AlignVCenter)
        self.lot_label.move(47, 192)

        self.b1 = QCheckBox("Generate CofA?", self)
        self.b1.move(75, 165)

        self.show()

    def close_application(self):

        sys.exit()

    def file_open(self):

        library_box_value = self.textbox.text()
        file_box_value = self.textbox2.text()
        output_box = "{}.xls".format(self.textbox3.text())
        lot = self.textbox4.text()
        generate = False

        if self.b1.isChecked() == True:
            generate = True


        if library_box_value.find(".csv") == -1 or file_box_value.find(".txt") == -1 or len(output_box) == 0:
            return

        cofa = []
        if generate and len(lot) == 0:
            return
        elif generate and len(lot) != 0:

            cofa = CofA_format_builder()
            cofa = CofA_Static_additions(cofa)
            cofa = CofA_variable_additions(cofa, lot)

            if cofa == "Missing":
                self.rumple_missing()
                return
            elif cofa == "Doesn't Exist":
                self.error_window_lot()
                return
        try:
            main(generate, file_box_value, library_box_value, output_box, cofa)

            self.textbox2.setText("")
            self.textbox3.setText("")
            self.textbox4.setText("")
            return
        except IOError:
            self.error_window_lib()
            return

    def library_text(self):

        filters = "*.csv"
        selected_filter = "*.csv"
        lib = ""
        lib, _ = QFileDialog.getOpenFileName(self, 'Choose Library',"", filters, selected_filter)
        self.textbox.setText(lib)

    def file_text(self):

        filters = "*.txt"
        selected_filter = "*.txt"
        file_name_text = ""
        file_name_text, _ = QFileDialog.getOpenFileName(self, 'Choose Input File', filters, selected_filter)
        self.textbox2.setText(file_name_text)

    def error_window_lib(self):

        choice = QMessageBox.question(self, 'Error',
                                      "Invalid File Path", QMessageBox.Ok)

        if choice == QMessageBox.Ok:
            self.textbox.setText("")
            self.textbox2.setText("")
            self.textbox3.setText("")
            self.textbox4.setText("")
            pass

    def error_window_lot(self):

        choice4 = QMessageBox.question(self, 'Error',
                                      "That Lot# doesn't exist", QMessageBox.Ok)

        if choice4 == QMessageBox.Ok:
            pass

    def rumple_missing(self):

        choice5 = QMessageBox.question(self, 'Error',
                                      "Rumplestilskin.xls is missing!", QMessageBox.Ok)

        if choice5 == QMessageBox.Ok:
            pass

    def checkbox(self, state):
        if state == Qt.Checked:
            print("Yeah")
            return True
        else:
            return False

def run():

    app = QApplication(sys.argv)
    GUI = window()
    sys.exit(app.exec_())

run()

