import sys
from Compound_Finder_Functions import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton, QAction, QFileDialog, QLineEdit, QMessageBox, QLabel, QCheckBox
from PyQt5.QtGui import *

library_imported = False

class window(QMainWindow):


    def __init__(self):

        super(window, self).__init__()
        self.setGeometry(50, 50, 440, 280)
        self.setWindowTitle("Compound Finder")
        #self.setWindowIcon(QIcon("pythonlogo.png"))



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
        output_box = self.textbox3.text()
        lot = self.textbox4.text()
        generate = False

        if self.b1.isChecked() == True:
            generate = True


        if library_box_value.find(".csv") == -1 or file_box_value.find(".txt") == -1 or len(output_box) == 0:
            return

        if output_box.find("/") == -1 and output_box.find("\\") == -1:

            compound_list = import_library(library_box_value)
            file_converter(file_box_value)
            output_box += ".xls"
            x = Final_File_Creator(compound_list, output_box, generate, lot)

            # --------------------------------------------------------------
            # Error handling for lot.
            # Final file returns False if lot# is not matched
            # --------------------------------------------------------------

            print(x)
            if x == False:
                self.error_window_lot()
                return


            self.textbox2.setText("")
            self.textbox3.setText("")
            self.textbox4.setText("")
            return

        else:
            self.error_window_output()
            self.textbox3.setText("")
            return

    def library_text(self):

        filters = "*.csv"
        selected_filter = "*.csv"
        lib = ""
        lib, _ = QFileDialog.getOpenFileName(self, 'Choose Library',"", filters, selected_filter)
        if lib.find(".csv") == -1 and len(lib) > 0:
            self.error_window_lib()
            return
        elif lib == "":
            return
        else:
            self.textbox.setText(lib)

    def file_text(self):

        filters = "*.txt"
        selected_filter = "*.txt"
        file_name_text = ""
        file_name_text, _ = QFileDialog.getOpenFileName(self, 'Choose Input File', filters, selected_filter)
        if file_name_text.find(".txt") == -1 and len(file_name_text) > 0:
            self.error_window_file()
            return
        elif file_name_text == "":
            return
        else:
            self.textbox2.setText(file_name_text)

    def error_window_lib(self):

        choice = QMessageBox.question(self, 'Error',
                                      "Please only load .csv files", QMessageBox.Ok)

        if choice == QMessageBox.Ok:
            self.textbox.setText("")
            pass

    def error_window_file(self):

        choice2 = QMessageBox.question(self, 'Error',
                                       "Please only load .txt files", QMessageBox.Ok)

        if choice2 == QMessageBox.Ok:
            self.textbox2.setText("")
            return

    def error_window_output(self):

        choice3 = QMessageBox.question(self, 'Error',
                                      "File name cannot contain slashes", QMessageBox.Ok)

        if choice3 == QMessageBox.Ok:
            pass

    def error_window_lot(self):

        choice4 = QMessageBox.question(self, 'Error',
                                      "That Lot# doesn't exist", QMessageBox.Ok)

        if choice4 == QMessageBox.Ok:
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

