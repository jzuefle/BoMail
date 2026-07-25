import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QMessageBox, QVBoxLayout, QLineEdit, QLabel,
    QFormLayout, QFileDialog, QHBoxLayout, QComboBox, QTableView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator

from create_outlook_mail import create_outlook_mail
from import_piece_table import import_pice_table
from pandasToQT import PandasModel


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.setWindowTitle('Bestelltool FSA')
        self.setGeometry(200, 200, 700, 400)

        self.xls_adress = QLineEdit(r"W:\KONSTRUKTION-3D\PROJEKTE\4491 Leonardo Matrizen BW\CAD-Daten\S31020-0709-00.iam")
        self.file_selection = QPushButton("open file")
        self.file_selection.clicked.connect(self.choose_file)
        self.creat_mail_button = QPushButton("Erstelle eMail")
        self.creat_mail_button.clicked.connect(self.create_mail)
        self.number_to_build = QLineEdit('1')
        self.number_to_build.setValidator(QIntValidator(1, 1000, self))
        self.project_number = QLineEdit('4491')
        self.parts_to_order = QComboBox()
        self.parts_to_order.addItems(['Kaufteile', 'Aluminium', 'Kunststoff', 'Kupfer', 'Edelstahl', 'Federkontakte', 'Sonstiges'])

        self.piece_table_viewed = QTableView()
        
        hbox = QHBoxLayout()
        hbox.addWidget(self.xls_adress)
        hbox.addWidget(self.file_selection)
        
        # Layout
        layout = QFormLayout()
        layout.addRow(QLabel("Teile Tabelle:"), hbox)
        layout.addRow(QLabel("Zu bauen:"), self.number_to_build)
        layout.addRow(QLabel("Projektnummer"), self.project_number)
        layout.addRow(QLabel("Zu beschaffen"), self.parts_to_order)
        
        layout.addRow(self.creat_mail_button) #, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addRow(self.piece_table_viewed)
        self.setLayout(layout)

    def show_message(self):
        """Show a message box when the button is clicked."""
        QMessageBox.information(self, "Hello", "You clicked the button!")

    def create_mail(self):
        piecetable = import_pice_table(self.xls_adress.text())
        piecetable.change_bouild_number(int(self.number_to_build.text()))
        
        model = PandasModel(piecetable.df)
        self.piece_table_viewed.setModel(model)

        projectNumber = self.project_number.text()
        toOrderString = self.parts_to_order.currentText()
        toFillInTable = piecetable.create_list_to_order(toOrderString)

        message_text = f"""Hallo Daniela,
            bitte gemäß Stückliste und Zeichnungen {toOrderString} anfragen und für das Projekt {projectNumber} bestellen:
            """
        end_text = """Beste Grüße
            Johannes
            """
    
        html_piecetable = create_outlook_mail.create_html_body(message_text, toFillInTable, end_text)
    
        addFile = piecetable.all_drawings_to_order(toOrderString)
    
        create_outlook_mail.create_outlook_email(
           to_recipients = ["d.rakonic@fsant.de"],
            subject = projectNumber + ' - ' + toOrderString,
            body = html_piecetable,
            attachments= addFile
        )
    
    def choose_file(self):
        if self.xls_adress != "":
            folderpath = "\\".join(self.xls_adress.text().split('\\')[0:-1])
        else:
            folderpath = ""
        file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select a File",
                folderpath,  # Start directory (empty = current dir)
                "All Files (*);;Tabelle (*.xlsx);;Baugruppe (*.iam)"
            )
        self.xls_adress.setText(file_path)
        
        piecetable = import_pice_table(self.xls_adress.text())
        piecetable.change_bouild_number(int(self.number_to_build.text()))
        
        model = PandasModel(piecetable.df)
        self.piece_table_viewed.setModel(model)


def main():
    # Create the application
    app = QApplication(sys.argv)

    # Create and show the main window
    window = MyWindow()
    window.show()

    # Run the application event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()