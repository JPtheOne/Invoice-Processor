import os

import sys
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from processor import process_cfdi, unzip_folder  
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,QWidget, QFileDialog, QLabel, QLineEdit, QMessageBox)


class FolderSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extractor de CFDIs")
        self.setGeometry(300, 200, 400, 350)

        # Main layout
        layout = QVBoxLayout()

        # Folder inputs
        self.label_folder1 = QLabel("Carpeta con archivos ZIP:")
        self.input_folder1 = QLineEdit(self)
        self.button_folder1 = QPushButton("Seleccionar carpeta origen")
        self.button_folder1.clicked.connect(self.select_folder1)

        self.label_folder2 = QLabel("Carpeta de salida para Excel:")
        self.input_folder2 = QLineEdit(self)
        self.button_folder2 = QPushButton("Seleccionar carpeta destino")
        self.button_folder2.clicked.connect(self.select_folder2)

        # Input for output file name
        self.label_output = QLabel("Nombre del archivo de salida (Excel):")
        self.input_output = QLineEdit(self)
        self.input_output.setPlaceholderText("Ejemplo: Excel_final")

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)

        # Button to execute the script
        self.run_button = QPushButton("Ejecutar Script")
        self.run_button.setIcon(QIcon("play_icon.png"))  # Change "play_icon.png" to your icon
        self.run_button.clicked.connect(self.run_script)

        # Add widgets to layout
        layout.addWidget(self.label_folder1)
        layout.addWidget(self.input_folder1)
        layout.addWidget(self.button_folder1)
        layout.addWidget(self.label_folder2)
        layout.addWidget(self.input_folder2)
        layout.addWidget(self.button_folder2)
        layout.addWidget(self.label_output)
        layout.addWidget(self.input_output)
        layout.addWidget(self.status_label)
        layout.addWidget(self.run_button)

        # Configure the main container
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_folder1(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta origen")
        if folder:
            self.input_folder1.setText(folder)

    def select_folder2(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta destino")
        if folder:
            self.input_folder2.setText(folder)

    def run_script(self):
        folder1 = self.input_folder1.text()
        folder2 = self.input_folder2.text()
        output_filename_obtained = self.input_output.text()
        output_filename = output_filename_obtained + ".xlsx"

        if not folder1 or not folder2 or not output_filename:
            QMessageBox.warning(self, "Error", "Por favor selecciona ambas carpetas y un nombre para el archivo de salida.")
            return

        try:
            # Change status to "Processing in progress"
            self.status_label.setText("Ejecución en progreso... Por favor espera.")
            QApplication.processEvents()  # Update the GUI while the script is running

            # Ensure the output folder exists
            os.makedirs(folder2, exist_ok=True)
            output_path = os.path.join(folder2, output_filename)

            # Counters for CFDI types
            counters = {"Total": 0, "I/E": 0, "P": 0, "N": 0, "Desconocido": 0}
            unzipped_folder = os.path.join(folder2, "unzipped")
            os.makedirs(unzipped_folder, exist_ok=True)

            # Process each zip file in the source folder
            for zip_file in os.listdir(folder1):
                zip_path = os.path.join(folder1, zip_file)
                if zip_file.endswith(".zip"):
                    print(f"Procesando archivo ZIP: {zip_file}")
                    extracted_files = unzip_folder(zip_path, unzipped_folder)
                    for cfdi_file in extracted_files:
                        if cfdi_file.endswith(".xml"):
                            counters["Total"] += 1
                            process_cfdi(cfdi_file, output_path, counters)

            # Show success message
            QMessageBox.information(self, "Éxito", "El procesamiento se completó con éxito.")
            print("\nResumen de procesamiento:")
            print(f"Total XML procesados: {counters['Total']}")
            print(f"I/E: {counters['I/E']}, P: {counters['P']}, N: {counters['N']}, Desconocidos: {counters['Desconocido']}")

            # Ask if user wants to perform another operation
            self.ask_for_restart()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error: {e}")
        finally:
            # Clear status
            self.status_label.setText("")

    def ask_for_restart(self):
        reply = QMessageBox.question(
            self,
            "Finalizado",
            "¿Deseas realizar otra operación?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            # Clear all fields
            self.input_folder1.clear()
            self.input_folder2.clear()
            self.input_output.clear()
            self.status_label.setText("")
        else:
            self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FolderSelectorApp()
    window.show()
    sys.exit(app.exec_())
