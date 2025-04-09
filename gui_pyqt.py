import sys
import os
import json
import pandas as pd
from docx import Document
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QFileDialog, QMessageBox
from utils.json_writer import file_to_json


class FileConverterApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("File Converter")
        self.setGeometry(200, 200, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Select a File:")
        layout.addWidget(self.label)

        self.select_button = QPushButton("Browse")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.file_label = QLabel("")
        layout.addWidget(self.file_label)

        self.preview_area = QTextEdit()
        self.preview_area.setReadOnly(True)
        layout.addWidget(self.preview_area)

        self.convert_button = QPushButton("Convert to JSON")
        self.convert_button.clicked.connect(self.convert_file)
        layout.addWidget(self.convert_button)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select a file",
                                                   "", "All Files (*.*);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;Text Files (*.txt *.log *.md);;Word Files (*.docx)")
        if file_path:
            self.file_label.setText(f"Selected: {file_path}")
            self.file_path = file_path
            self.preview_data(file_path)

    def preview_data(self, file_path):
        _, file_extension = os.path.splitext(file_path)

        # Preview Excel Files
        if file_extension.lower() in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets
                first_sheet_name = list(df.keys())[0]  # Get the first sheet name
                preview_text = df[first_sheet_name].head().to_string(index=False)  # Preview first few rows
                self.preview_area.setText(preview_text)
            except Exception as e:
                self.preview_area.setText(f"Error previewing Excel file: {str(e)}")
            return

        # Preview Word (.docx) Files
        elif file_extension.lower() == '.docx':
            try:
                doc = Document(file_path)
                paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
                preview_text = "\n".join(paragraphs[:5])  # Show first 5 paragraphs
                self.preview_area.setText(preview_text)
            except Exception as e:
                self.preview_area.setText(f"Error previewing Word file: {str(e)}")
            return

        # Preview Normal Text Files
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                data = file.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, "r", encoding="ISO-8859-1") as file:
                    data = file.read()
            except UnicodeDecodeError:
                self.preview_area.setText("Error: Unsupported file encoding.")
                return

        self.preview_area.setText(data[:1000])  # Limit preview size

    def convert_file(self):
        if not hasattr(self, 'file_path'):
            QMessageBox.critical(self, "Error", "No file selected!")
            return

        output_file = self.file_path + ".json"
        file_to_json(self.file_path, output_file)
        self.status_label.setText(f"Conversion complete! JSON saved as: {output_file}")
        QMessageBox.information(self, "Success", f"File converted successfully!\nSaved as: {output_file}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec())
