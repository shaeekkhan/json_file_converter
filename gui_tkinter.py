import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd
from docx import Document
from utils.json_writer import file_to_json


class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")
        self.root.geometry("600x400")

        tk.Label(root, text="Select a File:", font=("Arial", 12)).pack(pady=10)
        self.select_button = tk.Button(root, text="Browse", command=self.select_file)
        self.select_button.pack()

        self.file_label = tk.Label(root, text="", fg="blue", font=("Arial", 10))
        self.file_label.pack()

        tk.Label(root, text="Preview Data:", font=("Arial", 12)).pack(pady=10)
        self.preview_area = scrolledtext.ScrolledText(root, width=70, height=10)
        self.preview_area.pack()

        self.convert_button = tk.Button(root, text="Convert to JSON", command=self.convert_file)
        self.convert_button.pack(pady=10)

        self.status_label = tk.Label(root, text="", fg="green", font=("Arial", 10))
        self.status_label.pack()

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select a file",
                                               filetypes=[("All files", "*.*"),
                                                          ("Excel files", "*.xlsx;*.xls"),
                                                          ("CSV files", "*.csv"),
                                                          ("Text files", "*.txt;*.log;*.md"),
                                                          ("Word files", "*.docx")])
        if file_path:
            self.file_label.config(text=f"Selected: {file_path}")
            self.file_path = file_path
            self.preview_data(file_path)

    def preview_data(self, file_path):
        _, file_extension = os.path.splitext(file_path)

        # Preview Excel Files
        if file_extension.lower() in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets
                first_sheet_name = list(df.keys())[0]  # Get the first sheet
                preview_text = df[first_sheet_name].head().to_string(index=False)  # Show first few rows
                self.preview_area.delete("1.0", tk.END)
                self.preview_area.insert(tk.END, preview_text)
            except Exception as e:
                self.preview_area.delete("1.0", tk.END)
                self.preview_area.insert(tk.END, f"Error previewing Excel file: {str(e)}")
            return

        # Preview Word (.docx) Files
        elif file_extension.lower() == '.docx':
            try:
                doc = Document(file_path)
                paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
                preview_text = "\n".join(paragraphs[:5])  # Show first 5 paragraphs
                self.preview_area.delete("1.0", tk.END)
                self.preview_area.insert(tk.END, preview_text)
            except Exception as e:
                self.preview_area.delete("1.0", tk.END)
                self.preview_area.insert(tk.END, f"Error previewing Word file: {str(e)}")
            return

        # Preview Normal Text-Based Files
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                data = file.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, "r", encoding="ISO-8859-1") as file:
                    data = file.read()
            except UnicodeDecodeError:
                self.preview_area.delete("1.0", tk.END)
                self.preview_area.insert(tk.END, "Error: Unsupported file encoding.")
                return

        self.preview_area.delete("1.0", tk.END)
        self.preview_area.insert(tk.END, data[:1000])  # Limit preview size

    def convert_file(self):
        if not hasattr(self, 'file_path'):
            messagebox.showerror("Error", "No file selected!")
            return

        output_file = self.file_path + ".json"
        file_to_json(self.file_path, output_file)
        self.status_label.config(text=f"Conversion complete! JSON saved as: {output_file}")
        messagebox.showinfo("Success", f"File converted successfully!\nSaved as: {output_file}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()
