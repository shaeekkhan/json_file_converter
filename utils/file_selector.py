import tkinter as tk
from tkinter import filedialog


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select a file",
                                           filetypes=[("Excel files", "*.xlsx;*.xls"),
                                                      ("CSV files", "*.csv"),
                                                      ("Text files", "*.txt;*.log;*.md"),
                                                      ("All files", "*.*")])
    return file_path
