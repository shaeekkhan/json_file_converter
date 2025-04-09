# JSON File Converter 🚀

## 🔥 Overview
The **JSON File Converter** is a powerful tool designed to convert multiple file formats (**Excel, CSV, Text, Word**) into structured **JSON** files. It features **dual UI options** (Tkinter & PyQt) for a professional and user-friendly experience.

## 🎯 Features
✔ Supports **Excel (.xlsx, .xls), CSV (.csv), Text (.txt, .log, .md), and Word (.docx)**  
✔ **Preview file contents** before conversion  
✔ **Handles merged cells & table formats** in Excel  
✔ **Detects encoding issues automatically**  
✔ **Dual UI options**: Tkinter for simplicity, PyQt for a modern look  

## 💾 Installation
### **Step 1: Clone the repository**
git clone https://github.com/YOUR_USERNAME/json_file_converter.git
cd json_file_converter

**Step 2: Install dependencies**
pip install -r requirements.txt

**🚀 Usage**
**Using the Tkinter GUI**
python gui.py

**Using the PyQt GUI**
python gui_pyqt.py

**Command-Line Mode**
python main.py <input_file> <output_file.json>

**📂 Project Structure**
json_file_converter/
│── main.py                  # Command-line execution
│── gui_tkinter.py           # Tkinter-based GUI
│── gui_pyqt.py              # PyQt-based GUI
│── utils/
│   ├── excel_utils.py       # Excel processing
│   ├── csv_utils.py         # CSV processing
│   ├── text_utils.py        # Text processing
│   ├── docx_utils.py        # Word processing
│   ├── file_selector.py     # File selection logic
│   ├── json_writer.py       # JSON conversion handler
│── README.md                # Project documentation

**🔍 File Conversion Process**
1️⃣ Select a file (from GUI or CLI) 2️⃣ Preview the file content before conversion 3️⃣ Convert file into structured JSON 4️⃣ Save the JSON file to the same directory

**🛠 Dependencies**
openpyxl 📊 (Excel handling)
pandas 📈 (Data processing)
python-docx 📝 (Word document reading)
PyQt6 🎨 (Modern GUI)
tkinter 🖥️ (Simple GUI)

**🤝 Contributing**
🔸 Fork the repository 🔸 Create a new branch 🔸 Make improvements & test them 🔸 Submit a pull request