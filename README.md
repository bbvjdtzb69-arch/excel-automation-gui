# Excel Automation GUI Tool

A Python GUI tool for merging, filtering, and organizing Excel/CSV files automatically.  
This tool allows users to process multiple Excel and CSV files at once with just a few clicks.

---

## 🚀 Main Features

- **Batch import**: Supports CSV and Excel files (.xlsx / .xls)  
- **Merge files**: Combine multiple files into one DataFrame  
- **Remove unwanted columns**: Specify columns to delete easily  
- **Filter data**: Extract rows based on column values  
- **Split sheets**: Automatically divide output into separate sheets based on column values  
- **GUI-based operation**: Easy-to-use interface via Tkinter

---

## 📦 Output File


- If a sheet-splitting column is specified → separate sheets for each unique value  
- If no split column → output all data in a single sheet named `result`

---

## 🖥 Demo / GUI Screenshot

**Screenshot placeholder:**  
<img src="images/screenshot.png" width="400">

> Replace `images/screenshot.png` with your actual screenshot file.

---

## 🛠 Installation Requirements

- Python 3.x
- pip

### Install dependencies
```bash
pip install pandas openpyxl tkinterdnd2

git clone https://github.com/<your-username>/excel-automation-gui.git

pip install pandas openpyxl tkinterdnd2

python main.py

excel-automation-gui/
 ├─ main.py
 ├─ README.md
 ├─ LICENSE
 ├─ .gitignore
 └─ images/
      └─ screenshot.png


---
