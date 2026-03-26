# 📘 Excel → LaTeX Manual Generator

Desktop application that generates **technical manuals in PDF** starting from structured Excel files.

The tool reads an Excel workbook, lets the user choose machine/version/language via GUI, converts the content to LaTeX and compiles the final PDF automatically.

---

## ✨ Features

- GUI interface (Tkinter)
- Excel → LaTeX → PDF pipeline
- Automatic table conversion (merged cells + formatting)
- Image integration from images folder
- Multi-language manuals
- Machine / Gamma / Version selection
- Preview mode for debugging

---

## 🗂 Project structure
manual-generator/
│
├── src/ → application source code
├── assets/ → images used inside manuals
├── data/ → example Excel file
├── requirements.txt
├── README.md
└── .gitignore


---

## 🚀 Run the project

### 1. Clone repository
git clone https://github.com/YOUR_USERNAME/manual-generator.git

cd manual-generator


### 2. Create virtual environment
python -m venv venv
venv\Scripts\activate # Windows


### 3. Install dependencies
pip install -r requirements.txt


### 4. Run the app
python src/main.py


---

## 📄 Example input file

Example Excel file available in:
data/Manuale_esempio.xlsm