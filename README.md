# 📊 Excel Sheet Splitter

A Python web application built with Streamlit that splits Excel files containing multiple sheets into separate Excel files, one file per sheet. **Preserves 100% of original formatting** including colors, fonts, borders, merged cells, and all Excel features.

## ✨ Features

- 📊 Upload Excel files with multiple sheets (`.xlsx`, `.xls`)
- 🔄 Automatically split each sheet into a separate Excel file
- 🎨 **100% Formatting Preservation** (colors, fonts, borders, merged cells, column widths, row heights, etc.)
- 📥 Download all files as ZIP or individually
- 💾 Memory-based processing (no automatic file saves)
- 🔄 Persistent download buttons (available anytime)

## 🚀 Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Usage

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

1. Upload an Excel file with multiple sheets
2. Click "Split Sheets into Separate Files"
3. Download the files (ZIP or individually)

## 📝 Example

**Input:**
- File: `data.xlsx` with sheets: `nov`, `oc`, `p`, `c`

**Output:**
- `nov.xlsx`
- `oc.xlsx`
- `p.xlsx`
- `c.xlsx`

Each file maintains all original formatting!

## 📋 Requirements

- Python 3.7+
- Dependencies: `streamlit`, `pandas`, `openpyxl`, `xlrd`

## 🛠️ Technologies

- **Streamlit** - Web framework
- **Pandas** - Excel file reading
- **OpenPyXL** - Excel formatting preservation
- **xlrd** - Legacy Excel support

## 📁 Project Structure

```
.
├── app.py              # Main application
├── requirements.txt    # Dependencies
└── README.md          # This file
```

## 📌 Notes

- Files are processed entirely in memory (no disk writes)
- Download buttons persist across page interactions
- All formatting is preserved: colors, fonts, borders, merged cells, etc.
- Sheet names are automatically cleaned for valid filenames

## 👨‍💻 Developer

**Developed by:** Ahmed Saeed  
**Last Updated:** 2025

---

⭐ If you find this useful, consider giving it a star!
