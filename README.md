# Excel Village-Wise Splitter

This tool reads an Excel file and automatically creates
separate sheets for each village, along with:

- An **Index sheet** with hyperlinks
- A **Master sheet** containing all data
- Unique sheet-name cleanup and collision handling

Built using Pandas + XlsxWriter.

## 🚀 Features
- Splits rows by target column (`शहर/गांव`)
- Auto-creates sheet hyperlinks
- Prevents duplicate sheet names
- Supports large datasets

## 🛠️ Usage
1. Update the config values in `ocr.py`:
   - `file_path`
   - `output_file`
   - `target_column`

2. Run:
├─ src/
│  └─ ocr.py
├─ samples/
│  └─ example_input.xlsx   (optional)
├─ .gitignore
├─ README.md
├─ LICENSE
└─ requirements.txt
pandas
xlsxwriter
openpyxl
