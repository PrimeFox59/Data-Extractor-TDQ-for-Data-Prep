# 📊 [DATA EXTRACTOR] TECH. DATA QUOTATION – Excel VBA Automation Tool

![Built with VBA](https://img.shields.io/badge/Built_with-VBA-blue)
![Platform](https://img.shields.io/badge/Platform-Excel-lightgrey)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)


A smart Excel-based automation tool to scan folders, extract cost data from multiple TDQ Excel files, and summarize them in one structured dashboard. Built with pure VBA for internal use in manufacturing cost control and data traceability.

## ✅ Features

- 🔍 Recursively scan folders and extract `.xls`/`.xlsx` files
- 📂 Automatically skip excluded folders and temporary files (`~$`)
- 📎 Generate clickable hyperlinks to file locations
- 📊 Extract cost data and metadata from specific sheets with revision logic
- 📈 Predict and summarize MCT, profit, productivity, and more
- ⏳ Background loading with a visual userform progress bar

## 🧠 Built With

- VBA (Visual Basic for Applications)
- Excel `.xlsm`
- FileSystemObject for folder traversal
- Dynamic sheet detection with revision pattern search
- Custom UserForm loading screen

## 📂 Folder Structure

- `vba_modules/` → Contains exported `.bas` modules:
  - `ListFileDenganTable.bas` → For file scanning & table listing
  - `AmbilDataBiaya.bas` → For multi-sheet, multi-revision data extraction
- `TRIAL_RESUME_TDQ.xlsm` → Main working file
- `screenshots/` → Visual previews (optional)

## 💡 Usage

1. Open `TRIAL_RESUME_TDQ.xlsm`
2. Configure folder path & exclusions in the **Setting** sheet
3. Run `ListFileDenganTable` to populate file list
4. Select file(s) and run `AmbilDataBiayaMultiSelect`
5. Extracted data will appear in columns C–S


## 🛡️ License

MIT (You may freely use and modify, attribution appreciated)

