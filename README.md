# DragDrop2PDF 🚀  
**Excel & Word をドラッグ＆ドロップするだけで PDF に変換！**  
_Convert Excel & Word files to PDF by drag & drop!_

---

## 🌟 特徴 | Features  
✅ **ドラッグ＆ドロップだけでOK！** | **Just drag & drop to convert to PDF!**  
✅ **Acrobat 不要！ Windows で動作！** | **No need for Acrobat! Works on Windows!**  
✅ **Excel や Word の印刷設定をそのまま反映！** | **Keeps print settings from Excel & Word!**  

---

## 🛠 インストール方法（Windows 用） | Installation (Windows)  
1. **Python をインストール | Install Python**（[公式サイト | Python Official Site](https://www.python.org/downloads/)）  
2. **必要なライブラリをインストール | Install required libraries**  
   ```sh
   pip install pywin32
   pyinstaller --onefile --noconsole convert_to_pdf.py
