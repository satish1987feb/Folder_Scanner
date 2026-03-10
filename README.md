# Folder Scanner

Scan any folder structure on your computer and get a clean, formatted Excel report — listing every file and subfolder with its type, depth level, and location.

## Two Ways to Use

### Option 1: Desktop App (.exe) — No Coding Required

Download `FolderScanner.exe` from the [Releases](https://github.com/satish1987feb/Folder_Scanner/releases) page. Double-click to open, pick your folder, and save the report.

### Option 2: Run with Python (Developers)

```bash
pip install -r requirements.txt
streamlit run folder_scanner.py
```

Open http://localhost:8501, paste a folder path, and generate the report.

## What's in the Report?

| Column | Description |
|--------|-------------|
| **Name** | File or folder name |
| **Type** | Auto-detected (Excel, PDF, Image, Word, Code, etc.) |
| **Level** | Depth in the folder tree |
| **Parent Folder** | Name of the containing folder |
| **Full Path** | Complete path to the item |

## Building the .exe Yourself

```bash
pip install pandas openpyxl pyinstaller
pyinstaller --onefile --windowed --name "FolderScanner" desktop_app.py
```

The `.exe` will be in the `dist/` folder.

## Project Structure

```
folder_scanner.py   ← Streamlit web app (run locally with Python)
desktop_app.py      ← Desktop GUI app (builds into .exe)
build_exe.bat       ← One-click build script for Windows
requirements.txt    ← Python dependencies
```

## Tech Stack

- **Web UI**: Python, Streamlit, Pandas, openpyxl
- **Desktop UI**: Python, tkinter, Pandas, openpyxl
- **Packaging**: PyInstaller
