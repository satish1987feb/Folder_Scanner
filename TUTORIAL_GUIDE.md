# YouTube Tutorial Guide — Folder Scanner

> **Recommended format**: 1 video, 2 parts with timestamps.
> Total estimated length: 15–20 minutes.

---

## Video Title Ideas

- "Build a Folder Scanner App in Python — From Code to .exe"
- "Python Project: Scan Any Folder & Generate Excel Reports (+ Free .exe)"
- "Turn a Python Script into a Desktop App (.exe) — Folder Scanner Tutorial"

## Video Description Template

```
Build a Python app that scans any folder on your computer and generates
a formatted Excel report — then package it as a standalone .exe anyone
can use. No coding knowledge needed to use the .exe!

TIMESTAMPS:
00:00 - What we're building
XX:XX - Part 1: Using the app (no coding needed)
XX:XX - Part 2: Building it from scratch with Python
XX:XX - Step 1: Project setup
XX:XX - Step 2: Core scanning logic
XX:XX - Step 3: Streamlit web UI
XX:XX - Step 4: Desktop GUI with tkinter
XX:XX - Step 5: Packaging as .exe with PyInstaller
XX:XX - Wrap up

SOURCE CODE: https://github.com/satish1987feb/Folder_Scanner
DOWNLOAD .EXE: https://github.com/satish1987feb/Folder_Scanner/releases

#python #programming #tutorial #automation #excel
```

---

## INTRO (1–2 min)

**What to say:**
- "Have you ever needed to get a list of every file and folder inside a
  directory? Maybe you manage client projects, or organize a media library,
  or just want to know what's inside a big folder."
- "Today I'll show you a tool that scans any folder and gives you a clean
  Excel report — with file types, folder levels, and full paths."
- "If you're NOT a programmer — stay for Part 1. I'll show you how to
  download and use the app in under 2 minutes."
- "If you ARE a developer — stick around for Part 2 where we build the
  entire thing from scratch in Python."

**What to show on screen:**
- Quick preview of the final Excel report (open a sample .xlsx)
- Quick flash of the desktop app UI
- Quick flash of the Streamlit web UI (single page with Browse button)

---

## PART 1: For Non-Coders — Using the .exe (2–3 min)

> **Audience**: Anyone. No coding, no setup, no Python.

### Step 1: Download the .exe

**What to say:**
- "Go to the GitHub releases page — link in the description."
- "Download `FolderScanner.exe`. That's all you need — one file."

**What to show:**
- Open browser → GitHub releases page
- Click download → show the file in Downloads folder

### Step 2: Run the App

**What to say:**
- "Double-click the .exe to open it."
- "You'll see a simple window with a Browse button."

**What to show:**
- Double-click `FolderScanner.exe`
- The app window opens

### Step 3: Select a Folder

**What to say:**
- "Click Browse, pick any folder on your computer."
- "I'll scan this project folder as an example."

**What to show:**
- Click "Browse..." → navigate to a sample folder → select it
- The path appears in the text field

### Step 4: Generate the Report

**What to say:**
- "Click Generate Report. It asks you where to save the Excel file."
- "Pick a location, hit Save, and... done! It even asks if you want to
  open it right away."

**What to show:**
- Click "Generate Report"
- Save dialog → choose Desktop → Save
- "Open file?" dialog → click Yes
- Excel opens → scroll through the report
- Point out: Name, Type, Level, Parent Folder, Full Path columns
- Point out: formatted headers, filters, auto-sized columns

### Wrap up Part 1

**What to say:**
- "That's it. No Python, no terminal, no coding. Just download, run, scan."
- "If you want to know how this was built — keep watching."

---

## PART 2: For Developers — Building From Scratch (12–15 min)

> **Audience**: Python beginners to intermediate. Teaches os.walk, pandas,
> openpyxl, Streamlit, tkinter, and PyInstaller.

### Step 1: Project Setup (1 min)

**What to say:**
- "Let's build this from zero. Create a new folder and set up the project."

**What to show / type:**

```bash
mkdir Folder_Scanner
cd Folder_Scanner
```

Create `requirements.txt`:
```
streamlit
pandas
openpyxl
```

Install:
```bash
pip install -r requirements.txt
```

---

### Step 2: Core Logic — folder_scanner.py (4–5 min)

**What to say:**
- "The heart of this app is `os.walk()` — it recursively walks through
  every folder and file in a directory tree."
- "We'll classify each file by its extension, track the folder depth,
  and collect everything into a list."

**What to type (explain as you go):**

#### 2a. Imports and file type detection

```python
import os
import pandas as pd
from pathlib import Path
import io
import streamlit as st
```

- "We use `pathlib` to handle file extensions cleanly."

```python
def get_file_type(filename):
    ext = Path(filename).suffix.lower()
    type_map = {
        '.xlsx': 'Excel', '.xls': 'Excel', '.xlsm': 'Excel', '.xlsb': 'Excel',
        '.pdf': 'PDF',
        '.ppt': 'PowerPoint', '.pptx': 'PowerPoint', '.pptm': 'PowerPoint',
        '.doc': 'Word', '.docx': 'Word',
        '.txt': 'Text', '.csv': 'CSV', '.json': 'JSON', '.xml': 'XML',
        '.jpg': 'Image', '.jpeg': 'Image', '.png': 'Image', '.gif': 'Image',
        '.bmp': 'Image', '.svg': 'Image', '.webp': 'Image',
        '.mp4': 'Video', '.avi': 'Video', '.mov': 'Video', '.mkv': 'Video',
        '.mp3': 'Audio', '.wav': 'Audio', '.flac': 'Audio',
        '.zip': 'Archive', '.rar': 'Archive', '.7z': 'Archive',
        '.tar': 'Archive', '.gz': 'Archive',
        '.py': 'Code', '.js': 'Code', '.html': 'Code', '.css': 'Code',
        '.java': 'Code', '.cpp': 'Code', '.c': 'Code', '.ts': 'Code',
    }
    return type_map.get(ext, 'Folder' if ext == '' else 'Other')
```

- "Simple dictionary lookup — fast and easy to extend."

#### 2b. Directory scanning

```python
def scan_directory(root_path):
    items = []
    root_path = Path(root_path)

    for root, dirs, files in os.walk(root_path):
        current_path = Path(root)
        level = len(current_path.relative_to(root_path).parts)

        for dir_name in dirs:
            items.append({
                'Name': dir_name,
                'Type': 'Folder',
                'Level': level + 1,
                'Full Path': str(current_path / dir_name),
                'Parent Folder': current_path.name if level > 0 else 'Root'
            })

        for file_name in files:
            items.append({
                'Name': file_name,
                'Type': get_file_type(file_name),
                'Level': level + 1,
                'Full Path': str(current_path / file_name),
                'Parent Folder': current_path.name if level > 0 else 'Root'
            })

    return items
```

**Key points to explain:**
- `os.walk()` gives you (root, dirs, files) for every level
- `relative_to()` calculates the depth/level
- We store everything as a list of dicts — easy to convert to DataFrame

#### 2c. Excel report generation

```python
def create_excel_report(root_folder):
    if not os.path.exists(root_folder):
        raise ValueError(f"Folder '{root_folder}' does not exist.")

    items = scan_directory(root_folder)
    if not items:
        raise ValueError("The selected folder is empty.")

    items.sort(key=lambda x: (x['Level'], x['Name']))
    df = pd.DataFrame(items)[['Name', 'Type', 'Level', 'Parent Folder', 'Full Path']]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Folder Structure', index=False)
        worksheet = writer.sheets['Folder Structure']

        # Auto-size columns
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.freeze_panes = 'A2'

        # Format header row
        from openpyxl.styles import PatternFill, Font
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font

    buffer.seek(0)
    return buffer.getvalue(), len(items)
```

**Key points to explain:**
- `io.BytesIO()` creates the Excel in memory (no temp files)
- `openpyxl` gives us formatting: auto-width, filters, frozen header, colors
- Returns bytes + count — Streamlit can serve bytes directly as a download

---

### Step 3: Streamlit Web UI (3 min)

**What to say:**
- "Now let's add a web interface using Streamlit — just a few lines of Python."

**What to type (continue in same file):**

```python
st.set_page_config(page_title="Folder Scanner", page_icon="📁", layout="wide")
st.title("📁 Folder Scanner")
st.write("Scan any folder structure and download a detailed Excel report.")
```

- "Streamlit makes it dead simple — a Browse button opens a native
  folder picker, then one click to generate, and a download button
  for the Excel file."
- Build up the single-page UI step by step (Browse button, path display,
  generate button, results with metrics and preview, download)

**What to demo:**

```bash
streamlit run folder_scanner.py
```

- Browser opens at localhost:8501
- Click Browse → pick a folder from the native dialog
- Click Generate Report → see metrics (folders, files, file types)
- Expand the preview to see the data → click Download
- "That's the developer version — runs in your browser, full control."

---

### Step 4: Desktop GUI with tkinter (3 min)

**What to say:**
- "The web version is great for developers, but what about people who
  don't have Python? Let's build a desktop app."
- "We'll use tkinter — it comes built-in with Python, no extra install needed."

**What to show:**
- Create `desktop_app.py`
- Walk through the key parts:
  - `filedialog.askdirectory()` — native folder picker
  - `filedialog.asksaveasfilename()` — save dialog
  - `threading.Thread` — keeps UI responsive during scan
  - `os.startfile()` — auto-opens the report

**What to demo:**

```bash
python desktop_app.py
```

- The GUI window opens
- Browse → pick folder → Generate → Save → Excel opens
- "Same result, but now it's a standalone window app."

---

### Step 5: Package as .exe with PyInstaller (2 min)

**What to say:**
- "Final step — turn this into a single .exe file that anyone can run
  without installing Python."

**What to type:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "FolderScanner" desktop_app.py
```

**What to explain:**
- `--onefile` → everything bundled into one .exe
- `--windowed` → no console window pops up
- `--name` → the name of the output file
- Takes a few minutes to build

**What to show:**
- Build completes → open `dist/` folder → show `FolderScanner.exe`
- Double-click it → works exactly like before
- "This is what you share with people. One file, no setup."

---

## OUTRO (30 sec)

**What to say:**
- "So that's it — we built a folder scanner from scratch in Python,
  gave it both a web UI and a desktop UI, and packaged it as a
  standalone .exe."
- "If you're not a programmer, just download the .exe from the link
  in the description."
- "If you are — the full source code is on GitHub. Fork it, extend it,
  make it your own."
- "If this was helpful, like and subscribe. See you in the next one."

---

## Checklist Before Recording

- [ ] Have a sample folder ready with mixed file types (Excel, PDF, images, subfolders)
- [ ] Pre-built .exe is ready in `dist/` folder
- [ ] GitHub repo is up to date
- [ ] Create a GitHub Release with the .exe attached
- [ ] Test both the Streamlit app and .exe work end-to-end
- [ ] Have the terminal ready with a clean working directory
