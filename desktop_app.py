"""
Folder Scanner — Desktop Edition
Double-click to run. Pick a folder, get an Excel report.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import io
import pandas as pd


# ─── Core scanning logic (shared with the web version) ───────────────────────

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


def create_excel_report(root_folder, output_path):
    """Scan a folder and save an Excel report to disk."""
    items = scan_directory(root_folder)
    if not items:
        raise ValueError("The selected folder is empty.")

    items.sort(key=lambda x: (x['Level'], x['Name']))
    df = pd.DataFrame(items)[['Name', 'Type', 'Level', 'Parent Folder', 'Full Path']]

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Folder Structure', index=False)
        worksheet = writer.sheets['Folder Structure']

        for column in worksheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.freeze_panes = 'A2'

        from openpyxl.styles import PatternFill, Font
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font

    return len(items)


# ─── GUI ──────────────────────────────────────────────────────────────────────

class FolderScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Scanner")
        self.root.geometry("620x420")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 10))

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=24)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="📁  Folder Scanner", style="Header.TLabel").pack(pady=(0, 4))
        ttk.Label(main, text="Select a folder and generate a detailed Excel report.").pack(pady=(0, 20))

        # Folder selection
        folder_frame = ttk.LabelFrame(main, text="  Folder to Scan  ", padding=12)
        folder_frame.pack(fill="x", pady=(0, 12))

        path_row = ttk.Frame(folder_frame)
        path_row.pack(fill="x")

        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(path_row, textvariable=self.folder_var, font=("Segoe UI", 10))
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(path_row, text="Browse...", command=self._browse_folder).pack(side="right")

        # Output file name
        output_frame = ttk.LabelFrame(main, text="  Report File Name  ", padding=12)
        output_frame.pack(fill="x", pady=(0, 16))

        self.output_var = tk.StringVar(value="folder_structure.xlsx")
        ttk.Entry(output_frame, textvariable=self.output_var, font=("Segoe UI", 10)).pack(fill="x")

        # Generate button
        self.generate_btn = ttk.Button(
            main, text="Generate Report", style="Accent.TButton", command=self._generate
        )
        self.generate_btn.pack(fill="x", ipady=6, pady=(0, 12))

        # Progress bar
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))

        # Status label
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, style="Status.TLabel").pack()

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select a folder to scan")
        if folder:
            self.folder_var.set(folder)

    def _generate(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("No folder selected", "Please select a folder first.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Invalid folder", f"'{folder}' is not a valid folder.")
            return

        output_name = self.output_var.get().strip()
        if not output_name.endswith('.xlsx'):
            output_name += '.xlsx'

        save_path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".xlsx",
            initialfile=output_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not save_path:
            return

        self.generate_btn.config(state="disabled")
        self.progress.start(15)
        self.status_var.set("Scanning...")

        def worker():
            try:
                count = create_excel_report(folder, save_path)
                self.root.after(0, lambda: self._on_success(count, save_path))
            except Exception as exc:
                self.root.after(0, lambda: self._on_error(str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, count, path):
        self.progress.stop()
        self.generate_btn.config(state="normal")
        self.status_var.set(f"Done — {count} items saved to {Path(path).name}")
        open_it = messagebox.askyesno(
            "Report Generated",
            f"Found {count} items.\nSaved to:\n{path}\n\nOpen the file now?"
        )
        if open_it:
            os.startfile(path)

    def _on_error(self, msg):
        self.progress.stop()
        self.generate_btn.config(state="normal")
        self.status_var.set("Error")
        messagebox.showerror("Error", msg)


if __name__ == "__main__":
    app_root = tk.Tk()
    FolderScannerApp(app_root)
    app_root.mainloop()
