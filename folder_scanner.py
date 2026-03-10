import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pathlib import Path
import io
import streamlit as st


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

        from openpyxl.styles import PatternFill, Font
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font

    buffer.seek(0)
    return buffer.getvalue(), len(items), df


def open_folder_dialog():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder = filedialog.askdirectory(master=root, title="Select a folder to scan")
    root.destroy()
    return folder


# ─── Streamlit App ───────────────────────────────────────────────────────────

st.set_page_config(page_title="Folder Scanner", page_icon="📁", layout="centered")

st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 720px; }

    .app-header { text-align: center; margin-bottom: 2rem; }
    .app-header h1 { font-size: 2.2rem; margin-bottom: 0.2rem; }
    .app-header p { opacity: 0.6; font-size: 1.05rem; }

    .folder-path-box {
        background: rgba(128,128,128,0.08); border: 1px solid rgba(128,128,128,0.2);
        border-radius: 8px; padding: 0.75rem 1rem; font-size: 0.95rem;
        margin: 0.5rem 0 1rem 0; word-break: break-all; min-height: 1.2rem;
    }

    .result-card {
        background: linear-gradient(135deg, #1b5e20, #2e7d32);
        border-radius: 12px; padding: 1.5rem; text-align: center; margin: 1rem 0;
    }
    .result-card h2 { margin: 0; font-size: 2.4rem; color: #fff; font-weight: 700; }
    .result-card p { margin: 0.3rem 0 0 0; color: rgba(255,255,255,0.8); font-size: 0.95rem; }

    div[data-testid="stMetric"] {
        background: rgba(128,128,128,0.08);
        border: 1px solid rgba(128,128,128,0.15);
        border-radius: 10px; padding: 1rem 0.8rem; text-align: center;
    }
    div[data-testid="stMetric"] label {
        font-size: 0.8rem !important; text-transform: uppercase;
        letter-spacing: 0.05em; opacity: 0.6;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 1.8rem !important; font-weight: 700 !important;
    }

    section[data-testid="stSidebar"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
    <h1>📁 Folder Scanner</h1>
    <p>Select a folder, generate a formatted Excel report.</p>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────

if "folder_path" not in st.session_state:
    st.session_state.folder_path = ""
if "scan_result" not in st.session_state:
    st.session_state.scan_result = None

# ── Folder selection ──────────────────────────────────────────────────────────

st.markdown("##### Select Folder")

col_btn, col_name = st.columns([1, 2])

with col_btn:
    if st.button("📂 Browse Folder", use_container_width=True):
        folder = open_folder_dialog()
        if folder:
            st.session_state.folder_path = folder
            st.session_state.scan_result = None

with col_name:
    output_name = st.text_input(
        "Report file name",
        value="folder_structure.xlsx",
        label_visibility="collapsed",
        placeholder="Report file name (e.g. folder_structure.xlsx)",
    )

selected = st.session_state.folder_path
if selected:
    st.markdown(f'<div class="folder-path-box">📁 {selected}</div>', unsafe_allow_html=True)
else:
    st.markdown(
        '<div class="folder-path-box" style="color:#999;">No folder selected — click Browse to pick one</div>',
        unsafe_allow_html=True,
    )

# ── Generate ──────────────────────────────────────────────────────────────────

st.markdown("")  # spacer

if st.button("📊 Generate Report", use_container_width=True, type="primary"):
    if not selected:
        st.warning("Please select a folder first.")
    else:
        with st.spinner("Scanning..."):
            try:
                normalized = selected.strip().replace('\\', '/')
                excel_bytes, total_items, df = create_excel_report(normalized)
                st.session_state.scan_result = {
                    "bytes": excel_bytes,
                    "count": total_items,
                    "df": df,
                    "folder": Path(selected).name,
                }
            except ValueError as e:
                st.error(f"**Error:** {e}")
            except Exception as e:
                st.error(f"**Unexpected error:** {e}")

# ── Results ───────────────────────────────────────────────────────────────────

result = st.session_state.scan_result
if result:
    st.markdown("---")

    st.markdown(
        f'<div class="result-card">'
        f'<h2>{result["count"]}</h2>'
        f'<p>items found in <strong>{result["folder"]}</strong></p>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Summary metrics
    type_counts = result["df"]["Type"].value_counts()
    folders = type_counts.get("Folder", 0)
    files = result["count"] - folders

    c1, c2, c3 = st.columns(3)
    c1.metric("Folders", f"{folders:,}")
    c2.metric("Files", f"{files:,}")
    c3.metric("File Types", f"{len(type_counts):,}")

    # Preview
    with st.expander("Preview report", expanded=False):
        st.dataframe(result["df"].head(50), use_container_width=True, hide_index=True)

    # Download
    st.download_button(
        label="⬇️  Download Excel Report",
        data=result["bytes"],
        file_name=output_name if output_name.endswith(".xlsx") else output_name + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Footer ────────────────────────────────────────────────────────────────────

st.markdown("---")
st.markdown(
    '<p style="text-align:center; color:#999; font-size:0.82rem;">'
    'Folder Scanner &nbsp;·&nbsp; '
    '<a href="https://github.com/satish1987feb/Folder_Scanner" style="color:#999;">GitHub</a>'
    '</p>',
    unsafe_allow_html=True,
)
