import os
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
    return buffer.getvalue(), len(items)


# ─── Streamlit App ───────────────────────────────────────────────────────────

st.set_page_config(page_title="Folder Scanner", page_icon="📁", layout="wide")

st.title("📁 Folder Scanner")
st.write("Scan any folder structure and download a detailed Excel report.")

tab_scan, tab_info = st.tabs(["📊 Scan Folder", "ℹ️ About"])

# ── Tab 1: Scan ──────────────────────────────────────────────────────────────

with tab_scan:
    col_input, col_help = st.columns([2, 1])

    with col_input:
        root_folder = st.text_input(
            "Folder path to scan",
            placeholder="e.g., C:\\Users\\YourName\\Documents",
            help="Paste the full folder path from File Explorer."
        )
        output_name = st.text_input(
            "Report file name",
            value="folder_structure.xlsx",
        )

    with col_help:
        st.markdown("**How to copy a folder path:**")
        st.markdown(
            "1. Open the folder in **File Explorer**\n"
            "2. Click the **address bar** at the top\n"
            "3. Press **Ctrl+C** to copy\n"
            "4. Paste it here"
        )

    if st.button("📊 Generate Report", use_container_width=True):
        if root_folder:
            with st.spinner("Scanning folder..."):
                try:
                    normalized_path = root_folder.strip().replace('\\', '/')
                    excel_bytes, total_items = create_excel_report(normalized_path)

                    st.success(f"Scan complete! Found **{total_items}** items.")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Items", total_items)

                    st.download_button(
                        label="⬇️ Download Excel Report",
                        data=excel_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except ValueError as e:
                    st.error(f"Folder not found: {e}")
                    st.info("Make sure the folder path is correct and accessible.")
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.info("Try using forward slashes (/) in the path.")
        else:
            st.warning("Please enter a folder path.")

# ── Tab 2: About ─────────────────────────────────────────────────────────────

with tab_info:
    st.markdown("""
### What This App Does

Scans any folder on your computer recursively and generates a formatted Excel
report listing every file and subfolder with its type, depth level, and location.

### What's in the Report?

| Column | Description |
|--------|-------------|
| **Name** | File or folder name |
| **Type** | Auto-detected category (Excel, PDF, Image, etc.) |
| **Level** | Depth in the folder tree |
| **Parent Folder** | The containing folder |
| **Full Path** | Complete path to the item |

### Supported File Types
    """)

    categories = {
        "Excel": ".xlsx, .xls, .xlsm, .xlsb",
        "PDF": ".pdf",
        "Word": ".doc, .docx",
        "PowerPoint": ".ppt, .pptx, .pptm",
        "Images": ".jpg, .jpeg, .png, .gif, .bmp, .svg, .webp",
        "Videos": ".mp4, .avi, .mov, .mkv",
        "Audio": ".mp3, .wav, .flac",
        "Archives": ".zip, .rar, .7z, .tar, .gz",
        "Code": ".py, .js, .html, .css, .java, .cpp, .c, .ts",
        "Data": ".csv, .json, .xml",
    }

    col1, col2 = st.columns(2)
    for idx, (category, exts) in enumerate(categories.items()):
        target = col1 if idx % 2 == 0 else col2
        with target:
            st.markdown(f"**{category}**: {exts}")

    st.markdown("---")
    st.markdown(
        "**GitHub**: [github.com/satish1987feb/Folder_Scanner]"
        "(https://github.com/satish1987feb/Folder_Scanner)"
    )
