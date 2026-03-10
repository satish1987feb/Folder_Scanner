import os
import pandas as pd
from pathlib import Path
import io
import streamlit as st

def get_file_type(filename):
    """Determine file type based on extension"""
    ext = Path(filename).suffix.lower()
    
    if ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
        return 'Excel'
    elif ext in ['.pdf']:
        return 'PDF'
    elif ext in ['.ppt', '.pptx', '.pptm']:
        return 'PowerPoint'
    elif ext in ['.doc', '.docx']:
        return 'Word'
    elif ext in ['.txt']:
        return 'Text'
    elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
        return 'Image'
    elif ext in ['.mp4', '.avi', '.mov', '.mkv']:
        return 'Video'
    elif ext in ['.mp3', '.wav', '.flac']:
        return 'Audio'
    elif ext in ['.zip', '.rar', '.7z', '.tar', '.gz']:
        return 'Archive'
    elif ext == '':
        return 'Folder'
    else:
        return 'Other'

def scan_directory(root_path):
    """Scan directory and collect information about all items"""
    items = []
    root_path = Path(root_path)
    
    for root, dirs, files in os.walk(root_path):
        current_path = Path(root)
        level = len(current_path.relative_to(root_path).parts)
        
        # Add folders
        for dir_name in dirs:
            items.append({
                'Name': dir_name,
                'Type': 'Folder',
                'Level': level + 1,  # +1 because current level is for parent
                'Full Path': str(current_path / dir_name),
                'Parent Folder': current_path.name if level > 0 else 'Root'
            })
        
        # Add files
        for file_name in files:
            file_type = get_file_type(file_name)
            items.append({
                'Name': file_name,
                'Type': file_type,
                'Level': level + 1,
                'Full Path': str(current_path / file_name),
                'Parent Folder': current_path.name if level > 0 else 'Root'
            })
    
    return items

def create_excel_report(root_folder):
    """Create Excel report with folder/file structure and return as bytes"""
    
    # Check if root folder exists
    if not os.path.exists(root_folder):
        raise ValueError(f"Folder '{root_folder}' does not exist.")
    
    print(f"Scanning folder: {root_folder}")
    
    # Get all items
    items = scan_directory(root_folder)
    
    # Sort by level and name
    items.sort(key=lambda x: (x['Level'], x['Name']))
    
    # Create DataFrame
    df = pd.DataFrame(items)
    
    # Reorder columns
    df = df[['Name', 'Type', 'Level', 'Parent Folder', 'Full Path']]
    
    # Create Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Folder Structure', index=False)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Folder Structure']
        
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Add filters
        worksheet.auto_filter.ref = worksheet.dimensions
        
        # Freeze header row
        worksheet.freeze_panes = 'A2'
        
        # Add color coding for different levels (optional)
        from openpyxl.styles import PatternFill, Font
        
        # Header formatting
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
    
    buffer.seek(0)
    print(f"Excel report created successfully in memory")
    print(f"Total items found: {len(items)}")
    return buffer.getvalue(), len(items)

# Streamlit App
st.set_page_config(page_title="Folder Scanner", layout="wide")
st.title("📁 Folder Scanner SaaS")
st.write("Scan folder structures and download reports as Excel files.")

# Create tabs
tab1, tab2 = st.tabs(["📍 Local Use", "ℹ️ Information"])

with tab1:
    st.subheader("Scan Your Local Folder")
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.info("⚠️ **Important**: This deployed version works best when run locally on your machine. If you're on Streamlit Cloud, please run this app locally to access your folders.")
        
        root_folder = st.text_input(
            "Enter the root folder path to scan",
            placeholder="e.g., C:\\Users\\YourName\\Documents or D:\\KI\\02. Client Projects",
            help="Use forward slashes (/) or backslashes (\\). Backslashes need to be properly escaped."
        )
        
        output_name = st.text_input(
            "Output Excel file name",
            value="folder_structure.xlsx",
            help="Give your report a meaningful name"
        )
    
    with col2:
        st.write("**Path Formats:**")
        st.code("Windows:\nD:/KI/02. Client Projects\nor\nD:\\\\KI\\\\02. Client Projects\n\nMac/Linux:\n/Users/YourName/Documents")
    
    if st.button("🔍 Generate Report", use_container_width=True):
        if root_folder:
            with st.spinner("Scanning folder..."):
                try:
                    # Normalize path
                    normalized_path = root_folder.replace('\\', '/')
                    excel_bytes, total_items = create_excel_report(normalized_path)
                    
                    st.success(f"✅ Report generated successfully!")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Items", total_items)
                    
                    st.download_button(
                        label="⬇️ Download Excel Report",
                        data=excel_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel",
                        use_container_width=True
                    )
                except ValueError as e:
                    st.error(f"❌ Folder Not Found: {str(e)}")
                    st.info("💡 Make sure the folder path is correct and accessible from your computer.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    st.info("💡 Try using forward slashes (/) in the path instead of backslashes.")
        else:
            st.warning("Please enter a valid folder path.")

with tab2:
    st.subheader("How to Use")
    st.markdown("""
    ### Deployment Options
    
    **Option 1: Run Locally (Recommended)**
    - Install Python 3.7+
    - Run: `pip install -r requirements.txt`
    - Run: `streamlit run folder_scanner.py`
    - Access at http://localhost:8501
    - ✅ Can access local folders and network drives
    
    **Option 2: Streamlit Cloud (Current)**
    - Visit the GitHub repo and deploy independently
    - ⚠️ Cannot access your local file system
    - Best for testing the interface
    
    ### Features
    - 📊 Scan any folder structure recursively
    - 🏷️ Automatic file type classification
    - 📈 Detailed Excel reports with formatting
    - 🎯 Organized by folder level and parent folder
    - ⚡ Fast processing for large directories
    
    ### GitHub Repository
    - **Repository**: https://github.com/satish1987feb/Folder_Scanner
    - Clone and run locally for full functionality
    """)
    
    st.subheader("File Type Categories")
    categories = {
        "Excel": ".xlsx, .xls, .xlsm, .xlsb",
        "PDF": ".pdf",
        "Word": ".doc, .docx",
        "PowerPoint": ".ppt, .pptx, .pptm",
        "Images": ".jpg, .jpeg, .png, .gif, .bmp",
        "Videos": ".mp4, .avi, .mov, .mkv",
        "Audio": ".mp3, .wav, .flac",
        "Archives": ".zip, .rar, .7z, .tar, .gz",
    }
    
    col1, col2 = st.columns(2)
    for idx, (category, exts) in enumerate(categories.items()):
        if idx % 2 == 0:
            with col1:
                st.text(f"**{category}**: {exts}")
        else:
            with col2:
                st.text(f"**{category}**: {exts}")