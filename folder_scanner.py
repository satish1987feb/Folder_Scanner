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
st.title("Folder Scanner SaaS")
st.write("Scan a folder structure and download the report as an Excel file.")

root_folder = st.text_input("Enter the root folder path to scan (e.g., C:\\Users\\YourName\\Documents):")
output_name = st.text_input("Enter the output Excel file name:", "folder_structure.xlsx")

if st.button("Generate Report"):
    if root_folder:
        try:
            excel_bytes, total_items = create_excel_report(root_folder)
            st.success(f"Report generated successfully! Total items found: {total_items}")
            st.download_button(
                label="Download Excel Report",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel"
            )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    else:
        st.error("Please enter a valid root folder path.")