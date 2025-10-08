import pandas as pd
import pdfplumber
from datetime import datetime
import os

def process_single_pdf(pdf_file):
    """
    Process a single PDF file and extract technical report data.
    """
    try:
        # Initialize data dictionary with None values
        data = {
            'TECHNICAL REPORT': None,
            'Reported Date': None,
            'TEST TYPE': None,
            'PRODUCT CODE': None,
            'FINISH NAME/LOOK NAME': None,
            'SEASON': None,
            'FFC (5 DIGIT)/LOOK CODE': None,
            'END USE': None,
            'MILL STYLE / SAMPLE STYLE': None,
            'FABRIC TYPE': None,
            'FIBER CONTENT': None,
            'GARMENT FINISH DETAIL': None,
        }
        
        # Dictionary to track if we've already found a value for each field
        found_fields = {key: False for key in data.keys()}

        with pdfplumber.open(pdf_file) as pdf:
            text = ''
            tables = []
            
            # Extract both text and tables from each page
            for page in pdf.pages:
                text += page.extract_text() + '\n'
                tables.extend(page.extract_tables())
            
            # Split text into lines
            lines = text.split('\n')
            
            # First try to find TECHNICAL REPORT number from the first few lines
            for i, line in enumerate(lines[:10]):  # Check first 10 lines
                if 'TECHNICAL REPORT' in line:
                    # Try to extract the report number
                    parts = line.split('TECHNICAL REPORT')
                    if len(parts) > 1:
                        report_num = parts[1].strip().strip('#').strip()
                        if report_num:
                            data['TECHNICAL REPORT'] = report_num
                    # If no number after, check next line
                    elif i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line and next_line.startswith('#'):
                            data['TECHNICAL REPORT'] = next_line.strip('#').strip()
                    break

            # Process tables to build a mapping of fields and values
            field_values = {}
            for table in tables:
                for row_idx, row in enumerate(table):
                    if not row:
                        continue
                    # Clean row values
                    row = [str(cell).strip() if cell else '' for cell in row]
                    
                    # Check for field matches
                    for field in data.keys():
                        for col_idx, cell in enumerate(row):
                            if field in cell:
                                # Look for value in the next row, same column
                                if row_idx + 1 < len(table):
                                    next_row = table[row_idx + 1]
                                    if col_idx < len(next_row) and next_row[col_idx]:
                                        value = str(next_row[col_idx]).strip()
                                        
                                        # Special handling for GARMENT FINISH DETAIL
                                        if field == 'GARMENT FINISH DETAIL':
                                            # Collect all non-empty lines until we hit another field or empty line
                                            finish_details = [value]
                                            next_idx = row_idx + 2  # Start from two rows down
                                            while next_idx < len(table):
                                                next_row_value = table[next_idx][col_idx] if col_idx < len(table[next_idx]) else ''
                                                next_row_value = str(next_row_value).strip()
                                                if not next_row_value or any(key in next_row_value for key in data.keys()):
                                                    break
                                                finish_details.append(next_row_value)
                                                next_idx += 1
                                            value = '\n'.join(finish_details)
                                        else:
                                            # For other fields, clean the value
                                            value = value.split('\n')[0] if '\n' in value else value
                                            for other_field in data.keys():
                                                if other_field in value and other_field != field:
                                                    value = value.replace(other_field, '').strip()
                                        
                                        if value and value != 'None':
                                            field_values[field] = value

            # Special handling for Reported Date
            for i, line in enumerate(lines):
                if 'Reported Date' in line:
                    if ':' in line:
                        date_value = line.split('Date:')[-1].strip()
                        if date_value:
                            field_values['Reported Date'] = date_value
                    elif i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line and not any(key in next_line for key in data.keys()):
                            field_values['Reported Date'] = next_line
                    break

            # Update data with cleaned values
            for field, value in field_values.items():
                if value:
                    # Additional cleaning for specific fields
                    if field == 'TEST TYPE' and 'Deni' in value:
                        value = value.split('Deni')[0].strip()
                    data[field] = value

            # Backup: Process tables for any missing values
            for table in tables:
                for row_idx, row in enumerate(table):
                    if not row:
                        continue
                        
                    row = [str(cell).strip() if cell else '' for cell in row]
                    for col_idx, cell in enumerate(row):
                        for field in data.keys():
                            if field in cell and not data[field]:
                                # Try to get value from next row same column
                                if row_idx + 1 < len(table):
                                    next_row = table[row_idx + 1]
                                    if len(next_row) > col_idx and next_row[col_idx]:
                                        value = str(next_row[col_idx]).strip()
                                        if value and value != 'None':
                                            data[field] = value
                            
            # Special handling for specific fields
            if not data['PRODUCT CODE']:
                for line in lines:
                    if 'PRODUCT CODE' in line.upper():
                        parts = line.split()
                        if len(parts) > 2:
                            data['PRODUCT CODE'] = parts[-1].strip()

        return data
        
    except Exception as e:
        print(f"Error processing PDF {pdf_file}: {str(e)}")
        return None

def extract_technical_report_data(pdf_files, save_excel=True):
    """
    Extract technical report data from multiple PDF files.
    Args:
        pdf_files: List of PDF file paths or single PDF file path
        save_excel: If True, saves to a new Excel file. If False, only returns the data
    Returns:
        tuple: (excel_path or None, list of extracted data dictionaries)
    """
    # Convert single file to list if string is provided
    if isinstance(pdf_files, str):
        pdf_files = [pdf_files]
    
    # List to store all extracted data
    all_data = []
    errors = []
    
    # Process each PDF file
    for pdf_file in pdf_files:
        try:
            data = process_single_pdf(pdf_file)
            if data:
                all_data.append(data)
        except Exception as e:
            errors.append(f"Error processing {pdf_file}: {str(e)}")
    
    if not all_data:
        raise Exception("No data could be extracted from any of the PDF files")
    
            # Print any errors that occurred
    if errors:
        print("\nWarnings:")
        for error in errors:
            print(error)
    
    # Only save to Excel if save_excel is True
    if save_excel:
        # Convert to DataFrame for saving
        df = pd.DataFrame(all_data)
        
        # Save to Excel with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.dirname(pdf_files[0])  # Use first file's directory
        excel_path = os.path.join(output_dir, f"technical_reports_combined_{timestamp}.xlsx")
        df.to_excel(excel_path, index=False)
        return excel_path, all_data    # If save_excel is False, just return the data
    return None, all_data

