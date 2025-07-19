import streamlit as st
from chatbot_api import get_chat_response
import os
import PyPDF2
import docx
from docx import Document
import json
import csv
import pandas as pd
import numpy as np
from io import StringIO, BytesIO
import zipfile
from datetime import datetime
import base64
from PIL import Image
import xml.etree.ElementTree as ET
import yaml
import markdown
import html2text

# Advanced PDF processing libraries
from pdf2docx import Converter
import pdfplumber
import fitz  # PyMuPDF

# Additional libraries for professional conversions
import xlsxwriter
from bs4 import BeautifulSoup
try:
    import camelot
except ImportError:
    camelot = None
try:
    import tabula
except ImportError:
    tabula = None

st.set_page_config(page_title="rAI bot", layout="centered")

# Initialize session state for ChatGPT-like conversation memory
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

# Multi-document support (background collection)
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = {}  # Dictionary to store multiple files {filename: file_object}

if "active_file" not in st.session_state:
    st.session_state.active_file = None  # Currently selected file for conversation

if "show_switcher" not in st.session_state:
    st.session_state.show_switcher = False  # Document switcher visibility

if "conversation_context" not in st.session_state:
    # This will help maintain conversation context across interactions
    st.session_state.conversation_context = []

# Function to extract file content
def extract_file_content(uploaded_file):
    """Extract text content from uploaded files"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        if file_extension == '.txt':
            return uploaded_file.read().decode('utf-8')
        
        elif file_extension == '.pdf':
            # Use advanced PDF extraction for better text quality
            try:
                # Try pdfplumber first for better table and structure extraction
                with pdfplumber.open(uploaded_file) as pdf:
                    text = ""
                    for page_num, page in enumerate(pdf.pages, 1):
                        page_text = f"--- Page {page_num} ---\n\n"
                        
                        # Extract tables first
                        tables = page.extract_tables()
                        if tables:
                            for table in tables:
                                if table:
                                    page_text += "TABLE:\n"
                                    for row in table:
                                        if row:
                                            page_text += " | ".join(str(cell) if cell else "" for cell in row) + "\n"
                                    page_text += "\n"
                        
                        # Extract regular text
                        regular_text = page.extract_text()
                        if regular_text:
                            page_text += regular_text
                        
                        text += page_text + "\n\n"
                    
                    return text.strip()
            except Exception:
                # Fallback to PyMuPDF
                try:
                    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                    text = ""
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        page_text = f"--- Page {page_num + 1} ---\n\n"
                        page_text += page.get_text()
                        text += page_text + "\n\n"
                    doc.close()
                    return text.strip()
                except Exception:
                    # Final fallback to PyPDF2
                    uploaded_file.seek(0)  # Reset file pointer
                    pdf_reader = PyPDF2.PdfReader(uploaded_file)
                    text = ""
                    for page in pdf_reader.pages:
                        page_text = page.extract_text()
                        page_text = page_text.replace('\n', ' ')
                        page_text = ' '.join(page_text.split())
                        text += page_text + "\n\n"
                    return text.strip()
        
        elif file_extension == '.docx':
            doc = docx.Document(uploaded_file)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text
        
        elif file_extension == '.csv':
            csv_content = uploaded_file.read().decode('utf-8')
            csv_reader = csv.reader(StringIO(csv_content))
            rows = list(csv_reader)
            # Return first 20 rows as preview
            preview = "\n".join([",".join(row) for row in rows[:20]])
            if len(rows) > 20:
                preview += f"\n... and {len(rows)-20} more rows"
            return preview
        
        elif file_extension == '.json':
            json_content = uploaded_file.read().decode('utf-8')
            data = json.loads(json_content)
            return json.dumps(data, indent=2)
        
        elif file_extension in ['.py', '.js', '.html', '.css', '.md', '.yaml', '.yml', '.xml']:
            return uploaded_file.read().decode('utf-8')
        
        else:
            return f"File type {file_extension} content extraction not supported yet."
            
    except Exception as e:
        return f"Error reading file: {str(e)}"

# Safe function to extract file content from stored file data
def get_file_content_safe(file_key):
    """Safely extract content from stored file data without consuming streams"""
    if file_key not in st.session_state.uploaded_files:
        return f"Document '{file_key}' not found in uploaded files."
    
    file_data = st.session_state.uploaded_files[file_key]
    
    # Check if content is already extracted
    if file_data.get('file_content') is not None:
        return file_data['file_content']
    
    # Extract content from stored bytes
    try:
        from io import BytesIO
        
        # Create a new file-like object from stored bytes
        file_bytes = file_data['file_bytes']
        file_name = file_data['file_name']
        file_extension = os.path.splitext(file_name)[1].lower()
        
        # Debug info
        # print(f"[DEBUG] Extracting content for: {file_name} ({len(file_bytes)} bytes)")
        
        if file_extension == '.txt':
            content = file_bytes.decode('utf-8')
        
        elif file_extension == '.pdf':
            # Use pdfplumber with BytesIO
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                text = ""
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = f"--- Page {page_num} ---\n\n"
                    
                    # Extract tables first
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            page_text += "TABLE:\n"
                            for row in table:
                                if row and any(cell for cell in row if cell):  # Skip empty rows
                                    page_text += " | ".join(str(cell) if cell else "" for cell in row) + "\n"
                            page_text += "\n"
                    
                    # Extract regular text
                    regular_text = page.extract_text()
                    if regular_text:
                        page_text += regular_text.strip()
                    
                    text += page_text + "\n\n"
                
                content = text.strip()
        
        elif file_extension == '.docx':
            # Use python-docx with BytesIO
            try:
                from docx import Document
                doc = Document(BytesIO(file_bytes))
                content = ""
                for para in doc.paragraphs:
                    content += para.text + "\n"
            except:
                content = "Error reading DOCX file"
        
        elif file_extension in ['.xlsx', '.xls']:
            import pandas as pd
            try:
                df = pd.read_excel(BytesIO(file_bytes))
                content = df.to_string(index=False)
            except:
                content = "Error reading Excel file"
        
        elif file_extension == '.csv':
            content = file_bytes.decode('utf-8')
        
        elif file_extension == '.json':
            json_content = file_bytes.decode('utf-8')
            data = json.loads(json_content)
            content = json.dumps(data, indent=2)
        
        elif file_extension in ['.py', '.js', '.html', '.css', '.md', '.yaml', '.yml', '.xml']:
            content = file_bytes.decode('utf-8')
        
        else:
            content = f"File type {file_extension} content extraction not supported yet."
        
        # Store extracted content to avoid re-extraction
        file_data['file_content'] = content
        # print(f"[DEBUG] Successfully extracted {len(content)} characters from {file_name}")
        return content
        
    except Exception as e:
        error_msg = f"Error extracting content from {file_name}: {str(e)}"
        # print(f"[DEBUG] {error_msg}")
        return error_msg

# File conversion functions
def convert_pdf_to_word(uploaded_file):
    """Convert PDF to Word document with advanced formatting preservation (like iLovePDF)"""
    try:
        # Save uploaded file temporarily
        temp_pdf_path = f"/tmp/{uploaded_file.name}"
        temp_docx_path = f"/tmp/{os.path.splitext(uploaded_file.name)[0]}.docx"
        
        # Write uploaded file to temp location
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Method 1: Try pdf2docx for best layout preservation
        try:
            cv = Converter(temp_pdf_path)
            cv.convert(temp_docx_path, start=0, end=None)
            cv.close()
            
            # Read the converted file
            with open(temp_docx_path, "rb") as f:
                buffer = BytesIO(f.read())
            
            # Clean up temp files
            os.remove(temp_pdf_path)
            os.remove(temp_docx_path)
            
            return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.docx"
            
        except Exception as e1:
            print(f"pdf2docx failed: {e1}, trying alternative method...")
            
            # Method 2: Use PyMuPDF for better text extraction with formatting
            try:
                doc = fitz.open(temp_pdf_path)
                word_doc = Document()
                
                # Add metadata
                word_doc.add_heading('Converted from PDF', 0)
                word_doc.add_paragraph(f'Original file: {uploaded_file.name}')
                word_doc.add_paragraph(f'Converted on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
                word_doc.add_paragraph('')
                
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    
                    # Add page header
                    word_doc.add_heading(f'Page {page_num + 1}', level=1)
                    
                    # Extract text blocks with formatting
                    blocks = page.get_text("dict")
                    
                    for block in blocks["blocks"]:
                        if "lines" in block:
                            for line in block["lines"]:
                                line_text = ""
                                for span in line["spans"]:
                                    line_text += span["text"]
                                
                                if line_text.strip():
                                    # Try to preserve some formatting
                                    p = word_doc.add_paragraph()
                                    
                                    # Check if it looks like a header (short, all caps, etc.)
                                    if len(line_text.strip()) < 50 and (line_text.isupper() or line_text.strip().endswith(':')):
                                        p.add_run(line_text.strip()).bold = True
                                    else:
                                        p.add_run(line_text.strip())
                    
                    # Add page break except for last page
                    if page_num < len(doc) - 1:
                        word_doc.add_page_break()
                
                doc.close()
                
                # Save to buffer
                buffer = BytesIO()
                word_doc.save(buffer)
                buffer.seek(0)
                
                # Clean up temp file
                os.remove(temp_pdf_path)
                
                return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.docx"
                
            except Exception as e2:
                print(f"PyMuPDF failed: {e2}, using fallback method...")
                
                # Method 3: Fallback to basic extraction with better structure
                try:
                    with pdfplumber.open(temp_pdf_path) as pdf:
                        word_doc = Document()
                        
                        # Add metadata
                        word_doc.add_heading('Converted from PDF', 0)
                        word_doc.add_paragraph(f'Original file: {uploaded_file.name}')
                        word_doc.add_paragraph(f'Converted on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
                        word_doc.add_paragraph('')
                        
                        for page_num, page in enumerate(pdf.pages):
                            # Add page header
                            word_doc.add_heading(f'Page {page_num + 1}', level=1)
                            
                            # Extract tables first
                            tables = page.extract_tables()
                            if tables:
                                for table in tables:
                                    # Add table to Word document
                                    if table and len(table) > 0:
                                        word_table = word_doc.add_table(rows=len(table), cols=len(table[0]))
                                        word_table.style = 'Table Grid'
                                        
                                        for i, row in enumerate(table):
                                            for j, cell in enumerate(row):
                                                if cell:
                                                    word_table.cell(i, j).text = str(cell)
                                
                                word_doc.add_paragraph('')
                            
                            # Extract text
                            text = page.extract_text()
                            if text:
                                # Split into paragraphs and preserve structure
                                paragraphs = text.split('\n')
                                current_para = ""
                                
                                for line in paragraphs:
                                    line = line.strip()
                                    if not line:
                                        if current_para:
                                            # Check formatting
                                            p = word_doc.add_paragraph()
                                            if len(current_para) < 50 and (current_para.isupper() or current_para.endswith(':')):
                                                p.add_run(current_para).bold = True
                                            else:
                                                p.add_run(current_para)
                                            current_para = ""
                                    else:
                                        if current_para:
                                            current_para += " " + line
                                        else:
                                            current_para = line
                                
                                # Add remaining paragraph
                                if current_para:
                                    p = word_doc.add_paragraph()
                                    if len(current_para) < 50 and (current_para.isupper() or current_para.endswith(':')):
                                        p.add_run(current_para).bold = True
                                    else:
                                        p.add_run(current_para)
                            
                            # Add page break except for last page
                            if page_num < len(pdf.pages) - 1:
                                word_doc.add_page_break()
                    
                    # Save to buffer
                    buffer = BytesIO()
                    word_doc.save(buffer)
                    buffer.seek(0)
                    
                    # Clean up temp file
                    os.remove(temp_pdf_path)
                    
                    return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.docx"
                    
                except Exception as e3:
                    # Clean up temp file
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)
                    return None, f"All conversion methods failed: {e1}, {e2}, {e3}"
        
    except Exception as e:
        return None, f"Error in PDF to Word conversion: {str(e)}"

def convert_to_pdf(uploaded_file, content):
    """Convert various formats to PDF"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.utils import simpleSplit
        
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter
        
        # Title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, f"Converted from {uploaded_file.name}")
        
        # Content
        c.setFont("Helvetica", 12)
        y_position = height - 100
        line_height = 14
        
        # Split content into lines that fit the page width
        lines = content.split('\n')
        for line in lines:
            if y_position < 50:  # Start new page
                c.showPage()
                y_position = height - 50
                c.setFont("Helvetica", 12)
            
            # Wrap long lines
            if len(line) > 80:
                wrapped_lines = simpleSplit(line, "Helvetica", 12, width - 100)
                for wrapped_line in wrapped_lines:
                    if y_position < 50:
                        c.showPage()
                        y_position = height - 50
                    c.drawString(50, y_position, wrapped_line)
                    y_position -= line_height
            else:
                c.drawString(50, y_position, line)
                y_position -= line_height
        
        c.save()
        buffer.seek(0)
        
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.pdf"
        
    except Exception as e:
        return None, f"Error converting to PDF: {str(e)}"

def convert_to_csv(uploaded_file, content):
    """Convert various formats to CSV"""
    try:
        buffer = StringIO()
        
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        if file_extension == '.json':
            # Convert JSON to CSV
            data = json.loads(content)
            if isinstance(data, list) and len(data) > 0:
                # If it's a list of objects, convert to CSV
                if isinstance(data[0], dict):
                    df = pd.DataFrame(data)
                    df.to_csv(buffer, index=False)
                else:
                    # Simple list
                    pd.DataFrame({'values': data}).to_csv(buffer, index=False)
            elif isinstance(data, dict):
                # Convert dict to CSV with key-value pairs
                df = pd.DataFrame(list(data.items()), columns=['Key', 'Value'])
                df.to_csv(buffer, index=False)
            else:
                return None, "JSON format not suitable for CSV conversion"
        
        elif file_extension in ['.txt', '.pdf', '.docx']:
            # Convert text content to CSV (split by lines)
            lines = content.split('\n')
            # Try to detect if it's structured data
            if any(',' in line or '\t' in line for line in lines[:5]):
                # Looks like it might already be CSV-like
                buffer.write(content)
            else:
                # Create simple CSV with line numbers
                df = pd.DataFrame({'Line_Number': range(1, len(lines)+1), 'Content': lines})
                df.to_csv(buffer, index=False)
        
        buffer.seek(0)
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.csv"
        
    except Exception as e:
        return None, f"Error converting to CSV: {str(e)}"

def convert_to_excel(uploaded_file, content):
    """Convert to Excel format with professional formatting and advanced features"""
    try:
        buffer = BytesIO()
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        # Method 1: Try advanced Excel creation with xlsxwriter
        try:
            workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})
            
            # Create formats for professional styling
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'
            })
            
            number_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'num_format': '#,##0.00'
            })
            
            if file_extension == '.csv':
                # Advanced CSV to Excel conversion
                df = pd.read_csv(StringIO(content))
                worksheet = workbook.add_worksheet('Data')
                
                # Write headers with formatting
                for col, column_name in enumerate(df.columns):
                    worksheet.write(0, col, column_name, header_format)
                    # Auto-adjust column width
                    worksheet.set_column(col, col, max(len(str(column_name)) + 2, 12))
                
                # Write data with appropriate formatting
                for row, data_row in enumerate(df.itertuples(index=False), 1):
                    for col, value in enumerate(data_row):
                        # Try to detect if it's a number
                        if pd.api.types.is_numeric_dtype(type(value)):
                            worksheet.write(row, col, value, number_format)
                        else:
                            worksheet.write(row, col, value, cell_format)
                
                # Add autofilter
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                
            elif file_extension == '.json':
                data = json.loads(content)
                
                if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                    # Convert list of objects to Excel
                    df = pd.DataFrame(data)
                    worksheet = workbook.add_worksheet('Data')
                    
                    # Write headers
                    for col, column_name in enumerate(df.columns):
                        worksheet.write(0, col, column_name, header_format)
                        worksheet.set_column(col, col, max(len(str(column_name)) + 2, 12))
                    
                    # Write data
                    for row, data_row in enumerate(df.itertuples(index=False), 1):
                        for col, value in enumerate(data_row):
                            if isinstance(value, (int, float)):
                                worksheet.write(row, col, value, number_format)
                            else:
                                worksheet.write(row, col, str(value), cell_format)
                    
                    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                    
                elif isinstance(data, dict):
                    # Convert dict to key-value Excel
                    worksheet = workbook.add_worksheet('Key-Value')
                    worksheet.write(0, 0, 'Key', header_format)
                    worksheet.write(0, 1, 'Value', header_format)
                    worksheet.set_column(0, 0, 20)
                    worksheet.set_column(1, 1, 30)
                    
                    for row, (key, value) in enumerate(data.items(), 1):
                        worksheet.write(row, 0, str(key), cell_format)
                        if isinstance(value, (int, float)):
                            worksheet.write(row, 1, value, number_format)
                        else:
                            worksheet.write(row, 1, str(value), cell_format)
                    
                    worksheet.autofilter(0, 0, len(data), 1)
                
            elif file_extension == '.pdf':
                # Try to extract tables from PDF for Excel conversion
                extracted_tables = []
                
                # Method 1: Try camelot for table extraction
                if camelot:
                    try:
                        temp_pdf_path = f"/tmp/{uploaded_file.name}"
                        with open(temp_pdf_path, "wb") as f:
                            f.write(uploaded_file.getvalue())
                        
                        tables = camelot.read_pdf(temp_pdf_path, pages='all')
                        for i, table in enumerate(tables):
                            df = table.df
                            if not df.empty:
                                extracted_tables.append((f'Table_{i+1}', df))
                        
                        os.remove(temp_pdf_path)
                    except Exception:
                        pass
                
                # Method 2: Try tabula as fallback
                if not extracted_tables and tabula:
                    try:
                        temp_pdf_path = f"/tmp/{uploaded_file.name}"
                        with open(temp_pdf_path, "wb") as f:
                            f.write(uploaded_file.getvalue())
                        
                        tables = tabula.read_pdf(temp_pdf_path, pages='all', multiple_tables=True)
                        for i, table in enumerate(tables):
                            if not table.empty:
                                extracted_tables.append((f'Table_{i+1}', table))
                        
                        os.remove(temp_pdf_path)
                    except Exception:
                        pass
                
                if extracted_tables:
                    # Create worksheets for each table
                    for sheet_name, df in extracted_tables:
                        worksheet = workbook.add_worksheet(sheet_name)
                        
                        # Write headers
                        for col, column_name in enumerate(df.columns):
                            worksheet.write(0, col, str(column_name), header_format)
                            worksheet.set_column(col, col, 15)
                        
                        # Write data
                        for row, data_row in enumerate(df.itertuples(index=False), 1):
                            for col, value in enumerate(data_row):
                                if pd.api.types.is_numeric_dtype(type(value)):
                                    worksheet.write(row, col, value, number_format)
                                else:
                                    worksheet.write(row, col, str(value), cell_format)
                        
                        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                else:
                    # Fallback: Convert text content to Excel
                    lines = [line for line in content.split('\n') if line.strip()]
                    worksheet = workbook.add_worksheet('Content')
                    worksheet.write(0, 0, 'Line Number', header_format)
                    worksheet.write(0, 1, 'Content', header_format)
                    worksheet.set_column(0, 0, 12)
                    worksheet.set_column(1, 1, 50)
                    
                    for row, line in enumerate(lines, 1):
                        worksheet.write(row, 0, row, number_format)
                        worksheet.write(row, 1, line, cell_format)
                    
                    worksheet.autofilter(0, 0, len(lines), 1)
            
            else:
                # Generic text-based content to Excel
                lines = [line for line in content.split('\n') if line.strip()]
                worksheet = workbook.add_worksheet('Content')
                worksheet.write(0, 0, 'Line Number', header_format)
                worksheet.write(0, 1, 'Content', header_format)
                worksheet.set_column(0, 0, 12)
                worksheet.set_column(1, 1, 50)
                
                for row, line in enumerate(lines, 1):
                    worksheet.write(row, 0, row, number_format)
                    worksheet.write(row, 1, line, cell_format)
                
                worksheet.autofilter(0, 0, len(lines), 1)
            
            # Add metadata worksheet
            meta_worksheet = workbook.add_worksheet('Metadata')
            meta_worksheet.write(0, 0, 'Property', header_format)
            meta_worksheet.write(0, 1, 'Value', header_format)
            meta_worksheet.write(1, 0, 'Original File', cell_format)
            meta_worksheet.write(1, 1, uploaded_file.name, cell_format)
            meta_worksheet.write(2, 0, 'Conversion Date', cell_format)
            meta_worksheet.write(2, 1, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), cell_format)
            meta_worksheet.write(3, 0, 'File Size (KB)', cell_format)
            meta_worksheet.write(3, 1, len(uploaded_file.getvalue()) / 1024, number_format)
            meta_worksheet.set_column(0, 0, 20)
            meta_worksheet.set_column(1, 1, 30)
            
            workbook.close()
            buffer.seek(0)
            
            return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.xlsx"
            
        except Exception as e1:
            print(f"xlsxwriter method failed: {e1}, trying fallback...")
            
            # Method 2: Fallback to openpyxl
            buffer = BytesIO()
            
            if file_extension == '.csv':
                df = pd.read_csv(StringIO(content))
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Data')
                    
            elif file_extension == '.json':
                data = json.loads(content)
                if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Data')
                elif isinstance(data, dict):
                    df = pd.DataFrame(list(data.items()), columns=['Key', 'Value'])
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Key-Value')
                else:
                    return None, "JSON format not suitable for Excel conversion"
            else:
                lines = [line for line in content.split('\n') if line.strip()]
                df = pd.DataFrame({'Content': lines})
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Content')
            
            buffer.seek(0)
            return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.xlsx"
        
    except Exception as e:
        return None, f"Error converting to Excel: {str(e)}"

def convert_to_json(uploaded_file, content):
    """Convert various formats to JSON with professional structure and metadata"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        # Create comprehensive metadata
        metadata = {
            "_metadata": {
                "source_file": uploaded_file.name,
                "original_format": file_extension.upper(),
                "file_size_bytes": len(uploaded_file.getvalue()),
                "file_size_kb": round(len(uploaded_file.getvalue()) / 1024, 2),
                "conversion_timestamp": datetime.now().isoformat(),
                "conversion_date": datetime.now().strftime('%B %d, %Y'),
                "conversion_time": datetime.now().strftime('%H:%M:%S'),
                "converter_version": "Professional JSON Converter v2.0"
            }
        }
        
        if file_extension == '.csv':
            # Enhanced CSV to JSON conversion
            try:
                df = pd.read_csv(StringIO(content))
                
                # Create structured JSON with metadata
                json_data = {
                    **metadata,
                    "data_summary": {
                        "total_records": len(df),
                        "total_columns": len(df.columns),
                        "column_names": df.columns.tolist(),
                        "data_types": {col: str(dtype) for col, dtype in df.dtypes.items()},
                        "memory_usage_mb": round(df.memory_usage(deep=True).sum() / 1024 / 1024, 3)
                    },
                    "statistics": {},
                    "records": df.to_dict('records')
                }
                
                # Add statistics for numeric columns
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    json_data["statistics"] = {
                        "numeric_columns": numeric_cols.tolist(),
                        "summary_stats": df[numeric_cols].describe().to_dict()
                    }
                
                # Sample data for quick preview
                json_data["sample_records"] = df.head(5).to_dict('records')
                
            except Exception as e:
                # Fallback for malformed CSV
                lines = [line.strip() for line in content.split('\n') if line.strip()]
                json_data = {
                    **metadata,
                    "parse_error": str(e),
                    "raw_content": {
                        "lines": lines,
                        "total_lines": len(lines)
                    }
                }
        
        elif file_extension == '.xlsx':
            # Enhanced Excel to JSON conversion
            try:
                df = pd.read_excel(BytesIO(uploaded_file.getvalue()))
                
                json_data = {
                    **metadata,
                    "workbook_info": {
                        "total_records": len(df),
                        "total_columns": len(df.columns),
                        "column_names": df.columns.tolist(),
                        "data_types": {col: str(dtype) for col, dtype in df.dtypes.items()}
                    },
                    "data": df.to_dict('records'),
                    "sample_data": df.head(3).to_dict('records')
                }
                
                # Add statistics for numeric data
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    json_data["statistics"] = df[numeric_cols].describe().to_dict()
                    
            except Exception as e:
                json_data = {
                    **metadata,
                    "error": f"Failed to parse Excel file: {str(e)}",
                    "fallback_content": "Binary Excel content - conversion failed"
                }
        
        elif file_extension == '.html':
            # Enhanced HTML to JSON conversion
            try:
                soup = BeautifulSoup(content, 'html.parser')
                
                # Extract structured data from HTML
                html_structure = {
                    "title": soup.find('title').text.strip() if soup.find('title') else "No title",
                    "headings": {},
                    "paragraphs": [],
                    "links": [],
                    "images": [],
                    "tables": [],
                    "lists": []
                }
                
                # Extract headings
                for i in range(1, 7):
                    headings = [h.text.strip() for h in soup.find_all(f'h{i}') if h.text.strip()]
                    if headings:
                        html_structure["headings"][f"h{i}"] = headings
                
                # Extract paragraphs
                html_structure["paragraphs"] = [p.text.strip() for p in soup.find_all('p') if p.text.strip()]
                
                # Extract links
                html_structure["links"] = [{"text": a.text.strip(), "href": a.get('href', '')} 
                                         for a in soup.find_all('a') if a.text.strip()]
                
                # Extract images
                html_structure["images"] = [{"alt": img.get('alt', ''), "src": img.get('src', '')} 
                                          for img in soup.find_all('img')]
                
                # Extract tables
                for table in soup.find_all('table'):
                    table_data = []
                    rows = table.find_all('tr')
                    if rows:
                        headers = [th.text.strip() for th in rows[0].find_all(['th', 'td'])]
                        for row in rows[1:]:
                            cells = [td.text.strip() for td in row.find_all(['th', 'td'])]
                            if cells:
                                table_data.append(dict(zip(headers, cells)))
                    html_structure["tables"].append(table_data)
                
                # Extract lists
                for ul in soup.find_all(['ul', 'ol']):
                    list_items = [li.text.strip() for li in ul.find_all('li') if li.text.strip()]
                    if list_items:
                        html_structure["lists"].append({
                            "type": ul.name,
                            "items": list_items
                        })
                
                json_data = {
                    **metadata,
                    "html_structure": html_structure,
                    "raw_text": soup.get_text()[:1000] + ("..." if len(soup.get_text()) > 1000 else "")
                }
                
            except Exception as e:
                json_data = {
                    **metadata,
                    "error": f"HTML parsing failed: {str(e)}",
                    "raw_content": content[:500] + ("..." if len(content) > 500 else "")
                }
        
        elif file_extension == '.xml':
            # Enhanced XML to JSON conversion
            try:
                root = ET.fromstring(content)
                
                def xml_to_dict(element):
                    result = {}
                    # Add attributes
                    if element.attrib:
                        result["@attributes"] = element.attrib
                    
                    # Add text content
                    if element.text and element.text.strip():
                        if len(element) == 0:  # No child elements
                            return element.text.strip()
                        result["@text"] = element.text.strip()
                    
                    # Add child elements
                    for child in element:
                        child_data = xml_to_dict(child)
                        if child.tag in result:
                            if not isinstance(result[child.tag], list):
                                result[child.tag] = [result[child.tag]]
                            result[child.tag].append(child_data)
                        else:
                            result[child.tag] = child_data
                    
                    return result
                
                xml_dict = xml_to_dict(root)
                
                json_data = {
                    **metadata,
                    "xml_structure": {
                        "root_element": root.tag,
                        "namespace": root.tag.split('}')[0][1:] if '}' in root.tag else None,
                        "attributes": root.attrib,
                        "total_elements": len(list(root.iter()))
                    },
                    "data": {root.tag: xml_dict}
                }
                
            except Exception as e:
                json_data = {
                    **metadata,
                    "error": f"XML parsing failed: {str(e)}",
                    "raw_content": content[:500] + ("..." if len(content) > 500 else "")
                }
        
        elif file_extension in ['.txt', '.pdf', '.docx']:
            # Enhanced text-based content to JSON
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            
            # Analyze content structure
            analysis = {
                "total_lines": len(lines),
                "total_characters": len(content),
                "total_words": len(content.split()),
                "average_line_length": round(sum(len(line) for line in lines) / len(lines), 2) if lines else 0,
                "longest_line": max(lines, key=len) if lines else "",
                "shortest_line": min(lines, key=len) if lines else ""
            }
            
            # Try to detect structure
            potential_headers = [line for line in lines if len(line) < 50 and (line.isupper() or line.endswith(':'))]
            potential_sections = []
            
            current_section = {"title": "Introduction", "content": []}
            for line in lines:
                if line in potential_headers:
                    if current_section["content"]:
                        potential_sections.append(current_section)
                    current_section = {"title": line, "content": []}
                else:
                    current_section["content"].append(line)
            if current_section["content"]:
                potential_sections.append(current_section)
            
            json_data = {
                **metadata,
                "content_analysis": analysis,
                "detected_structure": {
                    "potential_headers": potential_headers,
                    "sections": potential_sections
                },
                "raw_content": {
                    "all_lines": lines,
                    "first_100_words": ' '.join(content.split()[:100]) + ("..." if len(content.split()) > 100 else "")
                }
            }
        
        elif file_extension == '.yaml':
            # YAML to JSON conversion
            try:
                yaml_data = yaml.safe_load(content)
                json_data = {
                    **metadata,
                    "yaml_structure": {
                        "data_type": type(yaml_data).__name__,
                        "keys": list(yaml_data.keys()) if isinstance(yaml_data, dict) else None,
                        "length": len(yaml_data) if hasattr(yaml_data, '__len__') else None
                    },
                    "data": yaml_data
                }
            except Exception as e:
                json_data = {
                    **metadata,
                    "error": f"YAML parsing failed: {str(e)}",
                    "raw_content": content
                }
        
        else:
            # Generic content conversion
            json_data = {
                **metadata,
                "content_info": {
                    "character_count": len(content),
                    "line_count": len(content.split('\n')),
                    "word_count": len(content.split())
                },
                "content": content[:1000] + ("..." if len(content) > 1000 else ""),
                "full_content": content
            }
        
        # Pretty print JSON with proper formatting
        json_string = json.dumps(json_data, indent=2, ensure_ascii=False, sort_keys=False)
        buffer = StringIO(json_string)
        
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.json"
        
    except Exception as e:
        return None, f"Error converting to JSON: {str(e)}"

def convert_to_yaml(uploaded_file, content):
    """Convert to YAML format with professional structure and metadata"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        # Create YAML header comment
        yaml_header = f"""# Professional YAML Conversion
# Source: {uploaded_file.name}
# Original Format: {file_extension.upper()}
# Converted: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}
# File Size: {len(uploaded_file.getvalue()) / 1024:.1f} KB
# Converter: Professional YAML Converter v2.0

"""
        
        if file_extension == '.json':
            # Enhanced JSON to YAML conversion
            try:
                data = json.loads(content)
                
                # Create structured YAML
                yaml_data = {
                    "metadata": {
                        "source_file": uploaded_file.name,
                        "conversion_date": datetime.now().strftime('%Y-%m-%d'),
                        "conversion_time": datetime.now().strftime('%H:%M:%S'),
                        "original_format": "JSON",
                        "data_type": type(data).__name__
                    },
                    "content": data
                }
                
                if isinstance(data, dict):
                    yaml_data["structure_info"] = {
                        "type": "object",
                        "keys_count": len(data),
                        "top_level_keys": list(data.keys())[:10]  # First 10 keys
                    }
                elif isinstance(data, list):
                    yaml_data["structure_info"] = {
                        "type": "array",
                        "items_count": len(data),
                        "item_types": list(set(type(item).__name__ for item in data[:10]))
                    }
                
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, 
                                                    allow_unicode=True, sort_keys=False, indent=2)
                
            except Exception as e:
                yaml_string = yaml_header + yaml.dump({
                    "error": f"JSON parsing failed: {str(e)}",
                    "raw_content": content
                }, default_flow_style=False, allow_unicode=True)
        
        elif file_extension == '.csv':
            # Enhanced CSV to YAML conversion
            try:
                df = pd.read_csv(StringIO(content))
                
                yaml_data = {
                    "dataset_info": {
                        "name": uploaded_file.name,
                        "format": "CSV Dataset",
                        "records_count": len(df),
                        "columns_count": len(df.columns),
                        "columns": df.columns.tolist(),
                        "data_types": {col: str(dtype) for col, dtype in df.dtypes.items()}
                    },
                    "summary_statistics": {},
                    "sample_data": df.head(5).to_dict('records'),
                    "full_dataset": df.to_dict('records')
                }
                
                # Add statistics for numeric columns
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    yaml_data["summary_statistics"] = {
                        "numeric_columns": numeric_cols.tolist(),
                        "stats": df[numeric_cols].describe().to_dict()
                    }
                
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, 
                                                    allow_unicode=True, sort_keys=False, indent=2)
                
            except Exception as e:
                lines = [line.strip() for line in content.split('\n') if line.strip()]
                yaml_data = {
                    "error": f"CSV parsing failed: {str(e)}",
                    "raw_data": {
                        "lines": lines,
                        "total_lines": len(lines)
                    }
                }
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, allow_unicode=True)
        
        elif file_extension == '.html':
            # Enhanced HTML to YAML conversion
            try:
                soup = BeautifulSoup(content, 'html.parser')
                
                yaml_data = {
                    "html_document": {
                        "title": soup.find('title').text.strip() if soup.find('title') else "No title",
                        "meta_info": {
                            "total_paragraphs": len(soup.find_all('p')),
                            "total_headings": len(soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])),
                            "total_links": len(soup.find_all('a')),
                            "total_images": len(soup.find_all('img')),
                            "total_tables": len(soup.find_all('table'))
                        },
                        "structure": {
                            "headings": [h.text.strip() for h in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']) if h.text.strip()][:10],
                            "paragraphs": [p.text.strip() for p in soup.find_all('p') if p.text.strip()][:5],
                            "links": [{"text": a.text.strip(), "href": a.get('href', '')} 
                                    for a in soup.find_all('a') if a.text.strip()][:10]
                        },
                        "content_preview": soup.get_text()[:500] + ("..." if len(soup.get_text()) > 500 else "")
                    }
                }
                
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, 
                                                    allow_unicode=True, sort_keys=False, indent=2)
                
            except Exception as e:
                yaml_data = {
                    "error": f"HTML parsing failed: {str(e)}",
                    "raw_content": content[:500] + ("..." if len(content) > 500 else "")
                }
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, allow_unicode=True)
        
        elif file_extension == '.xml':
            # Enhanced XML to YAML conversion
            try:
                root = ET.fromstring(content)
                
                def xml_to_dict(element):
                    result = {}
                    if element.attrib:
                        result["attributes"] = element.attrib
                    if element.text and element.text.strip():
                        if len(element) == 0:
                            return element.text.strip()
                        result["text"] = element.text.strip()
                    for child in element:
                        child_data = xml_to_dict(child)
                        if child.tag in result:
                            if not isinstance(result[child.tag], list):
                                result[child.tag] = [result[child.tag]]
                            result[child.tag].append(child_data)
                        else:
                            result[child.tag] = child_data
                    return result
                
                yaml_data = {
                    "xml_document": {
                        "root_element": root.tag,
                        "structure_info": {
                            "total_elements": len(list(root.iter())),
                            "root_attributes": root.attrib,
                            "namespace": root.tag.split('}')[0][1:] if '}' in root.tag else None
                        },
                        "content": xml_to_dict(root)
                    }
                }
                
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, 
                                                    allow_unicode=True, sort_keys=False, indent=2)
                
            except Exception as e:
                yaml_data = {
                    "error": f"XML parsing failed: {str(e)}",
                    "raw_content": content[:500] + ("..." if len(content) > 500 else "")
                }
                yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, allow_unicode=True)
        
        else:
            # Enhanced text to YAML conversion
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            
            # Analyze content
            content_analysis = {
                "total_lines": len(lines),
                "total_characters": len(content),
                "total_words": len(content.split()),
                "average_line_length": round(sum(len(line) for line in lines) / len(lines), 2) if lines else 0
            }
            
            # Try to detect structured content
            potential_headers = [line for line in lines if len(line) < 50 and (line.isupper() or line.endswith(':'))]
            
            yaml_data = {
                "document": {
                    "source_file": uploaded_file.name,
                    "analysis": content_analysis,
                    "structure": {
                        "detected_headers": potential_headers,
                        "content_lines": lines,
                        "preview": ' '.join(content.split()[:50]) + ("..." if len(content.split()) > 50 else "")
                    }
                }
            }
            
            yaml_string = yaml_header + yaml.dump(yaml_data, default_flow_style=False, 
                                                allow_unicode=True, sort_keys=False, indent=2)
        
        buffer = StringIO(yaml_string)
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.yaml"
        
    except Exception as e:
        return None, f"Error converting to YAML: {str(e)}"

def convert_to_markdown(uploaded_file, content):
    """Convert to Markdown format with professional formatting and structure"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        markdown_content = f"""# 📄 {uploaded_file.name}

---

**Document Information:**
- 📅 **Converted:** {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}
- 📝 **Original Format:** {file_extension.upper()}
- 📊 **File Size:** {len(uploaded_file.getvalue()) / 1024:.1f} KB
- ⚡ **Conversion:** Professional Markdown Converter

---

"""
        
        if file_extension == '.html':
            # Enhanced HTML to Markdown conversion
            try:
                soup = BeautifulSoup(content, 'html.parser')
                
                # Extract title
                title = soup.find('title')
                if title and title.text.strip():
                    markdown_content += f"## 🏷️ {title.text.strip()}\n\n"
                
                # Extract main content (avoid scripts, styles)
                for script in soup(["script", "style", "meta", "link"]):
                    script.decompose()
                
                # Convert headers
                for i in range(1, 7):
                    for header in soup.find_all(f'h{i}'):
                        if header.text.strip():
                            prefix = '#' * (i + 1)
                            markdown_content += f"{prefix} {header.text.strip()}\n\n"
                        header.decompose()
                
                # Convert paragraphs
                for p in soup.find_all('p'):
                    if p.text.strip():
                        markdown_content += f"{p.text.strip()}\n\n"
                    p.decompose()
                
                # Convert lists
                for ul in soup.find_all('ul'):
                    for li in ul.find_all('li'):
                        if li.text.strip():
                            markdown_content += f"- {li.text.strip()}\n"
                    markdown_content += "\n"
                    ul.decompose()
                
                for ol in soup.find_all('ol'):
                    for i, li in enumerate(ol.find_all('li'), 1):
                        if li.text.strip():
                            markdown_content += f"{i}. {li.text.strip()}\n"
                    markdown_content += "\n"
                    ol.decompose()
                
                # Convert tables
                for table in soup.find_all('table'):
                    rows = table.find_all('tr')
                    if rows:
                        # Header row
                        header_row = rows[0]
                        headers = [th.text.strip() for th in header_row.find_all(['th', 'td'])]
                        if headers:
                            markdown_content += "| " + " | ".join(headers) + " |\n"
                            markdown_content += "|" + "---|" * len(headers) + "\n"
                            
                            # Data rows
                            for row in rows[1:]:
                                cells = [td.text.strip() for td in row.find_all(['th', 'td'])]
                                if cells:
                                    # Pad cells to match header count
                                    while len(cells) < len(headers):
                                        cells.append("")
                                    markdown_content += "| " + " | ".join(cells[:len(headers)]) + " |\n"
                        markdown_content += "\n"
                    table.decompose()
                
                # Convert blockquotes
                for blockquote in soup.find_all('blockquote'):
                    if blockquote.text.strip():
                        lines = blockquote.text.strip().split('\n')
                        for line in lines:
                            if line.strip():
                                markdown_content += f"> {line.strip()}\n"
                        markdown_content += "\n"
                    blockquote.decompose()
                
                # Convert code blocks
                for pre in soup.find_all('pre'):
                    if pre.text.strip():
                        markdown_content += f"```\n{pre.text.strip()}\n```\n\n"
                    pre.decompose()
                
                # Convert inline code
                for code in soup.find_all('code'):
                    if code.text.strip():
                        markdown_content += f"`{code.text.strip()}`"
                    code.decompose()
                
                # Add remaining text
                remaining_text = soup.get_text()
                if remaining_text.strip():
                    markdown_content += f"## 📄 Additional Content\n\n{remaining_text.strip()}\n\n"
                    
            except Exception as e:
                # Fallback to simple conversion
                markdown_content += f"## 📄 HTML Content\n\n```html\n{content}\n```\n\n"
        
        elif file_extension == '.csv':
            # Enhanced CSV to Markdown table conversion
            try:
                df = pd.read_csv(StringIO(content))
                
                markdown_content += f"## 📊 Data Overview\n\n"
                markdown_content += f"- **Rows:** {len(df)}\n"
                markdown_content += f"- **Columns:** {len(df.columns)}\n"
                markdown_content += f"- **Data Types:** {', '.join(df.dtypes.astype(str).unique())}\n\n"
                
                # Summary statistics for numeric columns
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    markdown_content += f"## 📈 Summary Statistics\n\n"
                    summary_df = df[numeric_cols].describe()
                    markdown_content += summary_df.to_markdown()
                    markdown_content += "\n\n"
                
                # Full data table (limit to first 100 rows for readability)
                display_df = df.head(100)
                markdown_content += f"## 📋 Data Table\n\n"
                if len(df) > 100:
                    markdown_content += f"*Showing first 100 rows of {len(df)} total rows*\n\n"
                
                markdown_content += display_df.to_markdown(index=False)
                markdown_content += "\n\n"
                
                if len(df) > 100:
                    markdown_content += f"*... and {len(df) - 100} more rows*\n\n"
                    
            except Exception as e:
                # Fallback for malformed CSV
                lines = content.split('\n')
                markdown_content += f"## 📄 CSV Content\n\n"
                markdown_content += "```csv\n"
                for line in lines[:50]:  # Limit to first 50 lines
                    markdown_content += f"{line}\n"
                if len(lines) > 50:
                    markdown_content += f"... and {len(lines) - 50} more lines\n"
                markdown_content += "```\n\n"
        
        elif file_extension == '.json':
            # Enhanced JSON to Markdown conversion
            try:
                data = json.loads(content)
                
                markdown_content += f"## 🔗 JSON Structure\n\n"
                
                if isinstance(data, dict):
                    markdown_content += f"**Type:** Object with {len(data)} keys\n\n"
                    markdown_content += f"### 🔑 Keys Overview\n\n"
                    
                    # Create a summary table
                    markdown_content += "| Key | Type | Sample Value |\n"
                    markdown_content += "|-----|------|-------------|\n"
                    
                    for key, value in data.items():
                        value_type = type(value).__name__
                        if isinstance(value, (dict, list)):
                            sample_value = f"{value_type} with {len(value)} items"
                        else:
                            sample_value = str(value)[:50] + ("..." if len(str(value)) > 50 else "")
                        markdown_content += f"| `{key}` | {value_type} | {sample_value} |\n"
                    
                    markdown_content += "\n"
                    
                    # Detailed breakdown for each key
                    markdown_content += f"### 📊 Detailed Content\n\n"
                    for key, value in data.items():
                        markdown_content += f"#### {key}\n\n"
                        
                        if isinstance(value, dict):
                            markdown_content += f"**Type:** Object with {len(value)} properties\n\n"
                            if len(value) <= 10:
                                for subkey, subvalue in value.items():
                                    markdown_content += f"- **{subkey}:** {subvalue}\n"
                            else:
                                for i, (subkey, subvalue) in enumerate(list(value.items())[:5]):
                                    markdown_content += f"- **{subkey}:** {subvalue}\n"
                                markdown_content += f"- *... and {len(value) - 5} more properties*\n"
                        
                        elif isinstance(value, list):
                            markdown_content += f"**Type:** Array with {len(value)} items\n\n"
                            if len(value) <= 5:
                                for i, item in enumerate(value):
                                    markdown_content += f"{i+1}. {item}\n"
                            else:
                                for i, item in enumerate(value[:3]):
                                    markdown_content += f"{i+1}. {item}\n"
                                markdown_content += f"... and {len(value) - 3} more items\n"
                        
                        else:
                            markdown_content += f"**Value:** {value}\n"
                        
                        markdown_content += "\n"
                
                elif isinstance(data, list):
                    markdown_content += f"**Type:** Array with {len(data)} items\n\n"
                    
                    if len(data) > 0:
                        item_types = {}
                        for item in data[:10]:  # Sample first 10 items
                            item_type = type(item).__name__
                            item_types[item_type] = item_types.get(item_type, 0) + 1
                        
                        markdown_content += f"### 📊 Item Types\n\n"
                        for item_type, count in item_types.items():
                            markdown_content += f"- **{item_type}:** {count} items\n"
                        markdown_content += "\n"
                        
                        # If it's an array of objects, try to create a table
                        if isinstance(data[0], dict):
                            sample_df = pd.DataFrame(data[:10])  # First 10 items
                            markdown_content += f"### 📋 Sample Data (First 10 items)\n\n"
                            markdown_content += sample_df.to_markdown(index=False)
                            markdown_content += "\n\n"
                            if len(data) > 10:
                                markdown_content += f"*... and {len(data) - 10} more items*\n\n"
                        else:
                            markdown_content += f"### 📋 Sample Items\n\n"
                            for i, item in enumerate(data[:10]):
                                markdown_content += f"{i+1}. {item}\n"
                            if len(data) > 10:
                                markdown_content += f"... and {len(data) - 10} more items\n"
                            markdown_content += "\n"
                
                # Add raw JSON
                markdown_content += f"### 🔍 Raw JSON\n\n"
                markdown_content += "```json\n"
                markdown_content += json.dumps(data, indent=2)
                markdown_content += "\n```\n\n"
                
            except Exception as e:
                markdown_content += f"## 📄 JSON Content\n\n"
                markdown_content += f"```json\n{content}\n```\n\n"
        
        elif file_extension == '.pdf':
            # Enhanced PDF to Markdown conversion
            markdown_content += f"## 📄 PDF Document Content\n\n"
            
            # Split content by pages if page markers exist
            if "--- Page" in content:
                pages = content.split("--- Page")
                markdown_content += f"**Total Pages:** {len(pages) - 1}\n\n"
                
                for i, page_content in enumerate(pages[1:], 1):
                    markdown_content += f"### 📄 Page {i}\n\n"
                    
                    # Check for tables
                    if "TABLE:" in page_content:
                        parts = page_content.split("TABLE:")
                        
                        # Add text before table
                        if parts[0].strip():
                            text_lines = parts[0].strip().split('\n')
                            for line in text_lines:
                                if line.strip():
                                    # Detect potential headers
                                    if len(line) < 50 and (line.isupper() or line.endswith(':')):
                                        markdown_content += f"#### {line.strip()}\n\n"
                                    else:
                                        markdown_content += f"{line.strip()}\n\n"
                        
                        # Add tables
                        for j, table_part in enumerate(parts[1:], 1):
                            lines = table_part.split('\n')
                            table_lines = [line for line in lines if '|' in line]
                            
                            if table_lines:
                                markdown_content += f"**Table {j}:**\n\n"
                                
                                # Process table
                                for k, line in enumerate(table_lines[:20]):  # Limit table rows
                                    if k == 0:
                                        # Header row
                                        markdown_content += f"{line}\n"
                                        # Add separator
                                        cells = line.split('|')
                                        separator = "|" + "---|" * (len(cells) - 1)
                                        markdown_content += f"{separator}\n"
                                    else:
                                        markdown_content += f"{line}\n"
                                
                                if len(table_lines) > 20:
                                    markdown_content += f"*... and {len(table_lines) - 20} more rows*\n"
                                markdown_content += "\n"
                            
                            # Add remaining text after table
                            remaining_lines = [line for line in lines if '|' not in line]
                            remaining_text = '\n'.join(remaining_lines).strip()
                            if remaining_text:
                                text_lines = remaining_text.split('\n')
                                for line in text_lines:
                                    if line.strip():
                                        markdown_content += f"{line.strip()}\n\n"
                    else:
                        # Regular text content
                        text_lines = page_content.strip().split('\n')
                        for line in text_lines:
                            if line.strip():
                                # Detect potential headers
                                if len(line) < 50 and (line.isupper() or line.endswith(':')):
                                    markdown_content += f"#### {line.strip()}\n\n"
                                else:
                                    markdown_content += f"{line.strip()}\n\n"
            else:
                # No page markers, treat as single content
                text_lines = content.split('\n')
                for line in text_lines:
                    if line.strip():
                        # Detect potential headers
                        if len(line) < 50 and (line.isupper() or line.endswith(':')):
                            markdown_content += f"## {line.strip()}\n\n"
                        else:
                            markdown_content += f"{line.strip()}\n\n"
        
        else:
            # Enhanced plain text to Markdown conversion
            markdown_content += f"## 📄 Document Content\n\n"
            
            lines = content.split('\n')
            current_section = []
            
            for line in lines:
                line = line.strip()
                if not line:
                    if current_section:
                        markdown_content += '\n'.join(current_section) + '\n\n'
                        current_section = []
                else:
                    # Detect headers (short lines, all caps, or ending with :)
                    if len(line) < 50 and (line.isupper() or line.endswith(':')):
                        if current_section:
                            markdown_content += '\n'.join(current_section) + '\n\n'
                            current_section = []
                        markdown_content += f"### {line}\n\n"
                    else:
                        current_section.append(line)
            
            # Add remaining content
            if current_section:
                markdown_content += '\n'.join(current_section) + '\n\n'
        
        # Add footer
        markdown_content += f"""---

## 📊 Conversion Summary

- **Original File:** {uploaded_file.name}
- **File Size:** {len(uploaded_file.getvalue()) / 1024:.1f} KB
- **Conversion Tool:** Professional Markdown Converter
- **Conversion Time:** {datetime.now().strftime('%H:%M:%S')}
- **Date:** {datetime.now().strftime('%B %d, %Y')}

---

*Converted with ❤️ by Professional File Converter*
"""
        
        buffer = StringIO(markdown_content)
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.md"
        
    except Exception as e:
        return None, f"Error converting to Markdown: {str(e)}"

def convert_to_html(uploaded_file, content):
    """Convert to HTML format with professional styling and responsive design"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        # Professional HTML template with modern styling
        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{uploaded_file.name}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            line-height: 1.6;
            color: #333;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: white;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            border-radius: 10px;
            margin-top: 20px;
            margin-bottom: 20px;
        }}
        
        .header {{
            text-align: center;
            padding: 30px 0;
            border-bottom: 3px solid #667eea;
            margin-bottom: 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 10px 10px 0 0;
            margin: -20px -20px 30px -20px;
            padding: 30px 20px;
        }}
        
        h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .metadata {{
            font-size: 0.9em;
            opacity: 0.9;
            margin-top: 10px;
        }}
        
        h2, h3, h4 {{
            color: #667eea;
            margin: 25px 0 15px 0;
            padding-bottom: 5px;
            border-bottom: 2px solid #e1e5e9;
        }}
        
        h2 {{
            font-size: 1.8em;
        }}
        
        h3 {{
            font-size: 1.4em;
        }}
        
        p {{
            margin: 15px 0;
            text-align: justify;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }}
        
        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }}
        
        td {{
            padding: 12px;
            border-bottom: 1px solid #f0f0f0;
            transition: background-color 0.3s ease;
        }}
        
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        
        tr:hover td {{
            background-color: #e3f2fd;
        }}
        
        pre {{
            background: #f8f9fa;
            border: 1px solid #e1e5e9;
            border-radius: 8px;
            padding: 20px;
            overflow-x: auto;
            margin: 20px 0;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        code {{
            background: #f1f3f4;
            padding: 2px 6px;
            border-radius: 4px;
            font-family: 'Monaco', 'Menlo', 'Consolas', monospace;
            font-size: 0.9em;
        }}
        
        .content-section {{
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
            border-left: 4px solid #667eea;
        }}
        
        .footer {{
            text-align: center;
            padding: 20px;
            margin-top: 40px;
            border-top: 2px solid #e1e5e9;
            color: #666;
            font-size: 0.9em;
        }}
        
        blockquote {{
            border-left: 4px solid #667eea;
            padding-left: 20px;
            margin: 20px 0;
            color: #555;
            font-style: italic;
            background: #f8f9fa;
            padding: 15px 20px;
            border-radius: 0 8px 8px 0;
        }}
        
        @media (max-width: 768px) {{
            .container {{
                margin: 10px;
                padding: 15px;
            }}
            
            h1 {{
                font-size: 2em;
            }}
            
            table {{
                font-size: 0.9em;
            }}
            
            th, td {{
                padding: 8px;
            }}
        }}
        
        .highlight {{
            background: linear-gradient(120deg, #a8edea 0%, #fed6e3 100%);
            padding: 2px 4px;
            border-radius: 3px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 {uploaded_file.name}</h1>
            <div class="metadata">
                Converted on {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}
                <br>Professional HTML Conversion
            </div>
        </div>
"""
        
        if file_extension == '.md':
            # Enhanced Markdown to HTML conversion
            html_content += '<div class="content-section">'
            html_content += markdown.markdown(content, extensions=['tables', 'fenced_code', 'codehilite', 'toc'])
            html_content += '</div>'
            
        elif file_extension == '.csv':
            # Enhanced CSV to HTML table conversion
            try:
                df = pd.read_csv(StringIO(content))
                html_content += '<div class="content-section">'
                html_content += f'<h2>📊 Data Table ({len(df)} rows, {len(df.columns)} columns)</h2>'
                
                # Add summary statistics for numeric columns
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    html_content += '<h3>📈 Summary Statistics</h3>'
                    summary_html = df[numeric_cols].describe().to_html(classes='table table-striped', border=0)
                    html_content += summary_html
                
                html_content += '<h3>📋 Full Data</h3>'
                table_html = df.to_html(classes='table table-striped', border=0, escape=False, index=False)
                html_content += table_html
                html_content += '</div>'
                
            except Exception:
                # Fallback for malformed CSV
                lines = content.split('\n')
                html_content += '<div class="content-section">'
                html_content += '<h2>📄 Text Content</h2>'
                for line in lines:
                    if line.strip():
                        html_content += f"<p>{line}</p>"
                html_content += '</div>'
        
        elif file_extension == '.json':
            # Enhanced JSON to HTML conversion
            try:
                data = json.loads(content)
                html_content += '<div class="content-section">'
                html_content += '<h2>🔗 JSON Data Structure</h2>'
                
                if isinstance(data, dict):
                    html_content += '<h3>📋 Key-Value Pairs</h3>'
                    html_content += '<table>'
                    html_content += '<thead><tr><th>Key</th><th>Value</th><th>Type</th></tr></thead><tbody>'
                    for key, value in data.items():
                        value_type = type(value).__name__
                        if isinstance(value, (dict, list)):
                            display_value = f"<code>{json.dumps(value, indent=2)[:100]}{'...' if len(str(value)) > 100 else ''}</code>"
                        else:
                            display_value = str(value)
                        html_content += f'<tr><td><strong>{key}</strong></td><td>{display_value}</td><td><span class="highlight">{value_type}</span></td></tr>'
                    html_content += '</tbody></table>'
                    
                elif isinstance(data, list):
                    html_content += f'<h3>📊 Array Data ({len(data)} items)</h3>'
                    if len(data) > 0 and isinstance(data[0], dict):
                        # Convert to table if it's list of objects
                        df = pd.DataFrame(data)
                        table_html = df.to_html(classes='table table-striped', border=0, escape=False, index=False)
                        html_content += table_html
                    else:
                        html_content += '<ul>'
                        for i, item in enumerate(data[:50]):  # Limit to first 50 items
                            html_content += f'<li><strong>Item {i+1}:</strong> {item}</li>'
                        if len(data) > 50:
                            html_content += f'<li><em>... and {len(data)-50} more items</em></li>'
                        html_content += '</ul>'
                
                html_content += '<h3>🔍 Raw JSON</h3>'
                html_content += f'<pre><code>{json.dumps(data, indent=2)}</code></pre>'
                html_content += '</div>'
                
            except Exception:
                html_content += f'<div class="content-section"><h2>📄 Raw Content</h2><pre><code>{content}</code></pre></div>'
        
        elif file_extension == '.pdf':
            # Enhanced PDF content display
            html_content += '<div class="content-section">'
            html_content += '<h2>📄 PDF Document Content</h2>'
            
            # Split content by pages if page markers exist
            if "--- Page" in content:
                pages = content.split("--- Page")
                for i, page_content in enumerate(pages[1:], 1):  # Skip first empty split
                    html_content += f'<h3>📄 Page {i}</h3>'
                    
                    # Check for tables
                    if "TABLE:" in page_content:
                        parts = page_content.split("TABLE:")
                        # Add text before table
                        if parts[0].strip():
                            paragraphs = parts[0].strip().split('\n')
                            for para in paragraphs:
                                if para.strip():
                                    html_content += f"<p>{para.strip()}</p>"
                        
                        # Add tables
                        for table_part in parts[1:]:
                            lines = table_part.split('\n')
                            table_lines = [line for line in lines if '|' in line]
                            if table_lines:
                                html_content += '<table>'
                                for j, line in enumerate(table_lines[:10]):  # Limit table rows
                                    cells = [cell.strip() for cell in line.split('|')]
                                    if j == 0:
                                        html_content += '<thead><tr>'
                                        for cell in cells:
                                            if cell:
                                                html_content += f'<th>{cell}</th>'
                                        html_content += '</tr></thead><tbody>'
                                    else:
                                        html_content += '<tr>'
                                        for cell in cells:
                                            if cell:
                                                html_content += f'<td>{cell}</td>'
                                        html_content += '</tr>'
                                html_content += '</tbody></table>'
                            
                            # Add remaining text after table
                            remaining_text = '\n'.join([line for line in lines if '|' not in line])
                            if remaining_text.strip():
                                paragraphs = remaining_text.strip().split('\n')
                                for para in paragraphs:
                                    if para.strip():
                                        html_content += f"<p>{para.strip()}</p>"
                    else:
                        # Regular text content
                        paragraphs = page_content.strip().split('\n')
                        for para in paragraphs:
                            if para.strip():
                                html_content += f"<p>{para.strip()}</p>"
            else:
                # No page markers, treat as single content
                paragraphs = content.split('\n')
                for para in paragraphs:
                    if para.strip():
                        html_content += f"<p>{para.strip()}</p>"
            
            html_content += '</div>'
        
        else:
            # Enhanced plain text to HTML conversion
            html_content += '<div class="content-section">'
            html_content += '<h2>📄 Document Content</h2>'
            lines = content.split('\n')
            
            current_section = []
            for line in lines:
                line = line.strip()
                if not line:
                    if current_section:
                        html_content += f"<p>{'<br>'.join(current_section)}</p>"
                        current_section = []
                else:
                    # Detect headers (short lines, all caps, or ending with :)
                    if len(line) < 50 and (line.isupper() or line.endswith(':')):
                        if current_section:
                            html_content += f"<p>{'<br>'.join(current_section)}</p>"
                            current_section = []
                        html_content += f"<h3>{line}</h3>"
                    else:
                        current_section.append(line)
            
            # Add remaining content
            if current_section:
                html_content += f"<p>{'<br>'.join(current_section)}</p>"
            
            html_content += '</div>'
        
        # Add footer
        html_content += f"""
        <div class="footer">
            <p>🔄 Converted using Professional HTML Converter</p>
            <p>Original file: <strong>{uploaded_file.name}</strong> | 
               Size: <strong>{len(uploaded_file.getvalue()) / 1024:.1f} KB</strong> | 
               Conversion time: <strong>{datetime.now().strftime('%H:%M:%S')}</strong></p>
        </div>
    </div>
</body>
</html>"""
        
        buffer = StringIO(html_content)
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.html"
        
    except Exception as e:
        return None, f"Error converting to HTML: {str(e)}"

def convert_to_txt(uploaded_file, content):
    """Convert any format to plain text"""
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        
        txt_content = f"{uploaded_file.name}\n"
        txt_content += "=" * len(uploaded_file.name) + "\n"
        txt_content += f"Converted on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        if file_extension == '.html':
            # HTML to plain text
            h = html2text.HTML2Text()
            h.ignore_links = True
            h.ignore_images = True
            txt_content += h.handle(content)
        
        elif file_extension == '.json':
            # JSON to readable text
            data = json.loads(content)
            txt_content += json.dumps(data, indent=2)
        
        elif file_extension == '.csv':
            # CSV to formatted text
            df = pd.read_csv(StringIO(content))
            txt_content += df.to_string(index=False)
        
        else:
            # Already text or extract text
            txt_content += content
        
        buffer = StringIO(txt_content)
        return buffer, f"{os.path.splitext(uploaded_file.name)[0]}.txt"
        
    except Exception as e:
        return None, f"Error converting to TXT: {str(e)}"

def perform_file_conversion(user_input, uploaded_file, file_content):
    """Detect conversion request and perform the actual conversion"""
    user_lower = user_input.lower()
    
    conversions = []
    
    # PDF to Word conversion
    if ('pdf' in user_lower and ('word' in user_lower or 'docx' in user_lower)) or \
       ('convert' in user_lower and 'word' in user_lower and uploaded_file.name.endswith('.pdf')):
        buffer, filename = convert_pdf_to_word(uploaded_file)
        if buffer:
            conversions.append(('Word Document', buffer, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
    
    # To PDF conversion
    if ('pdf' in user_lower and 'convert' in user_lower) and not uploaded_file.name.endswith('.pdf'):
        buffer, filename = convert_to_pdf(uploaded_file, file_content)
        if buffer:
            conversions.append(('PDF Document', buffer, filename, 'application/pdf'))
    
    # To CSV conversion
    if 'csv' in user_lower and 'convert' in user_lower:
        buffer, filename = convert_to_csv(uploaded_file, file_content)
        if buffer:
            conversions.append(('CSV File', buffer, filename, 'text/csv'))
    
    # To Excel conversion
    if ('excel' in user_lower or 'xlsx' in user_lower) and 'convert' in user_lower:
        buffer, filename = convert_to_excel(uploaded_file, file_content)
        if buffer:
            conversions.append(('Excel File', buffer, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
    
    # To JSON conversion
    if 'json' in user_lower and 'convert' in user_lower:
        buffer, filename = convert_to_json(uploaded_file, file_content)
        if buffer:
            conversions.append(('JSON File', buffer, filename, 'application/json'))
    
    # To YAML conversion
    if 'yaml' in user_lower and 'convert' in user_lower:
        buffer, filename = convert_to_yaml(uploaded_file, file_content)
        if buffer:
            conversions.append(('YAML File', buffer, filename, 'text/yaml'))
    
    # To Markdown conversion
    if ('markdown' in user_lower or 'md' in user_lower) and 'convert' in user_lower:
        buffer, filename = convert_to_markdown(uploaded_file, file_content)
        if buffer:
            conversions.append(('Markdown File', buffer, filename, 'text/markdown'))
    
    # To HTML conversion
    if 'html' in user_lower and 'convert' in user_lower:
        buffer, filename = convert_to_html(uploaded_file, file_content)
        if buffer:
            conversions.append(('HTML File', buffer, filename, 'text/html'))
    
    # To TXT conversion
    if ('txt' in user_lower or 'text' in user_lower) and 'convert' in user_lower:
        buffer, filename = convert_to_txt(uploaded_file, file_content)
        if buffer:
            conversions.append(('Text File', buffer, filename, 'text/plain'))
    
    # Universal conversion requests (convert to multiple formats)
    if 'convert to all' in user_lower or 'all formats' in user_lower:
        # Convert to most common formats
        formats = [
            ('Word', convert_pdf_to_word if uploaded_file.name.endswith('.pdf') else None),
            ('PDF', convert_to_pdf),
            ('CSV', convert_to_csv),
            ('Excel', convert_to_excel),
            ('JSON', convert_to_json),
            ('Markdown', convert_to_markdown),
            ('HTML', convert_to_html),
            ('Text', convert_to_txt)
        ]
        
        for format_name, convert_func in formats:
            if convert_func:
                try:
                    if format_name == 'Word' and uploaded_file.name.endswith('.pdf'):
                        buffer, filename = convert_func(uploaded_file)
                    else:
                        buffer, filename = convert_func(uploaded_file, file_content)
                    
                    if buffer:
                        mime_types = {
                            'Word': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            'PDF': 'application/pdf',
                            'CSV': 'text/csv',
                            'Excel': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            'JSON': 'application/json',
                            'Markdown': 'text/markdown',
                            'HTML': 'text/html',
                            'Text': 'text/plain'
                        }
                        conversions.append((f'{format_name} File', buffer, filename, mime_types.get(format_name, 'application/octet-stream')))
                except:
                    continue  # Skip formats that can't be converted
    
    return conversions

# Custom CSS for ChatGPT-like styling
st.markdown("""
<style>
.bot-header {
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 2rem;
    gap: 15px;
}

.bot-logo {
    font-size: 3rem;
    background: linear-gradient(45deg, #667eea 0%, #764ba2 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    filter: drop-shadow(2px 2px 4px rgba(0,0,0,0.1));
}

.bot-title {
    font-size: 2.5rem;
    font-weight: 700;
    background: linear-gradient(45deg, #667eea 0%, #764ba2 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin: 0;
}

.attachment-pill {
    background: #e8f4fd;
    border: 1px solid #b3d9ff;
    border-radius: 20px;
    padding: 6px 12px;
    margin-bottom: 8px;
    display: inline-block;
    font-size: 13px;
    color: #0066cc;
    margin-right: 8px;
}

.chat-input-container {
    border: 1px solid #e1e5e9;
    border-radius: 12px;
    padding: 8px;
    background: white;
    margin-bottom: 16px;
}
</style>
""", unsafe_allow_html=True)

# Header with bot logo and title
st.markdown("""
<div class="bot-header">
    <div class="bot-logo">🤖</div>
    <h1 class="bot-title">rAI bot - Universal File Converter</h1>
</div>
""", unsafe_allow_html=True)

# Conversion help section
with st.expander("🔄 Advanced File Converter - Professional Quality", expanded=False):
    st.markdown("""
    **🚀 Professional-Grade File Conversion Platform!**
    
    **✨ High-Quality PDF Conversions (Like iLovePDF):**
    - **PDF → Word**: Preserves layout, formatting, tables, and structure
    - **Advanced OCR**: Multiple extraction methods for best quality
    - **Layout Preservation**: Maintains original document structure
    - **Table Detection**: Converts tables accurately to Word format
    
    **📄 Document Formats:**
    - PDF ↔ Word (DOCX) - **Professional Quality**
    - PDF ↔ Text (TXT)
    - HTML ↔ Markdown
    - Any format → PDF
    
    **📊 Data Formats:**
    - CSV ↔ Excel (XLSX)
    - JSON ↔ YAML
    - JSON ↔ CSV
    - Any format → JSON
    
    **🌐 Web Formats:**
    - Markdown ↔ HTML
    - Any format → HTML
    - HTML → Text
    
    **🎯 Example Commands:**
    - "Convert this PDF to Word" - **Advanced formatting preservation!**
    - "Make this into Excel format"
    - "Convert to JSON"
    - "Transform to Markdown"
    - "Convert to all formats" (bulk conversion)
    - "Export as HTML"
    
    **📊 Multi-Document Analysis:**
    - "Compare all documents" - **Analyze multiple files together!**
    - "Read both files" - **Cross-document insights**
    - "Analyze all my documents" - **Find patterns across files**
    - "What are the differences between all files?"
    - "Summarize all documents together"
    
    **🔧 Technology Stack:**
    - **pdf2docx**: Layout-preserving PDF conversion
    - **PyMuPDF**: Advanced text extraction
    - **pdfplumber**: Table and structure detection
    - **Multiple fallback methods** for best results
    
    **Supported File Types:** PDF, DOCX, CSV, XLSX, JSON, YAML, Markdown, HTML, TXT, Python, JavaScript, CSS, XML
    """)

st.markdown("---")

# ChatGPT-style File Upload (Clean & Minimal) - Multiple Files Supported
uploaded_files = st.file_uploader(
    "📎 Attach document(s)",
    type=['pdf', 'csv', 'txt', 'docx', 'xlsx', 'json', 'py', 'js', 'html', 'css', 'md', 'yaml', 'yml', 'xml'],
    help="Upload one or multiple documents to chat about them",
    accept_multiple_files=True
)

# Handle new file uploads (Multiple files support)
if uploaded_files:
    # Process multiple files
    new_files = {}
    for uploaded_file in uploaded_files:
        file_key = uploaded_file.name
        
        # Store file data permanently to avoid stream consumption issues
        file_data = {
            'file_obj': uploaded_file,
            'file_content': None,  # Will be extracted on first use
            'file_bytes': uploaded_file.getvalue(),  # Store raw bytes
            'file_name': uploaded_file.name,
            'file_type': uploaded_file.type
        }
        new_files[file_key] = file_data
    
    # Replace all existing files with new uploads (ChatGPT-like behavior)
    st.session_state.uploaded_files = new_files
    
    # Set the first file as active by default
    first_file_key = list(new_files.keys())[0]
    st.session_state.active_file = first_file_key
    st.session_state.uploaded_file = new_files[first_file_key]['file_obj']  # Legacy support
    
    # Show confirmation for multiple files
    if len(uploaded_files) == 1:
        st.success(f"📎 Now chatting about: **{uploaded_files[0].name}**")
    else:
        st.success(f"📎 Uploaded **{len(uploaded_files)} documents**. Currently active: **{first_file_key}**")
        st.info(f"💡 Use the document switcher below to switch between your {len(uploaded_files)} files.")

# ChatGPT-style attachment display (minimal pill)
if st.session_state.active_file:
    active_file_data = st.session_state.uploaded_files[st.session_state.active_file]
    file_name = active_file_data['file_name']
    file_size = len(active_file_data['file_bytes']) / 1024
    
    # Clean attachment pill like ChatGPT
    col_pill, col_switch, col_remove = st.columns([6, 1, 1])
    with col_pill:
        st.markdown(f"""
        <div class="attachment-pill">
            � {st.session_state.active_file} ({file_size:.1f} KB)
        </div>
        """, unsafe_allow_html=True)
    
    # Option to switch documents (only if multiple exist)
    if len(st.session_state.uploaded_files) > 1:
        with col_switch:
            if st.button("🔄", help=f"Switch document ({len(st.session_state.uploaded_files)} files)", key="switch_doc"):
                # Show document switcher
                st.session_state.show_switcher = not st.session_state.get('show_switcher', False)
    
    with col_remove:
        if st.button("❌", help="Remove current document", key="remove_current"):
            del st.session_state.uploaded_files[st.session_state.active_file]
            if st.session_state.uploaded_files:
                # Switch to another document
                st.session_state.active_file = list(st.session_state.uploaded_files.keys())[0]
                new_active_data = st.session_state.uploaded_files[st.session_state.active_file]
                st.session_state.uploaded_file = new_active_data['file_obj']
            else:
                # No documents left
                st.session_state.active_file = None
                st.session_state.uploaded_file = None
            st.rerun()

# Document switcher (only shown when requested)
if st.session_state.get('show_switcher', False) and len(st.session_state.uploaded_files) > 1:
    st.markdown("#### � Switch Document")
    
    for file_key in st.session_state.uploaded_files.keys():
        if file_key != st.session_state.active_file:
            col1, col2 = st.columns([5, 1])
            with col1:
                file_obj = st.session_state.uploaded_files[file_key]
                size = len(file_obj.getvalue()) / 1024
                st.write(f"� {file_key} ({size:.1f} KB)")
            with col2:
                if st.button("Switch", key=f"switch_to_{file_key}"):
                    file_data = st.session_state.uploaded_files[file_key]
                    st.session_state.active_file = file_key
                    st.session_state.uploaded_file = file_data['file_obj']
                    st.session_state.show_switcher = False
                    st.rerun()
    
    if st.button("✖️ Close", key="close_switcher"):
        st.session_state.show_switcher = False
        st.rerun()

# Multi-document helper (only shown when multiple files exist)
if len(st.session_state.uploaded_files) > 1:
    st.markdown("#### 🔗 Multi-Document Mode")
    st.info(f"💡 **Quick Tip**: To analyze all {len(st.session_state.uploaded_files)} documents together, use phrases like:")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📊 Compare All Documents", key="compare_all"):
            st.session_state.auto_multi_query = "compare all documents"
            st.rerun()
    with col2:
        if st.button("📖 Read Both Files", key="read_both"):
            st.session_state.auto_multi_query = "read both files together"
            st.rerun()
    with col3:
        if st.button("🔍 Analyze All Together", key="analyze_all"):
            st.session_state.auto_multi_query = "analyze all documents together"
            st.rerun()

st.markdown("---")

# Chat input
user_input = st.chat_input("Ask me anything...")

# Handle auto-generated multi-document queries
if st.session_state.get('auto_multi_query', None):
    user_input = st.session_state.auto_multi_query
    st.session_state.auto_multi_query = None  # Clear after use

# Display ALL chat history in chronological order (oldest to newest)
for i, message in enumerate(st.session_state.chat_history):
    with st.chat_message(message["role"]):
        content = message["content"]
        
        # Clean display for user messages (don't show file content in history)
        if message["role"] == "user" and "Here is the content from the attached file" in content:
            # Extract just the original user question
            parts = content.split("\n\nHere is the content from the attached file")
            display_content = parts[0]
            st.write(display_content)
            # Show attachment indicator
            if len(parts) > 1:
                file_part = parts[1]
                if "'" in file_part:
                    filename = file_part.split("'")[1]
                    st.caption(f"📎 {filename}")
        else:
            # For assistant messages, show as is
            st.write(content)

if user_input:
    # ChatGPT-like behavior: Focus on single active document
    clean_user_message = user_input
    
    # Detect explicit multi-document operations - expanded keywords for better detection
    multi_doc_keywords = [
        'all documents', 'all files', 'every document', 'all my documents', 'each document', 'all docs', 
        'compare documents', 'both documents', 'compare files', 'analyze all', 'read all', 'both files', 
        'all the documents', 'multiple documents', 'multiple files', 'both docs', 'all of them',
        'together', 'compare them', 'read them', 'analyze them', 'all these', 'these documents',
        'these files', 'other document', 'other file', 'second document', 'another document',
        'read both', 'check both', 'look at both', 'examine both', 'review both', 'scan both'
    ]
    is_multi_doc_request = any(keyword in user_input.lower() for keyword in multi_doc_keywords)
    
    # Debug output to help troubleshoot
    if len(st.session_state.uploaded_files) > 1:
        st.sidebar.write(f"🔍 **Multi-Doc Detection Debug:**")
        st.sidebar.write(f"- Available files: {len(st.session_state.uploaded_files)}")
        st.sidebar.write(f"- Your query: '{user_input[:50]}...'")
        st.sidebar.write(f"- Multi-doc detected: {is_multi_doc_request}")
        if is_multi_doc_request:
            matching_keywords = [kw for kw in multi_doc_keywords if kw in user_input.lower()]
            st.sidebar.write(f"- Matched keywords: {matching_keywords}")
        st.sidebar.write(f"- Files: {list(st.session_state.uploaded_files.keys())}")
    
    # Prepare context based on request type (ChatGPT-like intelligence + Multi-document support)
    # Enhanced detection: also check if user mentions "other" when multiple files exist
    implicit_multi_doc = (
        len(st.session_state.uploaded_files) > 1 and 
        ('other' in user_input.lower() or 'another' in user_input.lower() or 'second' in user_input.lower())
    )
    
    if (is_multi_doc_request or implicit_multi_doc) and len(st.session_state.uploaded_files) > 1:
        # Multi-document analysis mode
        st.sidebar.success(f"🔍 **MULTI-DOCUMENT MODE ACTIVATED!**")
        st.sidebar.write(f"Analyzing all {len(st.session_state.uploaded_files)} files together")
        
        all_files_content = []
        total_content_length = 0
        
        for file_key in st.session_state.uploaded_files.keys():
            file_content = get_file_content_safe(file_key)
            clean_content = file_content.replace('\n\n\n', '\n\n').strip()
            
            # Truncate individual files if too long
            if len(clean_content) > 1500:  # Smaller limit per file for multi-doc
                clean_content = clean_content[:1500] + "\n\n[Content truncated...]"
            
            all_files_content.append(f"=== FILE: {file_key} ===\n{clean_content}")
            total_content_length += len(clean_content)
        
        # Combine all files with clear separators
        combined_content = "\n\n" + "\n\n".join(all_files_content)
        
        # Final length check for API limits
        if total_content_length > 4000:
            combined_content = combined_content[:4000] + "\n\n[Combined content truncated for analysis...]"
        
        context_message = f"{user_input} [Analyzing {len(st.session_state.uploaded_files)} documents]\n\nFiles content:{combined_content}"
        attachment_info = f"📎 Analyzing {len(st.session_state.uploaded_files)} documents: {', '.join(st.session_state.uploaded_files.keys())}"
        
        # Set the first file as active for legacy support
        first_file = list(st.session_state.uploaded_files.values())[0]
        st.session_state.uploaded_file = first_file['file_obj']
        
    elif st.session_state.active_file and st.session_state.active_file in st.session_state.uploaded_files:
        # Single document mode (original behavior)
        file_content = get_file_content_safe(st.session_state.active_file)
        clean_content = file_content.replace('\n\n\n', '\n\n').strip()
        
        if len(clean_content) > 3000:
            clean_content = clean_content[:3000] + "\n\n[Content truncated...]"
        
        # Determine request type for context optimization
        conversion_keywords = ['convert', 'change', 'transform', 'save as', 'export', 'make into']
        is_conversion_request = any(keyword in user_input.lower() for keyword in conversion_keywords)
        
        if is_conversion_request:
            context_message = f"{user_input} [File: {st.session_state.active_file}]\n\nFile content (first 1000 chars): {clean_content[:1000]}..."
        else:
            context_message = f"{user_input} [File: {st.session_state.active_file}]\n\nFile content: {clean_content}"
        
        attachment_info = f"📎 {st.session_state.active_file}"
        
        # Update legacy support
        active_file_data = st.session_state.uploaded_files[st.session_state.active_file]
        st.session_state.uploaded_file = active_file_data['file_obj']
        
    else:
        # No documents - regular chat
        context_message = user_input
        attachment_info = None
    
    # Add to conversation context and history
    st.session_state.conversation_context.append({"role": "user", "content": context_message})
    st.session_state.chat_history.append({"role": "user", "content": clean_user_message})
    
    # Display user message (ChatGPT style)
    with st.chat_message("user"):
        st.write(user_input)
        if attachment_info:
            st.caption(attachment_info)
    
    # Get bot response
    with st.chat_message("assistant"):
        conversion_keywords = ['convert', 'change', 'transform', 'save as', 'export', 'make into']
        is_conversion_request = any(keyword in user_input.lower() for keyword in conversion_keywords)
        
        if is_conversion_request and st.session_state.uploaded_file is not None and not (is_multi_doc_request or implicit_multi_doc):
            # Single document conversion (ChatGPT behavior)
            conversion_info = st.empty()
            
            with st.spinner(f"Converting {st.session_state.uploaded_file.name}..."):
                file_content = extract_file_content(st.session_state.uploaded_file)
                
                # Special handling for PDF to Word conversion
                if ('pdf' in user_input.lower() and ('word' in user_input.lower() or 'docx' in user_input.lower()) and 
                    st.session_state.uploaded_file.name.endswith('.pdf')):
                    
                    conversion_info.info("🔄 Using advanced PDF conversion (preserving layout, tables, and formatting)...")
                    
                    # Reset file pointer for conversion
                    st.session_state.uploaded_file.seek(0)
                    buffer, filename = convert_pdf_to_word(st.session_state.uploaded_file)
                    
                    if buffer:
                        conversion_info.empty()
                        st.success("🎉 High-quality PDF to Word conversion completed!")
                        
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.write(f"📄 Word Document: `{filename}`")
                            st.caption("✨ Preserves formatting, tables, and layout structure")
                        with col2:
                            st.download_button(
                                label="⬇️ Download",
                                data=buffer.getvalue(),
                                file_name=filename,
                                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                key=f"download_{filename}"
                            )
                        
                        response = f"I've converted your PDF '{st.session_state.uploaded_file.name}' to a high-quality Word document with preserved formatting, tables, and layout structure - just like professional conversion services!"
                    else:
                        conversion_info.error("❌ PDF conversion failed")
                        response = f"I encountered an error converting your PDF: {filename}"
                
                else:
                    # Regular conversions
                    conversion_info.info("🔄 Processing conversion...")
                    conversions = perform_file_conversion(user_input, st.session_state.uploaded_file, file_content)
                    
                    if conversions:
                        conversion_info.empty()
                        st.success("🎉 Conversion completed! Download your files below:")
                        
                        for conv_type, buffer, filename, mime_type in conversions:
                            col1, col2 = st.columns([3, 1])
                            with col1:
                                st.write(f"📄 {conv_type}: `{filename}`")
                            with col2:
                                st.download_button(
                                    label="⬇️ Download",
                                    data=buffer.getvalue() if hasattr(buffer, 'getvalue') else buffer.read(),
                                    file_name=filename,
                                    mime=mime_type,
                                    key=f"download_{filename}"
                                )
                        
                        response = f"I've successfully converted your file '{st.session_state.uploaded_file.name}' to the requested format(s). You can download the converted file(s) using the download buttons above."
                    else:
                        conversion_info.error("❌ Conversion failed")
                        response = get_chat_response(st.session_state.conversation_context)
        else:
            # Regular AI response (not a conversion request) - handles both single and multi-document
            if (is_multi_doc_request or implicit_multi_doc) and len(st.session_state.uploaded_files) > 1:
                with st.spinner(f"Analyzing {len(st.session_state.uploaded_files)} documents..."):
                    response = get_chat_response(st.session_state.conversation_context)
            elif st.session_state.active_file:
                with st.spinner(f"Reading {st.session_state.active_file} and thinking..."):
                    response = get_chat_response(st.session_state.conversation_context)
            else:
                with st.spinner("Thinking..."):
                    response = get_chat_response(st.session_state.conversation_context)
        
        # Add response to conversation context
        st.session_state.conversation_context.append({"role": "assistant", "content": response})
        
        # Display response
        st.write(response)
    
    # Add to display history
    st.session_state.chat_history.append({"role": "assistant", "content": response})

# Clear chat button (ChatGPT-like)
if st.session_state.chat_history:
    if st.button("🗑️ Clear Conversation", key="clear_chat"):
        st.session_state.chat_history = []
        st.session_state.conversation_context = []
        # Keep documents but clear conversation context
        st.rerun()
