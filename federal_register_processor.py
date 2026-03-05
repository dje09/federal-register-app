"""
Federal Register Document Processing System
A comprehensive solution for extracting and managing Federal Register documents
"""

import os
import re
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pdfplumber
from dataclasses import dataclass, asdict
import json

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('fr_processing.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class FederalRegisterDocument:
    """Data structure for Federal Register document information"""
    upload_number: str
    title_number: Optional[str] = None
    volume: Optional[str] = None
    fr_date: Optional[str] = None
    fr_doc_number: Optional[str] = None
    processing_type: Optional[str] = None
    cfr_title: Optional[str] = None
    docket_number: Optional[str] = None
    agency: Optional[str] = None
    action: Optional[str] = None
    summary: Optional[str] = None
    effective_date: Optional[str] = None
    errors: List[str] = None

    def __post_init__(self):
        if self.errors is None:
            self.errors = []


class VolumeMapper:
    """Handles volume mapping from Excel file"""
    
    def __init__(self, excel_path: str):
        """
        Initialize volume mapper with Excel file
        
        Args:
            excel_path: Path to the Excel file containing volume mappings
        """
        self.excel_path = excel_path
        self.volume_data = None
        self.load_excel()
    
    def load_excel(self):
        """Load and parse Excel file with volume mappings"""
        try:
            self.volume_data = pd.read_excel(self.excel_path, sheet_name='DOV')
            logger.info(f"Successfully loaded Excel file: {self.excel_path}")
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise
    
    def get_volume(self, title: str, section: str) -> Optional[str]:
        """
        Get volume number based on title and section
        
        Args:
            title: CFR Title number
            section: Section number or range
            
        Returns:
            Volume number as string or None if not found
        """
        try:
            if self.volume_data is None:
                return None
            
            # Convert title to integer for comparison
            title_num = int(title)
            
            # Filter by title
            title_rows = self.volume_data[self.volume_data['Title'] == title_num]
            
            if title_rows.empty:
                logger.warning(f"No volume found for Title {title}")
                return None
            
            # Parse section to find matching volume
            for _, row in title_rows.iterrows():
                sections_range = str(row['Sections'])
                if self._section_in_range(section, sections_range):
                    return str(row['Volume'])
            
            logger.warning(f"No volume found for Title {title}, Section {section}")
            return None
            
        except Exception as e:
            logger.error(f"Error getting volume: {str(e)}")
            return None
    
    def _section_in_range(self, section: str, range_str: str) -> bool:
        """
        Check if section falls within a range
        
        Args:
            section: Section number to check
            range_str: Range string from Excel (e.g., "1-199", "All")
            
        Returns:
            True if section is in range, False otherwise
        """
        try:
            if range_str.lower() == 'all':
                return True
            
            # Extract numeric part from section
            section_num = int(re.findall(r'\d+', str(section))[0])
            
            # Parse range
            if '-' in range_str:
                parts = range_str.split('-')
                start = int(re.findall(r'\d+', parts[0])[0])
                end_match = re.findall(r'\d+', parts[1])
                if end_match:
                    end = int(end_match[0])
                    return start <= section_num <= end
            
            return False
            
        except Exception as e:
            logger.debug(f"Error parsing section range: {str(e)}")
            return False


class PDFProcessor:
    """Handles PDF parsing and data extraction"""
    
    # Processing type patterns
    PROCESSING_TYPES = {
        'Final Rule': [
            r'Rules and Regulations',
            r'Final Rule',
            r'ACTION:\s*Final rule'
        ],
        'Proposed Rule': [
            r'Proposed Rules',
            r'Proposed Rule',
            r'ACTION:\s*Proposed rule',
            r'Notice of proposed rulemaking'
        ],
        'Notice': [
            r'Notices',
            r'ACTION:\s*Notice'
        ],
        'Interim Final Rule': [
            r'Interim Final Rule',
            r'ACTION:\s*Interim final rule'
        ]
    }
    
    def __init__(self, volume_mapper: VolumeMapper):
        """
        Initialize PDF processor
        
        Args:
            volume_mapper: VolumeMapper instance for volume lookups
        """
        self.volume_mapper = volume_mapper
    
    def process_pdf(self, pdf_path: str, upload_number: str) -> FederalRegisterDocument:
        """
        Process a Federal Register PDF file
        
        Args:
            pdf_path: Path to PDF file
            upload_number: Manual upload number entry
            
        Returns:
            FederalRegisterDocument with extracted data
        """
        doc = FederalRegisterDocument(upload_number=upload_number)
        doc.fr_date = datetime.now().strftime('%Y-%m-%d')
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Extract text from first few pages (header info usually on first 2 pages)
                text = ""
                for i, page in enumerate(pdf.pages[:3]):
                    text += page.extract_text() or ""
                    if i < 2:  # Add page separator
                        text += "\n---PAGE_BREAK---\n"
                
                # Extract all fields
                doc.title_number = self._extract_title_number(text)
                doc.fr_doc_number = self._extract_fr_doc_number(text)
                doc.cfr_title = self._extract_cfr_title(text)
                doc.docket_number = self._extract_docket_number(text)
                doc.agency = self._extract_agency(text)
                doc.action = self._extract_action(text)
                doc.summary = self._extract_summary(text)
                doc.effective_date = self._extract_effective_date(text)
                doc.processing_type = self._determine_processing_type(text)
                
                # Get volume based on CFR title and section
                if doc.cfr_title:
                    section = self._extract_section_number(text)
                    if section:
                        doc.volume = self.volume_mapper.get_volume(doc.cfr_title, section)
                
                # Validate extracted data
                self._validate_document(doc)
                
                logger.info(f"Successfully processed PDF: {pdf_path}")
                
        except Exception as e:
            error_msg = f"Error processing PDF {pdf_path}: {str(e)}"
            logger.error(error_msg)
            doc.errors.append(error_msg)
        
        return doc
    
    def _extract_title_number(self, text: str) -> Optional[str]:
        """Extract title from Federal Register header"""
        patterns = [
            r'Federal Register\s*/\s*Vol\.\s*\d+,\s*No\.\s*\d+\s*/\s*\w+,?\s*\w+\s+\d+,\s*\d+\s*/\s*(.+?)(?:\n|\s{2,})',
            r'/\s*Rules and Regulations\s*$',
            r'/\s*Proposed Rules\s*$',
            r'/\s*Notices\s*$'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
            if match:
                if len(match.groups()) > 0:
                    return match.group(1).strip()
                else:
                    # Extract from context
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        if 'Federal Register' in line and i + 1 < len(lines):
                            # Look for section title after FR header
                            for j in range(i, min(i + 5, len(lines))):
                                if any(keyword in lines[j] for keyword in ['Rules', 'Notices', 'Proposed']):
                                    return lines[j].strip()
        
        return None
    
    def _extract_fr_doc_number(self, text: str) -> Optional[str]:
        """Extract FR Doc number"""
        patterns = [
            r'\[FR\s*Doc\.?\s*(\d{4}[-–]\d{4,6})',
            r'FR\s*Doc\.?\s*(\d{4}[-–]\d{4,6})',
            r'BILLING\s*CODE.*?\n.*?(\d{4}[-–]\d{4,6})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).replace('–', '-')  # Normalize dash
        
        return None
    
    def _extract_cfr_title(self, text: str) -> Optional[str]:
        """Extract CFR Title number"""
        patterns = [
            r'(\d+)\s*CFR\s*Part',
            r'Title\s*(\d+)',
            r'(\d+)\s*CFR\s*§',
            r'CFR\s*Title\s*(\d+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        
        return None
    
    def _extract_section_number(self, text: str) -> Optional[str]:
        """Extract section number for volume mapping"""
        patterns = [
            r'CFR\s*Part\s*(\d+)',
            r'Part\s*(\d+)',
            r'§\s*(\d+\.\d+)',
            r'Section\s*(\d+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)
        
        return None
    
    def _extract_docket_number(self, text: str) -> Optional[str]:
        """Extract docket number"""
        patterns = [
            r'Docket\s*(?:Number|No\.?|#)\s*([A-Z0-9\-]+)',
            r'\[Docket\s*([A-Z0-9\-]+)\]',
            r'DOCKET\s*(?:NUMBER|NO\.?)[:;]?\s*([A-Z0-9\-]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return None
    
    def _extract_agency(self, text: str) -> Optional[str]:
        """Extract agency name"""
        patterns = [
            r'AGENCY:\s*(.+?)(?:\n|ACTION:)',
            r'DEPARTMENT\s*OF\s*(.+?)(?:\n)',
            r'^([A-Z\s,]+(?:DEPARTMENT|AGENCY|COMMISSION|ADMINISTRATION))',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
            if match:
                agency = match.group(1).strip()
                # Clean up agency name
                agency = re.sub(r'\s+', ' ', agency)
                return agency[:200]  # Limit length
        
        return None
    
    def _extract_action(self, text: str) -> Optional[str]:
        """Extract action type"""
        pattern = r'ACTION:\s*(.+?)(?:\n|SUMMARY:)'
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            action = match.group(1).strip()
            action = re.sub(r'\s+', ' ', action)
            return action[:200]
        
        return None
    
    def _extract_summary(self, text: str) -> Optional[str]:
        """Extract summary section"""
        pattern = r'SUMMARY:\s*(.+?)(?:\n(?:DATES|FOR\s+FURTHER|EFFECTIVE|ADDRESSES):)'
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            summary = match.group(1).strip()
            summary = re.sub(r'\s+', ' ', summary)
            return summary[:500]  # Limit length
        
        return None
    
    def _extract_effective_date(self, text: str) -> Optional[str]:
        """Extract effective date"""
        patterns = [
            r'DATES?:\s*(?:Effective|This rule is effective)\s*(?:on\s*)?(\w+\s+\d+,\s*\d{4})',
            r'Effective\s*(?:Date|on)[:;]?\s*(\w+\s+\d+,\s*\d{4})',
            r'effective\s*(?:on\s*)?(\d{1,2}/\d{1,2}/\d{4})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # Try to parse and normalize date
                try:
                    if '/' in date_str:
                        date_obj = datetime.strptime(date_str, '%m/%d/%Y')
                    else:
                        date_obj = datetime.strptime(date_str, '%B %d, %Y')
                    return date_obj.strftime('%Y-%m-%d')
                except:
                    return date_str
        
        return None
    
    def _determine_processing_type(self, text: str) -> Optional[str]:
        """Determine processing type based on content"""
        for proc_type, patterns in self.PROCESSING_TYPES.items():
            for pattern in patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    return proc_type
        
        return None
    
    def _validate_document(self, doc: FederalRegisterDocument):
        """Validate extracted document data"""
        if not doc.fr_doc_number:
            doc.errors.append("FR Doc Number not found")
        
        if not doc.processing_type:
            doc.errors.append("Processing type could not be determined")
        
        if not doc.cfr_title:
            doc.errors.append("CFR Title not found")
        
        if not doc.agency:
            doc.errors.append("Agency not found")


class FederalRegisterApp:
    """Flask application for Federal Register processing"""
    
    def __init__(self, excel_path: str, upload_folder: str = 'uploads', output_folder: str = 'output'):
        """
        Initialize Flask application
        
        Args:
            excel_path: Path to Excel file with volume mappings
            upload_folder: Folder for uploaded PDFs
            output_folder: Folder for output files
        """
        self.app = Flask(__name__)
        self.app.config['UPLOAD_FOLDER'] = upload_folder
        self.app.config['OUTPUT_FOLDER'] = output_folder
        self.app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
        
        # Create folders if they don't exist
        Path(upload_folder).mkdir(exist_ok=True)
        Path(output_folder).mkdir(exist_ok=True)
        
        # Initialize processors
        self.volume_mapper = VolumeMapper(excel_path)
        self.pdf_processor = PDFProcessor(self.volume_mapper)
        
        # Storage for processed documents
        self.processed_documents: List[FederalRegisterDocument] = []
        
        # Setup routes
        self._setup_routes()
    
    def _setup_routes(self):
        """Setup Flask routes"""
        
        @self.app.route('/')
        def index():
            """Main page"""
            return render_template('index.html')
        
        @self.app.route('/upload', methods=['POST'])
        def upload_file():
            """Handle file upload and processing"""
            try:
                # Check if file is present
                if 'file' not in request.files:
                    return jsonify({'error': 'No file provided'}), 400
                
                file = request.files['file']
                upload_number = request.form.get('upload_number', '')
                
                if file.filename == '':
                    return jsonify({'error': 'No file selected'}), 400
                
                if not upload_number:
                    return jsonify({'error': 'Upload number is required'}), 400
                
                # Validate file type
                if not file.filename.lower().endswith('.pdf'):
                    return jsonify({'error': 'Only PDF files are allowed'}), 400
                
                # Save file
                filename = secure_filename(file.filename)
                filepath = os.path.join(self.app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                # Process PDF
                doc = self.pdf_processor.process_pdf(filepath, upload_number)
                self.processed_documents.append(doc)
                
                # Clean up uploaded file
                os.remove(filepath)
                
                return jsonify({
                    'success': True,
                    'data': asdict(doc),
                    'message': 'File processed successfully'
                })
                
            except Exception as e:
                logger.error(f"Error in upload route: {str(e)}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/batch-upload', methods=['POST'])
        def batch_upload():
            """Handle batch file upload"""
            try:
                files = request.files.getlist('files')
                upload_numbers = request.form.getlist('upload_numbers')
                
                if not files or not upload_numbers:
                    return jsonify({'error': 'Files and upload numbers required'}), 400
                
                if len(files) != len(upload_numbers):
                    return jsonify({'error': 'Number of files must match upload numbers'}), 400
                
                results = []
                
                for file, upload_number in zip(files, upload_numbers):
                    try:
                        filename = secure_filename(file.filename)
                        filepath = os.path.join(self.app.config['UPLOAD_FOLDER'], filename)
                        file.save(filepath)
                        
                        doc = self.pdf_processor.process_pdf(filepath, upload_number)
                        self.processed_documents.append(doc)
                        
                        results.append({
                            'filename': filename,
                            'success': True,
                            'data': asdict(doc)
                        })
                        
                        os.remove(filepath)
                        
                    except Exception as e:
                        results.append({
                            'filename': file.filename,
                            'success': False,
                            'error': str(e)
                        })
                
                return jsonify({
                    'success': True,
                    'results': results
                })
                
            except Exception as e:
                logger.error(f"Error in batch upload: {str(e)}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/documents', methods=['GET'])
        def get_documents():
            """Get all processed documents"""
            return jsonify([asdict(doc) for doc in self.processed_documents])
        
        @self.app.route('/export', methods=['POST'])
        def export_data():
            """Export processed documents to Excel"""
            try:
                if not self.processed_documents:
                    return jsonify({'error': 'No documents to export'}), 400
                
                # Convert to DataFrame
                data = [asdict(doc) for doc in self.processed_documents]
                df = pd.DataFrame(data)
                
                # Convert errors list to string
                df['errors'] = df['errors'].apply(lambda x: '; '.join(x) if x else '')
                
                # Generate filename
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'fr_documents_{timestamp}.xlsx'
                filepath = os.path.join(self.app.config['OUTPUT_FOLDER'], filename)
                
                # Export to Excel
                df.to_excel(filepath, index=False, engine='openpyxl')
                
                return send_file(filepath, as_attachment=True, download_name=filename)
                
            except Exception as e:
                logger.error(f"Error exporting data: {str(e)}")
                return jsonify({'error': str(e)}), 500
        
        @self.app.route('/clear', methods=['POST'])
        def clear_documents():
            """Clear all processed documents"""
            self.processed_documents.clear()
            return jsonify({'success': True, 'message': 'Documents cleared'})
    
    def run(self, debug=False, host='0.0.0.0', port=5000):
        """Run the Flask application"""
        self.app.run(debug=debug, host=host, port=port)


# HTML Template (save as templates/index.html)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Federal Register Document Processor</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2em;
            margin-bottom: 10px;
        }
        
        .tabs {
            display: flex;
            background: #f5f5f5;
            border-bottom: 2px solid #ddd;
        }
        
        .tab {
            flex: 1;
            padding: 15px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            border: none;
            background: transparent;
            font-size: 16px;
        }
        
        .tab:hover {
            background: #e0e0e0;
        }
        
        .tab.active {
            background: white;
            border-bottom: 3px solid #667eea;
            font-weight: bold;
        }
        
        .tab-content {
            display: none;
            padding: 30px;
        }
        
        .tab-content.active {
            display: block;
        }
        
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            margin-bottom: 20px;
            transition: all 0.3s;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #764ba2;
            background: #f9f9f9;
        }
        
        .upload-area.dragover {
            background: #e3e8ff;
            border-color: #764ba2;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
        }
        
        input[type="text"],
        input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 2px solid #ddd;
            border-radius: 5px;
            font-size: 16px;
        }
        
        input[type="file"] {
            cursor: pointer;
        }
        
        .btn {
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s;
            margin-right: 10px;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .btn-secondary {
            background: #6c757d;
            color: white;
        }
        
        .btn-success {
            background: #28a745;
            color: white;
        }
        
        .btn-danger {
            background: #dc3545;
            color: white;
        }
        
        .progress-container {
            display: none;
            margin: 20px 0;
        }
        
        .progress-bar {
            width: 100%;
            height: 30px;
            background: #f0f0f0;
            border-radius: 15px;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            transition: width 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
        }
        
        .documents-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        
        .documents-table th,
        .documents-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        
        .documents-table th {
            background: #f5f5f5;
            font-weight: bold;
        }
        
        .documents-table tr:hover {
            background: #f9f9f9;
        }
        
        .status-badge {
            padding: 5px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
        }
        
        .status-success {
            background: #d4edda;
            color: #155724;
        }
        
        .status-error {
            background: #f8d7da;
            color: #721c24;
        }
        
        .alert {
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 5px;
        }
        
        .alert-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .alert-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .batch-item {
            background: #f9f9f9;
            padding: 15px;
            margin-bottom: 10px;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        
        .batch-item-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }
        
        .remove-btn {
            background: #dc3545;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 3px;
            cursor: pointer;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 Federal Register Document Processor</h1>
            <p>Automated PDF Processing & Data Extraction System</p>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="switchTab('single')">Single Upload</button>
            <button class="tab" onclick="switchTab('batch')">Batch Upload</button>
            <button class="tab" onclick="switchTab('documents')">Processed Documents</button>
        </div>
        
        <!-- Single Upload Tab -->
        <div id="single-tab" class="tab-content active">
            <h2>Single Document Upload</h2>
            <div id="single-alert"></div>
            
            <div class="form-group">
                <label for="upload-number">Upload Number *</label>
                <input type="text" id="upload-number" placeholder="Enter upload number" required>
            </div>
            
            <div class="upload-area" id="upload-area" onclick="document.getElementById('file-input').click()">
                <h3>📁 Drop PDF file here or click to browse</h3>
                <p>Maximum file size: 50MB</p>
                <input type="file" id="file-input" accept=".pdf" style="display: none" onchange="handleFileSelect(event)">
            </div>
            
            <div id="file-info" style="display: none; margin-bottom: 20px;">
                <strong>Selected file:</strong> <span id="file-name"></span>
            </div>
            
            <div class="progress-container" id="progress-container">
                <div class="progress-bar">
                    <div class="progress-fill" id="progress-fill">0%</div>
                </div>
            </div>
            
            <button class="btn btn-primary" onclick="uploadFile()">Process Document</button>
            <button class="btn btn-secondary" onclick="resetForm()">Reset</button>
        </div>
        
        <!-- Batch Upload Tab -->
        <div id="batch-tab" class="tab-content">
            <h2>Batch Document Upload</h2>
            <div id="batch-alert"></div>
            
            <div class="upload-area" onclick="document.getElementById('batch-files').click()">
                <h3>📁 Select multiple PDF files</h3>
                <p>You can upload multiple files at once</p>
                <input type="file" id="batch-files" accept=".pdf" multiple style="display: none" onchange="handleBatchFiles(event)">
            </div>
            
            <div id="batch-items"></div>
            
            <div class="progress-container" id="batch-progress-container">
                <div class="progress-bar">
                    <div class="progress-fill" id="batch-progress-fill">0%</div>
                </div>
            </div>
            
            <button class="btn btn-primary" onclick="uploadBatch()">Process All Documents</button>
            <button class="btn btn-secondary" onclick="clearBatch()">Clear All</button>
        </div>
        
        <!-- Documents Tab -->
        <div id="documents-tab" class="tab-content">
            <h2>Processed Documents</h2>
            <div style="margin-bottom: 20px;">
                <button class="btn btn-success" onclick="exportDocuments()">📥 Export to Excel</button>
                <button class="btn btn-danger" onclick="clearDocuments()">🗑️ Clear All</button>
                <button class="btn btn-secondary" onclick="refreshDocuments()">🔄 Refresh</button>
            </div>
            
            <div id="documents-container">
                <p>No documents processed yet.</p>
            </div>
        </div>
    </div>
    
    <script>
        let selectedFile = null;
        let batchFiles = [];
        
        function switchTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });
            document.querySelectorAll('.tab').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            event.target.classList.add('active');
            
            // Refresh documents if switching to documents tab
            if (tabName === 'documents') {
                refreshDocuments();
            }
        }
        
        // Drag and drop functionality
        const uploadArea = document.getElementById('upload-area');
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.add('dragover');
            }, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.remove('dragover');
            }, false);
        });
        
        uploadArea.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                selectedFile = files[0];
                displayFileInfo();
            }
        });
        
        function handleFileSelect(event) {
            selectedFile = event.target.files[0];
            displayFileInfo();
        }
        
        function displayFileInfo() {
            if (selectedFile) {
                document.getElementById('file-name').textContent = selectedFile.name;
                document.getElementById('file-info').style.display = 'block';
            }
        }
        
        function showAlert(containerId, message, type) {
            const alertDiv = document.getElementById(containerId);
            alertDiv.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
            setTimeout(() => {
                alertDiv.innerHTML = '';
            }, 5000);
        }
        
        async function uploadFile() {
            const uploadNumber = document.getElementById('upload-number').value.trim();
            
            if (!uploadNumber) {
                showAlert('single-alert', 'Please enter an upload number', 'error');
                return;
            }
            
            if (!selectedFile) {
                showAlert('single-alert', 'Please select a PDF file', 'error');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('upload_number', uploadNumber);
            
            // Show progress
            document.getElementById('progress-container').style.display = 'block';
            document.getElementById('progress-fill').style.width = '50%';
            document.getElementById('progress-fill').textContent = 'Processing...';
            
            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                document.getElementById('progress-fill').style.width = '100%';
                document.getElementById('progress-fill').textContent = '100%';
                
                if (result.success) {
                    showAlert('single-alert', '✅ Document processed successfully!', 'success');
                    resetForm();
                } else {
                    showAlert('single-alert', '❌ Error: ' + result.error, 'error');
                }
            } catch (error) {
                showAlert('single-alert', '❌ Error: ' + error.message, 'error');
            } finally {
                setTimeout(() => {
                    document.getElementById('progress-container').style.display = 'none';
                    document.getElementById('progress-fill').style.width = '0%';
                }, 2000);
            }
        }
        
        function resetForm() {
            document.getElementById('upload-number').value = '';
            document.getElementById('file-input').value = '';
            document.getElementById('file-info').style.display = 'none';
            selectedFile = null;
        }
        
        function handleBatchFiles(event) {
            const files = Array.from(event.target.files);
            batchFiles = files.map(file => ({
                file: file,
                uploadNumber: ''
            }));
            displayBatchItems();
        }
        
        function displayBatchItems() {
            const container = document.getElementById('batch-items');
            container.innerHTML = '';
            
            batchFiles.forEach((item, index) => {
                const div = document.createElement('div');
                div.className = 'batch-item';
                div.innerHTML = `
                    <div class="batch-item-header">
                        <strong>${item.file.name}</strong>
                        <button class="remove-btn" onclick="removeBatchItem(${index})">Remove</button>
                    </div>
                    <input type="text" placeholder="Enter upload number" 
                           onchange="updateBatchUploadNumber(${index}, this.value)"
                           value="${item.uploadNumber}">
                `;
                container.appendChild(div);
            });
        }
        
        function updateBatchUploadNumber(index, value) {
            batchFiles[index].uploadNumber = value;
        }
        
        function removeBatchItem(index) {
            batchFiles.splice(index, 1);
            displayBatchItems();
        }
        
        function clearBatch() {
            batchFiles = [];
            document.getElementById('batch-files').value = '';
            document.getElementById('batch-items').innerHTML = '';
        }
        
        async function uploadBatch() {
            if (batchFiles.length === 0) {
                showAlert('batch-alert', 'Please select files to upload', 'error');
                return;
            }
            
            // Validate all have upload numbers
            const missingNumbers = batchFiles.filter(item => !item.uploadNumber);
            if (missingNumbers.length > 0) {
                showAlert('batch-alert', 'Please enter upload numbers for all files', 'error');
                return;
            }
            
            const formData = new FormData();
            batchFiles.forEach(item => {
                formData.append('files', item.file);
                formData.append('upload_numbers', item.uploadNumber);
            });
            
            // Show progress
            document.getElementById('batch-progress-container').style.display = 'block';
            document.getElementById('batch-progress-fill').style.width = '50%';
            document.getElementById('batch-progress-fill').textContent = 'Processing...';
            
            try {
                const response = await fetch('/batch-upload', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                document.getElementById('batch-progress-fill').style.width = '100%';
                document.getElementById('batch-progress-fill').textContent = '100%';
                
                if (result.success) {
                    const successCount = result.results.filter(r => r.success).length;
                    showAlert('batch-alert', `✅ Processed ${successCount} of ${batchFiles.length} documents successfully!`, 'success');
                    clearBatch();
                } else {
                    showAlert('batch-alert', '❌ Error: ' + result.error, 'error');
                }
            } catch (error) {
                showAlert('batch-alert', '❌ Error: ' + error.message, 'error');
            } finally {
                setTimeout(() => {
                    document.getElementById('batch-progress-container').style.display = 'none';
                    document.getElementById('batch-progress-fill').style.width = '0%';
                }, 2000);
            }
        }
        
        async function refreshDocuments() {
            try {
                const response = await fetch('/documents');
                const documents = await response.json();
                
                const container = document.getElementById('documents-container');
                
                if (documents.length === 0) {
                    container.innerHTML = '<p>No documents processed yet.</p>';
                    return;
                }
                
                let html = '<table class="documents-table"><thead><tr>';
                html += '<th>Upload #</th><th>FR Doc #</th><th>CFR Title</th>';
                html += '<th>Volume</th><th>Agency</th><th>Processing Type</th>';
                html += '<th>FR Date</th><th>Status</th></tr></thead><tbody>';
                
                documents.forEach(doc => {
                    const hasErrors = doc.errors && doc.errors.length > 0;
                    const statusClass = hasErrors ? 'status-error' : 'status-success';
                    const statusText = hasErrors ? 'Has Issues' : 'OK';
                    
                    html += '<tr>';
                    html += `<td>${doc.upload_number || '-'}</td>`;
                    html += `<td>${doc.fr_doc_number || '-'}</td>`;
                    html += `<td>${doc.cfr_title || '-'}</td>`;
                    html += `<td>${doc.volume || '-'}</td>`;
                    html += `<td>${doc.agency ? doc.agency.substring(0, 30) + '...' : '-'}</td>`;
                    html += `<td>${doc.processing_type || '-'}</td>`;
                    html += `<td>${doc.fr_date || '-'}</td>`;
                    html += `<td><span class="status-badge ${statusClass}">${statusText}</span></td>`;
                    html += '</tr>';
                });
                
                html += '</tbody></table>';
                container.innerHTML = html;
                
            } catch (error) {
                console.error('Error refreshing documents:', error);
            }
        }
        
        async function exportDocuments() {
            try {
                const response = await fetch('/export', {
                    method: 'POST'
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `fr_documents_${new Date().getTime()}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    a.remove();
                } else {
                    alert('Error exporting documents');
                }
            } catch (error) {
                alert('Error: ' + error.message);
            }
        }
        
        async function clearDocuments() {
            if (confirm('Are you sure you want to clear all processed documents?')) {
                try {
                    await fetch('/clear', { method: 'POST' });
                    refreshDocuments();
                } catch (error) {
                    alert('Error: ' + error.message);
                }
            }
        }
        
        // Auto-refresh documents every 30 seconds when on documents tab
        setInterval(() => {
            if (document.getElementById('documents-tab').classList.contains('active')) {
                refreshDocuments();
            }
        }, 30000);
    </script>
</body>
</html>
'''


# Main execution
if __name__ == '__main__':
    # Configuration
    EXCEL_PATH = 'List of Volumes.xlsx'  # Path to your Excel file
    UPLOAD_FOLDER = 'uploads'
    OUTPUT_FOLDER = 'output'
    
    # Create templates folder and save HTML
    templates_dir = Path('templates')
    templates_dir.mkdir(exist_ok=True)
    
    with open(templates_dir / 'index.html', 'w', encoding='utf-8') as f:
        f.write(HTML_TEMPLATE)
    
    # Initialize and run application
    print("Starting Federal Register Document Processor...")
    print(f"Excel file: {EXCEL_PATH}")
    print(f"Upload folder: {UPLOAD_FOLDER}")
    print(f"Output folder: {OUTPUT_FOLDER}")
    
    app = FederalRegisterApp(
        excel_path=EXCEL_PATH,
        upload_folder=UPLOAD_FOLDER,
        output_folder=OUTPUT_FOLDER
    )
    
    print("\n" + "="*60)
    print("Application ready!")
    print("Open your browser and navigate to: http://localhost:5000")
    print("="*60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)