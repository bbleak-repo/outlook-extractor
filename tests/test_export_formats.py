"""Tests for export format functionality."""
import json
import os
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any

import pandas as pd

from outlook_extractor.export.json_exporter import JSONExporter
from outlook_extractor.export.csv_exporter import CSVExporter
from outlook_extractor.export.excel_exporter import ExcelExporter
from outlook_extractor.export.pdf_exporter import PDFExporter
from outlook_extractor import constants


def create_mock_email(index: int) -> Dict[str, Any]:
    """Create a mock email dictionary for testing."""
    base_time = datetime.now() - timedelta(days=index)
    return {
        'id': f'email_{index}',
        'conversation_id': f'conv_{index // 2}',
        'subject': f'Test Email {index}',
        'sender': f'sender{index}@example.com',
        'to_recipients': [f'recipient{i}@example.com' for i in range(2)],
        'cc_recipients': [],
        'bcc_recipients': [],
        'received_date': (base_time - timedelta(hours=1)).isoformat(),
        'sent_date': base_time.isoformat(),
        'importance': 'Normal',
        'is_read': index % 2 == 0,
        'has_attachments': index % 3 == 0,
        'body': f'This is the body of test email {index}.\nIt contains multiple lines.\n\nAnd paragraphs.',
        'categories': ['Test', f'Category{index % 3}'],
        'size': 1024 * (index + 1),
        'is_draft': False,
        'is_encrypted': False,
        'is_signed': index % 4 == 0,
        'internet_message_headers': {
            'X-Test-Header': f'test-value-{index}',
            'Received': f'from server{index}.example.com (1.2.3.{index}) by mail.example.com'
        },
        'attachments': [
            {
                'id': f'attach_{index}_1',
                'name': f'document_{index}.pdf',
                'content_type': 'application/pdf',
                'size': 1024 * (index + 1),
                'is_inline': False
            }
        ] if index % 3 == 0 else []
    }


class TestExportFormats(unittest.TestCase):
    """Test export format functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = tempfile.mkdtemp()
        self.test_emails = [create_mock_email(i) for i in range(3)]  # Create 3 test emails
        
        # Initialize exporters
        self.json_exporter = JSONExporter()
        self.csv_exporter = CSVExporter()
        self.excel_exporter = ExcelExporter()
        self.pdf_exporter = PDFExporter()
    
    def tearDown(self):
        """Clean up test files."""
        # Clean up test files
        for file in Path(self.test_dir).glob('*'):
            try:
                if file.is_file():
                    file.unlink()
                else:
                    file.rmdir()
            except Exception as e:
                print(f"Warning: Could not delete {file}: {e}")
    
    def test_json_export_import(self):
        """Test exporting and importing JSON data."""
        # Export to JSON
        json_path = Path(self.test_dir) / 'test_export.json'
        success, message = self.json_exporter.export_emails(
            emails=self.test_emails,
            output_path=json_path
        )
        
        self.assertTrue(success, f"JSON export failed: {message}")
        self.assertTrue(json_path.exists(), "JSON file was not created")
        
        # Verify JSON structure
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Check metadata
        self.assertIn('metadata', data)
        self.assertEqual(data['metadata']['format'], 'json')
        self.assertEqual(data['metadata']['version'], '1.0')
        self.assertEqual(data['metadata']['record_count'], len(self.test_emails))
        
        # Check emails
        self.assertIn('emails', data)
        self.assertEqual(len(data['emails']), len(self.test_emails))
        
        # Verify data integrity
        for i, email in enumerate(data['emails']):
            self.assertEqual(email['subject'], f'Test Email {i}')
            self.assertEqual(email['sender'], f'sender{i}@example.com')
            self.assertEqual(len(email['to_recipients']), 2)
            
            # Check that all expected fields are present
            for field in ['id', 'conversation_id', 'subject', 'sender', 'received_date', 'body']:
                self.assertIn(field, email)
    
    def test_json_to_csv_conversion(self):
        """Test converting JSON export to CSV."""
        # First export to JSON
        json_path = Path(self.test_dir) / 'test_export.json'
        success, _ = self.json_exporter.export_emails(
            emails=self.test_emails,
            output_path=json_path
        )
        self.assertTrue(success)
        
        # Load the JSON data
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Export to CSV
        csv_path = Path(self.test_dir) / 'test_export.csv'
        success, message = self.csv_exporter.export_emails(
            emails=data['emails'],  # Pass the nested emails list
            output_path=csv_path,
            include_headers=True
        )
        
        self.assertTrue(success, f"CSV export failed: {message}")
        self.assertTrue(csv_path.exists(), "CSV file was not created")
        
        # Verify CSV content
        df = pd.read_csv(csv_path)
        self.assertEqual(len(df), len(self.test_emails))
        
        # Print actual fields for debugging
        print("\nCSV Fields:", df.columns.tolist())
        print("First row:", df.iloc[0].to_dict() if not df.empty else "No data in CSV")
        
        # Get display names from constants
        field_map = {field['name']: field['id'] for field in constants.EXPORT_FIELDS_V1}
        
        # Check some key fields - use .loc to avoid chained indexing
        for i in range(len(self.test_emails)):
            row = df.iloc[i]
            
            # Check subject (using display name 'Subject')
            if 'Subject' in row:
                self.assertEqual(row['Subject'], f'Test Email {i}')
            
            # Check sender (using display name 'From')
            if 'From' in row:
                self.assertEqual(row['From'], f'sender{i}@example.com')
            
            # Verify at least some fields are present
            self.assertTrue(any(field in row for field in ['Subject', 'From', 'Received']), 
                          f"No expected fields found in row {i}")
    
    def test_json_to_excel_conversion(self):
        """Test converting JSON export to Excel."""
        # First export to JSON
        json_path = Path(self.test_dir) / 'test_export.json'
        success, _ = self.json_exporter.export_emails(
            emails=self.test_emails,
            output_path=json_path
        )
        self.assertTrue(success)
        
        # Load the JSON data
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Export to Excel
        excel_path = Path(self.test_dir) / 'test_export.xlsx'
        success, message = self.excel_exporter.export_emails(
            emails=data['emails'],  # Pass the nested emails list
            output_path=excel_path
        )
        
        self.assertTrue(success, f"Excel export failed: {message}")
        self.assertTrue(excel_path.exists(), "Excel file was not created")
        
        # Verify Excel content
        df_emails = pd.read_excel(excel_path, engine='openpyxl')
        self.assertEqual(len(df_emails), len(self.test_emails))
        
        # Print actual fields for debugging
        print("\nExcel Fields:", df_emails.columns.tolist())
        print("First row:", df_emails.iloc[0].to_dict() if not df_emails.empty else "No data in Excel")
        
        # Check some key fields - use .loc to avoid chained indexing
        for i in range(len(self.test_emails)):
            row = df_emails.iloc[i]
            
            # Check subject (using display name 'Subject')
            if 'Subject' in row:
                self.assertEqual(row['Subject'], f'Test Email {i}')
            
            # Check sender (using display name 'From')
            if 'From' in row:
                self.assertEqual(row['From'], f'sender{i}@example.com')
            
            # Verify at least some fields are present
            self.assertTrue(any(field in row for field in ['Subject', 'From', 'Received']), 
                          f"No expected fields found in row {i}")
            
            # Check summary sheet
            df_summary = pd.read_excel(excel_path, sheet_name='Summary', engine='openpyxl')
            self.assertGreater(len(df_summary), 0)
    
    def test_pdf_export(self):
        """Test exporting to PDF format."""
        pdf_path = Path(self.test_dir) / 'test_export.pdf'
        success, message = self.pdf_exporter.export_emails(
            emails=self.test_emails,
            output_path=pdf_path,
            include_summary=True
        )
        
        self.assertTrue(success, f"PDF export failed: {message}")
        self.assertTrue(pdf_path.exists(), "PDF file was not created")
        self.assertGreater(pdf_path.stat().st_size, 1000, "PDF file is too small")
        
        # Verify the PDF contains expected content
        with open(pdf_path, 'rb') as f:
            pdf_content = f.read()
            self.assertIn(b'PDF-', pdf_content[:10], "File is not a valid PDF")
    
    def test_export_formats_consistency(self):
        """Test that all export formats produce consistent results."""
        # Export to JSON
        json_path = Path(self.test_dir) / 'consistency_test.json'
        success, message = self.json_exporter.export_emails(
            emails=self.test_emails,
            output_path=json_path
        )
        self.assertTrue(success, f"JSON export failed: {message}")
        
        # Load JSON data for comparison
        with open(json_path, 'r', encoding='utf-8') as f:
            json_emails = json.load(f)['emails']
        
        # Export to CSV
        csv_path = Path(self.test_dir) / 'consistency_test.csv'
        success, message = self.csv_exporter.export_emails(
            emails=json_emails,
            output_path=csv_path,
            include_headers=True
        )
        self.assertTrue(success, f"CSV export failed: {message}")
        
        # Export to Excel
        excel_path = Path(self.test_dir) / 'consistency_test.xlsx'
        success, message = self.excel_exporter.export_emails(
            emails=json_emails,
            output_path=excel_path
        )
        self.assertTrue(success, f"Excel export failed: {message}")
        
        # Export to PDF
        pdf_path = Path(self.test_dir) / 'consistency_test.pdf'
        success, message = self.pdf_exporter.export_emails(
            emails=json_emails,
            output_path=pdf_path,
            include_summary=True
        )
        self.assertTrue(success, f"PDF export failed: {message}")
        
        # Load all formats for comparison (except PDF which is binary)
        csv_data = pd.read_csv(csv_path)
        excel_data = pd.read_excel(excel_path, engine='openpyxl')
        
        # Check record counts
        self.assertEqual(len(json_emails), len(csv_data))
        self.assertEqual(len(json_emails), len(excel_data))
        
        # Create a mapping of field IDs to display names
        field_name_map = {field['id']: field['name'] for field in constants.EXPORT_FIELDS_V1}
        
        # Check data consistency between formats
        for i in range(len(json_emails)):
            # Get the current email from JSON
            json_email = json_emails[i]
            
            # Get the current rows from CSV and Excel
            csv_row = csv_data.iloc[i]
            excel_row = excel_data.iloc[i]
            
            # Check subject (using display name from constants)
            display_subject = field_name_map.get('subject', 'Subject')
            if display_subject in csv_row and 'subject' in json_email:
                self.assertEqual(str(json_email['subject']), str(csv_row[display_subject]))
            if display_subject in excel_row and 'subject' in json_email:
                self.assertEqual(str(json_email['subject']), str(excel_row[display_subject]))
            
            # Check sender (using display name from constants)
            display_sender = field_name_map.get('sender', 'From')
            if display_sender in csv_row and 'sender' in json_email:
                self.assertEqual(str(json_email['sender']), str(csv_row[display_sender]))
            if display_sender in excel_row and 'sender' in json_email:
                self.assertEqual(str(json_email['sender']), str(excel_row[display_sender]))
            
            # Verify at least some fields are present in all exports
            required_fields = [
                field_name_map.get('subject', 'Subject'),
                field_name_map.get('sender', 'From'),
                field_name_map.get('received_date', 'Received')
            ]
            
            self.assertTrue(any(field in csv_row for field in required_fields),
                          f"No expected fields found in CSV row {i}")
            self.assertTrue(any(field in excel_row for field in required_fields),
                          f"No expected fields found in Excel row {i}")
        
        # Verify PDF was created and has content
        self.assertTrue(pdf_path.exists(), "PDF file was not created")
        self.assertGreater(pdf_path.stat().st_size, 1000, "PDF file is too small")


if __name__ == '__main__':
    unittest.main()
