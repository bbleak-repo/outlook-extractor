"""Test script to verify export functionality with mock data."""
import os
import sys
import tempfile
from pathlib import Path

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent.absolute()))

from outlook_extractor.export import CSVExporter, ExcelExporter, JSONExporter, PDFExporter

def create_mock_emails(count=5):
    """Create mock email data for testing."""
    emails = []
    for i in range(1, count + 1):
        email = {
            'id': f'email_{i}',
            'subject': f'Test Email {i}',
            'sender': f'sender{i}@example.com',
            'to_recipients': [f'recipient{i}@example.com'],
            'cc_recipients': [f'cc{i}@example.com'],
            'bcc_recipients': [f'bcc{i}@example.com'],
            'received_date': f'2025-07-16T10:{i:02d}:00',
            'sent_date': f'2025-07-16T09:{i:02d}:00',
            'is_read': i % 2 == 0,
            'has_attachments': i % 3 == 0,
            'body': f'This is the body of test email {i}.\nIt has multiple lines.\n\nBest regards,\nSender {i}',
            'size': 1024 * i,
            'importance': 'Normal',
            'categories': ['Test', f'Category{i}'],
            'is_draft': False,
            'is_encrypted': False,
            'is_signed': True,
            'internet_message_headers': {f'Header-{i}': f'Value-{i}'},
            'attachments': [
                {
                    'name': f'file{i}.txt',
                    'size': 1024 * i,
                    'content_type': 'text/plain',
                    'content': f'This is attachment {i}'.encode('utf-8')
                }
            ] if i % 3 == 0 else []
        }
        emails.append(email)
    return emails

def test_export_formats():
    """Test exporting to different formats."""
    # Create a temporary directory for test exports
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        print(f"\nExporting test files to: {temp_dir}")
        
        # Create mock email data
        emails = create_mock_emails(3)
        
        # Test CSV export
        csv_path = temp_dir / 'test_export.csv'
        csv_exporter = CSVExporter()
        success, message = csv_exporter.export_emails(emails, str(csv_path))
        print(f"\nCSV Export: {message}")
        print(f"File exists: {csv_path.exists()}, Size: {csv_path.stat().st_size if csv_path.exists() else 0} bytes")
        
        # Test Excel export
        excel_path = temp_dir / 'test_export.xlsx'
        excel_exporter = ExcelExporter()
        success, message = excel_exporter.export_emails(emails, str(excel_path))
        print(f"\nExcel Export: {message}")
        print(f"File exists: {excel_path.exists()}, Size: {excel_path.stat().st_size if excel_path.exists() else 0} bytes")
        
        # Test JSON export
        json_path = temp_dir / 'test_export.json'
        json_exporter = JSONExporter()
        success, message = json_exporter.export_emails(emails, str(json_path))
        print(f"\nJSON Export: {message}")
        print(f"File exists: {json_path.exists()}, Size: {json_path.stat().st_size if json_path.exists() else 0} bytes")
        
        # Test PDF export
        pdf_path = temp_dir / 'test_export.pdf'
        pdf_exporter = PDFExporter()
        success, message = pdf_exporter.export_emails(emails, str(pdf_path))
        print(f"\nPDF Export: {message}")
        print(f"File exists: {pdf_path.exists()}, Size: {pdf_path.stat().st_size if pdf_path.exists() else 0} bytes")
        
        # Print the contents of the directory
        print("\nDirectory contents:")
        for f in temp_dir.glob('*'):
            print(f"- {f.name} ({f.stat().st_size} bytes)")

if __name__ == '__main__':
    test_export_formats()
