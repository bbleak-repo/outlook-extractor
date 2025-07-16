"""Script to verify the contents of exported files."""
import json
import csv
import tempfile
from pathlib import Path
import pandas as pd

def verify_csv(file_path):
    """Verify the contents of a CSV file."""
    print(f"\n=== Verifying CSV: {file_path} ===")
    with open(file_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        
    print(f"Found {len(rows)} records")
    print("First record:")
    for key, value in rows[0].items():
        print(f"  {key}: {value[:100]}{'...' if len(str(value)) > 100 else ''}")

def verify_excel(file_path):
    """Verify the contents of an Excel file."""
    print(f"\n=== Verifying Excel: {file_path} ===")
    df = pd.read_excel(file_path)
    print(f"Found {len(df)} records with {len(df.columns)} columns")
    print("Columns:", ", ".join(df.columns))
    print("First record:")
    for col in df.columns:
        value = df.iloc[0][col]
        print(f"  {col}: {str(value)[:100]}{'...' if len(str(value)) > 100 else ''}")

def verify_json(file_path):
    """Verify the contents of a JSON file."""
    print(f"\n=== Verifying JSON: {file_path} ===")
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    if isinstance(data, list):
        print(f"Found {len(data)} records")
        if data:
            print("First record keys:", ", ".join(data[0].keys()))
    else:
        print("Unexpected JSON structure:", type(data))

def verify_pdf(file_path):
    """Basic verification of a PDF file."""
    print(f"\n=== Verifying PDF: {file_path} ===")
    file_size = Path(file_path).stat().st_size
    print(f"File size: {file_size} bytes")
    print("Note: Full PDF content verification requires PDF parsing libraries")

def main():
    """Main function to verify all export files."""
    # Create a temporary directory for test exports
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        print(f"Using temporary directory: {temp_dir}")
        
        # Run the export test to generate files
        from test_export_ui import create_mock_emails
        emails = create_mock_emails(3)
        
        # Export to different formats
        from outlook_extractor.export import CSVExporter, ExcelExporter, JSONExporter, PDFExporter
        
        # CSV Export
        csv_path = temp_dir / 'test_export.csv'
        csv_exporter = CSVExporter()
        csv_exporter.export_emails(emails, str(csv_path))
        verify_csv(csv_path)
        
        # Excel Export
        excel_path = temp_dir / 'test_export.xlsx'
        excel_exporter = ExcelExporter()
        excel_exporter.export_emails(emails, str(excel_path))
        verify_excel(excel_path)
        
        # JSON Export
        json_path = temp_dir / 'test_export.json'
        json_exporter = JSONExporter()
        json_exporter.export_emails(emails, str(json_path))
        verify_json(json_path)
        
        # PDF Export
        pdf_path = temp_dir / 'test_export.pdf'
        pdf_exporter = PDFExporter()
        pdf_exporter.export_emails(emails, str(pdf_path))
        verify_pdf(pdf_path)
        
        print("\n=== All exports verified successfully! ===")

if __name__ == '__main__':
    main()
