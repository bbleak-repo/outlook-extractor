# outlook_extractor/export/csv_exporter.py
import csv
import re
import logging
import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime
import html
import email
from email import policy
from email.parser import BytesParser

logger = logging.getLogger(__name__)

class CSVExporter:
    """Handles the export of email data to CSV format with advanced text cleaning."""
    
    def __init__(self, config=None):
        self.config = config or {}
        self._setup_regex_patterns()
        
    def export_emails(self, emails, output_file, include_headers=True, encoding='utf-8'):
        """Export emails to a CSV file.
        
        Args:
            emails: List of email dictionaries to export
            output_file: Path to the output CSV file
            include_headers: Whether to include headers in the CSV
            encoding: File encoding to use
            
        Returns:
            bool: True if export was successful, False otherwise
        """
        try:
            if not emails:
                logger.warning("No emails to export")
                return False
                
            # Convert emails to a format suitable for CSV export
            rows = []
            for email_data in emails:
                row = {
                    'subject': email_data.get('subject', ''),
                    'sender': email_data.get('sender', ''),
                    'recipients': ', '.join(email_data.get('recipients', [])),
                    'date': email_data.get('received_time', '').isoformat() if email_data.get('received_time') else '',
                    'body': email_data.get('body', ''),
                    'folder': email_data.get('folder', '')
                }
                rows.append(row)
            
            # Create directory if it doesn't exist
            output_path = Path(output_file)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Write to CSV
            with open(output_file, 'w', newline='', encoding=encoding) as f:
                if not rows:
                    return False
                    
                fieldnames = list(rows[0].keys())
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                
                if include_headers:
                    writer.writeheader()
                
                for row in rows:
                    writer.writerow(row)
            
            logger.info(f"Successfully exported {len(emails)} emails to {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to export emails to CSV: {e}", exc_info=True)
            return False
    
    def _setup_regex_patterns(self):
        """Initialize regex patterns for text cleaning."""
        # Common email headers and footers to remove
        self.patterns = {
            'signature': re.compile(
                r'(?is)'  # Global flags at the start
                r'(?:'  # Start non-capturing group
                r'--\s*\n.*|'  # Standard signature separator
                r'^--\s*$.*|'  # Double dash separator
                r'(?:^[^\n]*[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}.*)|'  # Email in signature
                r'(?:^[^\n]*www\.[^\s]+\.[a-z]{2,}.*)|'  # URLs in signature
                r'(?:^[^\n]*\b(?:phone|mobile|tel|fax)[^\n:]*:.*)|'  # Contact info
                r'(?:^[^\n]*\b(?:regard|best|sincerely|cheers|thanks|thank you|br),?[^\n]*$)'  # Common closings
                r')'  # End non-capturing group
            ),
            'quoted_text': re.compile(
                r'(?m)'  # Global multiline flag at the start
                r'(?:'  # Start non-capturing group
                r'^>.*$|'  # Quoted text
                r'^On .*? wrote:$|'  # Email client quote
                r'^From:.*?$|'  # Email header
                r'^To:.*?$|'  # Email header
                r'^Sent:.*?$|'  # Email header
                r'^Subject:.*?$'  # Email header
                r')'  # End non-capturing group
            ),
            'whitespace': re.compile(r'\s+', re.UNICODE),
            'html_tags': re.compile(r'<[^>]+>'),
            'multiple_newlines': re.compile(r'\n{3,}'),
            'trailing_whitespace': re.compile(r'[ \t]+$', re.MULTILINE),
            'leading_whitespace': re.compile(r'^[ \t]+', re.MULTILINE),
            'confidentiality_notice': re.compile(
                r'(?is)confidential(?:ity)?(?: notice| statement| information).*?'
                r'(?:unintended recipient|do not use|unauthorized use)'
            )
        }
        
    def clean_body(self, body: str, is_html: bool = False) -> str:
        """Clean and normalize email body text."""
        if not body:
            return ""
            
        try:
            # Convert to string if needed
            if not isinstance(body, str):
                body = str(body)
                
            # Remove HTML tags if present
            if is_html:
                body = self.patterns['html_tags'].sub(' ', body)
                body = html.unescape(body)
                
            # Remove common email artifacts
            body = self.patterns['quoted_text'].sub('', body)
            body = self.patterns['signature'].sub('', body)
            body = self.patterns['confidentiality_notice'].sub('', body)
            
            # Normalize whitespace
            body = self.patterns['leading_whitespace'].sub('', body)
            body = self.patterns['trailing_whitespace'].sub('', body)
            body = self.patterns['multiple_newlines'].sub('\n\n', body)
            body = body.strip()
            
            return body
            
        except Exception as e:
            logger.error(f"Error cleaning email body: {str(e)}")
            return body or ""

    def extract_summary(self, body: str, max_sentences: int = 3) -> str:
        """Extract a summary from the email body."""
        if not body:
            return ""
            
        # Split into sentences (naive approach - could be enhanced with NLTK)
        sentences = re.split(r'(?<=[.!?])\s+', body)
        return ' '.join(sentences[:max_sentences])

    def export_emails_to_csv(
        self,
        emails: List[Dict],
        output_path: str,
        include_headers: bool = True
    ) -> str:
        """Export list of email dictionaries to CSV file."""
        if not emails:
            logger.warning("No emails provided for CSV export")
            return ""
            
        try:
            # Ensure output directory exists
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Define CSV fields
            fields = [
                'id', 'conversation_id', 'subject', 'sender', 'to_recipients',
                'cc_recipients', 'bcc_recipients', 'sent_datetime', 'received_datetime',
                'has_attachments', 'importance', 'is_read', 'body_preview',
                'web_link', 'parent_folder', 'categories', 'clean_body', 'summary'
            ]
            
            # Prepare data for CSV
            rows = []
            for email_data in emails:
                row = {field: email_data.get(field, '') for field in fields}
                
                # Clean and process body
                body = email_data.get('body', {}).get('content', '')
                is_html = email_data.get('body', {}).get('contentType', '').lower() == 'html'
                clean_body = self.clean_body(body, is_html)
                summary = self.extract_summary(clean_body)
                
                # Update row with processed data
                row.update({
                    'clean_body': clean_body,
                    'summary': summary,
                    'to_recipients': '; '.join(email_data.get('toRecipients', [])),
                    'cc_recipients': '; '.join(email_data.get('ccRecipients', [])),
                    'bcc_recipients': '; '.join(email_data.get('bccRecipients', [])),
                    'categories': '; '.join(email_data.get('categories', []))
                })
                
                rows.append(row)
            
            # Write to CSV
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fields)
                if include_headers:
                    writer.writeheader()
                writer.writerows(rows)
                
            logger.info(f"Successfully exported {len(emails)} emails to {output_path}")
            return str(output_path)
            
        except Exception as e:
            logger.error(f"Error exporting emails to CSV: {str(e)}")
            raise

    def export_subject_analysis(
        self,
        emails: List[Dict],
        output_path: str
    ) -> str:
        """Generate subject analysis report."""
        try:
            # Count subjects by folder
            subject_counts = {}
            for email_data in emails:
                folder = email_data.get('parent_folder', 'Unknown')
                subject = email_data.get('subject', '(No Subject)')
                
                if folder not in subject_counts:
                    subject_counts[folder] = {}
                subject_counts[folder][subject] = subject_counts[folder].get(subject, 0) + 1
            
            # Convert to DataFrame for easy CSV export
            rows = []
            for folder, subjects in subject_counts.items():
                for subject, count in subjects.items():
                    rows.append({
                        'folder': folder,
                        'subject': subject,
                        'count': count
                    })
            
            # Sort by count descending
            rows.sort(key=lambda x: x['count'], reverse=True)
            
            # Write to CSV
            df = pd.DataFrame(rows)
            df.to_csv(output_path, index=False)
            
            logger.info(f"Subject analysis exported to {output_path}")
            return str(output_path)
            
        except Exception as e:
            logger.error(f"Error generating subject analysis: {str(e)}")
            raise