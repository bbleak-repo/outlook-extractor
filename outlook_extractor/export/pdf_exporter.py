"""PDF export functionality for Outlook Extractor.

This module provides functionality to export email data to PDF format with
professional styling, tables, and support for custom templates.
"""
import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from threading import Event
from typing import Any, Dict, List, Optional, Tuple, Union, Callable

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.platypus.flowables import Image
from reportlab.platypus.paragraph import Paragraph

from .. import constants

logger = logging.getLogger(__name__)

class PDFExporter:
    """Handles the export of email data to PDF format.
    
    This exporter creates professional-looking PDF reports with proper formatting,
    tables, and support for custom templates.
    """
    
    def __init__(self, config: Optional[dict] = None):
        """Initialize the PDF exporter.
        
        Args:
            config: Optional configuration dictionary
        """
        self.config = config or {}
        self.fields = constants.EXPORT_FIELDS_V1
        self._setup_styles()
    
    def _setup_styles(self):
        """Set up the styles for the PDF document."""
        self.styles = getSampleStyleSheet()
        
        # Define custom styles
        custom_styles = [
            {
                'name': 'Title',
                'parent': 'Heading1',
                'fontSize': 18,
                'spaceAfter': 12,
                'alignment': 1
            },
            {
                'name': 'Heading2',
                'parent': 'Heading2',
                'fontSize': 14,
                'spaceAfter': 6,
                'textColor': colors.HexColor('#2E86C1')
            },
            {
                'name': 'BodyText',
                'parent': 'BodyText',
                'fontSize': 10,
                'leading': 14,
                'spaceAfter': 6
            },
            {
                'name': 'Footer',
                'parent': 'Italic',
                'fontSize': 8,
                'textColor': colors.grey,
                'alignment': 1
            },
            {
                'name': 'TableHeader',
                'parent': 'Heading3',
                'fontSize': 10,
                'textColor': colors.white,
                'alignment': 1,
                'backColor': colors.HexColor('#5D6D7E')
            },
            {
                'name': 'TableText',
                'parent': 'Normal',
                'fontSize': 9,
                'leading': 12,
                'spaceAfter': 2,
                'spaceBefore': 2
            }
        ]
        
        # Add custom styles if they don't already exist
        for style_def in custom_styles:
            style_name = style_def.pop('name')
            parent_style = style_def.pop('parent')
            
            # Only add the style if it doesn't already exist
            if style_name not in self.styles:
                self.styles.add(ParagraphStyle(
                    name=style_name,
                    parent=self.styles[parent_style],
                    **style_def
                ))
            # Restore the style definition for potential reuse
            style_def['name'] = style_name
            style_def['parent'] = parent_style
    
    def _create_header_footer(self, canvas, doc):
        """Create header and footer for each page."""
        # Save the state of our canvas so we can draw on it
        canvas.saveState()
        
        # Header
        header_text = "Outlook Email Export"
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawRightString(doc.width + doc.leftMargin, doc.height + doc.topMargin - 20, 
                             header_text)
        
        # Footer
        footer_text = f"Page {doc.page} • Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        canvas.setFont('Helvetica', 8)
        canvas.drawCentredString(doc.width/2.0, 0.5 * inch, footer_text)
        
        # Release the canvas
        canvas.restoreState()
    
    def _format_field_value(self, field: str, value: Any) -> str:
        """Format a field value for display in PDF."""
        if value is None:
            return ""
        
        # Handle datetime fields
        if isinstance(value, (datetime, str)) and field in ['received_date', 'sent_date']:
            try:
                if isinstance(value, str):
                    dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
                else:
                    dt = value
                return dt.strftime('%Y-%m-%d %H:%M:%S')
            except (ValueError, TypeError):
                return str(value)
        
        # Handle boolean fields
        if isinstance(value, bool):
            return "Yes" if value else "No"
        
        # Handle list/dict fields
        if isinstance(value, (list, dict)):
            return str(value)
        
        return str(value)
    
    def _create_email_section(self, email: Dict[str, Any]) -> list:
        """Create a section for a single email."""
        elements = []
        
        # Add email subject as title
        elements.append(Paragraph(email.get('subject', 'No Subject'), self.styles['Title']))
        
        # Add sender and date
        from_text = f"<b>From:</b> {email.get('sender', 'Unknown')}"
        date_text = f"<b>Date:</b> {self._format_field_value('received_date', email.get('received_date'))}"
        
        elements.append(Paragraph(from_text, self.styles['BodyText']))
        elements.append(Paragraph(date_text, self.styles['BodyText']))
        
        # Add recipients if available
        if 'to' in email and email['to']:
            to_text = f"<b>To:</b> {email['to']}"
            elements.append(Paragraph(to_text, self.styles['BodyText']))
        
        # Add CC if available
        if 'cc' in email and email['cc']:
            cc_text = f"<b>CC:</b> {email['cc']}"
            elements.append(Paragraph(cc_text, self.styles['BodyText']))
        
        # Add a separator
        elements.append(Spacer(1, 12))
        
        # Add email body
        body = email.get('body', '')
        if body:
            elements.append(Paragraph("<b>Message:</b>", self.styles['Heading2']))
            elements.append(Spacer(1, 6))
            
            # Convert line breaks to HTML breaks for proper formatting
            body = body.replace('\n', '<br/>')
            elements.append(Paragraph(body, self.styles['BodyText']))
        
        # Add attachments if available
        if 'attachments' in email and email['attachments']:
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>Attachments:</b>", self.styles['Heading2']))
            
            attachments = email['attachments']
            if isinstance(attachments, list) and attachments:
                for i, attachment in enumerate(attachments, 1):
                    if isinstance(attachment, dict):
                        name = attachment.get('name', f'Attachment {i}')
                        size = attachment.get('size', 0)
                        size_kb = f"({size/1024:.1f} KB)" if size else ""
                        elements.append(Paragraph(f"• {name} {size_kb}", self.styles['BodyText']))
        
        elements.append(PageBreak())
        return elements
    
    def _create_summary_table(self, emails: List[Dict[str, Any]]) -> Table:
        """Create a summary table of all emails."""
        # Prepare table data
        table_data = [
            ['Subject', 'From', 'Date', 'Size']
        ]
        
        for email in emails:
            table_data.append([
                email.get('subject', 'No Subject'),
                email.get('sender', 'Unknown'),
                self._format_field_value('received_date', email.get('received_date')),
                f"{len(str(email.get('body', '')))/1024:.1f} KB"
            ])
        
        # Create table
        table = Table(table_data, colWidths=[250, 150, 100, 80])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4F81BD')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        return table
    
    def export_emails(
        self,
        emails: List[Dict[str, Any]],
        output_path: Union[str, Path],
        include_summary: bool = True,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[Event] = None
    ) -> Tuple[bool, str]:
        """Export emails to a PDF file.
        
        Args:
            emails: List of email dictionaries to export
            output_path: Path to the output PDF file
            include_summary: Whether to include a summary table
            progress_callback: Optional callback for progress updates
            cancel_event: Optional event to cancel the export
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        if not emails:
            return False, "No emails to export"
        
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Create PDF document
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # Prepare story (content elements)
        story = []
        
        # Add title
        title = "Outlook Email Export"
        story.append(Paragraph(title, self.styles['Title']))
        story.append(Spacer(1, 24))
        
        # Add summary table
        if include_summary and len(emails) > 1:
            story.append(Paragraph("Summary", self.styles['Heading2']))
            story.append(Spacer(1, 12))
            story.append(self._create_summary_table(emails))
            story.append(PageBreak())
        
        # Add each email
        total_emails = len(emails)
        for i, email in enumerate(emails, 1):
            if cancel_event and cancel_event.is_set():
                return False, "Export cancelled by user"
                
            try:
                story.extend(self._create_email_section(email))
                
                # Update progress
                if progress_callback:
                    progress_callback(i, total_emails)
                    
            except Exception as e:
                logger.error(f"Error processing email {i}: {e}", exc_info=True)
                continue
        
        # Build the PDF
        doc.build(
            story,
            onFirstPage=self._create_header_footer,
            onLaterPages=self._create_header_footer
        )
        
        return True, f"Successfully exported {len(emails)} emails to {output_path}"

    def export_to_buffer(self, emails: List[Dict[str, Any]]) -> bytes:
        """Export emails to an in-memory PDF buffer.
        
        Args:
            emails: List of email dictionaries to export
            
        Returns:
            Bytes containing the PDF data
        """
        from io import BytesIO
        
        buffer = BytesIO()
        
        # Create PDF document in memory
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # Prepare story (content elements)
        story = []
        
        # Add title
        story.append(Paragraph("Outlook Email Export", self.styles['Title']))
        story.append(Spacer(1, 24))
        
        # Add summary table if multiple emails
        if len(emails) > 1:
            story.append(Paragraph("Summary", self.styles['Heading2']))
            story.append(Spacer(1, 12))
            story.append(self._create_summary_table(emails))
            story.append(PageBreak())
        
        # Add each email
        for email in emails:
            story.extend(self._create_email_section(email))
        
        # Build the PDF
        doc.build(
            story,
            onFirstPage=self._create_header_footer,
            onLaterPages=self._create_header_footer
        )
        
        # Get the PDF data
        buffer.seek(0)
        return buffer.getvalue()
