"""Email sending functionality for automated organization chart reports."""

from __future__ import annotations

import logging
import smtplib
import json
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, List, Tuple, Any, Optional

from .email_config import get_smtp_config, load_email_config
from .screenshot import is_playwright_available, generate_org_chart_png_via_export
from .settings import load_settings

logger = logging.getLogger(__name__)


def _get_chart_title() -> str:
    """Get the configured chart title, with fallback to default."""
    try:
        settings = load_settings()
        return settings.get('chartTitle', 'Organization Chart')
    except Exception:
        return 'Organization Chart'


def send_test_email(recipient: str) -> Tuple[bool, str]:
    """
    Send a test email to verify SMTP configuration.
    
    Args:
        recipient: Email address(es) to send test email to (comma-separated)
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    smtp_config = get_smtp_config()
    
    try:
        chart_title = _get_chart_title()
        
        # Parse recipients (comma-separated)
        recipients = _parse_recipients(recipient)
        if not recipients:
            return False, 'No valid recipient email addresses provided'
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = smtp_config['fromAddress']
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f'{chart_title} - Test Email'
        
        server = smtp_config['server']
        port = smtp_config['port']
        from_addr = smtp_config['fromAddress']
        test_date = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M %Z')
        
        body = f"""
        <html>
        <body>
            <h2>{chart_title} Email Reports - Test Email</h2>
            <p>This is a test email from your {chart_title} application.</p>
            <p>If you receive this email, your SMTP configuration is working correctly!</p>
            <p><strong>SMTP Server:</strong> {server}:{port}</p>
            <p><strong>From Address:</strong> {from_addr}</p>
            <p><strong>Date:</strong> {test_date}</p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html'))
        
        # Send email
        success = _send_email_smtp(smtp_config, msg, recipients)
        
        if success:
            return True, f'Test email sent successfully to {len(recipients)} recipient(s)!'
        else:
            return False, 'Failed to send test email. Check logs for details.'
            
    except Exception as e:
        logger.error("Error sending test email", exc_info=e)
        return False, 'Failed to send test email. Check server logs for details.'


def send_test_email_with_attachments(
    recipient: str,
    xlsx_content: bytes = None,
    base_url: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Send an email report with attachments (XLSX and/or PNG) immediately.
    
    This is used for manual "Send Now" functionality from the configure page.
    Unlike send_test_email, this sends a full report email with attachments
    based on the current email configuration settings.
    
    Args:
        recipient: Email address(es) to send email to (comma-separated)
        xlsx_content: XLSX file content as bytes (if available)
        base_url: Base URL for generating PNG screenshots
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    smtp_config = get_smtp_config()
    email_config = load_email_config()
    chart_title = _get_chart_title()
    
    try:
        # Parse recipients (comma-separated)
        recipients = _parse_recipients(recipient)
        if not recipients:
            return False, 'No valid recipient email addresses provided'
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = smtp_config['fromAddress']
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f'{chart_title} Report - {datetime.now().strftime("%Y-%m-%d")}'
        
        # Email body
        file_types = email_config.get('fileTypes', [])
        attachments_list = []
        if 'xlsx' in file_types and xlsx_content:
            attachments_list.append('XLSX (Excel)')
        if 'png' in file_types:
            attachments_list.append('PNG (Chart Image)')
        
        attachments_text = ', '.join(attachments_list) if attachments_list else 'No attachments configured'
        
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <h2 style="color: #0078D4;">{chart_title} - Organization Chart Report</h2>
            <p>This is your organization chart report with the configured attachments.</p>
            
            <h3>Included Attachments:</h3>
            <p>{attachments_text}</p>
            
            <p style="margin-top: 20px; color: #666; font-size: 12px;">
                Generated on: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M %Z')}<br>
                This is an automated email from {chart_title}.
            </p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html'))
        
        # Attach XLSX if provided and requested
        if 'xlsx' in file_types and xlsx_content:
            try:
                filename = f'org-chart-{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.xlsx'
                _attach_file(
                    msg, 
                    xlsx_content, 
                    filename, 
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                logger.info("XLSX file attached to email")
            except Exception as e:
                logger.warning(f"Failed to attach XLSX file to email: {e}")
                return False, 'Failed to attach XLSX file. Check server logs for details.'
        
        # Attach PNG screenshot if requested
        if 'png' in file_types and base_url:
            try:
                if not is_playwright_available():
                    logger.warning(
                        "PNG export requested but Playwright is not installed. "
                        "Install with: pip install playwright && playwright install chromium"
                    )
                    return False, 'PNG export requires Playwright. Install with: pip install playwright && playwright install chromium'
                else:
                    logger.info("Generating PNG export for email...")
                    png_content = generate_org_chart_png_via_export(
                        base_url=base_url,
                        timeout_ms=60000
                    )
                    
                    if png_content:
                        filename = f'org-chart-{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.png'
                        _attach_file(msg, png_content, filename, 'image/png')
                        logger.info(f"PNG export attached to email ({len(png_content)} bytes)")
                    else:
                        logger.warning("Failed to generate PNG export for email")
                        return False, 'Failed to generate PNG export'
            except Exception as e:
                logger.warning(f"Failed to attach PNG screenshot to email: {e}")
                return False, 'Failed to attach PNG. Check server logs for details.'
        
        # Send email
        success = _send_email_smtp(smtp_config, msg, recipients)
        
        if success:
            return True, f'Email sent successfully to {len(recipients)} recipient(s)!'
        else:
            return False, 'Failed to send email. Check logs for details.'
            
    except Exception as e:
        logger.error(f"Error sending email with attachments: {e}")
        return False, 'An internal error occurred while sending email with attachments.'


def send_report_email(
    xlsx_content: bytes = None,
    reports_data: Dict[str, Any] = None,
    base_url: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Send scheduled org chart report email with attachments.
    
    Supports XLSX and PNG exports. PNG requires Playwright for server-side rendering.
    
    Args:
        xlsx_content: XLSX file content as bytes (if available)
        reports_data: Optional dictionary of report data to include as attachments
        base_url: Base URL of the application for generating PNG screenshots (e.g., 'http://localhost:5000')
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    email_config = load_email_config()
    smtp_config = get_smtp_config()
    chart_title = _get_chart_title()
    
    if not email_config.get('enabled'):
        return False, 'Email reports are disabled'
    
    recipient = email_config.get('recipientEmail')
    if not recipient:
        return False, 'No recipient email configured'
    
    # Parse recipients (comma-separated)
    recipients = _parse_recipients(recipient)
    if not recipients:
        return False, 'No valid recipient email addresses configured'
    
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = smtp_config['fromAddress']
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f'{chart_title} Report - {datetime.now().strftime("%Y-%m-%d")}'
        
        # Email body
        body = _create_report_email_body(email_config, chart_title)
        msg.attach(MIMEText(body, 'html'))
        
        # Attach XLSX if provided and requested
        file_types = email_config.get('fileTypes', [])
        
        if 'xlsx' in file_types and xlsx_content:
            try:
                filename = f'org-chart-{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.xlsx'
                _attach_file(
                    msg, 
                    xlsx_content, 
                    filename, 
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                logger.info("XLSX file attached to email")
            except Exception as e:
                logger.warning(f"Failed to attach XLSX file: {e}")
        
        # Attach PNG screenshot if requested
        if 'png' in file_types and base_url:
            try:
                if not is_playwright_available():
                    logger.warning(
                        "PNG export requested but Playwright is not installed. "
                        "Install with: pip install playwright && playwright install chromium"
                    )
                else:
                    logger.info("Generating PNG export...")
                    png_content = generate_org_chart_png_via_export(
                        base_url=base_url,
                        timeout_ms=60000
                    )
                    
                    if png_content:
                        filename = f'org-chart-{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.png'
                        _attach_file(msg, png_content, filename, 'image/png')
                        logger.info(f"PNG export attached to email ({len(png_content)} bytes)")
                    else:
                        logger.warning("Failed to generate PNG export")
            except Exception as e:
                logger.warning(f"Failed to attach PNG screenshot: {e}")
        
        # Attach additional reports if configured
        if reports_data and email_config.get('includeReports'):
            _attach_reports(msg, reports_data, email_config.get('includeReports', []))
        
        # Send email
        success = _send_email_smtp(smtp_config, msg, recipients)
        
        if success:
            logger.info(f"Report email sent successfully to {len(recipients)} recipient(s): {', '.join(recipients)}")
            return True, f'Report email sent to {len(recipients)} recipient(s)'
        else:
            return False, 'Failed to send report email'
            
    except Exception as e:
        logger.error(f"Error sending report email: {e}")
        return False, f'Error: {str(e)}'


def _send_email_smtp(smtp_config: Dict[str, Any], msg: MIMEMultipart, recipients: List[str]) -> bool:
    """
    Send email using SMTP configuration.
    
    Args:
        smtp_config: SMTP configuration dictionary
        msg: Email message to send
        recipients: List of recipient email addresses
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Connect to SMTP server based on encryption type
        encryption = smtp_config.get('encryption', 'TLS')
        
        if encryption == 'SSL':
            # Use SSL/TLS from the start (typically port 465)
            server = smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'])
        else:
            # Use plain connection, optionally upgrade with STARTTLS
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
            
            if encryption == 'TLS':
                # Upgrade to TLS using STARTTLS (typically port 587)
                server.starttls()
        
        # Login
        server.login(smtp_config['username'], smtp_config['password'])
        
        # Send email to all recipients
        server.send_message(msg, to_addrs=recipients)
        server.quit()
        
        logger.info(f"Email sent successfully to {len(recipients)} recipient(s)")
        return True
        
    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"SMTP authentication failed: {e}")
        return False
    except smtplib.SMTPException as e:
        logger.error(f"SMTP error: {e}")
        return False
    except Exception as e:
        logger.error(f"Unexpected error sending email: {e}")
        return False


def _create_report_email_body(email_config: Dict[str, Any], chart_title: str = 'Organization Chart') -> str:
    """Create HTML body for report email."""
    frequency = email_config.get('frequency', 'weekly')
    file_types = email_config.get('fileTypes', [])
    
    # Build list of attachments
    attachments_html = ''
    if 'xlsx' in file_types:
        attachments_html += '<li>XLSX - Organization chart employee data</li>'
    if 'png' in file_types:
        attachments_html += '<li>PNG - Organization chart visual diagram</li>'
    
    include_reports = email_config.get('includeReports', [])
    if include_reports:
        for report in include_reports:
            report_name = report.replace('_', ' ').title()
            attachments_html += f'<li>JSON - {report_name} report data</li>'
    
    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #0078D4;">{chart_title} Report</h2>
        <p>This is your {frequency} automated organization chart report.</p>
        
        {f'<h3>Attached Files:</h3><ul>{attachments_html}</ul>' if attachments_html else '<p><em>No attachments configured.</em></p>'}
        
        <p style="margin-top: 20px; color: #666; font-size: 12px;">
            Generated on: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M %Z')}<br>
            Frequency: {frequency.capitalize()}<br>
            This is an automated email from {chart_title}.
        </p>
    </body>
    </html>
    """


def _attach_file(msg: MIMEMultipart, content: bytes, filename: str, mime_type: str) -> None:
    """Attach a file to the email message."""
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(content)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    part.set_type(mime_type)
    msg.attach(part)


def _attach_reports(msg: MIMEMultipart, reports_data: Dict[str, Any], report_types: List[str]) -> None:
    """Attach additional report files as JSON or CSV."""
    import json
    
    for report_type in report_types:
        if report_type in reports_data:
            try:
                data = reports_data[report_type]
                filename = f'{report_type}-{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.json'
                content = json.dumps(data, indent=2).encode('utf-8')
                _attach_file(msg, content, filename, 'application/json')
            except Exception as e:
                logger.warning(f"Failed to attach report {report_type}: {e}")


def _parse_recipients(recipient_string: str) -> List[str]:
    """
    Parse comma-separated email addresses and validate them.
    
    Args:
        recipient_string: Comma-separated email addresses
        
    Returns:
        List of valid email addresses
    """
    if not recipient_string:
        return []
    
    # Split by comma and clean up whitespace
    recipients = [email.strip() for email in recipient_string.split(',')]
    
    # Filter out empty strings and basic validation
    valid_recipients = []
    for email in recipients:
        if email and '@' in email and '.' in email:
            valid_recipients.append(email)
        elif email:
            logger.warning(f"Skipping invalid email address: {email}")
    
    return valid_recipients


__all__ = [
    'send_test_email',
    'send_report_email',
]
