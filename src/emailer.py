"""
Email Module
Sends PDF report(s) via Outlook COM API.
Auto-sends to configured recipient(s) -- no manual review needed.
"""

import os
import win32com.client


def send_report_email(subject, body_text, pdf_paths, recipient, cc=None):
    """
    Send an email with PDF attachments via Outlook.

    Args:
        subject: Email subject line
        body_text: Plain text email body
        pdf_paths: List of PDF file paths to attach
        recipient: Email address (or semicolon-separated list)
        cc: Optional CC address(es)
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    mail.To = recipient
    if cc:
        mail.CC = cc
    mail.Subject = subject
    mail.Body = body_text

    # Attach PDFs
    for pdf_path in pdf_paths:
        abs_path = os.path.abspath(pdf_path)
        if os.path.exists(abs_path):
            mail.Attachments.Add(abs_path)
        else:
            print(f"  WARNING: Attachment not found: {abs_path}")

    mail.Send()
    print(f"  Email sent to: {recipient}")
    if cc:
        print(f"  CC: {cc}")
    print(f"  Subject: {subject}")
    print(f"  Attachments: {len(pdf_paths)}")
