# Outlook Email Automation
Send personalized emails with attachments directly from an Excel file.
### Features
This script reads email data from an Excel spreadsheet using pandas, extracting the fields; Name, Email, Attachment Path, and CC for each contact. It validates attachment paths with Python’s os module to ensure the files exist before proceeding. With win32com.client, the script connects to Outlook and generates individual draft emails, preserving your default HTML signature. Each message is personalized with the recipient’s name, includes the relevant attachment, and is opened in Outlook for manual review, giving you the chance to double-check everything before sending.
### Useage
After configuration, run from any terminal or active environment:
```bash
python program.py
