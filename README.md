# Email List Cleaning Automation

This project was developed during my internship to automate email list maintenance tasks.  
All data used in this repository is anonymized, and no company-specific content is included.

## Project Overview

Maintaining clean mailing lists is essential for business communication. Manual cleanup of undelivered
emails and rejected addresses across multiple Excel sheets is error-prone and time-consuming.

This tool provides automated methods to:

- Identify undelivered or bounced emails from raw text input
- Extract and move them to a reject list for review
- Remove rejected emails from other mailing lists
- Highlight affected rows for easier verification

## Scripts Included

### 1) `undelivered_email_cleaner.py`

This script lets users paste raw undelivered email text and automatically:

- Parses email addresses from the text
- Searches mailing lists in selected Excel sheets
- Extracts and highlights undelivered rows
- Adds them to a dedicated reject sheet

It uses a simple GUI interface for file and sheet selection.

### 2) `reject_list_cleaner.py`

This script removes rejected email addresses from mailing lists based on a provided reject list.

It scans all specified sheets and deletes rows with matched emails.

## How to Use

1. Install required Python packages:
   ```bash
   pip install pandas openpyxl
