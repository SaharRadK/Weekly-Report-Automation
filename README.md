# Weekly APN Report Automation

This Python script automates the generation and distribution of weekly APN activation reports by querying an SQL database, processing the data, and sending the results via email.

## Features
- Connects to an SQL Server database using `pyodbc`.
- Queries and processes data for two weekly reports.
- Exports results to Excel files with dynamically generated filenames.
- Sends an email with the generated reports as attachments.

## Prerequisites
To run this script, you need:
- Python 3.x installed on your system.
- Required Python libraries:
  - `pyodbc`
  - `pandas`
  - `win32com`
  - `smtplib`
  - `email`
  - `matplotlib` (if required for further extensions)
- Access to an SQL Server database.
- A configured SMTP server for sending emails.

## How to Use
1. Update the database connection string to match your SQL Server credentials and database:
   ```python
   conn = pyodbc.connect('Driver={SQL Server};'
                         'Server=your_server;'
                         'Database=your_database;'
                         'Trusted_Connection=yes;')

Update the email configuration in the send_email function