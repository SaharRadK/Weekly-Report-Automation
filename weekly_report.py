from datetime import datetime
import pyodbc
import pandas as pd
import win32com.client as win32
from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import matplotlib.pyplot as plt
import numpy as np
import io
import base64
import pickle
import pyodbc
import pandas as pd

# Connect to your SQL database
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=your_server;'
                      'Database=your_database;'
                      'Trusted_Connection=yes;')

# Execute your SQL query
current_date = datetime.now().strftime('%Y-%m-%d')

query1 = "select * from [dbo].[Weekly_SS_report_1]"
df1 = pd.read_sql(query1, conn)

query2 = "select * from [dbo].[Weekly_SS_report_2]"
df2 = pd.read_sql(query2, conn)

# Close the connection
conn.close()

#Defin file names
from datetime import datetime, timedelta
today = datetime.today()
offset_to_last_thursday = (today.weekday() - 3) % 7
last_thursday = today - timedelta(days=offset_to_last_thursday)
friday_before_last_thursday = last_thursday - timedelta(days=6)

formatted_last_thursday = last_thursday.strftime("%Y%m%d")
formatted_friday_before_last_thursday = friday_before_last_thursday.strftime("%Y%m%d")

filename1 = f"WeeklyAddedNodes_{formatted_friday_before_last_thursday}_{formatted_last_thursday}.xlsx"
filename2 = f"WeeklyNewAPNs_{formatted_friday_before_last_thursday}_{formatted_last_thursday}.xlsx"

# Write DataFrame to an Excel file
df1.to_excel(filename1, index=False)
df2.to_excel(filename2, index=False)

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email(to_email, subject, attachment_paths):
    html = """<p>Dears,</p>
<p> Please find the APN activation weekly report attached.</p>
<p>Regards,</p>
<p>DS_reporting team</p>"""

    me = 'Your mail address'
    server = 'Your SMTP server details'
    you = ['first Recipient's email address','second Recipient's email address']

    message = MIMEMultipart("alternative")
    message['Subject'] = subject
    message['From'] = me
    message['To'] = ", ".join(you)

    message.attach(MIMEText(html, 'html'))
    
    for attachment_path in attachment_paths:
        with open(attachment_path,'rb') as binary:
            part = MIMEBase('application', 'octet-stream', Name=attachment_path)
            part.set_payload(binary.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
            message.attach(part)

    with smtplib.SMTP(host=server, port=25) as server:
        server.sendmail(me, you, message.as_string())


send_email(me, 'APN activation report', [filename1,filename2])