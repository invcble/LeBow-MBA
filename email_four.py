import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd

def send_email_with_attachment(to_email, attachment_path):
    smtp_server = "smtp.gmail.com"
    smtp_port = 465
    from_email = "noreply.leadership.report@gmail.com"
    password = ""

    sender_display_name = "LeBow Leadership Report"
    subject = "Your Personalized Leadership Development Report - ORGB 511"
    body = """Dear Student,

We hope you're doing well. As part of your ORGB 511 course, you completed the LeBow Leadership Development Survey.

Attached is your personalized feedback report, providing insights into your leadership style and characteristics. This report is essential for your course development and will be used in group discussions, reflections, and assignments. Please review it carefully to support your learning and leadership growth throughout the quarter.

If you have any questions, feel free to contact your course coordinator.

Best regards,
LeBow Leadership Development Team
LeBow College of Business
Drexel University"""

    msg = MIMEMultipart()
    msg['From'] = f"{sender_display_name} <{from_email}>"
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    filename = os.path.basename(attachment_path)
    try:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {filename}')
            msg.attach(part)
    except Exception as e:
        print(f"Error attaching file: {e}")
        return

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(from_email, password)
        server.sendmail(from_email, to_email, msg.as_string())
        print("Email sent successfully!")

    except Exception as e:
        print(f"Error sending email: {e}")

    finally:
        if 'server' in locals():
            server.quit()


file_path = 'Fall 99_dataframe.xlsx'
dataframe = pd.read_excel(file_path)

pdf_folder = "Generated PDFs"

for index, row in dataframe.iterrows():
    name = row['Name']
    email = row['Email']
    
    pdf_path = os.path.join(pdf_folder, f"Workbook {name}.pdf")
    
    if os.path.exists(pdf_path):
        send_email_with_attachment(email, pdf_path)
        print(f"PDF for {name} not found at {pdf_path}")
    else:
        print(f"PDF for {name} not found at {pdf_path}")