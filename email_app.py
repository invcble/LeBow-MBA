import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox


class EmailApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("Email Sender with Attachments")
        self.geometry("600x300")

        # Scorelist file path input
        self.scorelist_label = ctk.CTkLabel(self, text="Scorelist File Path (.xlsx):")
        self.scorelist_label.pack(pady=10)
        self.scorelist_entry = ctk.CTkEntry(self, width=500)
        self.scorelist_entry.pack()

        # PDF folder path input
        self.pdf_folder_label = ctk.CTkLabel(self, text="PDF Folder Path:")
        self.pdf_folder_label.pack(pady=10)
        self.pdf_folder_entry = ctk.CTkEntry(self, width=500)
        self.pdf_folder_entry.pack()

        # Submit button
        self.submit_button = ctk.CTkButton(self, text="Send Emails", command=self.send_emails)
        self.submit_button.pack(pady=20)

    def send_emails(self):
        scorelist_file = self.scorelist_entry.get().replace('"','')
        pdf_folder = self.pdf_folder_entry.get().replace('"','')

        # Check if input paths are valid
        if not os.path.exists(scorelist_file) or not os.path.isdir(pdf_folder):
            messagebox.showerror("Error", "Invalid file path or folder path.")
            return

        try:
            dataframe = pd.read_excel(scorelist_file)
            for index, row in dataframe.iterrows():
                name = row['Name']
                email = row['Email']

                pdf_path = os.path.join(pdf_folder, f"Workbook {name}.pdf")

                if os.path.exists(pdf_path):
                    self.send_email_with_attachment(email, pdf_path)
                else:
                    print(f"PDF for {name} not found at {pdf_path}")
                    messagebox.showwarning("Missing PDF", f"PDF for {name} not found.")
                messagebox.showinfo("Success", "Reports generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {e}")

    def send_email_with_attachment(self, to_email, attachment_path):
        smtp_server = "smtp.gmail.com"
        smtp_port = 465
        from_email = "noreply.leadership.report@gmail.com"
        password = ""  # Use your app password here

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
        msg['X-Priority'] = '1'  # High Priority
        msg['X-MSMail-Priority'] = 'High'
        msg['Importance'] = 'High'

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


if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()
