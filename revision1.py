import pandas as pd
import smtplib
import os
from email.message import EmailMessage
from pathlib import Path

# Load HR emails from the Excel sheet
excel_file = Path("C:/Users/onkar/OneDrive/Documents/hr_emails.xlsx") # Update with your file path
df = pd.read_excel(excel_file, engine="openpyxl")  # Read the Excel file
email_list = df["Email"].dropna().tolist()  # Extract email column and remove NaNs

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "onkarpatil272@gmail.com"  # Replace with your email
EMAIL_PASSWORD = "gmzx sudx sbsf szwy"  # Use an app password

# Resume file path (update as necessary)
resume_path = "C:/Users/onkar/OneDrive/Desktop/Onkar patil resume.pdf"


# Function to send email with attachment
def send_email(receiver_email):
    msg = EmailMessage()
    msg["From"] = EMAIL_SENDER
    msg["To"] = receiver_email
    msg["Subject"] = "DevOps Engineer Job Application - Onkar Patil"
    msg.set_content(
        """Dear HR,

I hope this email finds you well. I am writing to express my interest in the DevOps Engineer position at your company.
With 3.2 years of experience in AWS, Terraform, Ansible, and CI/CD tools, I am eager to contribute my expertise.

Please find my resume attached for your reference. I would love the opportunity to discuss my qualifications further.

Looking forward to your response.

Best Regards,  
Onkar Patil  
Contact: +91 8421529035  
Email: onkarpatil272@gmail.com
"""
    )

    # Attach resume
    with open(resume_path, "rb") as resume:
        msg.add_attachment(resume.read(), maintype="application", subtype="pdf", filename="Onkar_Patil_Resume.pdf")

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")

# Send emails to all HR contacts
for email in email_list:
    send_email(email)
