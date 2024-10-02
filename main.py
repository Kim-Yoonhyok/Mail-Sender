import pandas as pd
import win32com.client
import os

companies = []
recruiters = []
emails = []

def read_excel_file(file_path):
    # Read the excel file
    df = pd.read_excel(file_path, index_col=0)

    for company in df["Company"]:
        companies.append(company)
    for recruiter in df["Recruiter"]:
        recruiters.append(recruiter)
    for email in df["Email"]:
        emails.append(email)

def read_message_template(file_path):
    with open(file_path, 'r') as file:
        return file.read()

def send_email(companies, recruiters, emails):
    outlook = win32com.client.Dispatch('outlook.application')
    min_num = min(len(companies), len(recruiters), len(emails))

    for i in range(min_num):
        mail = outlook.CreateItem(0) 
        subject = "Fall 2024 Internship"
        body = message_template.format(recruiter=recruiters[i], company=companies[i])

        mail.Subject = subject
        mail.Body = body
        mail.To = emails[i]

        script_dir = os.path.dirname(os.path.abspath(__file__))
        resume_path = os.path.join(script_dir, 'Resume.pdf')
        mail.Attachments.Add(resume_path)

        mail.Send()
    print("Email sent!")

message_template = read_message_template('message.txt')
read_excel_file('Book1.xlsx')
send_email(companies, recruiters, emails)