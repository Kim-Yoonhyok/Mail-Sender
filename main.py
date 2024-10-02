import pandas as pd
import win32com.client
import os


companies = []
recruiters = []
emails = []


def read_excel_file(file_path):
    # Read the excel file
    df = pd.read_excel(file_path, index_col=0)
    #print("Column names:", df.columns.tolist())

    for company in df["Company"]:
        companies.append(company)

    for recruiter in df["Recruiter"]:
        recruiters.append(recruiter)

    for email in df["Email"]:
        emails.append(email)




def send_email(companies, recruiters, emails):
    # Create an Outlook application object
    outlook = win32com.client.Dispatch('outlook.application')
    
    # Set email parameters
    min_num = min(len(companies), len(recruiters), len(emails))

    for i in range(min_num):
        mail = outlook.CreateItem(0) 

        subject = "Fall 2024 Internship"
        body = f"""Hi {recruiters[i]}, 

I hope you're having a great week so far! Recognizing your likely busy schedule, I'll keep this brief. 

I'm currently looking for a fall 2024 internship in operations, business, or any other relevant field at {companies[i]}.
I realize it's late in the term, but unfortunately, my internship planned for this fall fell through. I would love the
opportunity to discuss any openings you may have. Looking forward to hearing from you soon! 

Sincerely, 
Jasmine Tey"""

    mail.Subject = subject
    mail.Body = body
    mail.To = emails[i]

    script_dir = os.path.dirname(os.path.abspath(__file__))
    resume_path = os.path.join(script_dir, 'Resume.pdf')
    mail.Attachments.Add(resume_path)

    # Send the email
    mail.Send()
    print("Email sent!")

excel_file = 'Book1.xlsx'
read_excel_file(excel_file)
send_email(companies, recruiters, emails)