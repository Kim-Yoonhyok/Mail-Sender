# Mail Sender Script
### Overview
This script automates the process of sending personalized emails to a list of recruiters from an Excel file. It uses the win32com.client library to interact with Microsoft Outlook and send emails.

### Prerequisites
- Python 3.x
- Microsoft Outlook installed and configured
- Required Python libraries:
    - pandas
    - pywin32

### Installation
1. Clone the repository:
git clone <repository-url>
cd <repository-directory>

2. Install the required libraries:
pip install pandas pywin32

3. Ensure the Excel file and the resume are in the same directory as the script

### Usage
1. Prepare the Excel file:
- The excel file should have columns named Company, Recruiter, and Email. The first column must be kept empty
- Make sure to change the message file to whatever message you want to send

2. Run the script:
python main.py


