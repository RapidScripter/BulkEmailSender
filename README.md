# Bulk Email Sender

Bulk Email Sender is a user-friendly application for sending bulk emails with attachments. It allows you to easily manage email campaigns by selecting recipients from an Excel file, composing messages, and adding attachments effortlessly. Ideal for businesses, organizations, and individuals who need to reach a large audience quickly and effectively.

## Features
- Send bulk emails to multiple recipients
- Attach multiple files to emails
- Read recipient email addresses from an Excel file
- User-friendly interface with responsive design

## Installation

### Prerequisites
- Python 3.7+
- `pip` (Python package installer)

### Usage
1. Clone the repository or download the script.
   ```bash
   git clone https://github.com/RapidScripter/BulkEmailSender
   cd BulkEmailSender

2. Install Required Packages
   ```bash
   pip install -r requirements.txt

3. Run the Application
   ```bash
   python bulk_email_sender.py

### User Interface
- Enter the sender email and password. (To use the Bulk Email Sender with Gmail, you need to create an app password.)
- Enter the email subject and body.
- Select the Excel file containing the recipient email addresses. (Do not change header in excel file)
- Select attachments to include in the emails.
- Click "Send Emails" to start the email campaign.

### Steps to Creating an App Password for Gmail
1. **Log in to your Google account**: Open a web browser and log in to your Google account.

2. **Go to My Account/App Passwords**: Point your web browser to Google App Passwords (URL - https://myaccount.google.com/apppasswords). Even though you just logged in to your account, you'll most likely be prompted to type your account password again.

3. **Create your first app password**: At the bottom of the page, you should see two drop-downs, one titled "Select app" and the other "Select device." Click the "Select app" drop-down and choose Other. You will then be prompted to name the app password. I would suggest naming the app password for the app or service you'll use it for. After naming the app password, click GENERATE.
![Alt text](/Screenshots/apppassword.jpg?raw=true "App Password")

4. **Save Your app password**: Use this app password instead of your regular Gmail password when setting up the Bulk Email Sender.

### Screenshots

![Alt text](/Screenshots/BulkEmailTool1.jpg?raw=true "Bulk Email Sender Tool")

![Alt text](/Screenshots/BulkEmailTool2.jpg?raw=true "Bulk Email Sender Tool")
