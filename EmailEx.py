from flask import Flask, render_template, request
from flask_restful import Resource, Api
import csv
import imaplib
import email
import re

app = Flask(__name__)  # Creating Flask app
api = Api(app)  # Creating Flask-RESTful API

# Function to extract text from email
def get_text_from_email(msg):
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                try:
                    text = part.get_payload(decode=True).decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        text = part.get_payload(decode=True).decode('iso-8859-1')
                    except UnicodeDecodeError:
                        try:
                            text = part.get_payload(decode=True).decode('utf-16')
                        except UnicodeDecodeError:
                            text = part.get_payload(decode=True).decode('latin-1')
    else:
        try:
            text = msg.get_payload(decode=True).decode('utf-8')
        except UnicodeDecodeError:
            try:
                text = msg.get_payload(decode=True).decode('iso-8859-1')
            except UnicodeDecodeError:
                try:
                    text = msg.get_payload(decode=True).decode('utf-16')
                except UnicodeDecodeError:
                    text = msg.get_payload(decode=True).decode('latin-1')
    return text

# Function to extract data from text
def extract_data_from_text(text):
    extracted_data = {}
    # Extract name
    name_match = re.search(r"Name:\s*([^\r\n]+)", text)
    if name_match:
        extracted_data["Name"] = name_match.group(1).strip()
    # Extract phone number
    phone_match = re.search(r"Phone:\s*([\d\s-]+)", text)
    if phone_match:
        extracted_data["Phone"] = phone_match.group(1).strip()
    # Extract email
    email_match = re.search(r"Email:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", text)
    if email_match:
        extracted_data["Email"] = email_match.group(1).strip()
    # Extract company
    company_match = re.search(r"Company:\s*([^\r\n]+)", text)
    if company_match:
        extracted_data["Company"] = company_match.group(1).strip()
    # Extract subject
    subject_match = re.search(r"Subject:\s*([^\r\n]+)", text)
    if subject_match:
        extracted_data["Subject"] = subject_match.group(1).strip()
    return extracted_data

# Resource class for email extraction
class EmailExtraction(Resource):
    def post(self):
        # Extracting email credentials from form data
        email_address = request.form['T1']
        password = request.form['P1']
        extracted_data = []  # List to store extracted data
        
        # Defining IMAP settings for different email providers
        imap_settings = {
            "gmail.com": ("imap.gmail.com", "inbox"),
            "one.com": ("imap.one.com", "inbox"),
            "outlook.com":("imap.outlook.com","inbox")
            # Add more providers and their settings as needed
        }
        
        domain = email_address.split('@')[-1]  # Extracting domain from email address
        if domain in imap_settings:
            server, mailbox = imap_settings[domain]
            mail = imaplib.IMAP4_SSL(server)  # Connecting to IMAP server
            try:
                mail.login(email_address, password)  # Logging in to email account
            except imaplib.IMAP4.error as e:
                return {"error": f"Failed to log in: {str(e)}"}  # Returning error if login fails
            
            mail.select(mailbox)  # Selecting mailbox

            # Searching and retrieving emails
            status, email_ids = mail.search(None, "ALL")
            email_ids = email_ids[0].split()
            for email_id in email_ids:
                status, email_data = mail.fetch(email_id, "(RFC822)")
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                text = get_text_from_email(msg)  # Extracting text from email
                extracted_data.append(extract_data_from_text(text))  # Extracting data from text
            
            mail.logout()  # Logging out from email account
            
            # Writing extracted data to a CSV file
            with open("extracted_data.csv", "w", newline="", encoding="utf-8") as csv_file:
                fieldnames = ["Name", "Phone", "Email", "Company","Subject"]
                csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                csv_writer.writeheader()
                for data in extracted_data:
                    csv_writer.writerow(data)
        
        return {"message": "Email extraction completed."}

# Adding EmailExtraction resource to API
api.add_resource(EmailExtraction, '/process-credentials')

# Route for home page
@app.route('/')
def home():
    return render_template('index.html')  # Rendering index.html template

if __name__ == '__main__':
    app.run(debug=True)  # Running Flask app
