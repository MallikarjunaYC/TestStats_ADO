############################################################################
# Program to 
# import data from ADO into Excel and
# Send the generated Excel file by mail
# Create a scheduler to send mail
# Program by Mallikarjuna YC
############################################################################
############################################################################
# Import necessary python libraries
############################################################################

import requests
import base64
import json
import pandas as pd
import smtplib
import os
from email.message import EmailMessage

############################################################################
# Step 1: Fetch data from Azure DevOps (ADO) API
############################################################################

def fetch_ado_data():
    """
    Connects to Azure DevOps API and retrieves work items.
    Returns the JSON response if successful; otherwise, prints an error message.
    """

    # Organization and project details
    org = 'AIdevOpsMallikOrg'
    project = 'AIDevOpsTestProject'
    
    # API endpoint to fetch work items (IDs should be updated as required)
    api_url = f'https://dev.azure.com/{org}/{project}/_apis/wit/workitems?ids=1,2,3,4,5&api-version=7.1'
    
    # Personal Access Token (PAT) for authentication (replace with a valid PAT)
    pat = '3ooyl8JnjV4qquLCKRMUFGSAteikWe8E1D5n39FztHGCZRDLekJuJQQJ99BAACAAAAAAAAAAAAASAZDOgqs6'
    encoded_pat = base64.b64encode(f':{pat}'.encode()).decode()

    # Request headers
    headers = {
        'Authorization': f'Basic {encoded_pat}',
        'Content-Type': 'application/json'
    }

    # Making the API request
    response = requests.get(api_url, headers=headers)

    # Checking response status and returning data
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching ADO data: {response.status_code}")
        return None

############################################################################
# Step 2: Generate Excel report from fetched data
############################################################################

def generate_excel(data):
    """
    Converts JSON response from ADO API to an Excel report.
    Extracts key fields and saves the report in the specified location.
    """

    # Define the output Excel file path
    output_file = 'C:/Users/ycmal/Documents/ADO_output/ADO_Bug_Report.xlsx'

    if data:
        # Extracting required fields from the JSON response
        extracted_data = []
        for item in data.get('value', []):
            extracted_data.append({
                'ID': item['id'],
                'Area Path': item['fields'].get('System.AreaPath', 'N/A'),
                'Iteration Path': item['fields'].get('System.IterationPath', 'N/A'),
                'Work Item Type': item['fields'].get('System.WorkItemType', 'N/A'),
                'State': item['fields'].get('System.State', 'N/A'),
                'Reason': item['fields'].get('System.Reason', 'N/A'),
                'Created Date': item['fields'].get('System.CreatedDate', 'N/A'),
                'Created By': item['fields'].get('System.CreatedBy', 'N/A')
            })

        # Convert extracted data to a Pandas DataFrame
        df = pd.DataFrame(extracted_data)

        # Save data to Excel file
        df.to_excel(output_file, index=False)
        print("Excel report generated successfully.")
    else:
        print("No data available to generate report.")

############################################################################
# Step 3: Send an email with the Excel report as an attachment
############################################################################

def send_email():
    """
    Sends an email with the generated Excel report attached.
    Uses SMTP to send email from a Gmail account.
    """

    # SMTP server configuration for Gmail
    SMTP_SERVER = 'smtp.gmail.com'
    SMTP_PORT = 587

    # Email credentials and recipient details
    EMAIL_SENDER = 'your email'  # Sender email
    EMAIL_PASSWORD = 'pwd'  # Gmail app password
    EMAIL_RECIPIENT = 'Recipient email'  # Recipient email
    EMAIL_SUBJECT = 'Automated Bug Report'
    EMAIL_BODY = 'Please find the attached bug report.'

    # Path of the generated Excel report
    attachment_path = 'C:/Users/ycmal/Documents/ADO_output/ADO_Bug_Report.xlsx'

    # Create email message
    msg = EmailMessage()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECIPIENT
    msg.set_content(EMAIL_BODY)

    # Attach the Excel report to the email
    try:
        with open(attachment_path, 'rb') as attachment:
            file_data = attachment.read()
            file_name = os.path.basename(attachment_path)
            msg.add_attachment(file_data, maintype='application',
                              subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                              filename=file_name)

        # Connect to SMTP server and send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure the connection
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
            print("Email sent successfully!")

    except Exception as e:
        print(f"Failed to send email: {e}")

############################################################################
# Step 4: Execute the workflow
############################################################################

# Fetch data from ADO API
ado_data = fetch_ado_data()

# Generate Excel report
generate_excel(ado_data)

# Send email with report attached
send_email()
