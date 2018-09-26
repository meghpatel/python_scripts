"""
This scripts helps you to send personalised emails using CSV files.
Just import CSV file here and it will send individual personalised mails.
For example if you want to start a mail starting from "Hi, ${name}" 
Well this script takes name as variable and replaces it.
"""

import smtplib
import csv

from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = 'Enter Email'
PASSWORD = 'Use password or app password by gmail'
"""
If you are sending from gmail, you have to use app password.
Just google how to generate app password.
"""

def get_contacts(filename):
    """
    Return two lists names, emails containing names and email addresses
    You can add subtract columns as per your requirements of personalised variables
    For example "Hi ${name}, Your order ${orderno}"
    You can add order no column in the CSV file.
    read from a file specified by filename.
    """
    
    names = []
    emails = []
    
    with open(filename) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            emails.append(row[0])
            names.append(row[1])
            line_count += 1
    return names,emails

"""
Most of the scripts you find online will be using normal text file for 
Contacts and database.
I've used CSV to make it more convenient.

NOTE: In case you edit CSV file manually.  
Don't keep the cursor in the new line at the end of file.
Keep it at the end of the last line or else you'll get error. 
"""

def read_template(filename):
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)
"""
You can send the message as a text or an HTML. 
"""
def main():
    names, emails = get_contacts('contacts_final.txt') # read contacts
    message_template = read_template('Message.html')

    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)

    # For each contact, send the email:
    for name, email in zip(names, emails):
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())

        # Prints out the emails to whom the mail is sent.
        print(email)

        # setup the parameters of the message
        msg['From']='YOUR NAME'
        msg['To']=email
        msg['Subject']="WRITE SUBJECT HERE"
        
        """ add in the message body
            add text instead of HTML below to send simple text messages.
        """
        msg.attach(MIMEText(message, 'html'))
        
        # send the message via the server set up earlier.
        s.send_message(msg)
        del msg
        
    # Terminate the SMTP session and close the connection
    s.quit()
    
if __name__ == '__main__':
    main()
