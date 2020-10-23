import smtplib
import pandas as pd
from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

FROM_EMAIL = 'mail@gmail.com'
MY_PASSWORD = 'password'
subject = 'People, Data and Reputation , This is what Matter the Most'
def get_users():

    excel_read = pd.read_excel("email.xlsx")
    all_user_email = excel_read['Email']

#    names = []
#    emails = []
#    with open(file_name, mode='r', encoding='utf-8') as user_file:
#        for user_info in user_file:
#            names.append(user_info.split()[0])
#            emails.append(user_info.split()[1])
    return all_user_email

def parse_template(file_name):
    with open(file_name, 'r', encoding='utf-8') as msg_template:
        msg_template_content = msg_template.read()
    return Template(msg_template_content)

def send_email(smtp_server, message, bcc_list):
   
    print("BCC:", bcc_list) 
    multipart_msg = MIMEMultipart()  #create a message

    multipart_msg['From']=FROM_EMAIL
    multipart_msg['To']=FROM_EMAIL
    multipart_msg['Bcc']=bcc_list
    multipart_msg['Subject']=subject

    multipart_msg.attach(MIMEText(message, 'html'))
        
    #send the message via the server set up earlier.
    smtp_server.ehlo()
    smtp_server.send_message(multipart_msg)
    del multipart_msg
 


def main():
    message_template = parse_template('messagehtml.txt')

    smtp_server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    smtp_server.starttls()
    smtp_server.login(FROM_EMAIL, MY_PASSWORD)

    emails = get_users() # read user details

    # add in the actual person name to the message template
    name = FROM_EMAIL
    message = message_template.substitute(USER_NAME=name.title())
    print(message)

    new_email = ''
    for idx in range(len(emails)):
        print(emails[idx])
        if idx%3 == 0:
            #got the 10 contact send the email
            new_email = emails[idx] + ',' + new_email
            send_email(smtp_server, message, new_email) 
            new_email = ''
        else:
            new_email = emails[idx] + ',' + new_email

    #out of for loop send last email
    print("Sending last email")
    send_email(smtp_server, message, new_email) 

    # Terminate the SMTP session and close the connection
    smtp_server.quit()
    
if __name__ == '__main__':
    main()
