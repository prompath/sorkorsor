import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

my_username = 'sender@email.com'
my_password = 'sender_password'
send_from = 'sender@email.com'
sender_name = 'Peter Parkingspace'
email_subject = 'Happy New Year!'
    
def send_email(send_from, sender_name, subject, body, recipients, username, password):
    """Perform email transaction to emails in a dataframe
    """
    recipients_list = [i for i in zip(recipients['email'], recipients['message'], recipients['recipient'])]
    for (email, message, recipient) in recipients_list:
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['Subject'] = subject
        msg['Date'] = formatdate(localtime=True)
        body_temp = body.replace('{sender name}', sender_name)
        body_temp = body_temp.replace('{recipient name}', recipient)
        body_temp = body_temp.replace('{text}', message)
        msg['To'] = email
        msgText = MIMEText(body_temp, 'html')
        msg.attach(msgText)
        smtp = smtplib.SMTP('smtp.office365.com', 587)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, [email], msg.as_string())
        smtp.quit()

def run():
    df = pd.read_excel('./sorkorsor/recipients.xlsx')
    body = open('./sorkorsor/html/card.html', 'r', encoding='utf8').read()
    send_email(send_from, sender_name, email_subject, body, df, username=my_username,
               password=my_password)
    
if __name__ == '__main__':
    df = pd.read_excel('./recipients.xlsx')
    body = open('./html/card.html', 'r', encoding='utf8').read()
    send_email(send_from, sender_name, email_subject, body, df, username=my_username,
               password=my_password)