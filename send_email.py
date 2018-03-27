import smtplib, datetime, os, logging, configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def email(toaddr, subject, body, **kwargs):
    config = configparser.ConfigParser()
    config.read('config.ini')
    from_email = config['gmail']['addr']
    pw = config['gmail']['pw']
    mailserver = config['gmail']['server']
    port = config['gmail']['port']
    to_email = toaddr 

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject 
    
    if 'html' in kwargs:
        msg.attach(MIMEText(body, 'html'))
    else:
        msg.attach(MIMEText(body, 'plain'))

    if 'filename' in kwargs and 'filepath' in kwargs:
        # open the file to be sent
        filename = kwargs.get('filename')
        filepath = kwargs.get('filepath')
        try:
            attachment = open(filepath + '\\' + filename, 'rb')
        except:
            logging.warning("file could not be opened")
            return False

        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

        # to change the payload into an encoded form
        p.set_payload((attachment).read())

        # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # attache the instance p to instance msg
        msg.attach(p)
    else:
        print('no attachment')
 
    # send the email
    server = smtplib.SMTP(mailserver, port)
    server.starttls()
    server.login(from_email, pw)
    text  = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    return True