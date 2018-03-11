import smtplib, datetime, os, logging, configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def email(toaddr, reportnm, filename, filepath):
    config = configparser.ConfigParser()
    config.read('config.ini')
    from_email = config['gmail']['addr']
    pw = config['gmail']['pw']
    mailserver = config['gmail']['server']
    port = config['gmail']['port']
    
    dt_str = datetime.datetime.now().strftime("%Y_%m_%d_%H.%M.%S")
    to_email = toaddr # 'gavingreer@mac.com'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = "Attached Report: " + reportnm #"This is the subject line: " + dt_str
    
    body = 'This the email message ' + dt_str
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    try:
        attachment = open(filepath + "\\" + filename, 'rb')
    except:
        # print("file could not be opened")
        # logging.warning("file could not be opened")
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

    # send the email
    server = smtplib.SMTP(mailserver, port)
    server.starttls()
    server.login(from_email, pw)
    text  = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    # return True