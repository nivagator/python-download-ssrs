import smtplib, datetime, os, logging, configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def email(toaddr, subject, body, **kwargs):
    config = configparser.ConfigParser()
    config.read('email_creds.ini')
    from_email = config['gmail']['addr']
    pw = config['gmail']['pw']
    mailserver = config['gmail']['server']
    port = config['gmail']['port']
    # py_file_path = os.path.dirname(os.path.realpath(__file__))
    # dt_str = datetime.datetime.now().strftime("%Y_%m_%d_%H.%M.%S")
    to_email = toaddr # 'gavingreer@mac.com'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = "Attached Report: " + subject 
    
    # body = 'This the email message ' + dt_str
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
    # server = smtplib.SMTP(mailserver, port)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, pw)
    text  = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    return True