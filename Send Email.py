import smtplib, os, datetime,logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

"""Email Logging Setting"""
LOG_FORMAT = '%(levelname)s %(asctime)s - %(message)s'
logging.basicConfig(filename='PycharmProjects\\PyFun\\Email_Log.log',
                    level=logging.INFO, format=LOG_FORMAT, datefmt='%d/%m/%Y %H:%M:%S', filemode='a')
logger = logging.getLogger()

"""Current Date"""
CD = datetime.datetime.now().strftime('%d/%m/%Y_%H:%M:%S')

"""SMTP Setting"""
sender_email = os.environ.get('EmailName')
send_pass = os.environ.get('EmailPass')
smtp_server = 'smtp.office365.com'
port = 587


def send_email(receiver_email, log_file_path, screenshots_path):
    """Email Content"""
    logger.info("Drafting Email")
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receiver_email)
    msg['Subject'] = "Error Found During Testing"
    # Email Body
    msg.attach(MIMEText(
        '<b>%s</b>' % (
                "Hi all, "
                "<br><br>Email Sent at: " + CD), 'html'))
    # Attach log file
    try:
        if len(os.listdir(log_file_path)) > 0:
            with open(log_file_path, 'rb') as attach_log:
                logfile = attach_log.read()
                attach_text = MIMEText(logfile, 'plain', 'utf-8')
                attach_text.add_header('Content-Disposition', 'attachment', filename='Log.txt')
                msg.attach(attach_text)
                logger.info("Attached log file")
        else:
            logger.info("Can not find log file")
    except IOError as error:
        logger.info("Log Attachment Error: "+error)

    # Attach screenshots
    try:
        if len(os.listdir(screenshots_path)) > 0:
            files = os.listdir(screenshots_path)
            for image in files:
                file_path = os.path.join(screenshots_path, image)
                image_name = os.path.basename(image)
                image = MIMEImage(open(file_path, 'rb').read())
                image.add_header('Content-Disposition', 'attachment', filename=image_name)
                msg.attach(image)
                logger.info("Attached screenshots")
        else:
            logger.info("Can not find any screenshots to attach")
    except IOError as error:
        logger.info("Screenshots Attachment Error: "+error)

    try:
        with smtplib.SMTP(smtp_server, port) as smtpObj:
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.login(sender_email, send_pass)
            smtpObj.sendmail(sender_email, receiver_email, msg.as_string())
            logger.info("Email sent\n")
    except Exception as error:
        logger.info(error)
    finally:
        smtpObj.quit()