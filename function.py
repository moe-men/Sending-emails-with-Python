from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import smtplib
import yaml
import os
import zipfile



def load_config():
    with open('config.yaml', "r", encoding="utf-8") as stream:
        return yaml.safe_load(stream)


def send_email(csv_file="path/to/csv"):

    msg = MIMEMultipart()

    config = load_config()

    EMAIL_FROM = config["sender"]["email"]

    EMAIL_PASSWORD = config["sender"]["password"]

    EMAIL_TO = [ config["receivers"][res]["email"] for res in config["receivers"] ]

    EMAIL_SUBJECT = config["subject"]

    MESSAGE_BODY = config["message_body"]

    body_part = MIMEText(MESSAGE_BODY, 'plain')
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO

    # Add body to email
    msg.attach(body_part)



    zidfile_name = compress([csv_file])

    with open(csv_file,'rb') as file:
        # Attach the file with filename to the email
        compress(file)
        msg.attach(MIMEApplication(file.read(), Name="data"))

    try:

        smtp_server = smtplib.SMTP('smtp.office365.com',587)

        smtp_server.ehlo()

        smtp_server.starttls()

        smtp_server.login(EMAIL_FROM, EMAIL_PASSWORD)

        smtp_server.sendmail(EMAIL_FROM, msg['To'], msg.as_string())

        smtp_server.close()

        print ("Email sent successfully!")

    except Exception as ex:

        print ("Something went wrong….",ex)


def compress(archive_list=[],zfilename=''):

    zout = zipfile.ZipFile(zfilename, "w", zipfile.ZIP_DEFLATED)

    for fname in archive_list:
        print ("writing: ", fname)
        zout.write(fname)
    zout.close()

    return zfilename
