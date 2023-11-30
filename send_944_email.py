from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import os


def send_email_csv(file, subject, message, recipients):

    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login('noreply@gpalogisticsgroup.com', 'Turn*17300')
    # Craft message (obj)
    msg = MIMEMultipart()

    # message = f'{message}\nSend from Hostname: '
    message = f'{message}\n'
    msg['Subject'] = subject
    msg['From'] = 'noreply@gpalogisticsgroup.com'
    # msg['To'] = destination
    msg['To'] = '; '.join(recipients)
    # Insert the text to the msg going by e-mail
    msg.attach(MIMEText(message, 'html'))
    # Attach the pdf to the msg going by e-mail
    # for file in files:

    with open(file, "rb") as csv:
        attach = MIMEApplication(csv.read(), _subtype="csv")
        file = os.path.basename(file)
        attach.add_header('Content-Disposition', 'attachment', filename=str(file))
        msg.attach(attach)
        server.send_message(msg)
        server.quit()
