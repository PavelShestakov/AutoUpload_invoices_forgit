import config
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_email(sender, password, recipient, inbox_dict_path):
    server = smtplib.SMTP("smtp.yandex.ru", 999)
    server.starttls()
    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = recipient

        for file in os.listdir(inbox_dict_path):
            file = config.inbox_dict_path + '\\' + file
            filename = os.path.basename(file)
            with open(file, 'rb') as f:  # encoding='Windows-1251'
                file = MIMEApplication(f.read())

            file.add_header('content-disposition', 'attachment', filename=filename)
            msg.attach(file)

        server.sendmail(sender, recipient, msg.as_string())

        print("The message was sent successfully!")

    except Exception as _ex:
        print(f"{_ex}\nCheck your login or password please!")

