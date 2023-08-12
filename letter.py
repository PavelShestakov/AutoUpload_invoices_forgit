import email
from email.header import decode_header
import os
import datetime


class Letter():
    def __init__(self, uid, def_connect):
        self.UID = uid
        self.def_connect = def_connect

        self.res, self.msg = self.def_connect.uid('fetch', self.UID, '(RFC822)')#После этой операции в почтовом ящике письмо будет отмечено как прочитанное.
        self.letter = email.message_from_bytes(self.msg[0][1])#декодирует письмо

        self.letter_date = email.utils.parsedate_tz(self.letter["Date"])
        self.letter_date = datetime.date(self.letter_date[0],self.letter_date[1],self.letter_date[2])

        self.letter_uid = self.letter["Message-ID"]
        self.letter_from = self.letter["Return-path"]#e-mail отправителя

        self.letter_theme = self.letter["Subject"]

        self.sender = ''

        try:
            if self.letter_theme is None:
                self.letter_theme = "NoneType"

            else:
                self.letter_theme = decode_header(self.letter_theme[0][0])
                self.letter_theme = decode_header(self.letter["Subject"])[0][0].decode('koi8-r')
        except:
            self.letter_theme = "Not_theme"



    def download_inbox_letters(self, dict_path):
        for part in self.letter.walk():
            if "application" in part.get_content_type():
                filename = part.get_filename()
                filename = str(email.header.make_header(email.header.decode_header(filename)))
                if not (filename):
                    filename = "invoice"

                if 'ТОРГ-12' in filename or filename == "invoice":
                    with open(os.path.join(dict_path, filename), 'wb') as fp:
                        fp.write(part.get_payload(decode=1))


    def download_spam_letters(self, dict_path):
        for part in self.letter.walk():
            if "application" in part.get_content_type():
                filename = part.get_filename()
                filename = str(email.header.make_header(email.header.decode_header(filename)))
                if not (filename):
                    filename = "invoice"

                with open(os.path.join(dict_path, filename), 'wb') as fp:
                    fp.write(part.get_payload(decode=1))

