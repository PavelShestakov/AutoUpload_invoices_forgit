import pandas as pd
import config
import os
import datetime
import restructure_foo
from letter import Letter
import functions
import json

def download_inbox_invoices(imap):
    """Скачивает накладные из входящих"""
    tomorrow_date = datetime.date.today() - datetime.timedelta(days=1)
    today_date = datetime.date.today()

    uid_list = functions.last_uid_list(imap, 30, 'SEEN') #список uid последних 30 прочитанных писем
    all_letters = 1
    for uid in uid_list:
        LastLetter = Letter(uid, imap)
        letter_from = str(LastLetter.letter_from)
        if "@email.com" not in letter_from and "@eml.com" not in letter_from:
            if LastLetter.letter_date == tomorrow_date or LastLetter.letter_date == today_date:
                if LastLetter.letter_theme != "NoneType":
                    all_letters += 1
                    LastLetter.download_inbox_letters(config.inbox_dict_path)

    print(f'Всего за {tomorrow_date} обработано {all_letters} писем')

def create_request_inbox_invoices():
    request = {
    "secretWord":"*****************************",
    "systemInfo":"system=*********************************",
    "payload": {"documents": []} }

    for name in (os.listdir(config.inbox_dict_path)):
        invoice_path = config.inbox_dict_path + '\\' + name

        # формат json_data торг-12 отличается от формата поставщика(спам)
        df = pd.read_excel(invoice_path, engine='openpyxl', skiprows=0, header=None)
        json_data = df.to_json()
        data = json.loads(json_data)

        correct_data = functions.correct_names_columns(data, num_correct=1)

        request_inbox_document = restructure_foo.create_json_request_inbox_document(correct_data)

        request["payload"]["documents"].append(request_inbox_document)

    return request

