import os
import config
import datetime
from letter import Letter
import functions
import restructure_foo


def create_request_spam_invoices(imap):
    """Формирует словарь с данными из накладных папки Spam, для запроса request"""
    request_spam = {
        "secretWord":"*****************************",
        "systemInfo":"system=*********************************",
        "payload": {"documents": []} }

    tomorrow_date = datetime.date.today() - datetime.timedelta(days=1)
    today_date = datetime.date.today()

    uid_list = functions.last_uid_list(imap, 30, 'ALL')
    for uid in uid_list:
        # Очищаем папку от всех файлов
        functions.delete_all_files(dir=config.spam_dict_path)
        LastLetter = Letter(uid, imap)
        # Скачиваем накладную
        try:
            if LastLetter.letter_date == tomorrow_date or LastLetter.letter_date == today_date:
                LastLetter.download_spam_letters(config.spam_dict_path)
                if len(os.listdir(config.spam_dict_path)) <= 0:
                    print(f"В письме {LastLetter.UID} накладная не обнаружена")
                    continue

                #сохраняем имя накладной
                invoice_name = [name for name in (os.listdir(config.spam_dict_path))][0]
                #форматируем накладную в xlsx
                file_path = config.spam_dict_path + '\\' + invoice_name
                functions.convert_xls_to_xlsx(file_path)
                os.remove(file_path)
                invoice_name = [name for name in (os.listdir(config.spam_dict_path))][0]
                file_path = config.spam_dict_path + '\\' + invoice_name

                data = functions.xlsx_to_json(file_path)
                correct_data = functions.correct_names_columns(data, num_correct=2)

                # +формируем json документ из только что скачанной накладной
                request_document = restructure_foo.create_json_request_spam_document(correct_data)

                #Выгрузка только для КОНКРЕТНОГО КЛИЕНТА
                #------------------------------------------------------------------
                buyer_name = request_document['buyer']['name']
                if 'КОНКРЕТНЫЙ КЛИЕНТ'.upper() in buyer_name.upper():
                    #добавляем json документ в json запрос
                    print('Есть накладная для КОНКРЕТНЫЙ КЛИЕНТ')
                    request_spam["payload"]["documents"].append(request_document)
                else:
                    print(f'{buyer_name} не выгружаем.')
                # ------------------------------------------------------------------
                # удаляем накладную
                functions.delete_all_files(dir=config.spam_dict_path)
        except:
            print(f'При обработке письма {LastLetter.UID} от {LastLetter.letter_date} произошла ошибка. '
                  f'Возможно в письме оказался файл не похожий на накладную.')

    return request_spam


