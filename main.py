import time
import config
import os
import invoice
import functions
import upload_spam
import upload_inbox


if __name__ == "__main__":
    #Создаем папки для обработки накладных
    if not os.path.isdir(config.inbox_dict_path):
        os.mkdir(config.inbox_dict_path)
    if not os.path.isdir(config.spam_dict_path):
        os.mkdir(config.spam_dict_path)

    imap_inbox = functions.def_connect('INBOX')

    print('Начало выгрузки накладных из INBOX')

    '''Удаляем все файлы из папки с накладными'''
    functions.delete_all_files(config.inbox_dict_path)

    '''Скачиваем накладные из почты за вчерашний день'''
    try:
        upload_inbox.download_inbox_invoices(imap_inbox)
        print('Накладные скачаны')
    except:
        raise SystemExit('Вероятно, почтовый сервер сейчас недоступен. Попробуй позже.')

    '''получаем список всех накладных в папке'''
    invoice_list, xlsx_list = invoice.invoices_lists()

    '''Конвертируем все накладные в xlsx и удаляем xls файлы'''
    for file in invoice_list:
        if file[-1] == 's':
            try:
                invoice.convert_xls_to_xlsx(file)
                os.remove(file)
            except:
                os.remove(file)

    invoice_list, xlsx_list = invoice.invoices_lists()

    '''Проверяем накладные на покупателей, ненужные(из списка) удаляем'''
    functions.check_invoice()

    try:
        '''Формируем реквест из инбокс накладных'''
        inbox_request = upload_inbox.create_request_inbox_invoices()
        '''отправляем inbox_request'''
        functions.post_json_request(inbox_request)
    except:
        print(f'Не удалось сформировать request, '
              f'вероятно в папке {config.inbox_dict_path} оказался файл не похожий на накладную.')

    print(f'Накладные из папки "Входящие" выгружены в систему через POST-запрос (Api выгрузка)')

#-----------------------------------------------------------------------------------------------------------------------

    print('Начало выгрузки накладных из Spam')
    imap_spam = functions.def_connect('Spam')

    try:
        '''Составляем json_request !!! ALL ПИСЬМА !!!'''
        spam_request = upload_spam.create_request_spam_invoices(imap_spam)
        '''Отправляем spam_request'''
        functions.post_json_request(spam_request)
        print(f'Накладные из папки "Спам" выгружены в систему через POST-запрос (Api выгрузка)')
    except:
        print('Что-то прошло не по плану... не удалось выгрузить накладные из папки "Спам":(, возможно их нет...')

    print("Окно закроется через 5 секунд")
    time.sleep(5)
    #input('Для завершения работы программы нажми "Enter"')

