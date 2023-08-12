import functions

request = {
        "secretWord": "*****************************",
        "systemInfo": "system_info******************",
        "payload": {"documents": []}}

def create_json_request_spam_document(correct_data):
    request_document = {}

    number_data_cell = correct_data['E'][2]
    incomingNumber = number_data_cell.split(' ')[2]
    date = number_data_cell.split(' ')[4:7]
    date = functions.refactor_date_data(date)

    global sum_string
    column_with_sum = correct_data['D']
    for key, value in column_with_sum.items():
        if value == 'Всего к оплате:':
            sum_string = key
    sum = correct_data['Q'][sum_string]

    #seller_info
    #seller_inn = correct_data['E'][7].split(':')[-1].split('/')[0].strip()
    seller_inn = '1234567890'
    seller_name = correct_data['E'][5].split(':')[-1].strip()
    seller_additionalId = ''
    #seller_kpp = correct_data['E'][7].split(':')[-1].split('/')[-1].strip()
    seller_kpp = '123451001'
    seller = dict(inn=seller_inn, name=seller_name, additionalId=seller_additionalId, kpp=seller_kpp)

    #buyer_info
    buyer_inn = correct_data['E'][14].split(':')[-1].split('/')[0].strip()
    buyer_name = correct_data['E'][12].split(':')[-1].strip()
    buyer_additionalId = correct_data['E'][9].split(':')[1::]
    buyer_additionalId = ''.join(buyer_additionalId).strip()
    buyer_kpp = correct_data['E'][14].split(':')[-1].split('/')[-1].strip()
    buyer = dict(inn=buyer_inn, name=buyer_name, additionalId=buyer_additionalId, kpp=buyer_kpp)

    #строки между которыми находится номенклатура
    start_iter_string = 0
    last_iter_string = 0

    column_with_start = correct_data['B']
    for key, value in column_with_start.items():
        if value == 'Код  товара/ работ, услуг':
            start_iter_string = key + 3
            break

    column_with_sum = correct_data['D']
    for key, value in column_with_sum.items():
        if value == 'Всего к оплате:':
            last_iter_string = key
            break

    example = {"name": "Крахмал кукурузный(весовой) кг",
               "code": "00-00000",
               "measure": "кг",
               "amount": 1.0,
               "ndsPercent": 0.0,
               "ndsSum": 0.0,
               "totalSum": 99.9,
               "storeCode": ""}

    items = []
    for s in range(start_iter_string, last_iter_string):
        if correct_data['C'][s] == None:
            continue
        example.update(name=correct_data['E'][s],
                       code=correct_data['B'][s],
                       measure=correct_data['H'][s],
                       amount=correct_data['J'][s],
                       ndsPercent=correct_data['O'][s],
                       ndsSum=correct_data['P'][s],
                       totalSum=correct_data['Q'][s], )

        ndsPercent = float(example['ndsPercent'].split()[-1][0:-1])
        example.update(ndsPercent=ndsPercent)
        items.append(example.copy())

    request_document.update(seller=seller,
                            buyer=buyer,
                            incomingNumber=incomingNumber,
                            ttnNumber="",
                            invoiceNumber="",
                            date=date,
                            sum=sum,
                            items=items
                            )

    return request_document


def create_json_request_inbox_document(correct_data):
    request_document = {}

    incomingNumber = correct_data['H'][14]

    date = correct_data['J'][14].split('.')
    date[-1] = '20' + date[-1]
    date = '.'.join(date)
    #'27.06.23' ---> '27.06.2023'

    global sum_string
    column_with_sum = correct_data['I']
    for key, value in column_with_sum.items():
        if value == 'Всего по накладной:':
            sum_string = key
    sum = correct_data['Q'][sum_string]

    # seller_info
    #seller_inn = correct_data['B'][8].split('ИНН')[-1].split(' ')[1].strip()
    seller_inn = '1234567890'
    seller_name = correct_data['B'][8].split(':')[1].split('Адрес')[0].strip()
    seller_additionalId = ''
    #seller_kpp = correct_data['B'][8].split('КПП')[-1].split(' ')[1].strip()
    seller_kpp = '123451001'
    seller = dict(inn=seller_inn, name=seller_name, additionalId=seller_additionalId, kpp=seller_kpp)

    # buyer_info
    buyer_inn = correct_data['D'][7].split('Р/с')[0].split('ИНН:')[1].split(',')[0].strip()
    buyer_name = correct_data['D'][7].split('Р/с')[0].split(',')[0].split(',')[0].split(')')[-1].strip()
    buyer_additionalId = correct_data['D'][7].strip()
    buyer_additionalId = ''.join(buyer_additionalId).strip()
    buyer_kpp = correct_data['D'][7].split('Р/с')[0].split('КПП:')[1].split(',')[0].strip()
    buyer = dict(inn=buyer_inn, name=buyer_name, additionalId=buyer_additionalId, kpp=buyer_kpp)

    # строки между которыми находится номенклатура
    start_iter_string, last_iter_string = 0, 0

    column_with_start = correct_data['C']
    for key, value in column_with_start.items():
        if value == 'Наименование, характеристика, сорт, артикул':
            start_iter_string = key + 2
            break

    column_with_sum = correct_data['I']
    for key, value in column_with_sum.items():
        if value == 'Всего по накладной:':
            last_iter_string = key
            break

    example = {"name": "Крахмал кукурузный(весовой) кг",
               "code": "00-00000119",
               "measure": "кг",
               "amount": 2.0,
               "ndsPercent": 0.0,
               "ndsSum": 0.0,
               "totalSum": 184.8,
               "storeCode": ""}

    items = []
    for s in range(start_iter_string, last_iter_string):
        if correct_data['B'][s] == None:
            continue
        example.update(name=correct_data['C'][s],
                       code=correct_data['E'][s],
                       measure=correct_data['F'][s],
                       amount=correct_data['L'][s],
                       ndsPercent=correct_data['O'][s],
                       ndsSum=correct_data['P'][s],
                       totalSum=correct_data['Q'][s], )

        ndsPercent = float(example['ndsPercent'])
        example.update(ndsPercent=ndsPercent)
        items.append(example.copy())

    request_document.update(seller=seller,
                            buyer=buyer,
                            incomingNumber=incomingNumber,
                            ttnNumber="",
                            invoiceNumber="",
                            date=date,
                            sum=sum,
                            items=items
                            )

    return request_document







