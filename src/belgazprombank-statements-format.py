"""
Выписки белгазпромбанка генерируются в HTML формате.
Таблицу выписки можно скопировать и вставить в LibreOffice Calc, но данные\яйчейки будут иметь неккоректный формат.
Для форматирования данных и сортировки транзакций по датам в корректном порядке был написан этот скрипт.

## Usage
```sh
pipenv run python belgazprombank-stateents-format.py *.xlsx
```

## Dependencies
```sh
apt install -yq pipenv
```

"""

import sys
import openpyxl
import datetime
import dataclasses

@dataclasses.dataclass
class Record:
    date: datetime.datetime
    amount: float
    currency: str
    comment: str

def process():
    records = []

    for file_path in sys.argv[1:]:
        workbook = openpyxl.load_workbook(filename=file_path)
        worksheet = workbook.active
        for row in worksheet:
            # 0 Дата операции 
            # 1 Дата отражения 
            # 2 Операция
            # 3 Тип операции 	
            # 4 Сумма в валюте операции 	
            # 5 Валюта операции 	
            # 6 Сумма в валюте счета 	
            # 7 Валюта счета 	
            # 8 Место операции (страна, наименование точки, город) 	
            # 9 Код авторизации 	
            # 10 МСС 	
            # 11 Вознаграждение клиенту по операции в валюте счета
            if len(row) >= 11 and type(row[0].value) is str:
                try:
                    column = map(lambda x: x.value, row)
                    date1, date2, comment, operation_type, amount1, currency1, amount2, currency2, location, code, mcc, _ = column
                    
                    date = datetime.datetime.strptime(date1, '%d.%m.%Y %H:%M')
                    sign = -1 if operation_type == "СПИСАНИЕ" else 1
                
                    amount = sign * float(amount2.replace(' ', '').replace(',', '.'))
                    currency = currency2
                    
                    records.append((Record(
                        date=date, 
                        amount=amount,
                        currency=currency,
                        comment=comment
                    )))

                    if len(row) > 11:
                        cashback = row[11].value
                        if type(cashback) is str:
                            cashback_f = float(cashback.replace(',', '.'))
          
                            date = datetime.datetime.strptime(date2, '%d.%m.%Y')
                            records.append(Record(
                                date=date,
                                amount=cashback_f,
                                currency=currency,
                                comment="Cashback"
                            ))
                except ValueError as e:
                    print(e)
                    continue
        
        records.sort(key=lambda x: x.date)
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        for record in records:
            worksheet.append([record.date, record.comment, record.amount, record.currency])

    workbook.save(filename=f"{datetime.datetime.now()}.xlsx")

process()
