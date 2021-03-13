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
            if len(row) == 6 and type(row[0].value) is str:
                try:
                    date = datetime.datetime.strptime(row[0].value, '%d.%m.%Y')
                    tokens = row[4].value.split(' ')
                    amount = ''.join(tokens[:-1])
                    currency = tokens[-1]
                    records.append((Record(
                        date=date, 
                        amount=float(amount),
                        currency=currency,
                        comment=row[2].value
                    )))
                except ValueError:
                    continue
        
        records.sort(key=lambda x: x.date)
        
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        for record in records:
            worksheet.append([record.date, record.comment, record.amount, record.currency])

    workbook.save(filename=f"{datetime.datetime.now()}.xlsx")

process()