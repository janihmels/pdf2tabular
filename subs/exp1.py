from Pdf_To_Text import pdf_To_text
import re

pdf_text = pdf_To_text(path='../exempleAudit/Wixen Music 006246 Qtr2 2018 Stmt.pdf',
                       pages=[0])

rows = pdf_text.split('\n')

rows = [item.strip() for item in rows]
rows = [item for item in rows if item != '']

print(rows)

payee_account_number = rows[0].split(':')[1].split()[0][1:-1]

period_idx = [i for i in range(len(rows)) if 'ForthePeriod:' in rows[i]][0]
period = rows[period_idx].split(':')[1]
to_index = period.index('to')
from_year = period[to_index-4: to_index]
from_month = period[0: to_index-4]
to_year = period[-4:]
to_month = period[to_index+2: -4]

period = from_month + ' ' + from_year + ' - ' + to_month + ' ' + to_year

royalties_row = rows[[i for i in range(len(rows)) if 'ROYLTSRoyaltiesforperiodending' in rows[i]][0]]
royalties = re.search("([0-9]*)['.']([0-9]*)", royalties_row.split()[-1]).group(0)

currency_row = rows[[i for i in range(len(rows)) if 'Balancethisperiod' in rows[i]][0]]
original_currency = currency_row.split(':')[1].strip()[0]

print(period)
print(royalties)
print(original_currency)

