from Pdf_To_Text import pdf_To_text
import re

pdf_text = pdf_To_text(path='../exempleAudit/Royalty Summary-J Bereal (June 2019).pdf',
                       pages=[0])

rows = pdf_text.split('\n')

rows = [item.strip() for item in rows]
rows = [item for item in rows if item != '']

print(rows)

payee_account_number_idx = [i for i in range(len(rows)) if 'In Account with:' in rows[i]][0]
payee_account_number = rows[payee_account_number_idx].split(':')[-1].split()[-1][1:-1]

period_row = rows[[i for i in range(len(rows)) if 'for period' in rows[i]][0]].split()
period_start = period_row[-3]
period_end = period_row[-1]

royalties_idx = rows.index('TOTAL ROYALTIES') - 1
royalties = rows[royalties_idx]

original_currency_row = rows[[i for i in range(len(rows)) if 'BALANCE CARRIED FORWARD' in rows[i]][0]]
original_currency = re.search('[A-Z]*', original_currency_row).group(0)

print(period_start)
print(period_end)
print(royalties)
print(payee_account_number)
print(original_currency)

