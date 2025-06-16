import xlwings as xw
import requests
from datetime import datetime

# === Configuration ===
api_key = '32558d25b55fb4c669d981d1'
currencies = ['AUD', 'BHD', 'CAD', 'CNY', 'INR', 'IDR', 'JPY', 'KRW', 'MYR', 'PKR',
              'QAR', 'SAR', 'CHF', 'AED', 'GBP', 'USD']
file_path = r"C:\Users\Administrator\Documents\final Project datacraft__excel.xlsx"

# === Fetch exchange rates ===
url = f'https://v6.exchangerate-api.com/v6/{api_key}/latest/USD'
response = requests.get(url)
data = response.json()

if data['result'] != 'success':
    raise Exception("Failed to retrieve exchange rates.")

today = datetime.today().strftime('%Y-%m-%d %H:%M:%S')

# === Open Excel ===
app = xw.App(visible=False)
wb = xw.Book(file_path)

# === Update 'Live Rates' ===
if 'Live Rates' not in [s.name for s in wb.sheets]:
    wb.sheets.add('Live Rates')
live = wb.sheets['Live Rates']
live.range('A1:D100').clear_contents()
live.range('A1').value = ['Date', 'Base Currency', 'Target Currency', 'Exchange Rate']

for i, curr in enumerate(currencies, start=2):
    rate = data['conversion_rates'].get(curr)
    live.range(f'A{i}').value = [today, 'USD', curr, rate]

live.range('G1').value = f"Last updated: {today}"

# === Setup 'Converter' ===
if 'Converter' not in [s.name for s in wb.sheets]:
    wb.sheets.add('Converter')
conv = wb.sheets['Converter']

# Labels
conv.range('B14').value = 'Base Currency'
conv.range('D14').value = 'Target Currency'
conv.range('F14').value = 'Amount'
conv.range('C24').value = 'Converted Amount'

# Clear previous values and validation
conv.range("B15").value = ''
conv.range("D15").value = ''
conv.range("F15").value = ''
conv.range("C25").value = ''
conv.range("B15").api.Validation.Delete()
conv.range("D15").api.Validation.Delete()

# Set dropdowns
currency_list = ",".join(currencies)
conv.range("B15").api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=currency_list)
conv.range("D15").api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=currency_list)

# Formula for conversion
conv.range("C25").formula = (
    '=IFERROR((VLOOKUP(D15, \'Live Rates\'!C:D, 2, FALSE) / '
    'VLOOKUP(B15, \'Live Rates\'!C:D, 2, FALSE)) * F15, "")'
)

# === Save & close ===
wb.save()
wb.close()
app.quit()

print("✅ Workbook updated with dropdowns and formula.")
