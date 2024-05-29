import requests
import xlsxwriter
import json

start_date = "2023-05-01"
end_date = "2024-05-01"
ticker = "SBERP"

# columns:
# 1  - date
# 2  - company name
# 3  - stock ticker
# 6  - open
# 7  - low
# 8  - high
# 9  - legal close
# 10 - warprice
# 11 - close
# 12 - volume
columns = [1, 11]

link = f"https://iss.moex.com/iss/history/engines/stock/markets/shares/boardgroups/57/securities/{ticker}.jsonp?lang=en&from={start_date}&till={end_date}&iss.json=extended&iss.meta=off&sort_order=TRADEDATE&sort_order_desc=desc&start="

resp = requests.get(link + "0")
data_json = resp.json()
record_count = data_json[1]['history.cursor'][0]['TOTAL']

data = data_json[1]['history']

for page in range(0, record_count, 100):
    resp = requests.get(link + str(page))
    data += resp.json()[1]['history']

workbook = xlsxwriter.Workbook(f"{ticker}.xlsx")
worksheet = workbook.add_worksheet()

keys = list(enumerate(data[0].keys()))
for col, key in [keys[x] for x in columns]:
    worksheet.write(0, col, key)

for row, obj in enumerate(data):
    for col, key in [keys[x] for x in columns]:
        worksheet.write(row + 1, col, str(obj[key]))

workbook.close()
