import pygsheets
import pandas as pd

gc = pygsheets.authorize(service_file='credentials.json')
df = pd.read_csv("search_results_scraped.csv")

sh = gc.create('滴雞精搜尋結果（Scraped）')
wks = sh.sheet1
wks.set_dataframe(df, (1, 1))

print("✅ Uploaded to Google Sheets.")
