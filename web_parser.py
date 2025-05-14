import requests
import pandas as pd
import pygsheets
from newspaper import Article
import os
from dotenv import load_dotenv

load_dotenv()
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
SEARCH_ENGINE_ID = os.getenv("SEARCH_ENGINE_ID")
SHEET_TITLE = "關鍵字內容分析"
SERVICE_FILE = "service-account-key.json"

def google_search(query):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": GOOGLE_API_KEY,
        "cx": SEARCH_ENGINE_ID,
        "q": query,
        "num": 10
    }
    res = requests.get(url, params=params)
    return res.json().get("items", [])

def extract_article_text(url):
    try:
        article = Article(url)
        article.download()
        article.parse()
        return article.text
    except:
        return ""

def run():
    query = "滴雞精推薦"
    results = google_search(query)

    data = []
    for item in results:
        title = item.get("title", "")
        link = item.get("link", "")
        snippet = item.get("snippet", "")
        content = extract_article_text(link)
        data.append({
            "Title": title,
            "Link": link,
            "Snippet": snippet,
            "Content": content
        })

    df = pd.DataFrame(data)
    df.to_csv("滴雞精推薦_full_data.csv", index=False)
    print("CSV saved.")

    # Upload to Google Sheets
    gc = pygsheets.authorize(service_file=SERVICE_FILE)
    try:
        sh = gc.open(SHEET_TITLE)
    except pygsheets.SpreadsheetNotFound:
        sh = gc.create(SHEET_TITLE)
        sh.share('YOUR_EMAIL@gmail.com', role='writer')

    wks = sh.sheet1
    wks.set_dataframe(df, (1, 1))
    print("Data uploaded to Google Sheets.")

if __name__ == "__main__":
    run()
