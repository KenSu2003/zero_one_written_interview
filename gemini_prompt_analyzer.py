import os
import pygsheets
import pandas as pd
from dotenv import load_dotenv
import google.generativeai as genai

# Load environment
load_dotenv()
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Authorize Google Sheets access
gc = pygsheets.authorize(service_file='service-account-key.json')

# Open the exact sheet
sheet_id = '1glrObCE0Fhe6BqL9GDiFN-JY_lP8DmIc7WlE4RL-DNk'
sh = gc.open_by_key(sheet_id)
wks = sh.sheet1
df = wks.get_as_df(start='B2', include_tailing_empty=False)
print("Loaded columns:", df.columns.tolist())

# Check actual column names
print("Columns loaded:", df.columns.tolist())

# Rename if needed
df.columns = [col.strip() for col in df.columns]

# Build Gemini prompt from top 10 rows
top_keywords = df.head(10)
prompt_body = "\n".join([
    f"關鍵字：{row['Keyword']} | 點擊率：{row['CTR']} | 搜尋量：{row['Impressions']} | 平均排名：{row['Position']}"
    for _, row in top_keywords.iterrows()
])
prompt = (
    "以下是零一筆試的關鍵字模擬數據，請你：\n"
    "- 分析哪些關鍵字行銷潛力高\n"
    "- 建議放入內容中的方式與使用場景\n"
    "- 按優先順序排序推薦處理的關鍵字\n\n" + prompt_body
)

# Gemini call
def call_gemini_analysis(text):
    model = genai.GenerativeModel("models/gemini-2.5-pro-preview-05-06")
    response = model.generate_content(text)
    return response.candidates[0].content.parts[0].text

# Run analysis
result = call_gemini_analysis(prompt)
print("\nGemini 分析結果：\n")
print(result)

# Save back to the analysis sheet
sh = gc.open("關鍵字內容分析")

try:
    wks2 = sh.add_worksheet("Gemini分析結果")
    print("Created new worksheet: Gemini分析結果")
except Exception as e:
    if "already exists" in str(e):
        wks2 = sh.worksheet_by_title("Gemini分析結果")
        print("Worksheet already exists. Using existing one.")
    else:
        raise e  # re-raise other unexpected errors

# Clear and write output
wks2.clear(start='A1')
wks2.update_value("A1", "Gemini 分析摘要")
wks2.update_value("A2", result)
print("Gemini result written to Google Sheet.")
