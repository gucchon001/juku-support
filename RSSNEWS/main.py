import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup
import openai
import configparser
import traceback
from slack_notify import send_slack_error_message
import xml.etree.ElementTree as ET
import time
from urllib.parse import urlparse, parse_qs
from datetime import datetime

config = configparser.ConfigParser()
config.read('settings.ini')

spreadsheet_id = config['GOOGLE_SHEETS']['SPREADSHEET_ID']
sheet_name = config['GOOGLE_SHEETS']['SHEET_NAME']
data_sheet_name = config['GOOGLE_SHEETS']['DATA_SHEET_NAME']
credentials_file = config['GOOGLE_SHEETS']['CREDENTIALS_FILE']
openai.api_key = config['OPENAI']['API_KEY']

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
client = gspread.authorize(credentials)

sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
data_sheet = client.open_by_key(spreadsheet_id).worksheet(data_sheet_name)
data = sheet.get_all_values()[1:]  # 2行目からデータを読み込む

def extract_text(url):
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    actual_url = query_params['url'][0] if 'url' in query_params else url
    
    time.sleep(1)  # 1秒のスリープを追加
    response = requests.get(actual_url, allow_redirects=True)
    print(f"Actual URL: {actual_url}")  # 実際のURLを表示
    soup = BeautifulSoup(response.text, 'html.parser')
    body = soup.find('body')
    if body:
        text = body.get_text(separator=' ')
        max_tokens = 16000
        if len(text) > max_tokens:
            text = text[:max_tokens]
        return text
    return ""

def summarize_text(text, max_tokens=15000):
    text_tokens = len(text)
    
    if text_tokens <= max_tokens:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "あなたは役立つアシスタントで、テキストを要約します。"},
                {"role": "user", "content": f"以下のテキストを200文字程度で日本語で要約してください:\n\n{text}"}
            ],
            max_tokens=200,
            n=1,
            stop=None,
            temperature=0.5,
        )
        summary = response.choices[0].message['content'].strip()
    else:
        # Split the text into two halves
        mid = text_tokens // 2
        first_half = text[:mid]
        second_half = text[mid:]
        
        # Recursively summarize each half
        first_summary = summarize_text(first_half, max_tokens)
        second_summary = summarize_text(second_half, max_tokens)
        
        # Combine the summaries and summarize again
        combined_summary = first_summary + " " + second_summary
        summary = summarize_text(combined_summary, max_tokens)
    
    return summary

def post_to_slack(summaries):
    try:
        message = "記事要約:\n\n" + "\n\n".join(summaries)
        webhook_url = config['Slack']['SLACK_WEBHOOK_URL']
        bot_name = config['Slack']['BOT_NAME']
        icon_emoji = config['Slack'].get('ICON_EMOJI', ':memo:')
        payload = {
            "text": message,
            "username": bot_name,
            "icon_emoji": icon_emoji
        }
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("Slack notification sent.")
    except requests.exceptions.RequestException as e:
        send_slack_error_message(e, config)
        raise

summaries = []
rows_to_append = []

for row in data:
    rss_url = row[1]  # B列からURLを読み取る
    
    try:
        print(f"Processing RSS URL: {rss_url}")
        response = requests.get(rss_url)
        root = ET.fromstring(response.text)
        
        for entry in root.findall('{http://www.w3.org/2005/Atom}entry'):
            link = entry.find('{http://www.w3.org/2005/Atom}link')
            title = entry.find('{http://www.w3.org/2005/Atom}title').text
            published = entry.find('{http://www.w3.org/2005/Atom}published').text
            print(f"Processing entry: {title}")
            
            if link is not None:
                url = link.get('href')
                print(f"Extracting text from: {url}")
                text = extract_text(url)
                if text:
                    print("Summarizing text...")
                    summary = summarize_text(text)
                    summaries.append(f"Title: {title}\nURL: {url}\nSummary: {summary}")
                    rows_to_append.append([title, url, published, summary, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                else:
                    print("No text found. Skipping...")
            else:
                print("No link found. Skipping...")
    except requests.exceptions.MissingSchema as e:
        print(f"Invalid URL: {rss_url}. Skipping...")
        continue

# Slackに要約結果を通知
post_to_slack(summaries)

# スプレッドシートに要約結果を書き込む
data_sheet.append_rows(rows_to_append)