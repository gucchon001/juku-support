import os
import openpyxl
import configparser
import logging
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False

# ロガーの設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 設定ファイルを読み込む
config = configparser.ConfigParser()
config.read('settings.ini', encoding="utf-8")

# 設定値を取得する
folder_path = config.get('ExcelSettings', 'folder_path')
credentials_file = config.get('GoogleDriveAPI', 'credentials_file')

logger.info(f"Folder path: {folder_path}")
logger.info(f"Credentials File: {credentials_file}")

# スプレッドシートとシートに書き出す
spreadsheet_id = '1chN8KzdfBlEe-ud_3R-PKoiySTimNNFrO2G8S4PYuHs'
sheet_name = 'ファイル・シート一覧'

# GoogleドライブAPIの認証を行う
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
client = gspread.authorize(creds)

# スプレッドシートを開く
spreadsheet = client.open_by_key(spreadsheet_id)
logger.info(f"Spreadsheet opened: {spreadsheet.title}")

# 指定されたシートを取得する
sheet = spreadsheet.worksheet(sheet_name)
logger.info(f"Sheet opened: {sheet.title}")

# ヘッダ行を取得する
header = sheet.row_values(1)
logger.info(f"Header: {header}")

# 指定されたフォルダ内のエクセルファイルを処理する
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        file_path = os.path.join(folder_path, file_name)
        
        logger.info(f"Processing file: {file_name}")
        logger.info(f"File path: {file_path}")
        
        try:
            # エクセルファイルを開く
            workbook = openpyxl.load_workbook(file_path)
            logger.info(f"Workbook opened: {file_path}")
        except Exception as e:
            logger.error(f"Error opening workbook: {str(e)}")
            continue
        
        # シート名を書き出す
        for sheet in workbook.worksheets:
            logger.info(f"Processing sheet: {sheet.title}")
            
            # 最終行に新しい行を追加する
            new_row = [file_name, sheet.title]
            sheet_instance = spreadsheet.worksheet(sheet_name)
            sheet_instance.append_row(new_row)
            logger.info(f"New row appended: {new_row}")