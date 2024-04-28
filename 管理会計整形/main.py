# main.py
import configparser
from subcode import setup_logging, connect_to_spreadsheet, get_header_row, filter_target_rows, process_excel_files

# ロガーの設定
logger = setup_logging()

# 設定ファイルを読み込む
config = configparser.ConfigParser()
config.read('settings.ini', encoding="utf-8")

# 設定値を取得する
folder_path = config.get('ExcelSettings', 'folder_path')
files = config.get('ExcelSettings', 'files').split(';')
credentials_file = config.get('GoogleDriveAPI', 'credentials_file')

spreadsheet_id = '1chN8KzdfBlEe-ud_3R-PKoiySTimNNFrO2G8S4PYuHs'
sheet_name_moto = 'ファイル・シート一覧'
sheet_name_list = 'リスト'

# スプレッドシートに接続
client = connect_to_spreadsheet(credentials_file, spreadsheet_id)
spreadsheet = client.open_by_key(spreadsheet_id)  # スプレッドシートを開く
sheet = spreadsheet.worksheet(sheet_name_moto)  # ワークシートを取得

# ヘッダ行を取得
header = get_header_row(sheet)

# 対象シートの行をフィルタリング
target_rows = filter_target_rows(sheet, header)

# エクセルファイルを処理してデータを取得
process_excel_files(client, folder_path, files, target_rows, header, spreadsheet_id, sheet_name_list)
