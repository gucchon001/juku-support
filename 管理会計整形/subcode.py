# subcode.py
import openpyxl
import logging
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import os

# ロガー設定の関数
def setup_logging():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    return logging.getLogger(__name__)

# スプレッドシートに接続する関数
def connect_to_spreadsheet(credentials_file, spreadsheet_id):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)
    return client

# ヘッダ行を取得する関数
def get_header_row(sheet):
    return sheet.row_values(1)

# 対象シートの行をフィルタリングする関数
def filter_target_rows(sheet, header):
    rows = sheet.get_all_values()
    return [row for row in rows if row[header.index('対象シート')] == 'TRUE']

# エクセルファイルを処理してデータを取得し、出力する関数
def process_excel_files(client, folder_path, files, target_rows, header, spreadsheet_id, target_sheet_name):
    # clientを使用してスプレッドシートを開く
    spreadsheet = client.open_by_key(spreadsheet_id)
    for file_name in os.listdir(folder_path):
        if file_name in files and (file_name.endswith(".xlsx") or file_name.endswith(".xls")):
            file_path = os.path.join(folder_path, file_name)
            try:
                workbook = openpyxl.load_workbook(file_path, data_only=True)
                for sheet in workbook.worksheets:
                    if any(row[header.index('シート名')] == sheet.title for row in target_rows):
                        print(f"ファイル: {file_name}, シート: {sheet.title}")
                        unpivoted_data = unpivot_excel_data(file_name, workbook, sheet.title)  # file_name引数を追加
                        update_spreadsheet_data(spreadsheet, target_sheet_name, unpivoted_data)
            except Exception as e:
                logging.error(f"ワークブックを開く際にエラーが発生しました: {str(e)}")
                print(f"ワークブックを開く際にエラーが発生しました: {str(e)}")

'''
#2021以前用
def unpivot_excel_data(file_name, workbook, sheet_name):
    sheet = workbook[sheet_name]
    data = list(sheet.values)
    unpivoted_data = []
    months = data[1][1:]  # 月情報を取得
    # データ部分のみ処理、最大130行まで
    for i in range(2, min(len(data), 131)):
        account_item = data[i][0]  # 勘定項目
        row_number = i + 1  # 行番号を取得
        formatted_account_item = f"{row_number}_{account_item}"  # 行番号を付与した勘定科目
        for j in range(1, len(data[1])):  # 月情報開始列からデータ取得
            month = months[j-1]  # 月
            if "月" not in str(month):  # 月の列が"月"を含まない場合はスキップ
                continue
            value = data[i][j]  # 値
            if value is not None and value != '':  # 値が存在する場合のみ追加
                unpivoted_data.append([file_name, sheet_name, formatted_account_item, month, value])
    return unpivoted_data
'''

#2022以降用
def unpivot_excel_data(file_name, workbook, sheet_name):
    print(f"Processing file: {file_name}, sheet: {sheet_name}")
    sheet = workbook[sheet_name]
    data = list(sheet.values)
    print(f"Raw data: {data}")
    unpivoted_data = []
    months = data[1][1:]  # 月情報を取得（B2から）
    print(f"Months: {months}")
    # データ部分のみ処理（B3から）、最大130行まで
    for i in range(2, min(len(data), 132)):
        account_item = data[i][0]  # 勘定項目
        row_number = i + 1  # 行番号を取得
        formatted_account_item = f"{row_number}_{account_item}"  # 行番号を付与した勘定科目
        print(f"Account item: {formatted_account_item}")
        for j in range(1, len(data[i])):  # 月情報開始列からデータ取得
            month = months[j-1]  # 月
            print(f"Month: {month}")
            if not isinstance(month, (int, float)):  # 月情報が数値でない場合はスキップ
                print(f"Skipping column {j} (not a numeric month)")
                continue
            month_str = str(int(month)) + "月"  # 数値を文字列に変換し、"月" を追加
            value = data[i][j]  # 値
            print(f"Value: {value}")
            if value is not None and value != '':  # 値が存在する場合のみ追加
                unpivoted_data.append([file_name, sheet_name, formatted_account_item, month_str, value])
                print(f"Added data: {[file_name, sheet_name, formatted_account_item, month_str, value]}")
            else:
                print(f"Skipping empty value at row {i}, column {j}")
    print(f"Unpivoted data: {unpivoted_data}")
    return unpivoted_data


def update_spreadsheet_data(spreadsheet, target_sheet_name, data_to_update):
    sheet = spreadsheet.worksheet(target_sheet_name)
    last_row = len(sheet.col_values(1))  # 最終行を取得
    max_rows = sheet.row_count  # シートの最大行数を取得
    remaining_rows = max_rows - last_row  # 残りの行数を計算
    
    if remaining_rows >= len(data_to_update):
        range_to_update = f"A{last_row+1}"  # 最終行の次の行から更新開始
        sheet.update(range_to_update, data_to_update)  # データ更新
        print(f"スプレッドシート'{target_sheet_name}'を更新しました。")
    else:
        print(f"スプレッドシート'{target_sheet_name}'の行数が上限に達しているため、データを追加できません。")
        print(f"追加可能な行数: {remaining_rows}, 追加しようとした行数: {len(data_to_update)}")