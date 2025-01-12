from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from tkinter import messagebox

def fetch_sorted_expenses(service, spreadsheet_id, range_name="経費!A2:F"):
    """
    経費シートからソート済みデータを取得する
    
    Args:
        service: Google Sheets API サービスオブジェクト
        spreadsheet_id: スプレッドシートID
        range_name: データを取得する範囲 (既定値: 経費!A2:F)

    Returns:
        list: 月日でソートされた経費データ
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    values = result.get('values', [])
    if not values:
        return []

    # ソート用にデータを整形
    expenses = []
    for row in values:
        date, kind, subject, apply, means, amount = row
        year, month, day = date.split("/")
        amount = float(amount.replace(",", "")) if amount else 0  # 金額を数値に変換
        debit = amount if kind == "売上" else None
        credit = None if kind == "売上" else amount
        if means == "現金":
            expenses.append([int(month), int(day), subject, apply, debit, credit])
    
    # 月日でソート
    return sorted(expenses, key=lambda x: (x[0], x[1]))

def update_cash_sheet(service, spreadsheet_id, records, range_name="現金!B5:H"):
    """
    現金シートに経費データを追加する
    
    Args:
        service: Google Sheets API サービスオブジェクト
        spreadsheet_id: スプレッドシートID
        records: 経費データのリスト
        range_name: 書き込む範囲 (既定値: 現金!B5:H)
    """
    # 指定範囲を消去
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    # 前期繰越の残高を取得
    prev_balance_range = "現金!H4"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=prev_balance_range
    ).execute()
    prev_balance = float(result.get('values', [[0]])[0][0].replace(",", ""))

    # 現金データを準備
    cash_data = []
    for record in records:
        month, day, account, apply, debit, credit = record
        # 残高を計算
        if debit:
            prev_balance += debit
        elif credit:
            prev_balance -= credit

        # 新しい行を作成
        cash_data.append([month, day, account, apply, debit, credit, prev_balance])

    # シートに書き込む
    body = {'values': cash_data}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

def cash(service, SPREADSHEET_ID):
    try:
        # 経費データを取得
        sorted_expenses = fetch_sorted_expenses(service, SPREADSHEET_ID)

        # 現金シートを更新
        update_cash_sheet(service, SPREADSHEET_ID, sorted_expenses)
        messagebox.showinfo("成功", "現金データを登録しました！")
        print("現金シートの更新が完了しました。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
