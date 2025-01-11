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
        expenses.append([int(month), int(day), subject, apply, means, kind, amount])
    
    # 月日でソート
    return sorted(expenses, key=lambda x: (x[0], x[1]))

def journal(service, SPREADSHEET_ID):
    try:
        # 経費データを取得
        sorted_expenses = fetch_sorted_expenses(service, SPREADSHEET_ID)
        messagebox.showinfo("成功", "現金データを登録しました！")
        update_journal_sheet(service, SPREADSHEET_ID, sorted_expenses)
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def update_journal_sheet(service, spreadsheet_id, records, range_name="仕訳帳!B4:H"):
    journal_data = []
    for record in records:
            month, day, subject, apply, means, kind, amount = record
            # 新しい行を作成
            if kind == "経費":
                journal_data.append([month, day, subject, amount, means, amount, apply])
            elif kind == "売上":
                journal_data.append([month, day, means, amount, subject, amount, apply])

    # シートに書き込む
    body = {'values': journal_data}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

    