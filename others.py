from tkinter import messagebox

def fetch_sorted_expenses(service, spreadsheet_id, subjectif, range_name="経費!A2:F"):
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
    
    # subjectif_list が単一文字列の場合リストに変換
    if isinstance(subjectif, str):
        subjectif = [subjectif]

    # ソート用にデータを整形
    expenses = []
    for row in values:
        date, kind, subject, apply, means, amount = row
        year, month, day = date.split("/")
        amount = float(amount.replace(",", "")) if amount else 0  # 金額を数値に変換
        if subject in subjectif:
            expenses.append([int(month), int(day), subject, apply, means, kind, amount])
    
    # 月日でソート
    return sorted(expenses, key=lambda x: (x[0], x[1]))

def others(service, SPREADSHEET_ID, subjectif, range_name):
    # range_name="事業主貸・売上!A4:G"
    # subjectif="事業主貸"
    try:
        # 経費データを取得
        sorted_expenses = fetch_sorted_expenses(service, SPREADSHEET_ID, subjectif)
        # messagebox.showinfo("成功", "現金データを登録しました！")
        update_others_sheet(service, SPREADSHEET_ID, sorted_expenses, range_name)
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def update_others_sheet(service, spreadsheet_id, records, range_name):
    # 指定範囲を消去
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    others_data = []
    balance = 0
    for record in records:
            month, day, subject, apply, means, kind, amount = record
            balance += amount
            # 新しい行を作成
            others_data.append([month, day, subject, apply, amount, amount, balance])
    others_data.append(["総計", "", "", "", balance, "", ""])
    # シートに書き込む
    body = {'values': others_data}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    