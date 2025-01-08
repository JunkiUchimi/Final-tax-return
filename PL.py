from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# 定数
BLUE_DEDUCTION = 550000  # 青色申告特別控除額
EXPENSE_FIELDS = [
    "租税公課", "水道光熱費", "旅費交通費", "通信費", "接待交際費",
    "損害保険料", "修繕費", "消耗品費", "減価償却費", "利子割引料",
    "車両費", "支払手数料", "新聞図書費", "会議費", "研修費", "雑費"
]

def calculate_expenses(values):
    """
    経費計を計算する関数

    Args:
        values (list): スプレッドシートのデータ

    Returns:
        int: 経費計
    """
    expenses = 0
    for row in values[1:]:
        if len(row) < 4:  # データが不足している場合スキップ
            continue
        field, amount = row[2], int(row[5].replace(",", ""))
        if field in EXPENSE_FIELDS:
            expenses += amount
    return expenses

def update_pl_sheet(service, spreadsheet_id):
    """
    Googleスプレッドシートの「経費」シートからデータを取得し、「PL」シートを更新する関数。

    Args:
        service (Resource): Google Sheets API サービスオブジェクト
        spreadsheet_id (str): スプレッドシートID
    """
    try:
        # 「経費」シートの全データを取得
        range_name = "経費!A1:F"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        values = result.get("values", [])

        if not values or len(values) < 2:
            print("「経費」シートが空、またはデータが不足しています。")
            return
        
        # ヘッダーを取得
        headers = values[0]

        # 各経費項目の合計を計算
        expense_sums = {field: 0 for field in EXPENSE_FIELDS}
        for row in values[1:]:
            if len(row) < 6:  # データが不足している場合スキップ
                continue
            field, amount = row[2], int(row[5].replace(",", ""))
            if field in EXPENSE_FIELDS:
                expense_sums[field] += amount

        # PLデータを計算
        sales = sum(
            int(row[5].replace(",", ""))
            for row in values[1:]
            if row[1] == "売上" and len(row) >= 6
        )
        total_expenses = sum(expense_sums.values())
        net_income_before_deduction = sales - total_expenses
        net_income_after_deduction = net_income_before_deduction - BLUE_DEDUCTION

        pl_data = {
            "売上": sales,
            "経費計": total_expenses,
            "青色申告特別控除額": BLUE_DEDUCTION,
            "差引金額": net_income_before_deduction,
            "所得金額": net_income_after_deduction,
        }

        # 「PL」シート全体を取得
        pl_range = "PL!A1:Z50"  # PLシートの全範囲を取得
        result_pl = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=pl_range
        ).execute()
        pl_values = result_pl.get("values", [])

        if not pl_values:
            print("「PL」シートが空です。")
            return

        # 全セルをスキャンしてヘッダーの位置を記録
        header_positions = {}
        for row_idx, row in enumerate(pl_values):
            for col_idx, cell in enumerate(row):
                if cell in EXPENSE_FIELDS or cell in pl_data:
                    header_positions[cell] = (row_idx + 1, col_idx + 1)  # 1始まりの位置

        # 値を「PL」シートに格納
        for header, value in {**expense_sums, **pl_data}.items():
            if header in header_positions:
                row_idx, col_idx = header_positions[header]
                cell_range = f"PL!{chr(65 + col_idx)}{row_idx}"  # ヘッダーの右隣のセル
                body = {"values": [[value]]}
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=cell_range,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()
            else:
                print(f"ヘッダー '{header}' が PL シートで見つかりません。")

    except Exception as e:
        print(f"更新中にエラーが発生しました: {e}")
