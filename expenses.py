import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import os
from PL import update_pl_sheet
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials


# Google API設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'  # ダウンロードしたJSONファイルのパス
SPREADSHEET_ID = '10pGkyy7-qL6xMVdao97oxe-sE3nYedmtvyWGzb3GbFE'  # GoogleスプレッドシートのID
RANGE_NAME = '経費!A1:F'  # データを保存するシート名と範囲

# Google API認証
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# GUIの作成
root = tk.Tk()
root.title("データ入力フォーム")
# root.geometry("800x600")  # 必要に応じて幅と高さを調整

# 適用フィールドの選択肢
options_apply = [
    "作業着", "作業用品", "会食代", "手土産", "応接室", "事務室", "駐車料金",
    "鉄道運賃", "通行料金", "清掃用品", "事務用品", "車部品", "ケイコネクト", 
    "ガソリン代", "水道代", "ガス代", "電気代"
]
options_subject = [
    "消耗品費", "旅費交通費", "売上", "修繕費", "車両費",
    "接待交際費", "水道光熱費", "通信費", "利子割引料",
    "租税公課", "損害保険量", "会議費", "雑費", "事業主貸"
]
options_means = [ "現金", "普通預金" ]
options_kind = [ "経費", "事業主貸", "売上" ]

selected_option_apply = tk.StringVar(value=options_apply[0])  # 初期値を設定
selected_option_subject = tk.StringVar(value=options_subject[0])  # 初期値を設定
selected_option_means = tk.StringVar(value=options_means[0])  # 初期値を設定
selected_option_kind = tk.StringVar(value=options_kind[0])  # 初期値を設定

def create_radio_buttons(options, variable, row_start, column_start):
    for i, option in enumerate(options):
        tk.Radiobutton(
            root,
            text=option,
            variable=variable,
            value=option
        ).grid(row=row_start, column=column_start, padx=(10 + i * 95), pady=2, sticky="w")



# 日付フォーマット補完関数
def format_date(event):
    current_text = entry_date.get()
        # yyyy/mm/dd の形式で既に正しい場合はスルー
    if len(current_text) == 10 and current_text[:4].isdigit() and current_text[5:7].isdigit() and current_text[8:].isdigit() and current_text[4] == '/' and current_text[7] == '/':
        return
    
    if len(current_text) == 8 and current_text.isdigit():
        formatted_date = f"{current_text[:4]}/{current_text[4:6]}/{current_text[6:]}"
        entry_date.delete(0, tk.END)
        entry_date.insert(0, formatted_date)
    elif len(current_text) > 8:
        messagebox.showwarning("入力エラー", "日付は8桁の数字を入力してください！")
        entry_date.delete(0, tk.END)

# データを保存する関数
def save_data():
    global last_selected_item

    try:
        # 入力フィールドからデータを取得
        date = entry_date.get()
        kind = selected_option_kind.get()
        subject = selected_option_subject.get()
        apply = selected_option_apply.get()
        means = selected_option_means.get()

        # 金額を取得し、整形して数値に変換
        amount_text = entry_amount.get()
        amount_text = amount_text.replace(",", "")  # 「,」を削除
        try:
            amount = int(amount_text)
        except ValueError:
            messagebox.showwarning("入力エラー", "金額には数値を入力してください！")
            return

        # 保存データを準備
        values = [[date, kind, subject, apply, means, amount]]
        body = {'values': values}

        # 修正モード（既存行を上書き）
        if last_selected_item is not None:
            try:
                # Treeviewからインデックスを取得
                item_index = tree.index(last_selected_item) + 1  # ヘッダーを考慮
                update_range = f"経費!A{item_index + 1}:F{item_index + 1}"  # 対象行を指定

                # Googleスプレッドシートのデータを上書き
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()

                messagebox.showinfo("成功", "データを修正しました！")
            except Exception as e:
                messagebox.showerror("エラー", f"Googleスプレッドシートへの修正中にエラーが発生しました: {e}")
                return
        else:
            # 新規データの追加
            try:
                service.spreadsheets().values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range=RANGE_NAME,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()
                messagebox.showinfo("成功", "データを追加しました！")
            except Exception as e:
                messagebox.showerror("エラー", f"Googleスプレッドシートへの保存中にエラーが発生しました: {e}")
                return

        # フィールドをリセット
        entry_date.delete(0, tk.END)
        entry_amount.delete(0, tk.END)
        reset_fields()
        refresh_table()

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {e}")


# データを表示する関数
def refresh_table():
    """
    スプレッドシートのデータを取得してUIに表示する関数
    """
    try:
        # Googleスプレッドシートからデータを取得
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="経費!A1:F"  # 範囲を正確に指定
        ).execute()

        # データの取得結果を確認
        values = result.get("values", [])

        # データが空の場合
        if not values:
            messagebox.showwarning("データ取得エラー", "スプレッドシートが空です。")
            return
        
        # Treeview の列を定義
        tree["columns"] = ["月日", "取引分類", "科目", "適用", "取引手段", "金額"]
        
        # 各列の設定
        for col in tree["columns"]:
            tree.heading(col, text=col)  # ヘッダーを設定
            tree.column(col, anchor="center", width=150)  # 幅を調整

        # Treeview をクリア
        for row in tree.get_children():
            tree.delete(row)

        # データをTreeviewに挿入
        for row in values[1:]:  # ヘッダー行をスキップ
            tree.insert("", "end", values=row)

        # PLシートを更新
        update_pl_sheet(service, SPREADSHEET_ID)

    except Exception as e:
        # エラー内容を表示
        print("エラー詳細:", e)
        messagebox.showerror("エラー", f"データの取得中にエラーが発生しました: {e}")



def reset_fields():
    entry_date.delete(0, tk.END)
    entry_date.insert(0, "2024")
    entry_amount.delete(0, tk.END)
    selected_option_kind.set(options_kind[0])
    selected_option_subject.set(options_subject[0])
    selected_option_apply.set(options_apply[0])
    selected_option_means.set(options_means[0])
    entry_date.focus()

last_selected_item = None
def load_selected_record(event):
    """
    Treeviewで選択された行のデータをGoogleスプレッドシートから取得し、
    入力フィールドに表示する関数。

    Args:
        event: Treeviewの選択イベント。
    """
    global last_selected_item

    # 現在選択されているTreeviewのアイテム
    selected_item = tree.selection()

    # 同じアイテムを選択した場合は選択を解除し、フィールドをリセット
    if selected_item and selected_item[0] == last_selected_item:
        tree.selection_remove(selected_item[0])
        last_selected_item = None

        # フィールドをリセット
        reset_fields()
        return

    # 新しいアイテムが選択された場合、その内容を入力フィールドに表示
    if selected_item:
        last_selected_item = selected_item[0]  # 現在の選択を記録
        item_index = tree.index(selected_item[0])
        item_index += 1

        try:
            # Googleスプレッドシートからデータを取得
            result = service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
            values = result.get('values', [])
            # 選択された行のデータを取得
            if 0 <= item_index:
                record = values[item_index]

                # フィールドにデータを設定
                entry_date.delete(0, tk.END)
                entry_date.insert(0, record[0])  # 月日
                entry_amount.delete(0, tk.END)
                entry_amount.insert(0, record[5])  # 金額
                selected_option_kind.set(record[1])  # 取引分類
                selected_option_subject.set(record[2])  # 科目
                selected_option_apply.set(record[3])  # 適用
                selected_option_means.set(record[4])  # 取引手段
            else:
                messagebox.showerror("エラー", "選択されたデータがスプレッドシートに存在しません。")

        except Exception as e:
            messagebox.showerror("エラー", f"Googleスプレッドシートからデータを取得できませんでした: {e}")



def delete_data():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("削除エラー", "削除するデータを選択してください！")
        return

    try:
        # スプレッドシートのデータを取得
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])

        if not values:
            messagebox.showwarning("削除エラー", "スプレッドシートが空です。")
            return

        # 選択された行を削除（末尾から処理してインデックスのズレを防ぐ）
        for item in sorted(selected_item, key=tree.index, reverse=True):
            item_index = tree.index(item) + 1  # ヘッダー行を考慮
            if 0 <= item_index < len(values):
                values.pop(item_index)
            else:
                messagebox.showerror("エラー", "選択されたデータがスプレッドシートに存在しません。")
                return

        # 削除後のデータを更新
        if not values:  # データが空になった場合
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
        else:
            # スプレッドシート全体を一度クリアして新しいデータを書き込む
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()

            # 削除後のデータを書き込み
            body = {'values': values}
            updated_range = f"経費!A1:F{len(values)}"
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=updated_range,
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()

        # UIを更新
        refresh_table()
        messagebox.showinfo("成功", "データを削除しました！")

    except Exception as e:
        messagebox.showerror("エラー", f"データの削除中にエラーが発生しました: {e}")



# 各ラベルとエントリー
labels = ["月日", "金額"]
entries = []

# 課税所得を表示するラベル
taxable_income_label = tk.Label(root, text="課税所得: 計算中...", font=("Arial", 12))
taxable_income_label.grid(row=len(labels) + 11, column=0, columnspan=2, pady=(5, 10), padx=(20, 0), sticky = "w")

def update_taxable_income_label():
    try:
        # Googleスプレッドシートからデータを取得
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])
        
        # 売上と経費の合計を計算
        売上 = sum(int(row[5]) for row in values if row[1] == "売上")  # 取引分類が「売上」の金額を合計
        経費 = sum(int(row[5]) for row in values if row[1] == "経費")  # 取引分類が「経費」の金額を合計
        taxable_income = 売上 - 経費

        # ラベルを更新
        taxable_income_label.config(text=f"課税所得: {taxable_income}円")
    except Exception as e:
        taxable_income_label.config(text="課税所得: エラー発生")


# 既存のラベル設定を左詰めに変更
for i, label in enumerate(labels):
    tk.Label(root, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
    entry = tk.Entry(root, width=30)
    entry.grid(row=i, column=1, padx=12, pady=5, sticky="w")
    entries.append(entry)

# 日付と金額フィールドを取得
entry_date, entry_amount = entries

# 日付フィールドの初期値を設定
entry_date.insert(0, "2024")
entry_date.bind("<FocusOut>", format_date)

# 適用、科目、取引分類、取引手段のラベルを追加
tk.Label(root, text="適用").grid(row=len(labels), column=0, padx=10, pady=5, sticky="w")
# 最初の6つの選択肢を配置
create_radio_buttons(options_apply[:6], selected_option_apply, row_start=len(labels), column_start=1)
create_radio_buttons(options_apply[6:12], selected_option_apply, row_start=len(labels) + 1, column_start=1)
create_radio_buttons(options_apply[12:], selected_option_apply, row_start=len(labels) + 2, column_start=1)
tk.Label(root, text="科目").grid(row=len(labels) + 3, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_subject[:6], selected_option_subject, row_start=len(labels) + 3, column_start=1)
create_radio_buttons(options_subject[6:12], selected_option_subject, row_start=len(labels) + 4, column_start=1)
create_radio_buttons(options_subject[12:], selected_option_subject, row_start=len(labels) + 5, column_start=1)
tk.Label(root, text="取引分類").grid(row=len(labels) + 6, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_kind, selected_option_kind, row_start=len(labels) + 6, column_start=1)
tk.Label(root, text="取引手段").grid(row=len(labels) + 7, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_means, selected_option_means, row_start=len(labels) + 7, column_start=1)

# 保存ボタン
save_button = tk.Button(root, text="データを追加/修正", command=save_data, width=15)
save_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(20, 5), padx=(40, 0), sticky="w")

# 削除ボタン
delete_button = tk.Button(root, text="データを削除", command=delete_data, width=15)
delete_button.grid(row=len(labels) + 10, column=0, columnspan=2, pady=(5, 20), padx=(40, 0), sticky="w")

# データ表示用のTreeview
# データ表示用のFrameを作成
tree_frame = tk.Frame(root)
tree_frame.grid(row=len(labels) + 12, column=0, columnspan=2, padx=10, pady=10, sticky="w")

# Scrollbarを追加
tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scrollbar.pack(side="right", fill="y")

# データ表示用のTreeview
columns = ["月日", "取引分類", "科目", "適用", "取引手段", "金額"]
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=16, yscrollcommand=tree_scrollbar.set)
tree.pack(side="left", fill="both", expand=True)

# ScrollbarをTreeviewにリンク
tree_scrollbar.config(command=tree.yview)



# 各列のヘッダーと幅を設定
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)


# Treeviewに選択イベントをバインド
tree.bind("<<TreeviewSelect>>", load_selected_record)

# 初期データの表示
refresh_table()

# メインループの開始
root.mainloop()