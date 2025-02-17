import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import os
import threading
import math
from PL import update_pl_sheet
from openpyxl import load_workbook
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime
from utils import on_apply_change, show_auto_closing_popup
from cash import cash
from journal import journal
from others import others
import threading

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
default_font = ("Arial", 11)  # フォントサイズを11ポイントに設定

# 適用フィールドの選択肢
options_apply = [
    "会食代", "作業着", "作業用品", "手土産", "応接室", "事務室", "携帯料金", "駐車料金",
    "鉄道運賃", "通行料金", "清掃用品", "事務用品", "車部品", "インターネット料金", "ケイコネクト", 
    "ガソリン代", "水道代", "ガス代", "電気代"
]
# 科目フィールドの選択肢
options_subject = [
    "接待交際費", "消耗品費", "旅費交通費", "売上", "修繕費", 
    "車両費","水道光熱費", "通信費", "利子割引料",
    "租税公課", "損害保険料", "会議費", "雑費", "事業主貸"
]
# 取引手段フィールドの選択肢
options_means = [ "現金", "普通預金" ]
# 取引分類フィールドの選択肢
options_kind = [ "経費", "事業主貸", "売上" ]
# 適用フィールドに「その他」を追加
options_apply.append("その他")
options_subject.append("その他")

selected_option_apply = tk.StringVar(value=options_apply[0])  # 初期値を設定
selected_option_subject = tk.StringVar(value=options_subject[0])  # 初期値を設定
selected_option_means = tk.StringVar(value=options_means[0])  # 初期値を設定
selected_option_kind = tk.StringVar(value=options_kind[0])  # 初期値を設定

# グローバル変数として前回の選択状態を保持
previous_selection = {
    "apply": options_apply[0],  # 初期値
    "subject": options_subject[0],
    "means": options_means[0],
    "kind": options_kind[0]
}

def create_radio_buttons(options, variable, row_start, column_start):
    for i, option in enumerate(options):
        tk.Radiobutton(
            root,
            text=option,
            variable=variable,
            value=option,
            font=default_font  # フォントを指定
        ).grid(row=row_start, column=column_start, padx=(10 + i * 110), pady=2, sticky="w")

def sort_by_column(column_name):
    global sort_state

    try:
        # 現在のデータを取得
        rows = []
        for child in tree.get_children(''):
            value = tree.set(child, column_name)
            try:
                # 日付に変換可能か確認
                if column_name == "月日":
                    parsed_value = datetime.strptime(value, "%Y/%m/%d")
                else:
                    parsed_value = value
            except ValueError:
                parsed_value = value  # 日付でない場合そのまま
            rows.append((parsed_value, child))

        if not rows:
            return

        # 列が現在のソート列かを確認
        if sort_state["column"] == column_name:
            # ソート順を変更する
            if sort_state["order"] == "asc":
                sort_state["order"] = "desc"
                rows.sort(reverse=True, key=lambda x: x[0])
            elif sort_state["order"] == "desc":
                sort_state["order"] = None  # 登録順に戻す
                tree.delete(*tree.get_children())
                for row in original_data:
                    tree.insert("", "end", values=row)
                return
            else:
                sort_state["order"] = "asc"
                rows.sort(key=lambda x: x[0])
        else:
            # 新しい列でソートする
            sort_state["column"] = column_name
            sort_state["order"] = "asc"
            rows.sort(key=lambda x: x[0])

        # ソート後のデータをTreeviewに再挿入
        for index, (_, child) in enumerate(rows):
            tree.move(child, '', index)
    except Exception as e:
        print(f"ソート中にエラーが発生しました: {e}")

def bind_enter_to_save():
    """
    Enterキーをデータ追加/修正ボタンにバインドする
    """
    root.bind("<Return>", lambda event: save_data())

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
    global previous_selection  # 前回の選択状態を保持する変数
    format_date(None)  # 日付フォーマットを適用
    try:
        # 入力フィールドからデータを取得
        date = entry_date.get()
        kind = selected_option_kind.get()
        apply = apply_entry.get() if selected_option_apply.get() == "その他" else selected_option_apply.get()
        subject = subject_entry.get() if selected_option_subject.get() == "その他" else selected_option_subject.get()
        means = selected_option_means.get()

        # 金額を取得して数値に変換
        amount_text = entry_amount.get().replace(",", "")  # 「,」を削除

        try:
            amount = int(amount_text)
            if subject == "水道光熱費" or subject == "通信費":
                if apply == "携帯料金":
                    pass
                else:
                    amount = math.floor(int(amount_text) * 0.3)
        except ValueError:
            messagebox.showwarning("入力エラー", "金額には数値を入力してください！")
            return

        # 保存するデータを準備
        values = [[date, kind, subject, apply, means, amount]]
        body = {'values': values}

        # 修正モード（既存行を上書き）
        if last_selected_item is not None:
            try:
                # Treeviewからインデックスを取得
                item_index = tree.index(last_selected_item) + 1  # ヘッダーを考慮
                spre_index = len(original_data) - item_index + 1
                print(item_index)
                print(spre_index)
                update_range = f"経費!A{spre_index + 1}:F{spre_index + 1}"  # 対象行を指定

                # Googleスプレッドシートのデータを上書き
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()

                show_auto_closing_popup(root, "成功", "データを追加しました！")            
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
                show_auto_closing_popup(root, "成功", "データを追加しました！")
            except Exception as e:
                messagebox.showerror("エラー", f"Googleスプレッドシートへの保存中にエラーが発生しました: {e}")
                return

        # 保存が成功した場合、現在の選択状態を保持
        previous_selection = {
            "apply": selected_option_apply.get(),
            "subject": selected_option_subject.get(),
            "means": selected_option_means.get(),
            "kind": selected_option_kind.get()
        }

        # フィールドをリセット
        reset_fields()

        # テーブルを更新
        refresh_table()

        # 選択状態をクリア
        last_selected_item = None

        # ラベルを「計算中」に設定
        taxable_income_label.config(text="課税所得: 計算中...")
        lock = threading.Lock()
        def run_background_tasks():
            with lock:  # このブロック内の処理は他のスレッドと競合しない
                try:
                    update_pl_sheet(service, SPREADSHEET_ID)  # PLシート更新
                    update_taxable_income_label_from_pl(service, SPREADSHEET_ID)  # 課税所得ラベル更新
                except Exception as e:
                    print(f"エラー: {e}")

        def save_data():
            if lock.locked():
                messagebox.showwarning("処理中", "前回の登録が完了するまでお待ちください。")
                return

            threading.Thread(target=run_background_tasks, daemon=True).start()

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {e}")


# データを表示する関数
def refresh_table():
    """
    スプレッドシートのデータを取得してUIに表示する関数
    """
    global original_data  # 登録順データを保持するグローバル変数を使用

    try:
        # Googleスプレッドシートからデータを取得
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="経費!A1:F"  # 範囲を正確に指定
        ).execute()

        # データの取得結果を確認
        values = result.get("values", [])
        
        # データが空の場合
        if not values or len(values) <= 1:  # データが空か、ヘッダーのみの場合
            original_data = []  # 登録順データも空にする
            tree.delete(*tree.get_children())  # Treeviewをクリア
            messagebox.showwarning("データ取得エラー", "スプレッドシートが空です。")
            return
        
        # 登録順データを保存（ヘッダーをスキップ）
        original_data = values[1:]  # ヘッダー行を除いたデータを保持

        # Treeview の列を定義
        tree["columns"] = ["月日", "取引分類", "科目", "適用", "取引手段", "金額"]
        
        # 各列の設定
        for col in tree["columns"]:
            tree.heading(col, text=col)  # ヘッダーを設定
            tree.column(col, anchor="center", width=150)  # 幅を調整

        # Treeview をクリア
        tree.delete(*tree.get_children())  # Treeviewの既存データをすべて削除

        # データをTreeviewに挿入
        for row in reversed(original_data):
            tree.insert("", "end", values=row)
            
        # ツリービューのフォント設定
        style = ttk.Style()
        style.configure("Treeview", font=default_font)  # 本体のフォント
        style.configure("Treeview.Heading", font=default_font)  # ヘッダーのフォント

    except Exception as e:
        # エラー内容を表示
        print("エラー詳細:", e)
        messagebox.showerror("エラー", f"データの取得中にエラーが発生しました: {e}")


def reset_fields():
    entry_date.delete(0, tk.END)
    entry_date.insert(0, "2024")  # 必要に応じて変更
    entry_amount.delete(0, tk.END)
    # グローバル変数から選択状態を設定
    selected_option_apply.set(previous_selection["apply"])
    selected_option_subject.set(previous_selection["subject"])
    selected_option_means.set(previous_selection["means"])
    selected_option_kind.set(previous_selection["kind"])
    entry_date.focus()


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

    print(f"Selected items: {selected_item}")  # 選択アイテムを表示

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
        spreadsheet_index = len(original_data) - item_index  # スプレッドシートの正しいインデックスを計算


        try:
            # Googleスプレッドシートからデータを取得
            result = service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
            values = result.get('values', [])
            # 選択された行のデータを取得
            if 0 <= spreadsheet_index:
                record = values[spreadsheet_index]

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
        # Googleスプレッドシートのデータを取得
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])

        if not values:
            messagebox.showwarning("削除エラー", "スプレッドシートが空です。")
            return

        # 選択された行のスプレッドシート内の正しいインデックスを取得
        indexes_to_delete = []
        for item in selected_item:
            tree_index = tree.index(item)  # Treeview上でのインデックス
            spreadsheet_index = len(original_data) - tree_index  # スプレッドシートのインデックス調整
            if 0 <= spreadsheet_index < len(values):
                indexes_to_delete.append(spreadsheet_index)

        # インデックスを降順で削除（リストのズレを防ぐ）
        indexes_to_delete.sort(reverse=True)
        for index in indexes_to_delete:
            del values[index]

        # スプレッドシートを更新
        if not values:  # すべてのデータが削除された場合
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
        else:
            # スプレッドシートを一度クリアして新しいデータを書き込む
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()

            # 更新後のデータをスプレッドシートに書き込み
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

def update_proprietor_and_sales():
    # "事業主貸" の処理
    try:
        others(service, SPREADSHEET_ID, subjectif="事業主貸", range_name="事業主貸・売上!B21:H")
        others(service, SPREADSHEET_ID, subjectif="消耗品費", range_name="消耗品費!B4:H")
        others(service, SPREADSHEET_ID, subjectif="旅費交通費", range_name="旅費・接待・研修!B4:H")
        others(service, SPREADSHEET_ID, subjectif="接待交際費", range_name="旅費・接待・研修!J4:P")
        others(service, SPREADSHEET_ID, subjectif="研修費", range_name="旅費・接待・研修!R4:X")
        others(service, SPREADSHEET_ID, subjectif="会議費", range_name="会議費・修繕費・新聞・利子!B4:H")
        others(service, SPREADSHEET_ID, subjectif="新聞図書費", range_name="会議費・修繕費・新聞・利子!J4:P")
        others(service, SPREADSHEET_ID, subjectif="修繕費", range_name="会議費・修繕費・新聞・利子!R4:X")
        others(service, SPREADSHEET_ID, subjectif="利子割引料", range_name="会議費・修繕費・新聞・利子!Z4:AF")
        others(service, SPREADSHEET_ID, subjectif="租税公課", range_name="租税・雑費・水道・更新!B4:H")
        others(service, SPREADSHEET_ID, subjectif="雑費", range_name="租税・雑費・水道・更新!J4:P")
        others(service, SPREADSHEET_ID, subjectif="水道光熱費", range_name="租税・雑費・水道・更新!R4:X")
        others(service, SPREADSHEET_ID, subjectif="通信費", range_name="租税・雑費・水道・更新!Z4:AF")
        others(service, SPREADSHEET_ID, subjectif="車両費", range_name="車両・損保・減価!B4:H")
        others(service, SPREADSHEET_ID, subjectif="損害保険料", range_name="車両・損保・減価!J4:P")
        others(service, SPREADSHEET_ID, subjectif="減価償却費", range_name="車両・損保・減価!R4:X")
        messagebox.showinfo("成功", "その他シートを更新しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"Googleスプレッドシートへのデータ登録中にエラーが発生しました: {e}")

def update_taxable_income_label_from_pl(service, spreadsheet_id):
    """
    PLシートから課税所得を取得し、ラベルに反映する。

    Args:
        service: Google Sheets API サービスオブジェクト
        spreadsheet_id: スプレッドシートID
    """
    try:
        # PLシートのデータを取得
        pl_range = "PL!A1:K20"  # PLシートの範囲
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=pl_range
        ).execute()
        values = result.get("values", [])

        if not values:
            raise ValueError("PLシートにデータが存在しません。")

        # 所得金額の位置を特定
        taxable_income = None
        for row in values:
            if "所得金額" in row:  # 所得金額が含まれる列を探す
                col_index = row.index("所得金額") + 1  # 次の列（右隣）に値があると仮定
                row_index = values.index(row)
                taxable_income = values[row_index][col_index]
                break

        if taxable_income is None:
            raise ValueError("PLシートに課税所得が見つかりません。")

        # ラベルを更新
        taxable_income_label.config(text=f"課税所得: {taxable_income}円")

    except Exception as e:
        taxable_income_label.config(text="課税所得: エラー発生")
        print(f"課税所得取得エラー: {e}")

last_selected_item = None
# 各ラベルとエントリー
labels = ["月日"]
# 課税所得を表示するラベル
taxable_income_label = tk.Label(root, text="課税所得: 計算中...", font=default_font)
taxable_income_label.grid(row=len(labels) + 10, column=0, columnspan=2, pady=(5, 10), padx=(20, 0), sticky="w")

# 既存のラベル設定を左詰めに変更
# 「月日」と「金額」のラベルを横並びに表示
tk.Label(root, text="月日", font=default_font).grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_date = tk.Entry(root, width=15)
entry_date.grid(row=0, column=1, padx=0, pady=5, sticky="w")

tk.Label(root, text="金額", font=default_font).grid(row=0, column=1, padx=150, pady=5, sticky="w")
entry_amount = tk.Entry(root, width=15)
entry_amount.grid(row=0, column=1, padx=250, pady=5, sticky="w")

# 日付フィールドの初期値を設定
entry_date.insert(0, "2024")
entry_date.bind("<FocusOut>", format_date)

# 適用、科目、取引分類、取引手段のラベルを追加
tk.Label(root, text="適用", font=default_font).grid(row=len(labels) + 3, column=0, padx=10, pady=5, sticky="w")
# 最初の6つの選択肢を配置
create_radio_buttons(options_apply[:7], selected_option_apply, row_start=len(labels) + 3, column_start=1)
create_radio_buttons(options_apply[7:14], selected_option_apply, row_start=len(labels) + 4, column_start=1)
create_radio_buttons(options_apply[14:], selected_option_apply, row_start=len(labels) + 5, column_start=1)
# 自由入力用のエントリー（初期は無効）
apply_entry = tk.Entry(root, font=default_font, state="disabled", width=30)
apply_entry.grid(row=len(labels) + 5, column=1, padx=650, pady=2, sticky="w")
# ラジオボタンの選択変更時に動作を連動
selected_option_apply.trace_add(
    "write", 
    lambda *args: on_apply_change(*args, selected_option=selected_option_apply, entry=apply_entry)
)
tk.Label(root, text="科目", font=default_font).grid(row=len(labels), column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_subject[:6], selected_option_subject, row_start=len(labels), column_start=1)
create_radio_buttons(options_subject[6:12], selected_option_subject, row_start=len(labels) + 1, column_start=1)
create_radio_buttons(options_subject[12:], selected_option_subject, row_start=len(labels) + 2, column_start=1)
subject_entry = tk.Entry(root, font=default_font, state="disabled", width=30)
subject_entry.grid(row=len(labels) + 2, column=1, padx=350, pady=2, sticky="w")
# ラジオボタンの選択変更時に動作を連動
selected_option_subject.trace_add(
    "write", 
    lambda *args: on_apply_change(*args, selected_option=selected_option_subject, entry=subject_entry)
)
tk.Label(root, text="取引分類", font=default_font).grid(row=len(labels) + 6, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_kind, selected_option_kind, row_start=len(labels) + 6, column_start=1)
tk.Label(root, text="取引手段", font=default_font).grid(row=len(labels) + 7, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_means, selected_option_means, row_start=len(labels) + 7, column_start=1)

# 保存ボタン
save_button = tk.Button(root, text="データを追加/修正", command=save_data, font=default_font, width=15)
save_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(40, 0), sticky="w")

# 削除ボタン
delete_button = tk.Button(root, text="データを削除", command=delete_data, font=default_font, width=15)
delete_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(200, 0), sticky="w")

# # PLシート更新ボタン
# update_PL_button = tk.Button(root, text="PLシート更新", command=lambda: update_pl_sheet(service, SPREADSHEET_ID), font=default_font, width=15)
# update_PL_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(360, 0), sticky="w")

# PLシート更新ボタン
update_cash_button = tk.Button(root, text="現金シート更新", command=lambda: cash(service, SPREADSHEET_ID), font=default_font, width=15)
update_cash_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(360, 0), sticky="w")

# 仕訳帳シート更新ボタン
update_journal_button = tk.Button(root, text="仕訳帳シート更新", command=lambda: journal(service, SPREADSHEET_ID), font=default_font, width=15)
update_journal_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(520, 0), sticky="w")

# それ以外のシート更新ボタン
update_proprietor_button = tk.Button(root, text="それ以外全て更新", command=update_proprietor_and_sales, font=default_font, width=15)
update_proprietor_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(680, 0), sticky="w")

# データ表示用のTreeview
# データ表示用のFrameを作成
tree_frame = tk.Frame(root)
tree_frame.grid(row=len(labels) + 11, column=0, columnspan=2, padx=10, pady=10, sticky="w")

# Scrollbarを追加
tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scrollbar.pack(side="right", fill="y")

# データ表示用のTreeview
# Treeviewの列を定義
columns = ["月日", "取引分類", "科目", "適用", "取引手段", "金額"]
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12, yscrollcommand=tree_scrollbar.set)
tree.pack(side="left", fill="both", expand=True)

# ScrollbarをTreeviewにリンク
tree_scrollbar.config(command=tree.yview)

# 各列のヘッダーと幅を設定
for col in columns:
    tree.heading(col, text=col)  # ヘッダーを設定
    tree.column(col, width=150, anchor="center")  # 幅と配置を設定

# ソート状態を保持する変数
sort_state = {"column": None, "order": None}  # ソートの列と順序を記録

# スタイルの設定（Treeviewのヘッダーをカスタマイズ）
style = ttk.Style()
style.configure("Treeview.Heading", font=("Arial", 15), relief="flat")  # ヘッダーのフォントとスタイルを設定
style.map("Treeview.Heading", relief=[("active", "raised")])  # クリック時の効果

# ヘッダークリックイベントを処理する関数
def on_header_click(event):
    region = tree.identify_region(event.x, event.y)  # クリックされた場所を特定
    if region == "heading":  # ヘッダーがクリックされた場合
        column_id = tree.identify_column(event.x)  # クリックされた列を特定
        if column_id == "#1":  # #1は「月日」の列を指します
            sort_by_column("月日")

# Treeviewにヘッダークリックイベントをバインド
tree.bind("<Button-1>", on_header_click)
# Treeviewの選択イベントを処理する関数にバインド
tree.bind("<<TreeviewSelect>>", load_selected_record)

# 初期データの表示
refresh_table()
update_taxable_income_label_from_pl(service, SPREADSHEET_ID)
# Enterキーのバインド
bind_enter_to_save()
# メインループの開始
root.mainloop()