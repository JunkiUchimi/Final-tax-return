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
from utils import on_apply_change, show_auto_closing_popup, show_loading
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
options_means.append("その他")
options_kind.append("その他")

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

global processing
processing = False  # 処理中かどうかを判定するフラグ
# データを保存する関数
def save_data():
    global last_selected_item, previous_selection, processing

    if processing:
        messagebox.showinfo("処理中", "一つ前の処理が続いています。お待ち下さい。")
        return

    # --- バリデーション（処理開始前に確認） ---
    format_date(None)
    amount_text = entry_amount.get().replace(",", "")
    if not amount_text:
        messagebox.showwarning("入力エラー", "金額を入力してください！")
        return
    try:
        amount_raw = int(amount_text)
    except ValueError:
        messagebox.showwarning("入力エラー", "金額には数値を入力してください！")
        return

    # --- 処理開始 ---
    processing = True

    def set_buttons(state):
        """ボタンの有効/無効をまとめて切り替える"""
        save_button.config(state=state)
        delete_button.config(state=state)
        update_cash_button.config(state=state)
        update_journal_button.config(state=state)
        update_proprietor_button.config(state=state)

    def reset_processing():
        """処理終了時に必ず呼ばれるリセット処理"""
        global processing
        processing = False
        set_buttons("normal")
        bind_enter_to_save()

    set_buttons("disabled")
    root.unbind("<Return>")

    try:
        # --- 入力値の取得 ---
        date    = entry_date.get()
        kind    = selected_option_kind.get()
        subject = subject_entry.get() if selected_option_subject.get() == "その他" else selected_option_subject.get()
        apply   = apply_entry.get()   if selected_option_apply.get()   == "その他" else selected_option_apply.get()
        means   = selected_option_means.get()

        # 按分計算（携帯料金以外の水道光熱費・通信費・利子割引料は35%）
        if subject in ["水道光熱費", "通信費", "利子割引料"] and apply != "携帯料金":
            amount = math.floor(amount_raw * 0.35)
        else:
            amount = amount_raw

        values = [[date, kind, subject, apply, means, amount]]
        body   = {'values': values}

        # --- 保存処理（修正 or 新規） ---
        if last_selected_item is not None:
            # 修正モード
            item_index  = tree.index(last_selected_item) + 1
            spre_index  = len(original_data) - item_index + 1
            update_range = f"経費!A{spre_index + 1}:F{spre_index + 1}"
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=update_range,
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()
        else:
            # 新規追加モード
            service.spreadsheets().values().append(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME,
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()

        show_auto_closing_popup(root, "成功", "データを追加しました！")

        # 保存成功後の後処理
        previous_selection = {
            "apply":   selected_option_apply.get(),
            "subject": selected_option_subject.get(),
            "means":   selected_option_means.get(),
            "kind":    selected_option_kind.get()
        }
        reset_fields()
        refresh_table()
        last_selected_item = None
        taxable_income_label.config(text="課税所得: 計算中...")

        # PLシート更新はバックグラウンドで実行
        def update_processing():
            try:
                update_pl_sheet(service, SPREADSHEET_ID)
                root.after(0, lambda: update_taxable_income_label_from_pl(service, SPREADSHEET_ID))
            except Exception as e:
                print(f"PLシート更新エラー: {e}")
            finally:
                root.after(0, reset_processing)  # ✅ UIスレッドで安全にリセット

        threading.Thread(target=update_processing, daemon=True).start()
        # ※ バックグラウンド処理中はボタン無効のままにするため、ここでは reset_processing を呼ばない

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {e}")
        reset_processing()  # ✅ 例外発生時は即座にリセット

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
    entry_date.insert(0, "2025")  # 必要に応じて変更
    entry_amount.delete(0, tk.END)
    # グローバル変数から選択状態を設定
    selected_option_apply.set(previous_selection["apply"])
    selected_option_subject.set(previous_selection["subject"])
    selected_option_means.set(previous_selection["means"])
    selected_option_kind.set(previous_selection["kind"])
    entry_date.focus()


def load_selected_record(event):
    global last_selected_item

    selected_item = tree.selection()

    # 同じ行を再クリックしたら選択解除
    if selected_item and selected_item[0] == last_selected_item:
        tree.selection_remove(selected_item[0])
        last_selected_item = None
        reset_fields()
        return

    if not selected_item:
        return

    last_selected_item = selected_item[0]
    item_index = tree.index(last_selected_item)
    spreadsheet_index = len(original_data) - item_index - 1  # ✅ -1 に修正（0始まり）

    # ✅ API を叩かず original_data から直接取得
    try:
        record = original_data[spreadsheet_index]

        entry_date.delete(0, tk.END)
        entry_date.insert(0, record[0])
        entry_amount.delete(0, tk.END)
        entry_amount.insert(0, record[5])
        selected_option_kind.set(record[1])
        selected_option_subject.set(record[2])
        selected_option_apply.set(record[3])
        selected_option_means.set(record[4])

    except IndexError:
        messagebox.showerror("エラー", "選択されたデータが見つかりません。画面を更新してください。")
        last_selected_item = None
        reset_fields()

def delete_data():
    global last_selected_item

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("削除エラー", "削除するデータを選択してください！")
        return

    # 削除対象のデータ内容を事前に取得して確認ポップアップ
    item_values = tree.item(selected_item[0], "values")
    confirm_msg = (
        f"以下のデータを削除しますか？\n\n"
        f"  月日　　: {item_values[0]}\n"
        f"  取引分類: {item_values[1]}\n"
        f"  科目　　: {item_values[2]}\n"
        f"  適用　　: {item_values[3]}\n"
        f"  取引手段: {item_values[4]}\n"
        f"  金額　　: {item_values[5]}円"
    )
    if not messagebox.askyesno("削除確認", confirm_msg):
        return

    loading = show_loading(root, "削除中...")
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])

        if not values:
            messagebox.showwarning("削除エラー", "スプレッドシートが空です。")
            return

        indexes_to_delete = []
        for item in selected_item:
            tree_index = tree.index(item)
            spreadsheet_index = len(original_data) - tree_index
            if 0 <= spreadsheet_index < len(values):
                indexes_to_delete.append(spreadsheet_index)

        indexes_to_delete.sort(reverse=True)
        for index in indexes_to_delete:
            del values[index]

        if not values:
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
        else:
            service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=RANGE_NAME
            ).execute()
            body = {'values': values}
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"経費!A1:F{len(values)}",
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()

        refresh_table()
        last_selected_item = None
        reset_fields()
        # 削除成功：何を削除したか表示
        messagebox.showinfo(
            "削除完了",
            f"以下のデータを削除しました。\n\n"
            f"  月日　　: {item_values[0]}\n"
            f"  取引分類: {item_values[1]}\n"
            f"  科目　　: {item_values[2]}\n"
            f"  適用　　: {item_values[3]}\n"
            f"  取引手段: {item_values[4]}\n"
            f"  金額　　: {item_values[5]}円"
        )

    except Exception as e:
        messagebox.showerror("削除エラー", f"削除中にエラーが発生しました。\n\n{e}")
    finally:
        loading.destroy()  # ✅ 成功・失敗にかかわらずローディングを閉じる


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
    loading = show_loading(root, "シート更新中...")
    def run():
        try:
            others(service, SPREADSHEET_ID, subjectif="事業主貸",  range_name="事業主貸・売上!J4:P")
            others(service, SPREADSHEET_ID, subjectif="売上",      range_name="事業主貸・売上!B4:H")
            others(service, SPREADSHEET_ID, subjectif="消耗品費",  range_name="消耗品費!B4:H")
            others(service, SPREADSHEET_ID, subjectif="旅費交通費", range_name="旅費・接待・研修!B4:H")
            others(service, SPREADSHEET_ID, subjectif="接待交際費", range_name="旅費・接待・研修!J4:P")
            others(service, SPREADSHEET_ID, subjectif="研修費",    range_name="旅費・接待・研修!R4:X")
            others(service, SPREADSHEET_ID, subjectif="会議費",    range_name="会議費・修繕費・新聞・利子!B4:H")
            others(service, SPREADSHEET_ID, subjectif="新聞図書費", range_name="会議費・修繕費・新聞・利子!J4:P")
            others(service, SPREADSHEET_ID, subjectif="修繕費",    range_name="会議費・修繕費・新聞・利子!R4:X")
            others(service, SPREADSHEET_ID, subjectif="利子割引料", range_name="会議費・修繕費・新聞・利子!Z4:AF")
            others(service, SPREADSHEET_ID, subjectif="租税公課",  range_name="租税・雑費・水道・更新!B4:H")
            others(service, SPREADSHEET_ID, subjectif="雑費",      range_name="租税・雑費・水道・更新!J4:P")
            others(service, SPREADSHEET_ID, subjectif="水道光熱費", range_name="租税・雑費・水道・更新!R4:X")
            others(service, SPREADSHEET_ID, subjectif="通信費",    range_name="租税・雑費・水道・更新!Z4:AF")
            others(service, SPREADSHEET_ID, subjectif="車両費",    range_name="車両・損保・減価!B4:H")
            others(service, SPREADSHEET_ID, subjectif="損害保険料", range_name="車両・損保・減価!J4:P")
            others(service, SPREADSHEET_ID, subjectif="減価償却費", range_name="車両・損保・減価!R4:X")
            root.after(0, lambda: messagebox.showinfo("完了", "その他シートを更新しました。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("更新エラー", f"更新中にエラーが発生しました。\n\n{e}"))
        finally:
            root.after(0, loading.destroy)

    threading.Thread(target=run, daemon=True).start()

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

def run_cash():
    loading = show_loading(root, "現金シート更新中...")
    def run():
        try:
            cash(service, SPREADSHEET_ID)
            root.after(0, lambda: messagebox.showinfo("完了", "現金シートを更新しました。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("更新エラー", f"現金シート更新中にエラーが発生しました。\n\n{e}"))
        finally:
            root.after(0, loading.destroy)
    threading.Thread(target=run, daemon=True).start()

def run_journal():
    loading = show_loading(root, "仕訳帳シート更新中...")
    def run():
        try:
            journal(service, SPREADSHEET_ID)
            root.after(0, lambda: messagebox.showinfo("完了", "仕訳帳シートを更新しました。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("更新エラー", f"仕訳帳シート更新中にエラーが発生しました。\n\n{e}"))
        finally:
            root.after(0, loading.destroy)
    threading.Thread(target=run, daemon=True).start()

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
entry_date.insert(0, "2025")
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
kind_entry = tk.Entry(root, font=default_font, state="disabled", width=30)
kind_entry.grid(row=len(labels) + 6, column=1, padx=450, pady=2, sticky="w")
# ラジオボタンの選択変更時に動作を連動
selected_option_kind.trace_add(
    "write", 
    lambda *args: on_apply_change(*args, selected_option=selected_option_kind, entry=kind_entry)
)
tk.Label(root, text="取引手段", font=default_font).grid(row=len(labels) + 7, column=0, padx=10, pady=5, sticky="w")
create_radio_buttons(options_means, selected_option_means, row_start=len(labels) + 7, column_start=1)
means_entry = tk.Entry(root, font=default_font, state="disabled", width=30)
means_entry.grid(row=len(labels) + 7, column=1, padx=350, pady=2, sticky="w")
# ラジオボタンの選択変更時に動作を連動
selected_option_means.trace_add(
    "write", 
    lambda *args: on_apply_change(*args, selected_option=selected_option_means, entry=means_entry)
)
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
update_cash_button = tk.Button(root, text="現金シート更新", command=run_cash, font=default_font, width=15)
update_cash_button.grid(row=len(labels) + 9, column=0, columnspan=2, pady=(10, 5), padx=(360, 0), sticky="w")

# 仕訳帳更新ボタン
update_journal_button = tk.Button(root, text="仕訳帳シート更新", command=run_journal, font=default_font, width=15)
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

def on_treeview_click(event):
    """クリックされた行が選択済みなら選択解除する"""
    global last_selected_item

    # クリックされた行を特定
    clicked_item = tree.identify_row(event.y)

    if not clicked_item:
        return  # 行以外（余白など）をクリックした場合は無視

    if clicked_item == last_selected_item:
        # 選択済みの行を再クリック → 解除して新規追加モードに戻す
        tree.selection_remove(clicked_item)
        last_selected_item = None
        reset_fields()
        return  # <<TreeviewSelect>> を発火させないために早期リターン
    

# Treeviewにヘッダークリックイベントをバインド
tree.bind("<Button-1>", on_header_click)
# Treeviewの選択イベントを処理する関数にバインド
tree.bind("<<TreeviewSelect>>", load_selected_record)

tree.bind("<Button-1>", on_treeview_click)

# 初期データの表示
refresh_table()
update_taxable_income_label_from_pl(service, SPREADSHEET_ID)
# Enterキーのバインド
bind_enter_to_save()

# macOS 英数キーによるスペース挿入を抑制
def remove_unwanted_space(event):
    """Entry への意図しないスペース挿入を防ぐ"""
    focused = root.focus_get()
    if isinstance(focused, tk.Entry):
        # 現在のカーソル位置を取得
        pos = focused.index(tk.INSERT)
        content = focused.get()
        # カーソル直前がスペースなら削除
        if pos > 0 and content[pos - 1] == " ":
            focused.delete(pos - 1, pos)

root.bind("<KeyRelease>", remove_unwanted_space)

# メインループの開始
root.mainloop()
