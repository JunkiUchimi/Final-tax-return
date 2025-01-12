import pandas as pd
import os
from tkinter import messagebox
import tkinter as tk

def create_radio_buttons(root, options, variable, row_start, column_start, default_font):
    """
    ラジオボタンを作成する関数。
    
    Args:
        root: tk.Tkまたはtk.Frameオブジェクト
        options: ラジオボタンの選択肢リスト
        variable: tk.StringVarオブジェクト
        row_start: 配置開始の行
        column_start: 配置開始の列
        default_font: ラジオボタンのフォント設定
    """
    for i, option in enumerate(options):
        tk.Radiobutton(
            root,
            text=option,
            variable=variable,
            value=option,
            font=default_font
        ).grid(row=row_start, column=column_start, padx=(10 + i * 110), pady=2, sticky="w")

def on_apply_change(*args, selected_option, entry):
    """ラジオボタンの選択に応じて自由入力エントリーの有効化/無効化を切り替える"""
    if selected_option.get() == "その他":
        entry.config(state="normal")  # 入力可能にする
    else:
        entry.delete(0, tk.END)  # 内容をクリア
        entry.config(state="disabled")  # 入力を無効化する

def show_auto_closing_popup(root, title, message, duration=1500):
    """
    自動で消えるポップアップを表示する
    Args:
        title (str): ポップアップのタイトル
        message (str): ポップアップのメッセージ
        duration (int): ポップアップが消えるまでの時間（ミリ秒）
    """
    popup = tk.Toplevel()
    popup.title(title)
    popup.geometry("300x100")  # サイズを設定（必要に応じて調整）
    popup.resizable(False, False)

    # メッセージラベル
    tk.Label(popup, text=message, font=("Arial", 12), wraplength=280).pack(pady=20)

    # ポップアップを中央に配置
    x = root.winfo_x() + (root.winfo_width() // 2) - (300 // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (100 // 2)
    popup.geometry(f"+{x}+{y}")

    # 自動で消える設定
    popup.after(duration, popup.destroy)


