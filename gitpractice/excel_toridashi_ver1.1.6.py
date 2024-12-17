#横結合実装
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import tkinter.messagebox as messagebox
import os  # ファイルを開くために必要
import re
import sys
import threading  # スレッド処理用
from PIL import Image, ImageTk  # Pillowをインポート


# import time  # 時間計測のためのモジュール

processing_interrupted = False  # 処理中断フラグ
pd.set_option('future.no_silent_downcasting', True)

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

##############################
#ファイル選択機能
##############################

# ファイルパス入力用のリスト
file_paths = []

# 範囲選択用フィールドのリスト
range_frames = []

# ファイル選択用の関数
def select_files():
    global file_paths
    new_files = filedialog.askopenfilenames(filetypes=[
        ("Excel files (.xlsx, .xlsm)", "*.xlsx;*.xlsm"),
        ("Excel Workbook (.xlsx)", "*.xlsx"),
        ("Excel Macro-Enabled Workbook (.xlsm)", "*.xlsm"),
    ])

    # 順序を維持しつつ重複を排除
    file_paths = list(dict.fromkeys(file_paths + list(new_files)))  # dictを使って順序維持
    update_file_list()
# ファイルリストの表示を更新する関数
def update_file_list():
    file_list.delete(0, tk.END)
    for path in file_paths:
        file_list.insert(tk.END, path)

#ファイルリストをリセットする関数
def clear_file_list():
    global file_paths
    file_paths = []  # ファイルパスリストを空にする
    update_file_list()  # リストボックスの表示を更新

# リストから選択したファイルを削除する関数
def delete_selected_files(event):
    global file_paths
    selected_indices = file_list.curselection()
    selected_files = [file_list.get(i) for i in selected_indices]
    
    file_paths = [f for f in file_paths if f not in selected_files]
    update_file_list()


##############################
#範囲指定機能
##############################
# 列の入力バリデーション用関数
def validate_column_input(char):
    if re.match(r'^[A-Za-z]*$', char):  # 英字のみ許可（空白も許可）
        return True
    else:
        messagebox.showerror("エラー", "列には半角英字のみを入力してください")
        return False

# 行の入力バリデーション用関数
def validate_row_input(char):
    if re.match(r'^\d*$', char):  # 数字のみ許可（空白も許可）
        return True
    else:
        messagebox.showerror("エラー", "行には数字のみを入力してください")
        return False

# 範囲フィールドの順番を入れ替える関数
def move_range_frame_up(frame):
    idx = [entry['frame'] for entry in range_frames].index(frame)
    if idx > 0:
        range_frames[idx], range_frames[idx - 1] = range_frames[idx - 1], range_frames[idx]
        re_draw_range_fields()

def move_range_frame_down(frame):
    idx = [entry['frame'] for entry in range_frames].index(frame)
    if idx < len(range_frames) - 1:
        range_frames[idx], range_frames[idx + 1] = range_frames[idx + 1], range_frames[idx]
        re_draw_range_fields()

# 範囲フィールドを再描画する関数
def re_draw_range_fields():
    global separator_line
    # 既存の範囲フィールドを一旦すべて削除
    for entry in range_frames:
        entry['frame'].pack_forget()

    # 新しい順序で範囲フィールドを再描画
    for entry in range_frames:
        entry['frame'].pack(anchor="center", before=get_value_button)

# 範囲フィールドを追加する関数に【↑】【↓】ボタンを追加
def add_range_fields():
    global dynamic_separator
    frame = ttk.Frame(scrollable_frame)
    frame.pack(anchor="center", before=get_value_button)

    # 列の範囲のラベルと入力欄
    ttk.Label(frame, text="列の範囲（例: C ～ R）:", padding=(0, 10, 0, 0)).pack(anchor="center")
    col_frame = ttk.Frame(frame)
    col_frame.pack(anchor="center")

    # 列の入力バリデーションの設定
    vcmd_col = (root.register(validate_column_input), '%P')

    col_start_entry = ttk.Entry(col_frame, width=10, validate='key', validatecommand=vcmd_col)
    col_start_entry.pack(side="left", padx=(0, 5))
    ttk.Label(col_frame, text="～").pack(side="left")
    col_end_entry = ttk.Entry(col_frame, width=10, validate='key', validatecommand=vcmd_col)
    col_end_entry.pack(side="left", padx=(5, 0))

    # 行の入力バリデーションの設定
    vcmd_row = (root.register(validate_row_input), '%P')

    # 行の範囲のラベルと入力欄
    ttk.Label(frame, text="行の範囲（例: 57 ～ 61）:").pack(anchor="center")
    row_frame = ttk.Frame(frame)
    row_frame.pack(anchor="center")
    row_start_entry = ttk.Entry(row_frame, width=10,validate='key', validatecommand=vcmd_row)
    row_start_entry.pack(side="left", padx=(0, 5))
    ttk.Label(row_frame, text="～").pack(side="left")
    row_end_entry = ttk.Entry(row_frame, width=10,validate='key', validatecommand=vcmd_row)
    row_end_entry.pack(side="left", padx=(5, 0))

    # ボタンのフレームを追加
    button_frame = ttk.Frame(frame)
    button_frame.pack(anchor="center", pady=(10, 5))

    # 削除ボタン
    remove_button = ttk.Button(button_frame, text="△削除する", command=lambda: remove_range_fields(frame), style="Custom.TButton")
    remove_button.pack(side="left", padx=(5, 5))

    # 上に移動ボタン
    move_up_button = ttk.Button(button_frame, text="↑", command=lambda: move_range_frame_up(frame), style="Small.TButton", width=2)
    move_up_button.pack(side="left", padx=(2, 2))

    # 下に移動ボタン
    move_down_button = ttk.Button(button_frame, text="↓", command=lambda: move_range_frame_down(frame), style="Small.TButton", width=2)
    move_down_button.pack(side="left", padx=(2, 2))

    range_frames.append({
        'frame': frame,
        'col_start_entry': col_start_entry,
        'col_end_entry': col_end_entry,
        'row_start_entry': row_start_entry,
        'row_end_entry': row_end_entry
    })
    # 点線を最後の範囲フィールドの直下に描画
    draw_dynamic_separator(scrollable_frame, after=frame)
    # Canvasのスクロール領域を更新
    scrollable_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    toggle_merge_options_state()  # 結合方向の状態を更新

# 範囲フィールドを削除する関数
def remove_range_fields(frame):
    global dynamic_separator
    if len(range_frames) == 1:
        messagebox.showerror("エラー", "範囲をすべて削除することはできません")
    else:
        for entry in range_frames:
            if entry['frame'] == frame:
                range_frames.remove(entry)
                frame.destroy()
                break
        # 点線を再配置（最後の範囲フィールドの下に描画）
        if range_frames:
            last_frame = range_frames[-1]['frame']
            draw_dynamic_separator(scrollable_frame, after=last_frame)
        else:
            # すべての範囲フィールドが削除された場合、点線を非表示にする
            if dynamic_separator:
                dynamic_separator.pack_forget()
                dynamic_separator = None
        # 結合方向の状態を更新
        toggle_merge_options_state()
    scrollable_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))


######入力されたデータを変換
def excel_col_to_num(col_str):
    """Excelの列名（A, B, ..., Z, AA, AB, ...）を数値に変換"""
    num = 0
    for c in col_str.upper():  # 大文字に変換して処理
        if 'A' <= c <= 'Z':
            num = num * 26 + (ord(c) - ord('A') + 1)
    return num - 1  # Pandas は 0 ベースのため -1 を引く

######ファイル名を短く
def get_last_two_parts_of_path(filepath):
    # ファイル名とその直前のディレクトリ名を取得
    directory, filename = os.path.split(filepath)  # 最後のファイル名を取得
    _, last_directory = os.path.split(directory)  # 最後のディレクトリ名を取得
    return f".../{last_directory}/{filename}"

##############################
#入力バリデーション
##############################

def validate_range_fields():
    """
    範囲指定がすべて正しく入力されているかを確認する関数。
    空欄がある場合はエラーメッセージを表示し、False を返す。
    """
    error_messages = []

    # シート名入力のバリデーション
    if sheet_mode.get() == 1:  # 特定の名前を含むシートのみ
        sheet_name = sheet_entry.get()
        if not sheet_name.strip():
            error_messages.append("シート名を入力してください。")

    for idx, entry in enumerate(range_frames, start=1):
        col_start = entry['col_start_entry'].get()
        col_end = entry['col_end_entry'].get()
        row_start = entry['row_start_entry'].get()
        row_end = entry['row_end_entry'].get()

        # 空欄チェック
        if not col_start or not col_end or not row_start or not row_end:
            error_messages.append(f"範囲 {idx}: 列または行が未入力です。")

        # 列と行の大小関係を確認
        else:
            col_start_num = excel_col_to_num(col_start)
            col_end_num = excel_col_to_num(col_end)
            if col_start_num > col_end_num:
                error_messages.append(f"範囲 {idx}: 列の範囲が無効です ({col_start} ～ {col_end})。")
            if int(row_start) > int(row_end):
                error_messages.append(f"範囲 {idx}: 行の範囲が無効です ({row_start} ～ {row_end})。")

    if error_messages:
        messagebox.showerror("エラー", "\n".join(error_messages))
        return False
    return True

##############################
#データ取得機能
##############################

# Excelのセル値を取得する関数（pandasで処理）
def get_excel_values():
    """
    データを取得してExcelに保存する（キャッシュは使用しない）。
    """
    global processing_interrupted

    if not file_paths:
        messagebox.showerror("エラー", "ファイルが選択されていません。Excelファイルを選択してください。")
        return

    if not validate_range_fields():  # 範囲指定のバリデーション
        return

    process_excel_values(use_cache=False)  # キャッシュを使用しない
            
# 結果をExcelに保存する関数
def save_to_excel(results):
    # 結果を保存するファイル名を取得
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        # ExcelWriterオブジェクトを作成し、openpyxlエンジンを使用
        with pd.ExcelWriter(save_path, engine='openpyxl') as wb:
            # 結果をDataFrameとしてExcelシートに書き込み
            df = pd.DataFrame(results)
            df.to_excel(wb, index=False, header=False, sheet_name="結果")

        # 保存が完了したら、ファイルを開くかどうかをユーザーに確認
        open_file = messagebox.askyesno("保存完了", "保存しました！このままファイルを開きますか？")
        if open_file:
            # ユーザーが「はい」を選択した場合、ファイルを開く
            try:
                os.startfile(save_path)  # Windows環境の場合、Excelファイルを開く
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを開けませんでした: {str(e)}")

##############################
#プレビュー機能
##############################

sheet_names_cache = {}  # シート名情報をキャッシュする辞書
preview_window_ref = None  # プレビューウィンドウの参照を保持する変数

def get_sheet_names(file):
    """Excelファイルからシート名を取得し、キャッシュする"""
    if file not in sheet_names_cache:
        with pd.ExcelFile(file) as xls:
            sheet_names_cache[file] = xls.sheet_names
    return sheet_names_cache[file]

def get_excel_preview():
    """プレビュー: 必要最小限のデータ取得 + シート名キャッシュ"""
    global preview_window_ref, processing_interrupted,data_cache
    data_cache = {}  # キャッシュを初期化

    if not file_paths:
        messagebox.showerror("エラー", "ファイルが選択されていません。")
        return

    if not validate_range_fields():  # 範囲指定が正しいか確認
        return

    # 前回のプレビューウィンドウが存在する場合は閉じる
    if preview_window_ref is not None and preview_window_ref.winfo_exists():
        try:
            preview_window_ref.destroy()
        except:
            pass
        preview_window_ref = None  # 参照をリセット

    stop_indicator = show_processing_indicator("プレビュー")  # インジケーター表示

    def thread_task():
        try:
            sheet_name = sheet_entry.get()
            all_sheets = []
            for file in file_paths:
                if processing_interrupted:  # 中断チェック
                    return
                
                xls = pd.ExcelFile(file)
                sheet_names = xls.sheet_names  # シート名リストを取得
                
                for sheet in sheet_names:
                    if processing_interrupted:
                        return
                    # すべてのシート（非表示を含まない）処理
                    if sheet_mode.get() == 2:  # 非表示シートを除外する処理
                        if xls.book[sheet].sheet_state == 'hidden':  # 非表示シートならスキップ
                            continue
                    
                    # 特定のシート名処理
                    if sheet_mode.get() == 1 and sheet_name not in sheet:
                        continue
                    
                    all_sheets.append((file, sheet))

            if not all_sheets:
                messagebox.showinfo("情報", "指定した条件に一致するシートが見つかりませんでした。")
                stop_indicator()
                return

            file, sheet = all_sheets[0]
            data = fetch_sheet_data(file, sheet)

            if not processing_interrupted:
                root.after(0, lambda: preview_with_navigation(all_sheets, {all_sheets[0]: data}))

        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")
        finally:
            stop_indicator()

    threading.Thread(target=thread_task, daemon=True).start()

##############################
#プレビュー機能・値取得機能共通
##############################
def fetch_sheet_data(file, sheet):
    """指定範囲を強制的に取得し、欠損部分は空白で埋める"""
    combined_data = []

    for entry in range_frames:
        col_start_str = entry['col_start_entry'].get()
        col_end_str = entry['col_end_entry'].get()
        row_start_str = entry['row_start_entry'].get()
        row_end_str = entry['row_end_entry'].get()

        # Excel列名 → 数値に変換
        col_start = excel_col_to_num(col_start_str)  # 開始列
        col_end = excel_col_to_num(col_end_str)      # 終了列
        row_start = int(row_start_str) - 1           # Pandasは0ベースなので -1
        row_end = int(row_end_str)

        # 行数と列数を計算
        nrows = row_end - row_start  # 行数
        ncols = col_end - col_start + 1  # 列数

        try:
            # 指定した行数を取得（列は後で強制的に切り取る）
            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None,       # ヘッダーなし
                skiprows=row_start,  # 開始行をスキップ
                nrows=nrows          # 必要な行数のみ取得
            )

            # 列範囲を強制的に固定し、欠損部分は空白で埋める
            fixed_data = pd.DataFrame(index=range(nrows), columns=range(ncols))

            for r in range(nrows):
                for c in range(ncols):
                    try:
                        # データが範囲内なら値をセット
                        fixed_data.iloc[r, c] = df.iloc[r, col_start + c]
                    except (IndexError, KeyError):
                        # 範囲外は空白にする
                        fixed_data.iloc[r, c] = ""

            # 列名を設定（列1, 列2, ...）
            fixed_data.columns = [f"列{i+1}" for i in range(ncols)]
            combined_data.append(fixed_data)

        except Exception as e:
            # エラー時のデフォルトデータ (3列分空白で埋める)
            print(f"データ取得エラー: {e} (file: {file}, sheet: {sheet})")
            empty_cols = [f"列{i+1}" for i in range(3)]
            empty_rows = [["" for _ in range(3)] for _ in range(nrows)]
            combined_data.append(pd.DataFrame(empty_rows, columns=empty_cols))

    # 結合方法を切り替える
    if merge_mode.get() == "horizontal":  # 横に結合
        result = pd.concat(combined_data, axis=1, ignore_index=True).fillna("")
        result.columns = [f"列{i+1}" for i in range(result.shape[1])]  # 列名を再設定
        return result
    else:  # 縦に結合 (デフォルト)
        return pd.concat(combined_data, axis=0, ignore_index=True).fillna("")
    
    # # 結合して返す
    # if combined_data:
    #     return pd.concat(combined_data, ignore_index=True).fillna("")
    # return pd.DataFrame(columns=["列1", "列2", "列3"])

def process_excel_values(use_cache=False, cache_data=None):
    """
    データを取得し、指定されたシートデータを結合してExcelに保存する。
    :param use_cache: Trueの場合、キャッシュを利用する。
    :param cache_data: キャッシュデータ (use_cache=Trueの場合に使用)。
    """
    global processing_interrupted

    stop_indicator = show_processing_indicator("値の取り出し")

    def thread_task():
        results = []  # 結果データのリスト

        try:
            if use_cache and cache_data:  # キャッシュを利用する場合
                for (file, sheet), data in cache_data.items():
                    append_results(file, sheet, data, results)
            else:  # キャッシュを使用せずにデータ取得
                for file in file_paths:
                    if processing_interrupted: return
                    fetch_results_for_sheets(file, results)

            # 保存
            if not processing_interrupted:
                root.after(0, lambda: save_to_excel(results))

        except Exception as e:
            root.after(0, lambda: messagebox.showerror("エラー", f"エラーが発生しました: {e}"))

        finally:
            stop_indicator()  # インジケーター停止

    threading.Thread(target=thread_task, daemon=True).start()

def append_results(file, sheet, data, results):
    """ファイル名、シート名、データを結果に追加する関数"""
    results.append([f"ファイル名: {get_last_two_parts_of_path(file)}"])
    results.append([f"シート名: {sheet}"])
    results.extend(data.values.tolist())


def fetch_results_for_sheets(file, results):
    """シートモードに応じてデータを取得し、結果に追加する関数"""
    sheet_name = sheet_entry.get()
    xls = pd.ExcelFile(file)

    for sheet in xls.sheet_names:
        if processing_interrupted:
            return

        # すべてのシート（非表示含む）
        if sheet_mode.get() == 0:
            data = fetch_sheet_data(file, sheet)
            append_results(file, sheet, data, results)

        # 非表示シートを除外
        elif sheet_mode.get() == 2:
            if xls.book[sheet].sheet_state != 'hidden':
                data = fetch_sheet_data(file, sheet)
                append_results(file, sheet, data, results)

        # 特定の名前を含むシートのみ
        elif sheet_mode.get() == 1 and sheet_name:
            if sheet_name in sheet:
                data = fetch_sheet_data(file, sheet)
                append_results(file, sheet, data, results)

##############################
#プレビュー機能
##############################

# データキャッシュ用辞書
data_cache = {}
def preview_with_navigation(all_sheets, initial_data_cache):
    """ナビゲーション可能なプレビューウィンドウ"""
    global preview_window_ref, data_cache  # グローバル変数で参照を保持
    data_cache = initial_data_cache  # キャッシュを初期化

    def on_close():
        """ウィンドウを閉じた際に参照をリセットする"""
        global data_cache
        data_cache = {}  # キャッシュをクリア
        nonlocal preview_window
        preview_window_ref = None  # 参照をリセット
        preview_window.destroy()

    def update_preview(index):
        """現在のシートを表示"""
        nonlocal current_index
        current_index = index

        # 現在のファイルとシート
        file, sheet = all_sheets[current_index]

        # データを都度取得（キャッシュは使用しない）
        try:
            data = fetch_sheet_data(file, sheet)
        except Exception as e:
            messagebox.showerror("エラー", f"シートデータの取得に失敗しました: {e}")
            return

        # データが空の場合の処理
        if data.empty:
            # messagebox.showinfo("情報", "データが空です。範囲指定を確認してください。")
            data = pd.DataFrame(columns=["列1", "列2", "列3"])

        # ファイル名とシート名の表示を更新
        file_label["text"] = f"ファイル名： {get_last_two_parts_of_path(file)}"
        sheet_label["text"] = f"シート名： {sheet} ({current_index + 1}/{len(all_sheets)})"

        # Treeviewのデータをクリア
        tree.delete(*tree.get_children())  # 高速にTreeviewのデータをクリア
        
        # Treeviewの列を設定
        tree["columns"] = list(data.columns)
        tree["show"] = "headings"

        fixed_width = 80
        for col in tree["columns"]:
            # max_width = max(data[col].astype(str).map(len).max(), len(col)) * 10  # 幅を計算
            tree.heading(col, text=col)
            tree.column(col, width=fixed_width, anchor="center", stretch=False)  # stretch=Falseで固定  

        # データを一括で挿入（効率化）
        tree_data = [list(row) for _, row in data.iterrows()]
        for row in tree_data:
            tree.insert("", "end", values=row)

        # ボタンの有効/無効を設定
        prev_button["state"] = tk.NORMAL if current_index > 0 else tk.DISABLED
        next_button["state"] = tk.NORMAL if current_index < len(all_sheets) - 1 else tk.DISABLED

    # 【出力】ボタンの処理
    def export_preview_data():
        """
        プレビュー画面のデータをExcelに保存する。
        キャッシュが完全でない場合は改めて処理する。
        """
        # プレビュー画面が見ているシートリストを確認
        all_sheets_list = all_sheets  # グローバル変数からリストを取得

        if not all_sheets_list:
            messagebox.showerror("エラー", "プレビュー対象のシートが見つかりません。")
            return

        # キャッシュが利用可能かどうか確認
        if data_cache and len(data_cache) == len(all_sheets_list):
            # キャッシュが完全な場合
            results = []
            for (file, sheet), data in data_cache.items():
                results.append([f"ファイル名: {get_last_two_parts_of_path(file)}"])
                results.append([f"シート名: {sheet}"])
                results.extend(data.values.tolist())

            # Excelに保存
            save_to_excel(results)
        else:
            # キャッシュが不完全なら全シートを再処理
            process_excel_values(use_cache=False)

    # プレビューウィンドウの作成
    preview_window = tk.Toplevel(root)
    preview_window.title("データプレビュー")
    preview_window.geometry("600x400")
    center_window2(root, preview_window, 600, 400)

    preview_window.transient(root)  # 親ウィンドウの上に表示
    preview_window.protocol("WM_DELETE_WINDOW", on_close)  # 閉じる時に参照をリセット

    # ファビコン（アイコン画像）を設定
    if icon_image:  # icon_imageが正常に読み込まれていれば
       preview_window.iconphoto(False, icon_image)

    preview_window_ref = preview_window  # 参照を保持
   
    # ファイル名とシート名のラベル
    info_frame = ttk.Frame(preview_window)
    info_frame.pack(fill=tk.X, padx=10, pady=5)
    file_label = ttk.Label(info_frame, font=("Arial", 9))
    file_label.pack(anchor="w")
    sheet_label = ttk.Label(info_frame, font=("Arial", 9))
    sheet_label.pack(anchor="w")

    # データ表示部分（Treeview + スクロールバー）
    data_frame = ttk.Frame(preview_window)
    data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    y_scrollbar = ttk.Scrollbar(data_frame, orient=tk.VERTICAL)
    y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    x_scrollbar = ttk.Scrollbar(data_frame, orient=tk.HORIZONTAL)
    x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    tree = ttk.Treeview(data_frame, yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    tree.pack(fill=tk.BOTH, expand=True)

    y_scrollbar.config(command=tree.yview)
    x_scrollbar.config(command=tree.xview)

    # ナビゲーションボタン
    nav_frame = tk.Frame(preview_window, bg=preview_window["bg"])  # 親ウィンドウと同じ背景色
    nav_frame.pack(fill=tk.X, pady=10)

    prev_button = ttk.Button(nav_frame, text="◀ 前のシート", command=lambda: update_preview(current_index - 1), style="Custom.TButton")
    prev_button.pack(side=tk.LEFT, padx=10)
    next_button = ttk.Button(nav_frame, text="次のシート ▶", command=lambda: update_preview(current_index + 1), style="Custom.TButton")
    next_button.pack(side=tk.RIGHT, padx=10)

    # ボタンを含む共通フレーム
    button_frame = tk.Frame(preview_window, bg=preview_window["bg"])
    button_frame.pack(pady=10)

    # 閉じるボタン
    ttk.Button(button_frame, text="閉じる", command=preview_window.destroy, style="Custom.TButton").pack(side="left", padx=5)
    # 出力ボタン
    ttk.Button(button_frame, text="出力", command=export_preview_data, style="Highlight.TButton").pack(side="left", padx=5)

    # 初期化
    current_index = 0
    update_preview(current_index)
    # return preview_window  # ウィンドウの参照を返す

##############################
#進捗表示機能
##############################

def show_processing_indicator(task_title):
    """
    インジケーター型の進捗バーを表示する共通関数
    :param task_title: タスクのタイトル
    :return: stop関数（処理が終わったら呼ぶ）
    """
    global processing_interrupted  # グローバル変数を使用
    processing_interrupted = False  # フラグを初期化

    global icon_image  # 画像をグローバル変数として使用

    indicator_window = tk.Toplevel(root)
    indicator_window.title(task_title)
    indicator_window.geometry("300x100")
    # indicator_window.transient(root)  # 親ウィンドウの上に表示
    indicator_window.attributes("-topmost", True)  # 最前面表示
    indicator_window.resizable(False, False)  # リサイズ不可

    # 最前面表示は後回し
    indicator_window.lift()
    indicator_window.attributes("-topmost", True)


    # 最小化時の処理: 最前面解除
    def on_minimize(event):
        indicator_window.attributes("-topmost", False)

    # 復元時の処理: 再び最前面
    def on_restore(event):
        indicator_window.attributes("-topmost", True)

    # 「×」ボタンでウィンドウを閉じようとしたときの処理
    def on_close():
        global processing_interrupted
        if messagebox.askyesno("確認", "処理を中断しますか？"):
            processing_interrupted = True  # 中断フラグをTrueにする
            indicator_window.destroy()
        else:
            # いいえを選んだ場合、プログレスバーウィンドウを最前面に戻す
            indicator_window.lift()
            indicator_window.attributes("-topmost", True)

    
    indicator_window.protocol("WM_DELETE_WINDOW", on_close)

    # イベントバインド: 最小化と復元
    indicator_window.bind("<Unmap>", on_minimize)  # 最小化時
    indicator_window.bind("<Map>", on_restore)    # 復元時
    
    # 最小化を許可する
    # indicator_window.iconify()  # 初期状態で最小化も設定可能（任意）
    # indicator_window.grab_set()       # モーダルウィンドウとして動作

    # ファビコン（アイコン画像）を設定
    if icon_image:  # icon_imageが正常に読み込まれていれば
       indicator_window.iconphoto(False, icon_image)


    # 位置を中央にする（表示後に計算）
    root.after_idle(lambda: center_window2(root, indicator_window, 300, 100))

    # タイトルラベルと画像を同じ行に表示するフレーム
    title_frame = tk.Frame(indicator_window)
    title_frame.pack(pady=0)

    # 画像ラベル（左側）
    if icon_image:  # グローバル変数を使用
        icon_label = tk.Label(title_frame, image=icon_image, width=64, height=64)
        icon_label.pack(side="left", padx=(0, 0))

    # タイトルラベル（右側）
    title_label = tk.Label(title_frame, text=f"{task_title}中…", font=("Arial", 11))
    title_label.pack(side="left")

    # 進捗バーを中央に表示
    progress_frame = tk.Frame(indicator_window)
    progress_frame.pack(pady=(0, 5))  # 上に余白を設定

    # 進捗バー
    progress_bar = ttk.Progressbar(progress_frame, mode="indeterminate", length=250)
    progress_bar.pack()
    progress_bar.start(10)  # 進捗バーを開始

    # 停止用関数（呼び出されるとウィンドウを閉じる）
    def stop_indicator():
        try:
            if progress_bar.winfo_exists():  # ウィジェットが存在するか確認
                progress_bar.stop()
            if indicator_window.winfo_exists():  # ウィンドウが存在するか確認
                indicator_window.destroy()
        except:
            pass  # 何もしない

    return stop_indicator

##############################
#その他設定
##############################

# マウスホイールスクロールを有効にする関数
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

# Canvas の幅に追従させる関数
def resize_canvas(event):
    canvas_width = event.width
    canvas.itemconfig(canvas_window, width=canvas_width)

# シート名入力欄を有効・無効に切り替える関数
def toggle_sheet_entry():
    if sheet_mode.get() == 1:
        sheet_entry.config(state="normal")
    else:
        sheet_entry.config(state="disabled")

def toggle_merge_options_state():
    """
    範囲が1つなら「縦に結合」に固定して非活性化。
    範囲が2つ以上ならラジオボタンを活性化。
    """
    if len(range_frames) > 1:
        vertical_radio.config(state="normal")
        horizontal_radio.config(state="normal")
    else:
        merge_mode.set("vertical")  # 「縦に結合」に固定
        vertical_radio.config(state="disabled")
        horizontal_radio.config(state="disabled")

#画面中央寄せ
def center_window(root, width=570, height=730):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

#プレビューとプログレスバーは親ウィンドウのすぐそばに表示さす
def center_window2(parent, window, width=300, height=100):
    """
    指定したウィンドウを親ウィンドウの中央に表示する関数
    :param parent: 親ウィンドウ（rootなど）
    :param window: 表示する子ウィンドウ
    :param width: 子ウィンドウの幅
    :param height: 子ウィンドウの高さ
    """
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    x = parent_x + (parent_width // 2) - (width // 2)
    y = parent_y + (parent_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

# 最前面固定の状態を切り替える関数
def toggle_topmost():
    root.attributes("-topmost", is_topmost.get())

def draw_dashed_line(parent, width=2, color="#557CAE", pady=8, before=None):
    """セクション間に中央揃えの点線を引く"""
    canvas = tk.Canvas(parent, height=2, bg=background_color, highlightthickness=0)
    canvas.pack(fill="x", pady=pady, before=before)

    def draw_line():
        canvas.delete("all")
        canvas_width = canvas.winfo_width()
        canvas.create_line(50, 1, canvas_width - 50, 1, fill=color, width=width, dash=(4, 4))

    canvas.after(50, draw_line)
    canvas.bind("<Configure>", lambda event: draw_line())
    return canvas  # 作成したCanvasを返却する

dynamic_separator = None  # 動的な点線を管理するための変数
def draw_dynamic_separator(parent, after=None, pady=8):
    """最後の範囲入力フィールドの真下に1本だけ点線を引く"""
    global dynamic_separator

    # 既存の点線があれば削除
    if dynamic_separator:
        dynamic_separator.pack_forget()

    # 新しい点線を描画
    canvas = tk.Canvas(parent, height=2, bg=background_color, highlightthickness=0)
    canvas.pack(fill="x", pady=pady, after=after)
    dynamic_separator = canvas

    # キャンバス中央に線を描画
    def draw_line():
        canvas.delete("all")
        canvas_width = canvas.winfo_width()
        canvas.create_line(50, 1, canvas_width - 50, 1, fill="#557CAE", width=2, dash=(4, 4))

    canvas.after(50, draw_line)
    canvas.bind("<Configure>", lambda event: draw_line())

# ツールチップクラス
class Tooltip:
    """ツールチップを表示するクラス"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None

        # イベントバインド
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        """ツールチップを表示"""
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20

        # 新しいウィンドウでツールチップを作成
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  # ウィンドウの枠を非表示
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        self.tooltip_window.wm_attributes("-topmost", True)  # 常に最前面に表示

        # ツールチップ内のラベル
        label = tk.Label(self.tooltip_window, text=self.text, bg="#ffffe0", fg="black",
                         relief="solid", borderwidth=1, font=("Helvetica", 9, "normal"),anchor="w",justify="left")
        label.pack(ipadx=5, ipady=3, fill="both")

    def hide_tooltip(self, event=None):
        """ツールチップを非表示"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

####################################################################
#GUIのセットアップ
####################################################################

# GUIのセットアップ
root = tk.Tk()
root.title("Excelセルの値取り出し")
# photo = my_icon.get_photo_image4icon()  # PhotoImageオブジェクトの作成
# root.iconphoto(False, photo)         # アイコンの設定

# 背景色の設定
background_color = "#DEDEDE"  # ここで背景色を指定

#####アイコンとか、画像とか#####
#アイコン設定
def temp_path(relative_path):
    try:
        #Retrieve Temp Path
        base_path = sys._MEIPASS
    except Exception:
        #Retrieve Current Path Then Error 
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

#favicon読み込み
logo=temp_path('icon.ico')
# icon.png を読み込んで PhotoImage に変換
icon_path = temp_path("icon.png")
icon_image = None  # グローバル変数としてアイコン画像を保持

try:
    image_pil = Image.open(icon_path)
    image_pil = image_pil.resize((32, 32), Image.LANCZOS)  # サイズ調整
    icon_image = ImageTk.PhotoImage(image_pil)
except Exception as e:
    messagebox.showerror("エラー", f"画像の読み込みに失敗しました: {e}")

help_path = temp_path("help.png")  # 正しいパスを取得
help_image = None  # グローバル変数としてアイコン画像を保持
try:
    help_image_pil = Image.open(help_path)  # help.png を正しく読み込む
    help_image_pil = help_image_pil.resize((18, 18), Image.LANCZOS)  # サイズ調整
    help_image = ImageTk.PhotoImage(help_image_pil)  # PhotoImageオブジェクトを作成
except Exception as e:
    messagebox.showerror("エラー", f"画像の読み込みに失敗しました: {e}")

# ここで変数を定義（rootの作成後に実行）
sheet_mode = tk.IntVar(value=0)  # デフォルトは「すべてのシート」

# ttk.Styleを使ってテーマを設定
style = ttk.Style(root)
style.theme_use('clam')  # 'clam', 'alt', 'default', 'classic' などから選べます

# 背景色を全体に適用
root.configure(bg=background_color)

# ウィンドウを中央に配置
center_window(root)

# ウィンドウを常に最前面に表示
root.attributes("-topmost", True)

# ウィンドウを再表示する
root.deiconify()

# スクロール可能なフレームを作成
main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=1)

# Canvasの背景色を設定
canvas = tk.Canvas(main_frame, bg=background_color)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

# 縦スクロールバーの設定
scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# 背景色を設定したスクロール可能なフレーム
scrollable_frame = ttk.Frame(canvas)
scrollable_frame.config(style="TFrame")

# ウィンドウをCanvasに作成し、アンカー位置を適切に設定
canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

# Canvas の幅に追従させる関数
def resize_canvas(event):
    # 現在の Canvas の幅を取得し、scrollable_frame の幅を更新
    canvas_width = event.width
    canvas.itemconfig(canvas_window, width=canvas_width)

# Canvas のサイズが変更されたときに呼び出す
canvas.bind("<Configure>", resize_canvas)

# Canvasのスクロール領域を更新する関数
def update_scrollregion(event=None):
    # スクロール可能領域を更新（表示されているすべてのウィジェットに基づく）
    canvas.config(scrollregion=canvas.bbox("all"))

# scrollable_frame のサイズ変更時にスクロール領域を更新
scrollable_frame.bind("<Configure>", update_scrollregion)

# マウスホイールでスクロールを可能にする
canvas.bind_all("<MouseWheel>", on_mouse_wheel)

# 背景色を適用
style.configure("TFrame", background=background_color)
style.configure("TLabel", background=background_color)
style.configure("TButton", background=background_color)

# スタイル設定：buttonの色とか、いろいろ。です
style.configure("TEntry", fieldbackground="white", foreground="black")
style.map("TEntry", fieldbackground=[("disabled", "#DEDEDE")], foreground=[("disabled", "#a3a3a3")])
style.configure("Bold.TLabel", font=("Helvetica", 9, "bold"))  # フォントを太字に設定
style.configure("SmallFont.TLabel", font=("", 8))  # フォント名はデフォルトでサイズだけ小さく
style.configure("Small.TButton",background="#F7F4F0", font=("", 8), padding=(2,5))  # フォントサイズを小さく設定
style.map("Small.TButton",
          background=[("active", "#E4EAF3")],  # ホバー時の背景色
          foreground=[("active", "black")])   # ホバー時の文字色
style.configure("Highlight.TButton", background="#557CAE", foreground="white", font=("Helvetica", 9, "bold"))
style.map("Highlight.TButton",
          background=[('active', '#95C0D1')],   
          foreground=[('active', 'white')])    # ホバー時の文字色を白に設定
style.configure("Custom.TButton", 
                background="#F7F4F0",  # 背景色
                borderwidth=2,)
style.map("Custom.TButton", 
          background=[("active", "#E4EAF3"), ("disabled", "#D9D9D9")],  # ホバー時と非活性時の背景色
          foreground=[("active", "black"), ("disabled", "#A0A0A0")])   # ホバー時と非活性時の文字色
# ラベルとアイコンをまとめるフレーム
label_with_icon_frame_1 = ttk.Frame(scrollable_frame)
label_with_icon_frame_1.pack(anchor="center", pady=(0, 5))  # 余白を調整

# ラベルの追加
ttk.Label(label_with_icon_frame_1, text="➊ Excelファイルを選択してください(複数選択可)", 
          padding=(0, 5, 0, 0), style="Bold.TLabel").pack(side="left")

# ツールチップアイコン（help_imageを使用）
tooltip_icon_label_1 = ttk.Label(label_with_icon_frame_1, image=help_image)
tooltip_icon_label_1.pack(side="left", padx=(5, 0))

# ツールチップをアイコンに追加
Tooltip(tooltip_icon_label_1, (
    "【ファイル選択】ボタンからExcelファイルを選択してください。\n"
    "Ctrl+クリックで複数のファイルを選択できます。\n\n"
    "一度選択したあとでも、【ファイル選択】でさらに追加できます。\n"
    "選択したファイルはキーボードのDeleteキーで削除できます。\n"
    "【×クリア】を押すと、選択中のすべてのリストを削除します。\n\n"
    "※選択できるファイルは【.xlsx】【.xlsm】のみです\n"
    "※パスワード付きのファイルは処理できません"
))

# 「ファイル選択」と「クリア」ボタンを同じ行に配置
file_button_frame = ttk.Frame(scrollable_frame)
file_button_frame.pack(anchor="center", pady=5)
# ファイル選択ボタン
file_select_button = ttk.Button(file_button_frame, text="ファイル選択", command=select_files, style="Custom.TButton")
file_select_button.pack(side="left", padx=5)


# クリアボタン
ttk.Button(file_button_frame, text="✖　クリア", command=clear_file_list, style="Small.TButton").pack(side="left", padx=2)

# リストボックスのフレームを作成し、縦と横のスクロールバーを追加
listbox_frame = ttk.Frame(scrollable_frame)
listbox_frame.pack(anchor="center", pady=5, padx=30)

# 横スクロールバーの設定
x_scrollbar = tk.Scrollbar(listbox_frame, orient=tk.HORIZONTAL)
x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

# 縦スクロールバーの設定
y_scrollbar = tk.Scrollbar(listbox_frame)
y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# リストボックスの設定（縦と横スクロール対応）
file_list = tk.Listbox(listbox_frame, width=80, height=6, xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set, selectmode=tk.EXTENDED)
file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# キーイベントで選択したアイテムを削除
file_list.bind("<Delete>", delete_selected_files)

# スクロールバーの設定
x_scrollbar.config(command=file_list.xview)
y_scrollbar.config(command=file_list.yview)

# 点線を描画
draw_dashed_line(scrollable_frame)

####ファイル選択部分####
# ラベルとアイコンをまとめるフレーム
label_with_icon_frame = ttk.Frame(scrollable_frame)
label_with_icon_frame.pack(anchor="center", pady=2)
# ラベルの追加
ttk.Label(label_with_icon_frame, text="➋ 処理パターンを選択してください", padding=5, style="Bold.TLabel").pack(side="left")
# ツールチップアイコン（help_imageを使用）
tooltip_icon_label = ttk.Label(label_with_icon_frame, image=help_image)
tooltip_icon_label.pack(side="left", padx=(5, 0))
# ツールチップをアイコンに追加
Tooltip(tooltip_icon_label, (
    "処理するExcelシートのパターンを選択してください。\n"
    "すべてのシートまたは特定のシートを処理対象にできます。\n\n"
    "※「特定の名前を含むシートのみ」\n"
    "　 指定した文字列をシート名に含むシートが処理対象となります\n"
    "　 例）「損益」と入力\n"
    "　　➡「2.損益」「★損益表」「損益」などの名前がついたシートが対象になります"
))
# ラジオボタンを中央に配置しつつ左揃えにするためのフレーム
radio_frame = ttk.Frame(scrollable_frame)
radio_frame.pack(anchor="center")  # 全体として中央に配置
ttk.Radiobutton(radio_frame, text="すべてのシート(非表示シートを含む)", variable=sheet_mode, value=0, command=toggle_sheet_entry).pack(anchor="w")
ttk.Radiobutton(radio_frame, text="すべてのシート(非表示シートは含まない)", variable=sheet_mode, value=2, command=toggle_sheet_entry).pack(anchor="w")
ttk.Radiobutton(radio_frame, text="特定の名前を含むシートのみ", variable=sheet_mode, value=1, command=toggle_sheet_entry).pack(anchor="w")

# シート名入力部分のフレーム
sheet_frame = ttk.Frame(scrollable_frame)
sheet_frame.pack(anchor="center", pady=(10, 5))
# ラベルとエントリーフィールドを同じ行に配置
ttk.Label(sheet_frame, text="シート名を入力(部分一致):", padding=(0, 0, 5, 0)).pack(side="left")
sheet_entry = ttk.Entry(sheet_frame, style="TEntry", width=15)
sheet_entry.pack(side="left")
sheet_entry.config(state="disabled")  # 初期状態ではグレーアウト（無効化）

# 点線を描画
draw_dashed_line(scrollable_frame)

# ラベルとアイコンをまとめるフレーム
label_with_icon_frame_3 = ttk.Frame(scrollable_frame)
label_with_icon_frame_3.pack(anchor="center", pady=(5, 5))  # 余白を調整

# ラベルの追加
ttk.Label(label_with_icon_frame_3, text="➌ 取り出すセル範囲を入力してください",padding=(0, 5, 0, 0), style="Bold.TLabel").pack(side="left")

# ツールチップアイコン（help_imageを使用）
tooltip_icon_label_3 = ttk.Label(label_with_icon_frame_3, image=help_image)
tooltip_icon_label_3.pack(side="left", padx=(5, 0))

# ツールチップをアイコンに追加
Tooltip(tooltip_icon_label_3, (
    "取り出したいセル範囲を入力してください。\n"
    "例）C～R、10～20　/　W～W、3～3　\n\n"
    "【プレビュー】を押すと、指定した範囲がイメージ通りか確認できます。\n\n"
    "【＋範囲を追加】を押すと、範囲を複数指定できます。\n\n"
    "　※複数指定時の結合パターン\n"
    "　　「縦に結合」：取り出した範囲を縦(上→下)に結合します。\n"
    "　　「横に結合」：取り出した範囲を横(左→右)に結合します。\n"
    "　　 （上から指定した順番に結合されます）\n\n"
    "【△削除する】ボタンで、追加した範囲を削除できます。\n"
    "【↑】【↓】ボタンで、範囲の順位を調整できます。"
))

# 結合方向を選択するラジオボタン
merge_mode = tk.StringVar(value="vertical")  # 初期値は「縦に結合」

# 結合方向ラジオボタン用フレーム➌➌
merge_frame = ttk.Frame(scrollable_frame)
merge_frame.pack(anchor="center", pady=(5, 5))

# ラジオボタン: 縦に結合
vertical_radio = ttk.Radiobutton(merge_frame, text="縦に結合", variable=merge_mode, value="vertical")
vertical_radio.pack(side="left", padx=(0, 5))
# ラジオボタン: 横に結合
horizontal_radio = ttk.Radiobutton(merge_frame, text="横に結合", variable=merge_mode, value="horizontal")
horizontal_radio.pack(side="left", padx=(5, 0))

# ボタンを横に並べるためのフレーム
button_frame = ttk.Frame(scrollable_frame)
button_frame.pack(anchor="center", pady=(5, 5))

# 範囲を追加ボタン（左側）
ttk.Button(button_frame, text="＋範囲を追加", command=add_range_fields, style="Custom.TButton").pack(side="left", padx=(0, 5))

# プレビューボタン（右側）
preview_button = ttk.Button(button_frame, text="プレビュー", command=get_excel_preview, style="Highlight.TButton")
preview_button.pack(side="left", padx=(5, 0))

# 値を取得ボタンの中央揃え
get_value_button = ttk.Button(scrollable_frame, text=">> Excelで出力 <<", command=get_excel_values, style="Highlight.TButton")
get_value_button.pack(anchor="center",pady=(5,5),ipady=10,ipadx=20)

# 最前面固定のチェックボックスの状態を管理する変数
is_topmost = tk.BooleanVar(value=True)  # 初期値は「True」（最前面固定）
# 最前面固定のチェックボックスと画像をまとめるフレームを作成
bottom_frame = ttk.Frame(scrollable_frame, style="TFrame")  # フレーム作成
bottom_frame.pack(anchor="se",fill="x", pady=0)  # 下寄せ、横幅いっぱいに

# 最前面固定のチェックボックスを追加
is_topmost = tk.BooleanVar(value=True)  # 初期値はTrue（最前面固定）
topmost_checkbox = tk.Checkbutton( bottom_frame, text="ウィンドウを最前面に固定", variable=is_topmost, command=toggle_topmost, bg=background_color)
topmost_checkbox.pack(side="left", padx=(5, 5))  # 左寄せ

# 画像を読み込む
image = Image.open(icon_path)
image = image.resize((40, 40), Image.LANCZOS)  # サイズ調整
photo = ImageTk.PhotoImage(image)

# Labelに画像を配置（右寄せ）
image_label = tk.Label(bottom_frame, image=photo, bg=background_color)
image_label.photo = photo  # 参照を保持
image_label.pack(side="right", padx=(5, 5))  # 右寄せ

# 画像をクリックしてバージョン情報を表示する
def show_version_info():
    messagebox.showinfo("管理用", "1.2.0")
image_label.bind("<Button-1>", lambda event: show_version_info())

# + ボタンで範囲を追加
add_range_fields()

# # 結果のメッセージを表示するラベル
# message_label = ttk.Label(scrollable_frame, text="")
# message_label.pack(anchor="center")

# ウィンドウを表示させる
root.iconbitmap(logo)  # アイコンファイルのパスを設定
root.mainloop()