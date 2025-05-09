import sys
import json
import time
import os # osモジュールをインポート
import tkinter as tk
from tkinter import filedialog
import pandas as pd # pandas をインポート

from GoogleAdaptor import configure_gemini, translate_text_chunk, authenticate_google_apis, save_to_google_doc, save_to_google_sheet
from PdfEditor import split_text_by_bookmarks, split_text # extract_text_between_markers は削除
from tqdm import tqdm # tqdmライブラリをインポート


def select_pdf_file_old(): # 旧関数は不要になるためリネームまたは削除
    """GUIでPDFファイルを選択させます (旧バージョン)。

    @deprecated 新しい `select_file_and_format` 関数を使用してください。

    Returns:
        選択されたPDFファイルの絶対パス。キャンセルされた場合はNone。
    """
    root = tk.Tk()
    root.withdraw() # ルートウィンドウを非表示にする
    file_path = filedialog.askopenfilename(
        title="翻訳するPDFファイルを選択してください",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    root.destroy() # ダイアログが閉じたらルートウィンドウを破棄
    if file_path:
        print(f"選択されたPDFファイル: {file_path}")
        return file_path
    else:
        print("ファイル選択がキャンセルされました。")
        return None

def select_output_format_old(): # 旧関数は不要になるためリネームまたは削除
    """GUIで出力形式を選択させます (旧バージョン)。

    @deprecated 新しい `select_file_and_format` 関数を使用してください。

    Returns:
        選択された出力形式の識別子。キャンセルされた場合はNone。
    """
    root = tk.Tk()
    root.title("出力形式の選択")
    root.geometry("300x200") # ウィンドウサイズを調整

    output_format_var = tk.StringVar(value="google_doc") # デフォルト選択
    selected_format = None # 選択された形式を格納する変数

    tk.Label(root, text="出力形式を選択してください:", pady=10).pack()

    formats = {
        "Google ドキュメント": "google_doc",
        "Google スプレッドシート": "google_sheet",
        "AsciiDoc (ローカル)": "asciidoc",
        "Excel (ローカル)": "excel"
    }

    for display_text, value in formats.items():
        tk.Radiobutton(root, text=display_text, variable=output_format_var, value=value, anchor='w').pack(fill='x', padx=20)

    def on_ok():
        nonlocal selected_format
        selected_format = output_format_var.get()
        root.destroy()

    tk.Button(root, text="OK", command=on_ok, width=10).pack(pady=10)
    root.mainloop() # ウィンドウを表示し、ユーザーの操作を待つ
    return selected_format

def select_file_and_format(use_google_drive=True):
    """単一のGUIウィンドウでPDFファイルと出力形式を選択させます。

    Args:
        use_google_drive: Google Drive関連の出力形式を有効にするかどうか。
                          Falseの場合、該当するラジオボタンが無効化されます。

    Returns:
        (選択されたPDFファイルの絶対パス, 選択された出力形式の識別子) のタプル。
        キャンセルされた場合やファイル/形式が選択されなかった場合は (None, None)。
    """
    root = tk.Tk()
    root.title("PDF翻訳設定")
    # root.geometry("400x300") # 必要に応じてサイズ調整

    selected_pdf_path = tk.StringVar()
    # 選択された出力形式を保持するためのBooleanVarを格納する辞書
    selected_formats_vars = {}
    result = {"pdf_path": None, "output_formats": []} # 結果を格納する辞書 (複数形に変更)

    # --- PDFファイル選択部分 ---
    # (変更なし)
    # ... (省略) ...
    pdf_frame = tk.Frame(root, pady=10)
    pdf_frame.pack(fill='x', padx=10)

    tk.Label(pdf_frame, text="翻訳するPDFファイル:").pack(side=tk.LEFT, padx=5)
    # 選択されたパスを表示するラベル (読み取り専用風)
    pdf_path_label = tk.Label(pdf_frame, textvariable=selected_pdf_path, relief="sunken", width=40, anchor='w')
    pdf_path_label.pack(side=tk.LEFT, fill='x', expand=True, padx=5)

    def browse_pdf():
        file_path = filedialog.askopenfilename(
            title="翻訳するPDFファイルを選択してください",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            selected_pdf_path.set(file_path)
            print(f"選択されたPDFファイル: {file_path}")

    tk.Button(pdf_frame, text="参照...", command=browse_pdf).pack(side=tk.LEFT, padx=5)

    # --- 出力形式選択部分 ---
    format_frame = tk.LabelFrame(root, text="出力形式を選択してください", padx=10, pady=10)
    format_frame.pack(fill='x', padx=10, pady=5)

    formats = {
        "Google ドキュメント": "google_doc",
        "Google スプレッドシート": "google_sheet",
        "AsciiDoc (ローカル)": "asciidoc",
        "Excel (ローカル)": "excel"
    }

    checkboxes = {} # チェックボックスウィジェットを保持する辞書
    for display_text, value in formats.items():
        # 各形式に対応するBooleanVarを作成 (デフォルトはFalse=チェックなし)
        selected_formats_vars[value] = tk.BooleanVar(value=False)
        state = tk.NORMAL
        # Google Driveを使用しない設定の場合、Google関連の選択肢を無効化(グレーアウト)
        if not use_google_drive and value.startswith("google"):
            state = tk.DISABLED
        checkboxes[value] = tk.Checkbutton(format_frame, text=display_text, variable=selected_formats_vars[value], anchor='w', state=state)
        checkboxes[value].pack(fill='x')

    # --- OK/Cancel ボタン ---
    button_frame = tk.Frame(root, pady=10)
    button_frame.pack()

    def on_ok():
        if selected_pdf_path.get(): # PDFが選択されているか確認
            result["pdf_path"] = selected_pdf_path.get()
            # 選択された形式をリストに収集
            selected_list = [fmt for fmt, var in selected_formats_vars.items() if var.get()]
            if not selected_list:
                from tkinter import messagebox
                messagebox.showwarning("出力形式未選択", "少なくとも1つの出力形式を選択してください。")
                return # ウィンドウを閉じない
            result["output_formats"] = selected_list
            root.destroy()
        else:
            # messagebox を使うためにインポート
            from tkinter import messagebox
            messagebox.showwarning("PDF未選択", "翻訳するPDFファイルを選択してください。")

    tk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="キャンセル", command=root.destroy, width=10).pack(side=tk.LEFT, padx=10)

    root.mainloop()
    return result["pdf_path"], result["output_formats"]

def save_to_asciidoc(output_base_path, chapter_titles, translated_chunks, chapter_levels):
    """翻訳結果をAsciiDoc形式のローカルテキストファイルに保存します。

    Args:
        output_base_path: 出力ファイルのベースパス（拡張子なし）。
                          例: "output/mydoc_translated"
        chapter_titles: 各章のタイトルリスト。
        translated_chunks: 各章に対応する翻訳済みテキストのリスト。
        chapter_levels: 各章のブックマーク階層レベル (0始まり) のリスト。

    Returns:
        None
    """
    output_dir = os.path.dirname(output_base_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"出力ディレクトリ '{output_dir}' を作成しました。")

    output_adoc_path = f"{output_base_path}.adoc"
    print(f"翻訳結果をAsciiDocファイル '{output_adoc_path}' に保存中...")
    try:
        with open(output_adoc_path, 'w', encoding='utf-8') as f:
            # AsciiDocのドキュメントタイトルを設定 (ファイル名から)
            doc_title = os.path.basename(output_base_path)
            f.write(f"= {doc_title}\n\n") # ドキュメントタイトル

            for title, text, level in zip(chapter_titles, translated_chunks, chapter_levels):
                # AsciiDocの見出しレベルは = の数 (レベル0 -> ==, レベル1 -> ===, ...)
                header_prefix = "=" * (level + 2)
                f.write(f"{header_prefix} {title}\n\n")
                f.write(f"{text}\n\n") # 翻訳テキスト
        print(f"AsciiDocファイルへの保存完了: {output_adoc_path}")
    except Exception as e:
        print(f"エラー: AsciiDocファイルへの保存中にエラーが発生しました: {e}")

# --- Excelセル長超過対応ヘルパー ---
EXCEL_MAX_CELL_LENGTH = 32767 # Excelのセルあたりの最大文字数

def _split_text_for_excel(text, max_length=EXCEL_MAX_CELL_LENGTH):
    """テキストを指定された最大長以下のチャンクに分割します。"""
    if not isinstance(text, str): # 文字列でない場合はそのまま返す
        return [text]
    if len(text) <= max_length:
        return [text] # 最大長以下なら分割しない

    chunks = []
    start = 0
    while start < len(text):
        end = start + max_length
        chunks.append(text[start:end])
        start = end
    return chunks
def save_to_excel(output_base_path, chapter_titles, original_texts, translated_chunks):
    """翻訳結果をExcelファイル (.xlsx) に保存します。

    Args:
        output_base_path: 出力ファイルのベースパス（拡張子なし）。
                          例: "output/mydoc_translated"
        chapter_titles: 各章のタイトルリスト。
        original_texts: 各章に対応する原文テキストのリスト。
        translated_chunks: 各章に対応する翻訳済みテキストのリスト。

    Returns:
        None
    """
    output_xlsx_path = f"{output_base_path}.xlsx"
    print(f"翻訳結果をExcelファイル '{output_xlsx_path}' に保存中...")
    try:
        # --- セル長超過対応 ---
        excel_data = [] # Excelに出力するデータを格納するリスト
        for title, original, translated in zip(chapter_titles, original_texts, translated_chunks):
            # 原文と訳文を最大セル長で分割
            original_chunks = _split_text_for_excel(original)
            translated_chunks_split = _split_text_for_excel(translated) # 変数名を変更

            # 分割されたチャンクの最大数に合わせて行を作成
            max_rows = max(len(original_chunks), len(translated_chunks_split))

            for i in range(max_rows):
                # 最初の行にはタイトルを、それ以降は空文字を設定
                current_title = title if i == 0 else ""
                # 各行に対応する原文チャンクを取得 (なければ空文字)
                current_original = original_chunks[i] if i < len(original_chunks) else ""
                # 各行に対応する訳文チャンクを取得 (なければ空文字)
                current_translated = translated_chunks_split[i] if i < len(translated_chunks_split) else ""

                excel_data.append({
                    "タイトル": current_title,
                    "原文": current_original,
                    "訳文": current_translated
                })
        # --- ここまで ---

        df = pd.DataFrame(excel_data) # 整形されたデータからDataFrameを作成
        df.to_excel(output_xlsx_path, index=False, engine='openpyxl') # openpyxl をエンジンとして指定
        print(f"Excelファイルへの保存完了: {output_xlsx_path}")
    except Exception as e:
        print(f"エラー: Excelファイルへの保存中にエラーが発生しました: {e}")

def main():
    """PDF翻訳ツールのメイン処理。

    設定ファイルに基づいてPDFを読み込み、ブックマークで章に分割し、
    各章をGoogle Gemini APIで翻訳して、指定された形式で結果を保存します。

    処理フロー:
    1. 設定ファイル (config.json) を読み込みます。
    2. GUIでユーザーにPDFファイルと出力形式を選択させます。
    3. Google API (Gemini, Docs, Drive, Sheets) の認証と設定を行います (必要に応じて)。
    4. PDFをブックマークに基づいて章分割します。
    5. 各章のテキストを、必要であればさらにチャンク分割して翻訳APIを呼び出します。
    6. 翻訳結果を選択された形式 (Google Docs/Sheets, AsciiDoc, Excel) で保存します。
    @return None
    """
    # 設定ファイルを読み込む
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        print("エラー: config.json が見つかりません。")
        sys.exit(1)
    except json.JSONDecodeError:
        print("エラー: config.json の形式が不正です。")
        sys.exit(1)

    # --- 設定値の取得 ---
    # pdf_file_path = config.get("pdf_file_path") # config.json から読み込む代わりにGUIで選択
    max_chunk_size = config.get("max_chunk_size")
    # max_chunk_size が数値でない場合のデフォルト値設定やエラー処理を追加
    if not isinstance(max_chunk_size, int) or max_chunk_size <= 0:
        print(f"警告: config.json の max_chunk_size ({max_chunk_size}) が不正な値です。デフォルト値 10000 を使用します。")
        max_chunk_size = 10000 # デフォルト値

    google_api_key = config.get("google_api_key")
    sleep_time = config.get("sleep_time")
    model_name = config.get("model_name")
    token_pickle_file = config.get("token_pickle_file")
    credentials_file = config.get("credentials_file")
    scopes_from_config = config.get("scopes", []) # スコープリスト、なければ空リスト
    use_google_drive = config.get("use_google_drive", True) # デフォルトはTrue
    retry_count = config.get("retry_count", 2) # デフォルトリトライ回数を2に設定
    initial_retry_delay = config.get("initial_retry_delay", 5) # デフォルト初期遅延を5秒に設定
    # output_format はGUIで選択
    output_file_path = config.get("output_file_path", "output/result") # デフォルトパス設定

    # Google Driveを使用しない場合は、関連スコープを除外
    scopes = scopes_from_config
    if not use_google_drive:
        scopes = [s for s in scopes_from_config if 'drive' not in s and 'spreadsheets' not in s and 'documents' not in s]
        print("情報: Google Driveを使用しない設定のため、関連APIスコープを除外しました。")

    # --- 1. GUIでPDFファイルと出力形式を選択 ---
    pdf_file_path, selected_output_formats = select_file_and_format(use_google_drive)

    if not pdf_file_path or not selected_output_formats:
        print("PDFファイルまたは出力形式が選択されなかった（キャンセルされた）ため、処理を終了します。")
        sys.exit(0)
    print(f"選択された出力形式: {', '.join(selected_output_formats)}")
    # --- ここまで ---
    # --- 必須設定値のチェック ---
    # max_chunk_size, retry_count, initial_retry_delay はデフォルト値があるためチェック対象から外す
    if not all([google_api_key, sleep_time is not None, model_name, token_pickle_file, credentials_file, scopes_from_config]): # 元のscopes_from_configでチェック
        print("エラー: config.json に必須の設定項目が不足しています。")
        print("必要な項目: google_api_key, sleep_time, model_name, token_pickle_file, credentials_file, scopes, use_google_drive, output_file_path")
        sys.exit(1) # 設定不足はエラーとして 1 で終了

    filename = pdf_file_path.split('/')[-1].split('.')[0]   # PDFファイルパスから拡張子を除いたファイル名を取得
    # --- 出力ファイル名/タイトルの設定 (複数形式に対応) ---
    # Google Drive 用のタイトルと同じサフィックスを付ける
    filename_with_suffix = f"{filename}_translated"
    # output_file_path からディレクトリ部分を取得
    output_dir = os.path.dirname(output_file_path)
    output_base_path = os.path.join(output_dir, filename_with_suffix) # 例: "output/filename_translated"
    # Google Drive 用のドキュメント/スプレッドシートタイトル
    output_title = filename_with_suffix # ローカル保存と同じベース名を使用

    # --- 2. Gemini APIの設定 ---
    configure_gemini(google_api_key) # configure自体はリトライ不要と想定

    # --- 3. Google APIs (Docs, Drive, Sheets) の認証 (必要に応じて) ---
    # 渡すスコープは use_google_drive に基づいてフィルタリングされたもの
    docs_service, drive_service, gspread_client = authenticate_google_apis(token_pickle_file, credentials_file, scopes) # 認証プロセス内のAPI呼び出しはリトライ対象外とする
    # Google Driveを使用しない場合、drive_service は None になる可能性がある
    # Google関連の出力形式が選択されている場合、必要なサービスが認証されているか後でチェックする
    # ここでのチェックは簡略化（認証関数がNoneを返した場合のみエラーとする）
    if any(fmt.startswith("google") for fmt in selected_output_formats) and (docs_service is None or drive_service is None or gspread_client is None):
        print("Google API認証に失敗したため、処理を終了します。(Google出力形式選択時)")
        sys.exit(1)

    # --- 4. PDFをブックマークに基づいて章分割 ---
    # split_text_by_bookmarks は {章タイトル: {"text": 章テキスト, "level": 階層レベル}} の辞書を返す
    chapters_dict = split_text_by_bookmarks(pdf_file_path)
    if not chapters_dict:
        print("PDFから章を抽出できませんでした（ブックマークがないか、処理エラー）。処理を終了します。")
        sys.exit()

    # 抽出された章のデータからタイトル、テキスト、レベルをリストに分解
    chapter_titles = list(chapters_dict.keys()) # 章タイトルも保持しておく
    chapter_data = list(chapters_dict.values()) # [{"text": ..., "level": ...}, ...] のリスト
    chapter_texts = [data["text"] for data in chapter_data] # 翻訳対象のテキスト本体
    chapter_levels = [data["level"] for data in chapter_data] # 各章の階層レベル

    print(f"PDFは {len(chapter_texts)} 個の章（またはセクション）に分割されました。")

    # --- 5. 各章を翻訳 ---
    translated_chapters = [] # 変数名をより明確に
    original_texts_for_output = [] # 出力用の原文リスト (Excel/Spreadsheet用)
    print(f"合計 {len(chapter_texts)} 個の章を翻訳します...")

    # --- 動作確認用: 処理する最大章数を設定 ---
    # 通常実行時は None に設定するか、このブロック自体を削除/コメントアウト
    MAX_CHAPTERS_FOR_TEST = None # 通常実行時は None に設定するか、このブロック自体を削除
    # --- ここまで ---

    # tqdmのtotalを計算（章数またはMAX_CHAPTERS_FOR_TESTの小さい方、Noneの場合も考慮）
    total_chapters_to_process = min(len(chapter_texts), MAX_CHAPTERS_FOR_TEST) if 'MAX_CHAPTERS_FOR_TEST' in locals() and MAX_CHAPTERS_FOR_TEST is not None else len(chapter_texts)

    # tqdmで章ごとの進捗を表示
    # chapter_texts には split_text_by_bookmarks で抽出・絞り込みされたテキストが入っている
    for i, chapter_text_from_pdf in enumerate(tqdm(chapter_texts, desc="章翻訳処理", total=total_chapters_to_process, unit="章")):
        # --- 動作確認用: 上限に達したらループを抜ける ---
        # MAX_CHAPTERS_FOR_TEST が定義されていて None でなく、現在のインデックスがそれを超えた場合
        if 'MAX_CHAPTERS_FOR_TEST' in locals() and MAX_CHAPTERS_FOR_TEST is not None and i >= MAX_CHAPTERS_FOR_TEST:
            print(f"\n動作確認のため、{MAX_CHAPTERS_FOR_TEST}章で処理を中断します。")
            break
        # --- ここまで ---

        current_chapter_title = chapter_titles[i] # 現在の章タイトル
        # PdfEditor.py内で絞り込みが行われたテキストを使用
        text_to_translate = chapter_text_from_pdf # split_text_by_bookmarks から取得したテキストをそのまま使用
        print(f"\n--- 章 {i+1}/{total_chapters_to_process}: '{current_chapter_title}' (処理対象テキスト長: {len(text_to_translate)}) を処理中 ---")

        # スプレッドシート/Excelには、実際に翻訳APIに渡したテキスト(またはその元となった章テキスト)を記録
        # (split_text_by_bookmarksで絞り込まれたテキスト)
        original_texts_for_output.append(text_to_translate)

        translated_chapter = "" # この章全体の翻訳結果を格納する変数

        try:
            # 翻訳対象のテキスト (絞り込み後 or 元の章テキスト) が最大チャンクサイズを超えるかチェック
            if len(text_to_translate) > max_chunk_size:
                print(f"  テキスト長 ({len(text_to_translate)}) が最大チャンクサイズ ({max_chunk_size}) を超過。分割して翻訳します。")
                sub_chunks = split_text(text_to_translate, max_chunk_size)
                translated_sub_chunks = []
                # 分割されたサブチャンクを翻訳 (サブチャンクごとの進捗表示)
                for j, sub_chunk in enumerate(tqdm(sub_chunks, desc=f"    章{i+1}内チャンク翻訳", leave=False, unit="サブチャンク")):
                    # translate_text_chunk内でリトライが行われる
                    translated_sub = translate_text_chunk(sub_chunk, model_name, retry_count, initial_retry_delay)
                    translated_sub_chunks.append(translated_sub if translated_sub else "") # 翻訳失敗時は空文字を追加
                    # print(f"    サブチャンク {j+1} 翻訳完了。{sleep_time} 秒待機...")
                    time.sleep(sleep_time)
                # 分割された翻訳結果を結合
                translated_chapter = "\n".join(translated_sub_chunks)
                print(f"  章 '{current_chapter_title}' の分割翻訳完了。")
            else:
                # テキストがmax_chunk_size以下の場合、そのまま翻訳
                print(f"  テキスト長 ({len(text_to_translate)}) が最大チャンクサイズ以下。直接翻訳します。")
                # translate_text_chunk内でリトライが行われる
                translated_chapter = translate_text_chunk(text_to_translate, model_name, retry_count, initial_retry_delay)
                print(f"  章 '{current_chapter_title}' 翻訳完了。")
                time.sleep(sleep_time) # 翻訳APIへの負荷軽減のため待機

            translated_chapters.append(translated_chapter) # 翻訳結果をリストに追加

        except Exception as e: # 翻訳API呼び出し等でのエラー
            print(f"エラー: 章 '{current_chapter_title}' (インデックス {i}) の翻訳処理中にエラーが発生しました: {e}。この章をスキップします。")
            translated_chapters.append(f"--- 章 '{current_chapter_title}' 翻訳失敗: {e} ---") # エラーが発生した章のプレースホルダー
            # エラー時も原文リストには対応する要素を追加しておく（空文字など）
            # original_texts_for_output.append(text_to_translate) # 上で追加済み

    # --- 6. 選択された形式で翻訳結果を保存 ---
    # 処理された章の数を取得 (翻訳ループが途中で終了した場合を考慮)
    num_processed_chapters = len(translated_chapters) # 翻訳結果リストの長さが実際に処理された章数

    # デバッグ用に各リストの長さを表示 (必要に応じてコメントアウト)
    print(f"\n--- 保存処理前のリスト長確認 ---")
    print(f"chapter_titles: {len(chapter_titles)}")
    print(f"original_texts_for_output: {len(original_texts_for_output)}")
    print(f"translated_chapters: {len(translated_chapters)}")
    print(f"chapter_levels: {len(chapter_levels)}")
    print(f"実際に処理された章数 (num_processed_chapters): {num_processed_chapters}")
    print(f"---------------------------------")

    # 出力に使用するリストを、実際に処理された章の数に合わせてスライスする
    output_chapter_titles = chapter_titles[:num_processed_chapters]
    output_chapter_levels = chapter_levels[:num_processed_chapters]

    # --- 選択された各形式で保存処理を実行 ---
    for output_format in selected_output_formats:
        print(f"\n--- 出力形式 '{output_format}' で保存を開始 ---")
        try:
            if output_format == "google_doc":
                # Google Driveを使用する設定かつ、認証が成功している場合のみ実行
                if not use_google_drive or not docs_service or not drive_service:
                    print("エラー: Googleドキュメントへの保存に必要な設定または認証が不足しています。スキップします。")
                    continue # 次の形式へ
                save_to_google_doc(docs_service, drive_service, output_title, output_chapter_titles,
                                   translated_chapters, output_chapter_levels,
                                   max_retries=retry_count, initial_delay=initial_retry_delay)
            elif output_format == "google_sheet":
                # Google Driveを使用する設定の場合のみ実行 (SheetsもDrive APIを使うことがあるため)
                if not use_google_drive or not gspread_client or not drive_service: # 必要なサービスを確認
                     print("エラー: Googleスプレッドシートへの保存に必要な設定または認証が不足しています。スキップします。")
                     continue # 次の形式へ
                save_to_google_sheet(gspread_client, drive_service, output_title, output_chapter_titles,
                                     original_texts_for_output, translated_chapters,
                                     max_retries=retry_count, initial_delay=initial_retry_delay)
            elif output_format == "asciidoc":
                save_to_asciidoc(output_base_path, output_chapter_titles, translated_chapters, output_chapter_levels)
            elif output_format == "excel":
                save_to_excel(output_base_path, output_chapter_titles, original_texts_for_output, translated_chapters)
            else:
                # 基本的にここには到達しないはず (GUIで選択されたもののみのため)
                print(f"警告: 未知の出力形式 '{output_format}' が指定されました。スキップします。")
        except Exception as save_e:
            # 各保存処理中の予期せぬエラーをキャッチ
            print(f"エラー: 出力形式 '{output_format}' での保存中にエラーが発生しました: {save_e}")
            # エラーが発生しても、他の形式の保存は試みる

    print("--- PDF翻訳完了 ---")

# このスクリプトが直接実行された場合にmain()関数を呼び出す
if __name__ == "__main__":
    main()