import os
import pickle
import gspread # gspread をインポート

import google.generativeai as genai
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials # Credentials クラスをインポート

from utils import retry_api_call # リトライデコレータをインポート

def configure_gemini(google_api_key):
    """!
    @brief Gemini APIの初期設定を行います。
    @param google_api_key (str): Google CloudプロジェクトのAPIキー。
    @return None
    @exception Exception Gemini APIの設定に失敗した場合。
    """
    try:
        genai.configure(api_key=google_api_key)
        print("Gemini APIの設定が完了しました。")
    except Exception as e:
        print(f"エラー: Gemini APIの設定に失敗しました。APIキーを確認してください。{e}")
        # ここで exit() する代わりに、呼び出し元でエラーハンドリングできるよう例外を再送出するか、
        # None を返すなどの方法も考えられますが、現状維持とします。
        exit(1) # エラーコード 1 で終了

def authenticate_google_apis(token_pickle_file, credentials_file, scopes):
    """!
    @brief Google APIs (Docs, Drive, Sheets) の認証を行い、サービスオブジェクト/クライアントを返します。
           認証情報が存在しない、または無効な場合は、OAuth 2.0フローを実行して認証情報を取得・保存します。
    @param token_pickle_file (str): 認証トークンを保存/読み込みするファイルパス。
    @param credentials_file (str): Google Cloudからダウンロードした認証情報ファイル (JSON) のパス。
    @param scopes (list): APIアクセスに必要なスコープのリスト。
    @return tuple(googleapiclient.discovery.Resource, googleapiclient.discovery.Resource, gspread.Client) | tuple(None, None, None):
            認証成功時はDocsサービス、Driveサービス、gspreadクライアントのタプル。
            失敗時は (None, None, None)。
    """
    creds = None
    if os.path.exists(token_pickle_file):
        with open(token_pickle_file, 'rb') as token:
            print(f"デバッグ: 既存のトークンファイル '{token_pickle_file}' を読み込みます。")
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print("認証トークンが期限切れです。リフレッシュを試みます...")
                creds.refresh(Request())
                print("認証トークンのリフレッシュに成功しました。")
            except Exception as e:
                print(f"警告: 認証トークンのリフレッシュに失敗しました: {e}")
                print("再認証が必要です。ブラウザを開いて認証を行ってください。")
                print(f"デバッグ: リフレッシュ失敗のため、新規認証フローを開始します。Scopes: {scopes}")
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
                creds = flow.run_local_server(port=0)
                print("デバッグ: run_local_server が完了しました。")
        else:
            if not os.path.exists(credentials_file):
                print(f"エラー: 認証情報ファイル '{credentials_file}' が見つかりません。")
                print("Google Cloud Consoleから認証情報ファイル (credentials.json) をダウンロードし、指定されたパスに配置してください。")
                return None, None, None
            print("認証情報が見つからないか無効です。新規認証を開始します。")
            print(f"デバッグ: 新規認証フローを開始します。Scopes: {scopes}")
            print("ブラウザを開いて認証を行ってください。")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
            creds = flow.run_local_server(port=0)
            print("デバッグ: run_local_server が完了しました。")
    try:
        docs_service = build('docs', 'v1', credentials=creds)
        # gspread は credentials オブジェクトを直接使用する
        gspread_client = gspread.authorize(creds)
        drive_service = build('drive', 'v3', credentials=creds)
        print("Google API認証成功。")
        return docs_service, drive_service, gspread_client
    except gspread.exceptions.APIError as ge: # gspread固有のAPIエラーをキャッチ
        print(f"エラー: gspread APIエラーが発生しました: {ge}")
        return None, None, None
    except Exception as e:
        print(f"Google APIサービスオブジェクトの構築に失敗しました: {e}")
        # 認証情報が原因である可能性を考慮し、古いトークンファイルを削除
        if os.path.exists(token_pickle_file):
            try:
                os.remove(token_pickle_file)
                print(f"古い認証トークンファイル '{token_pickle_file}' を削除しました。")
                print("スクリプトを再実行して再認証してください。")
            except OSError as oe:
                print(f"警告: 古い認証トークンファイル '{token_pickle_file}' の削除に失敗しました: {oe}")
        return None, None, None
    finally:
        # 認証が成功した場合（またはリフレッシュ/新規認証後）にトークンを保存
        if creds and creds.valid: # 有効な認証情報がある場合のみ保存
            with open(token_pickle_file, 'wb') as token:
                pickle.dump(creds, token)
                print(f"認証情報を '{token_pickle_file}' に保存しました。")

def translate_text_chunk(text_chunk, model_name, max_retries=2, initial_delay=5):
    """!
    @brief Gemini APIを使用して指定されたテキストチャンクを英語から日本語に翻訳します。
           API呼び出し時にエラーが発生した場合、リトライ処理を行います。
    @param text_chunk (str): 翻訳対象の英語テキストチャンク。
    @param model_name (str): 使用するGeminiモデルの名前 (例: "gemini-pro")。
    @param max_retries (int): 最大リトライ回数。
    @param initial_delay (float): 初回のリトライ待機時間（秒）。
    @return str: 翻訳された日本語テキスト。翻訳に失敗した場合は空文字列。
    @exception Exception 翻訳API呼び出し中に予期せぬエラーが発生した場合。
           リトライしても成功しなかった場合も含む。
    """
    if not text_chunk or not text_chunk.strip():
        print("警告: 翻訳対象のテキストが空です。スキップします。")
        return ""
    # --- ここまで ---
    try:
        model = genai.GenerativeModel(model_name)
        # プロンプトを調整：翻訳のみを出力するように指示を追加
        # ソフトウェア開発ドキュメントの翻訳に特化したプロンプト
        prompt = f"""あなたはソフトウェア開発ドキュメントの翻訳を専門とするエキスパートです。
以下の英語のテキストを、日本のソフトウェア開発者が読むことを想定し、自然かつ正確な日本語に翻訳してください。

- 専門用語（例: design pattern, dependency injection, asynchronous processing）や技術的な概念は、文脈に合わせて最も適切で一般的に使われる日本語訳を選択してください。必要であれば、カタカナ表記や英語表記のままにする方が良い場合もあります。
- コードスニペット、変数名、関数名、クラス名、ファイルパス、APIエンドポイント、UI要素のラベルなどは、原則として翻訳せず原文のまま残してください。ただし、コメント部分はこの限りではありません。
- 全体として、技術文書として正確性を保ちつつ、読みやすい日本語になるようにしてください。
- 応答には翻訳された日本語のテキストのみを含めてください。挨拶、前置き、後書き、翻訳に関する注釈などは一切不要です。

--- English Text ---
{text_chunk}
--- End English Text ---

--- Japanese Translation ---
"""
        # --- リトライデコレータを適用した内部関数 ---
        @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        def _generate_content_with_retry():
            return model.generate_content(prompt)
        # --- 内部関数ここまで ---

        response = _generate_content_with_retry()

        # レスポンスオブジェクトから翻訳テキストを抽出
        raw_translation = ""
        if hasattr(response, 'text'):
            raw_translation = response.text
        elif response.parts:
            raw_translation = "".join(part.text for part in response.parts)
        else:
            print("警告: 予期しないレスポンス形式です。")
            print(f"レスポンス内容: {response}")
            return "" # 空文字を返す

        # --- 前置きを除去する処理 (注意: 簡易的な実装であり、誤作動の可能性あり) ---
        # プロンプトで応答形式を厳密に制御できている場合、この処理は不要/簡略化できる可能性が高い
        lines = raw_translation.strip().split('\n')
        # 最初の行が一般的な応答パターンに一致するか確認
        common_greetings = ["はい、承知いたしました。", "承知しました。", "以下に翻訳します。", "翻訳結果は以下の通りです。"]
        cleaned_translation = raw_translation.strip() # デフォルトは元のテキスト
        if lines and any(lines[0].strip().startswith(greeting) for greeting in common_greetings):
            # 最初の行が挨拶文パターンに一致する場合、それ以降の行を結合する (空行も考慮)
            cleaned_translation = "\n".join(line for i, line in enumerate(lines) if i > 0 or line.strip()).strip()

        return cleaned_translation # クリーニング後のテキストを返す

    except Exception as e:
        print(f"エラー: テキストチャンクの翻訳中にエラーが発生しました。{e}")
        # エラーの詳細（例：APIからのエラーメッセージ）があれば表示するとデバッグに役立つ
        # (genaiライブラリのエラーオブジェクト構造に依存するため、要確認)
        raise # キャッチした例外を再スローする
        # リトライデコレータが最終的なエラーを送出する

def save_to_google_doc(docs_service, drive_service, title, chapter_titles,
                       translated_chunks, chapter_levels, max_retries=2, initial_delay=5):
    """!
    @brief 翻訳されたテキストチャンクを章ごとに結合し、新しいGoogleドキュメントに保存します。
           章タイトルには階層レベルに応じたMarkdown風の見出し (#, ##, ...) を付けます。
           API呼び出し時にエラーが発生した場合、リトライ処理を行います。
    @param docs_service (googleapiclient.discovery.Resource): Google Docs APIサービスオブジェクト。
    @param drive_service (googleapiclient.discovery.Resource): Google Drive APIサービスオブジェクト。
    @param title (str): 作成するGoogleドキュメントのタイトル。
    @param chapter_titles (list[str]): 各章のタイトルリスト。
    @param translated_chunks (list[str]): 各章に対応する翻訳済みテキストのリスト。
    @param chapter_levels (list[int]): 各章のブックマーク階層レベル (0始まり) のリスト。
    @param max_retries (int): 最大リトライ回数。
    @param initial_delay (float): 初回のリトライ待機時間（秒）。
    @return None
    """
    if not docs_service or not drive_service:
        print("エラー: Google APIサービスが利用できません。保存をスキップします。")
        return

    print(f"翻訳結果を新しいGoogleドキュメント '{title}' に保存中...")
    try:
        # 1. 新しいGoogleドキュメントを作成 (Drive APIを使用)
        # --- リトライデコレータを適用した内部関数 ---
        @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        def _create_doc_with_retry():
            body = {
                'name': title,
                'mimeType': 'application/vnd.google-apps.document'
            }
            return drive_service.files().create(body=body).execute()
        # --- 内部関数ここまで ---

        new_doc = _create_doc_with_retry()
        document_id = new_doc['id']

        print(f"新しいGoogleドキュメントを作成しました。ID: {document_id}")
        print(f"ドキュメントURL: https://docs.google.com/document/d/{document_id}/edit")

        # 2. 章タイトルと翻訳されたテキストを交互に結合
        #    階層レベルに応じてタイトルの見出しレベル (# の数) を変更
        content_to_insert = []
        # chapter_titles, translated_chunks, chapter_levels の長さが同じであることを前提とする
        # (main.py でそのように処理されているはず)
        for chap_title, translated_text, level in zip(chapter_titles, translated_chunks, chapter_levels):
            header_prefix = "#" * (level + 1) # レベル0なら"#", レベル1なら"##", ...
            content_to_insert.append(f"{header_prefix} {chap_title}\n\n") # 階層に応じた見出し
            content_to_insert.append(f"{translated_text}\n\n")
            # 章の区切りとしてMarkdownの水平線を追加
            content_to_insert.append("---\n\n")

        full_content = "".join(content_to_insert)

        # 3. 結合したテキストをドキュメントに挿入 (Docs APIを使用)
        # --- リトライデコレータを適用した内部関数 ---
        @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        def _batch_update_with_retry():
            requests = [
                {
                    'insertText': {
                        'location': {
                            'index': 1, # ドキュメントの先頭
                        },
                        'text': full_content
                    }
                }
            ]
            return docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        # --- 内部関数ここまで ---

        _batch_update_with_retry()

        print("Googleドキュメントへの保存完了。")

    except Exception as e: # googleapiclient.errors.HttpError など、より具体的なエラーを捕捉することも検討
        print(f"エラー: Googleドキュメントへの保存中にエラーが発生しました。 {e}")
        # リトライデコレータが最終的なエラーを送出する
        if hasattr(e, 'content'):
            print(f"APIエラー詳細: {e.content}")

def save_to_google_sheet(gspread_client, drive_service, title, chapter_titles,
                         original_texts, translated_chunks, max_retries=2, initial_delay=5):
    """!
    @brief 翻訳結果を新しいGoogleスプレッドシートに保存します。
           'タイトル', '原文', '訳文' の3列で出力します。
    @param gspread_client (gspread.Client): 認証済みのgspreadクライアント。
    @param drive_service (googleapiclient.discovery.Resource): Google Drive APIサービスオブジェクト (パーミッション設定等に将来的に使用する可能性)。
    @param title (str): 作成するGoogleスプレッドシートのタイトル。
    @param chapter_titles (list[str]): 各章のタイトルリスト。
    @param original_texts (list[str]): 各章に対応する原文テキストのリスト。
    @param translated_chunks (list[str]): 各章に対応する翻訳済みテキストのリスト。
    @param max_retries (int): 最大リトライ回数。
    @param initial_delay (float): 初回のリトライ待機時間（秒）。
    @return None
    """
    if not gspread_client:
        print("エラー: Google Sheets APIクライアントが利用できません。保存をスキップします。")
        return

    print(f"翻訳結果を新しいGoogleスプレッドシート '{title}' に保存中...")
    try:
        # 1. 新しいスプレッドシートを作成
        # --- リトライデコレータを適用した内部関数 ---
        @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        def _create_sheet_with_retry():
            return gspread_client.create(title)
        # --- 内部関数ここまで ---

        spreadsheet = _create_sheet_with_retry()
        print(f"新しいGoogleスプレッドシートを作成しました。ID: {spreadsheet.id}")

        print(f"スプレッドシートURL: {spreadsheet.url}")

        # 2. 最初のワークシートを取得
        worksheet = spreadsheet.get_worksheet(0) # 最初のシート

        # 3. ヘッダー行を作成
        header = ["タイトル", "原文", "訳文"]
        # ヘッダー追加はリトライ不要と判断する場合が多いが、念のため含めることも可能
        # @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        # def _append_header():
        #     worksheet.append_row(header, value_input_option='USER_ENTERED')
        # _append_header()
        worksheet.append_row(header, value_input_option='USER_ENTERED') # ヘッダー追加はリトライ対象外とする

        # 4. データを準備して一括書き込み (効率的)
        # --- リトライデコレータを適用した内部関数 ---
        @retry_api_call(max_retries=max_retries, initial_delay=initial_delay)
        def _append_rows_with_retry():
            rows_to_insert = [list(row) for row in zip(chapter_titles, original_texts, translated_chunks)]
            return worksheet.append_rows(rows_to_insert, value_input_option='USER_ENTERED')
        # --- 内部関数ここまで ---
        _append_rows_with_retry()

        print("Googleスプレッドシートへの保存完了。")
    except Exception as e:
        # より詳細なエラー情報を表示
        print(f"エラー: Googleスプレッドシートへの保存中に予期せぬエラーが発生しました。")
        print(f"エラータイプ: {type(e).__name__}")
        print(f"エラー詳細: {e}")
        # traceback.print_exc() # 必要に応じてスタックトレース全体を表示
        # APIエラーの詳細を表示する場合 (属性が存在するか確認した方がより安全)
        if hasattr(e, 'content'):
            print(f"APIエラー詳細: {e.content}")
