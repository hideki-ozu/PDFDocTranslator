import time
import functools
import random
from googleapiclient.errors import HttpError
from gspread.exceptions import APIError

def retry_api_call(max_retries, initial_delay=1, backoff=2, jitter=0.1,
                   retry_status_codes=(500, 502, 503, 504)):
    """API呼び出しをリトライするデコレータ。

    指数バックオフとジッターを使用してリトライを行います。
    指定されたHTTPステータスコードを持つHttpErrorまたはAPIErrorが発生した場合に
    リトライを実行します。

    Args:
        max_retries: 最大リトライ回数。
        initial_delay: 初回のリトライ待機時間（秒）。デフォルトは1秒。
        backoff: 遅延時間を増加させる係数（例: 2は毎回2倍にする）。デフォルトは2。
        jitter: 待機時間に加えるランダムな変動の割合（例: 0.1は±10%）。デフォルトは0.1。
        retry_status_codes: リトライ対象とするHTTPステータスコードのタプル。
                              デフォルトは (500, 502, 503, 504)。

    Returns:
        デコレートされた関数。
    """
    def decorator(func):
        @functools.wraps(func) # 元の関数のメタデータを保持
        def wrapper(*args, **kwargs): # 可変長引数に対応
            retries = 0
            current_delay = initial_delay
            while retries <= max_retries:
                try:
                    return func(*args, **kwargs)
                except (HttpError, APIError) as e:
                    status_code = None
                    # エラーオブジェクトからステータスコードを取得
                    if isinstance(e, HttpError) and hasattr(e, 'resp') and hasattr(e.resp, 'status'):
                        status_code = e.resp.status
                    elif isinstance(e, APIError) and hasattr(e, 'response') and hasattr(e.response, 'status_code'):
                        # gspread の APIError は response.status_code を持つ場合がある
                        status_code = e.response.status_code
                    # TODO: 他の gspread エラーで status_code を取得する方法があればここに追加

                    if status_code in retry_status_codes:
                        if retries < max_retries:
                            retries += 1 # リトライ回数をインクリメント
                            # ジッターを加えた待機時間
                            wait_time = current_delay + random.uniform(-jitter * current_delay, jitter * current_delay)
                            print(f"警告: API呼び出しでエラー (ステータス: {status_code})。リトライします ({retries}/{max_retries})。{wait_time:.2f}秒待機...")
                            time.sleep(wait_time)
                            current_delay *= backoff # 次の遅延時間を計算
                        else:
                            print(f"エラー: リトライ上限 ({max_retries}回) に達しました。ステータス: {status_code}")
                            raise e # 最終的なエラーを再送出
                    else:
                        print(f"情報: リトライ対象外のエラーが発生しました (ステータス: {status_code})。エラー詳細: {e}")
                        raise e # リトライ対象外のエラーは再送出
                except Exception as e: # 予期しないその他の例外
                    print(f"エラー: API呼び出し中に予期しない例外が発生しました: {e}")
                    raise e # リトライ対象外の例外はそのまま送出
        return wrapper
    return decorator