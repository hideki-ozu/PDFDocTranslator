import PyPDF2
import PyPDF2.errors # PyPDF2のエラーをインポート
import os
from collections import OrderedDict # 順序付き辞書を使うためにインポート
from tqdm import tqdm # tqdmライブラリをインポート

def extract_text_from_pdf(pdf_path):
    """PDFファイルから全ページのテキストを抽出して結合します。

    @deprecated この関数は現在 main.py から直接使用されていません。
                ブックマークに基づいた分割処理には `split_text_by_bookmarks` を
                使用してください。

    Args:
        pdf_path: 読み込むPDFファイルのパス。

    Returns:
        抽出された全テキスト。エラー発生時やテキストが抽出できなかった場合はNone。

    Raises:
        PyPDF2.errors.PdfReadError: PDFファイルの読み込みに失敗した場合。
    """
    print(f"'{pdf_path}' からテキストを抽出中...")
    full_text = ""
    try:
        # ファイルが存在するか確認
        if not os.path.exists(pdf_path):
            print(f"エラー: PDFファイル '{pdf_path}' が見つかりません。")
            return None
        # ファイルがPDFか簡易的に確認 (拡張子)
        if not pdf_path.lower().endswith('.pdf'):
            print(f"エラー: 指定されたファイルはPDFファイルではない可能性があります - {pdf_path}")
            return None

        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = len(reader.pages)
            print(f"ページ数: {num_pages}")
            # tqdmを使ってページごとの読み込み進捗を表示
            for page_num in tqdm(range(num_pages), desc="PDF読み込み", unit="ページ"):
                page = reader.pages[page_num]
                page_text = page.extract_text()
                if page_text: # テキストが抽出できた場合のみ追加
                    full_text += page_text
                # else:
                #     print(f"デバッグ: ページ {page_num + 1} からテキストを抽出できませんでした。")

    except FileNotFoundError: # このブロックは通常 os.path.exists で捕捉されるはずだが、念のため残す
        print(f"エラー: PDFファイル '{pdf_path}' が見つかりません。")
        return None
    except PyPDF2.errors.PdfReadError as e: # より具体的なエラーをキャッチ
        print(f"エラー: PDFファイルの読み込みに失敗しました。ファイルが破損しているか、パスワードで保護されている可能性があります。 {e}")
        return None
    except Exception as e: # 予期せぬエラー
        print(f"エラー: PDFの読み込み中に予期せぬエラーが発生しました。{e}")
        return None

    if not full_text:
        print("警告: PDFからテキストを抽出できませんでした。画像ベースのPDFである可能性があります。")
        return None # テキストが空の場合もNoneを返すか、空文字を返すかは要件による

    print(f"PDFからのテキスト抽出完了。総文字数: {len(full_text)}")
    return full_text

# --- 新しい関数: ブックマークに基づいてテキストを分割 ---
def split_text_by_bookmarks(pdf_path):
    """PDFのブックマーク（アウトライン）に基づきテキストを分割します。

    ブックマークの階層構造を保持し、各ブックマークに対応するテキストを抽出します。
    ブックマークのタイトルをマーカーとして、テキストの絞り込みも試みます。

    Args:
        pdf_path: 処理対象のPDFファイルのパス。

    Returns:
        キーが章（ブックマーク）タイトル、値が {"text": 章のテキスト, "level": 階層レベル (0始まり)}
        の `collections.OrderedDict`。
        ブックマークが存在しない場合やエラー発生時はNoneを返します。

    """
    print(f"'{pdf_path}' のブックマーク（アウトライン）に基づいてテキストを分割中...")
    chapters = OrderedDict() # 章のタイトルとテキストを格納 (順序維持)

    try:
        # ファイルが存在するか確認
        if not os.path.exists(pdf_path):
            print(f"エラー: PDFファイル '{pdf_path}' が見つかりません。")
            return None
        # ファイルがPDFか簡易的に確認 (拡張子)
        if not pdf_path.lower().endswith('.pdf'):
            print(f"エラー: 指定されたファイルはPDFファイルではない可能性があります - {pdf_path}")
            return None

        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            bookmarks = reader.outline

            if not bookmarks:
                print("警告: このPDFにはブックマーク（目次）が見つかりませんでした。章分割はできません。")
                # ブックマークがない場合は、全テキストを一つの章として返すフォールバック処理
                full_text = "" # ここで全テキスト抽出を試みる
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text
                if full_text:
                    # ブックマークがない場合、タイトルを固定文字列にし、レベルを0とする
                    chapters["Full Text (No Bookmarks)"] = {"text": full_text, "level": 0}
                    return chapters # 単一要素の辞書を返す
                else:
                    print("警告: ブックマークがなく、テキストも抽出できませんでした。")
                    return None # テキストもなければNoneを返す

            # 再帰的にブックマーク情報を収集する内部関数
            def _get_bookmarks(outline_items, level=0):
                """ブックマークリストを再帰的に探索し、情報を収集する。"""
                for item in outline_items:
                    # PyPDF2のoutlineはリストとDestinationオブジェクトの混合リスト
                    # リストの場合は、それがサブブックマークのリストなので再帰呼び出し
                    # Destinationオブジェクトの場合は、それがブックマーク本体
                    # (参考: https://pypdf2.readthedocs.io/en/latest/user/bookmarks.html)
                    if isinstance(item, list):
                        # サブブックマークを再帰的に処理
                        _get_bookmarks(item, level + 1)
                    else:
                        # ブックマークアイテム
                        try:
                            page_index = reader.get_destination_page_number(item)
                            # ページ番号が取得できないブックマークはスキップ
                            if page_index is not None: # ページ番号が取得できた場合のみ追加
                                # (タイトル, 開始ページインデックス, 階層レベル) のタプルをリストに追加
                                all_bookmarks_info.append((item.title, page_index, level))
                            else:
                                print(f"警告: ブックマーク '{item.title}' は有効なページを指していません。スキップします。")
                        except Exception as e:
                            print(f"警告: ブックマーク '{item.title}' のページ番号取得に失敗しました。スキップします。 {e}")

            all_bookmarks_info = [] # (title, page_index, level) のタプルを格納するリスト
            _get_bookmarks(bookmarks) # ブックマーク情報収集の開始

            if not all_bookmarks_info:
                print("警告: 有効なブックマーク情報が見つかりませんでした。")
                return None # 有効なブックマークがなければNoneを返す

            # 収集した全てのブックマーク情報を、ページ番号を基準に昇順でソートする
            sorted_all_bookmarks_info = sorted(all_bookmarks_info, key=lambda item: item[1])

            # --- extract_text_between_markers のロジックで使用するヘルパー関数をここに移動 ---
            def _find_ignoring_whitespace(haystack, needle, start_offset=0):
                """
                haystack内でneedleを検索します。haystackとneedleの両方の空白文字は無視されます。
                見つかった場合、haystackにおける開始インデックス（空白無視前の元のインデックス）を返します。見つからない場合は-1を返します。
                """
                if not needle: return start_offset

                h_len, n_len = len(haystack), len(needle)
                h_idx, n_idx = start_offset, 0
                potential_start = -1

                while h_idx < h_len:
                    # haystackの空白をスキップ
                    while h_idx < h_len and haystack[h_idx].isspace():
                        if potential_start == -1: start_offset += 1 # マッチ開始前ならオフセットも進める
                        h_idx += 1
                    if h_idx == h_len: break

                    # needleの空白をスキップ
                    while n_idx < n_len and needle[n_idx].isspace(): n_idx += 1
                    if n_idx == n_len: return potential_start

                    # 非空白文字を比較
                    if h_idx < h_len and n_idx < n_len and haystack[h_idx] == needle[n_idx]:
                        if n_idx == 0: potential_start = h_idx # マッチ開始位置を記録
                        n_idx += 1
                    else: # ミスマッチ
                        if potential_start != -1:
                            h_idx = potential_start # マッチ途中だった場合、検索開始位置を戻す
                            potential_start = -1
                        n_idx = 0
                    h_idx += 1

                # ループ終了後、needleの残りが空白のみかチェック
                while n_idx < n_len and needle[n_idx].isspace(): n_idx += 1
                if n_idx == n_len: return potential_start # マッチ成功
                return -1 # 見つからなかった
            # --- ヘルパー関数ここまで ---

            # ソートされたブックマーク情報に基づき、各章のテキストを抽出
            for i, (title, start_page, level) in enumerate(sorted_all_bookmarks_info):
                # 次のブックマークの開始ページを、現在の章の終了ページとする
                # 最後のブックマークの場合は、PDFの最終ページまでを範囲とする
                next_title = "" # 次の章のタイトル（マーカー絞り込み用）
                if i + 1 < len(sorted_all_bookmarks_info):
                    end_page = sorted_all_bookmarks_info[i+1][1] # 次の章の開始ページ
                    next_title = sorted_all_bookmarks_info[i+1][0]
                else:
                    end_page = len(reader.pages) # PDFの総ページ数
                chapter_text = ""

                # 1. start_page の妥当性チェック
                if not (0 <= start_page < len(reader.pages)):
                    print(f"警告: ブックマーク '{title}' が指す開始ページ {start_page} はPDFの有効範囲外です。この章のテキストは空になります。")
                    # chapter_text は "" のまま
                else:
                    # ページ抽出の際の上限ページ番号（このページ番号自体はrangeに含まない）を設定します。
                    # 現在の章のテキストとして、次のブックマークの開始ページ(end_page)の内容まで含めて抽出します。
                    # end_page は次のブックマークの開始ページインデックス、または len(reader.pages) です。
                    # ループで page_num が start_page から end_page (またはPDF最終ページ) まで動くように、
                    # range の第二引数は page_extraction_upper_bound とします。
                    # page_extraction_upper_bound は end_page + 1 となりますが、len(reader.pages) を超えないようにします。
                    page_extraction_upper_bound = min(end_page + 1, len(reader.pages))

                    # 例:
                    # 1. 次のブックマークがページP (インデックス) にある場合:
                    #    end_page = P
                    #    page_extraction_upper_bound = min(P + 1, len(reader.pages))
                    #    range(start_page, P + 1) -> page_num は start_page ... P
                    # 2. これが最後のブックマークの場合:
                    #    end_page = len(reader.pages)
                    #    page_extraction_upper_bound = min(len(reader.pages) + 1, len(reader.pages)) = len(reader.pages)
                    #    range(start_page, len(reader.pages)) -> page_num は start_page ... len(reader.pages) - 1

                    if start_page < page_extraction_upper_bound: # 通常は true (start_page <= end_page のため)
                        for page_num in range(start_page, page_extraction_upper_bound):
                            # page_num は常に有効なインデックスになるはず
                            page = reader.pages[page_num]
                            page_text = page.extract_text()
                            if page_text:
                                chapter_text += page_text
                    # else の場合 (start_page >= page_extraction_upper_bound)、抽出するページなし。
                    # chapter_text は空のままなので問題ない。

                # --- マーカー（ブックマークタイトル）ベースのテキスト絞り込み処理 ---
                # 抽出したページテキスト(chapter_text)から、現在のブックマークタイトル(title)と
                # 次のブックマークタイトル(next_title)を使って、より正確な範囲を切り出す試み。
                # 注意: PDFのテキスト抽出精度やブックマークタイトルの完全一致に依存するため、
                #       必ずしも意図通りに絞り込めるとは限らない。
                refined_text = chapter_text # デフォルトはページ抽出テキスト
                if chapter_text: # ページ抽出テキストがある場合のみ絞り込み試行
                    try:
                        print(f"  章 '{title}': ページ抽出テキストに対しマーカー絞り込み試行...")
                        start_pos = _find_ignoring_whitespace(chapter_text, title) # 開始マーカー検索
                        if start_pos != -1:
                            start_index = start_pos + len(title) # 開始位置（マーカーの直後）※空白無視の影響あり
                            if next_title: # 次の章がある場合
                                end_pos = _find_ignoring_whitespace(chapter_text, next_title, start_index) # 終了マーカー検索
                                if end_pos != -1:
                                    refined_text = chapter_text[start_index:end_pos]
                                    print(f"    -> 開始/終了マーカーで絞り込み成功。")
                                else:
                                    refined_text = chapter_text[start_index:] # 終了マーカーが見つからない場合は最後まで
                                    print(f"    -> 開始マーカーで絞り込み成功（終了マーカーなし）。")
                            else: # 最終章
                                refined_text = chapter_text[start_index:]
                                print(f"    -> 開始マーカーで絞り込み成功（最終章）。")
                        else:
                            print(f"    -> 開始マーカー '{title}' がページ抽出テキスト内に見つからず。絞り込みスキップ。")
                            refined_text = chapter_text # 開始マーカーが見つからない場合は元のテキストを使用
                    except Exception as e_refine: # 念のため絞り込み中のエラーをキャッチ
                        print(f"    -> マーカー絞り込み中に予期せぬエラー: {e_refine}。絞り込みスキップ。")
                        refined_text = chapter_text # エラー時も元のテキストを使用
                chapters[title] = {"text": refined_text.strip(), "level": level} # 抽出/絞り込みしたテキストと階層レベルを辞書に格納 (前後の空白を除去)

    # --- エラーハンドリング ---
    except PyPDF2.errors.PdfReadError as e:
        print(f"エラー: PDFファイルの読み込みに失敗しました。ファイルが破損しているか、パスワードで保護されている可能性があります。 {e}")
        return None
    except Exception as e:
        print(f"エラー: PDFの処理中に予期せぬエラーが発生しました。{e}")
        return None

    print(f"ブックマークに基づいて {len(chapters)} 個の章（またはセクション）に分割完了。")
    return chapters # {章タイトル: {"text": 章テキスト, "level": 階層レベル}} の順序付き辞書を返す

def split_text(text, max_chunk_size):
    """長いテキストを指定された最大文字数以下のチャンクに分割します。

    Args:
        text: 分割対象のテキスト。
        max_chunk_size: 各チャンクの最大文字数。正の整数である必要があります。

    Returns:
        分割されたテキストチャンクのリスト。
        入力テキストが空の場合や max_chunk_size が不正な場合は空リスト。
    """
    if not text: # 入力テキストが空かNoneの場合のチェックを追加
        print("警告: 分割するテキストが空です。")
        return []
    if max_chunk_size <= 0:
        print("エラー: max_chunk_size は正の整数である必要があります。")
        return [] # またはエラーを発生させる

    print(f"テキストを最大 {max_chunk_size} 文字のチャンクに分割中...")
    chunks = [] # 分割後のチャンクを格納するリスト
    start = 0
    while start < len(text):
        end = min(start + max_chunk_size, len(text))
        chunks.append(text[start:end])
        start = end
    print(f"分割完了。チャンク数: {len(chunks)}")
    return chunks

def extract_text_between_markers(text: str, start_marker: str, end_marker: str) -> str:
    """文字列内から、指定された開始マーカーと終了マーカーの間のテキストを抽出します。

    @deprecated この関数は `split_text_by_bookmarks` 内にロジックが統合されたため、
                外部から呼び出す必要はありません。また、マーカーの完全一致に依存するため、
                ブックマークのタイトルがテキスト内に正確に含まれている保証がないため、意図しない動作をする可能性があります。

    Args:
        text: 処理対象の文字列。
        start_marker: 抽出範囲の開始を示す文字列（このマーカー自体は含まれない）。
        end_marker: 抽出範囲の終了を示す文字列（このマーカー自体は含まれない）。
                    空文字列の場合、start_marker以降の全てを抽出。
    Returns:
        抽出された部分文字列。マーカーが見つからない場合は空文字列を返す可能性あり。
    """
    # --- ヘルパー関数: haystack中の空白を無視してneedleを検索 ---
    def _find_ignoring_whitespace(haystack, needle, start_offset=0):
        """
        haystack内でneedleを検索します。haystackとneedleの両方の空白文字は無視されます。
        見つかった場合、haystackにおける開始インデックス（空白無視前の元のインデックス）を返します。見つからない場合は-1を返します。 (split_text_by_bookmarks内に移動済み)
        """
        if not needle:
            return start_offset # 空のneedleはオフセット位置で見つかったとみなす

        h_len = len(haystack)
        n_len = len(needle)
        h_idx = start_offset
        n_idx = 0
        potential_start = -1

        while h_idx < h_len:
            # haystackの空白をスキップ
            while h_idx < h_len and haystack[h_idx].isspace():
                if potential_start == -1: # マッチ開始前ならpotential_startも進める
                    start_offset += 1
                h_idx += 1
            if h_idx == h_len: break # スキップ中に終端に到達

            # needleの空白をスキップ
            while n_idx < n_len and needle[n_idx].isspace():
                n_idx += 1
            if n_idx == n_len: # needleの非空白文字を全てマッチした場合
                return potential_start # 発見

            # haystackとneedleの非空白文字を比較
            if h_idx < h_len and n_idx < n_len and haystack[h_idx] == needle[n_idx]:
                if n_idx == 0: potential_start = h_idx
                n_idx += 1
            else: # ミスマッチ
                if potential_start != -1: # マッチ途中だった場合、検索開始位置を戻す
                    h_idx = potential_start # 次の検索はpotential_startの次から
                    potential_start = -1
                # マッチ途中ではなかった場合、h_idxは次のループでインクリメントされる (下のh_idx += 1)
                n_idx = 0 # needleポインタをリセット
            h_idx += 1 # haystackポインタを進める
        # ループ終了後、needleの残りが空白のみかチェック
        while n_idx < n_len and needle[n_idx].isspace(): n_idx += 1
        if n_idx == n_len: return potential_start # マッチ成功
        return -1 # 見つからなかった

    # --- search_start_marker を使用して開始位置を検索 (空白無視) ---
    start_pos = _find_ignoring_whitespace(text, start_marker) # マーカーの前処理を削除したので、直接 start_marker を使用
    if start_pos == -1:
        # エラーを発生させる代わりに警告を出し、空文字列を返すように変更も検討可能
        raise ValueError(f"エラー: 開始マーカー '{start_marker}' がテキスト中に見つかりません（空白無視検索）。")

    # --- 開始インデックスは元の start_marker の長さを使って計算 ---
    # 注意: 空白無視検索で見つかった位置(start_pos)から元のマーカー長を加算するため、
    #       元のマーカーとテキストの間に予期せぬ空白が多い場合、意図した位置にならない可能性あり。
    #       より正確には、start_posから元のマーカーの非空白文字数分だけ進めるなどの調整が必要かもしれないが、
    #       ここでは元のロジックを踏襲し、見つかった位置 + 元のマーカー長とする。
    start_index = start_pos + len(start_marker)

    # --- end_marker が空の場合は、開始インデックス以降すべてを返す ---
    if not end_marker:
        return text[start_index:]

    # --- search_end_marker を使用して終了位置を検索 (start_index 以降、空白無視) ---
    end_pos = _find_ignoring_whitespace(text, end_marker, start_index) # マーカーの前処理を削除したので、直接 end_marker を使用

    # end_markerが見つからない場合は空、見つかればその手前までを返す
    return "" if end_pos == -1 else text[start_index:end_pos]