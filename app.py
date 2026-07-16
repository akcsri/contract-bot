import os
import io
import re
import time
import json
import logging

import requests
import fitz  # PyMuPDF
from docx import Document

from google import genai
from google.genai import types

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# --- ログ設定 -------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

# このチャンネルの投稿のみ処理する(テスト/本番の切り分け)
CONTRACT_CHANNEL_ID = os.environ["CONTRACT_CHANNEL_ID"]

# --- freee 設定 -------------------------------------------------------
FREEE_CLIENT_ID = os.environ["FREEE_CLIENT_ID"]
FREEE_CLIENT_SECRET = os.environ["FREEE_CLIENT_SECRET"]
FREEE_REFRESH_TOKEN = os.environ["FREEE_REFRESH_TOKEN"]
FREEE_COMPANY_ID = int(os.environ["FREEE_COMPANY_ID"])
FREEE_NDA_FORM_ID = int(os.environ.get("FREEE_NDA_FORM_ID", "87137"))

FREEE_API_BASE = "https://api.freee.co.jp"
FREEE_TOKEN_URL = "https://accounts.secure.freee.co.jp/public_api/token"

# 契約締結方法の選択肢(freee上のフォームの選択肢と一致させること。
# 増やす/変える場合はここを編集してください)
CONTRACT_METHODS = ["電子署名（CSRI発信）", "原本捺印"]

# freeeへの送信を承認したとみなすリアクション
APPROVE_REACTIONS = {"+1", "thumbsup", "white_check_mark", "heavy_check_mark"}

app = App(
    token=SLACK_BOT_TOKEN,
    logger=logger,
)

gemini_client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

# 二重処理防止(Slackのretry/イベント再送対策の簡易ガード)
_processed_file_ids = set()

# NDA判定〜freee申請までの一時状態(thread_ts をキーに保持)
# 注意: プロセス内メモリのみ。Renderの再起動で消える簡易実装。
_pending_nda = {}

PDF_MIMETYPES = {"application/pdf"}
PDF_FILETYPES = {"pdf"}
DOCX_MIMETYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}
DOCX_FILETYPES = {"docx"}

FIELD_REPLY_PATTERN = re.compile(
    r"部門[:：]\s*(?P<section>\S+)\s*/\s*締結方法[:：]\s*(?P<method>.+)"
)

NDA_ANALYSIS_PROMPT = """\
あなたは契約書を分析する専門家です。提供された契約書(PDFの場合は
スキャン画像のみでテキスト層が無いこともあるので、画像として内容を
読み取って判断してください)を確認し、以下の項目をJSON形式のみで
回答してください。説明文やコードブロックのマークダウンは付けず、
JSONオブジェクトのみを返すこと。

{
  "is_nda": true または false (秘密保持契約/NDAであればtrue),
  "document_type": "書類の種類(例: 秘密保持契約書, 業務委託契約書 等)",
  "parties": ["契約当事者1", "契約当事者2"],
  "contract_date": "YYYY-MM-DD形式の契約日。読み取れない場合はnull",
  "reason": "is_ndaと判定した理由の要約(1〜2文)"
}
"""


# =====================================================================
# ファイル判定・ダウンロード
# =====================================================================

def is_pdf_file(f: dict) -> bool:
    return f.get("mimetype") in PDF_MIMETYPES or f.get("filetype") in PDF_FILETYPES


def is_docx_file(f: dict) -> bool:
    return f.get("mimetype") in DOCX_MIMETYPES or f.get("filetype") in DOCX_FILETYPES


def download_slack_file(file_info: dict) -> bytes:
    """Slackのfiles:read権限でファイル本体(バイナリ)を取得する"""
    url = file_info["url_private"]
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.content


def get_pdf_page_count(pdf_bytes: bytes) -> int:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return doc.page_count
    finally:
        doc.close()


def extract_docx_text(docx_bytes: bytes) -> str:
    """python-docxで段落・表のテキストを抽出する"""
    doc = Document(io.BytesIO(docx_bytes))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return "\n".join(parts)


# =====================================================================
# Gemini解析
# =====================================================================

def analyze_contract_with_gemini(*, pdf_bytes: bytes = None, text: str = None) -> dict:
    """契約書(PDFバイナリ、またはWordから抽出済みのテキスト)をGeminiに渡し、
    NDA判定を含む解析結果を得る。
    PDFはスキャン画像(テキスト層無し)でもGeminiが画像として読み取れるため
    OCR処理を別途行う必要がない。Wordは事前にテキスト抽出したものを渡す。
    """
    if pdf_bytes is not None:
        contents = [
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            NDA_ANALYSIS_PROMPT,
        ]
    elif text is not None:
        contents = [
            NDA_ANALYSIS_PROMPT,
            f"\n--- 契約書本文(Wordファイルから抽出) ---\n{text}",
        ]
    else:
        raise ValueError("pdf_bytes または text のいずれかが必要です")

    response = gemini_client.models.generate_content(
        model=GEMINI_MODEL,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        ),
    )
    return json.loads(response.text)


def format_analysis_message(filename: str, meta_line: str, result: dict) -> str:
    is_nda = result.get("is_nda")
    judgement = "✅ NDA(秘密保持契約)と判定" if is_nda else "❌ NDAではないと判定"
    parties = "、".join(result.get("parties") or []) or "不明"
    return (
        f"ファイルを受信しました: *{filename}* ({meta_line})\n"
        f"{judgement}\n"
        f"書類種別: {result.get('document_type', '不明')}\n"
        f"契約当事者: {parties}\n"
        f"契約日: {result.get('contract_date') or '不明'}\n"
        f"判定理由: {result.get('reason', '-')}"
    )


# =====================================================================
# freee 連携
# =====================================================================

_freee_token_cache = {
    "access_token": None,
    "refresh_token": FREEE_REFRESH_TOKEN,
    "expires_at": 0,
}
_sections_cache = None


def get_freee_access_token() -> str:
    """freeeのアクセストークンを取得する(必要な時だけrefresh)。

    重要な注意点: freeeのrefresh_tokenは使用するたびに新しい値に
    ローテーション(再発行)される。このプロセスはローテーション後の
    refresh_tokenをメモリ内にしか保持していないため、Renderのプロセスが
    再起動すると環境変数FREEE_REFRESH_TOKENが指す値がすでに無効になっている
    可能性がある。本番運用では、ローテーション後の値を外部ストレージ
    (DB、Render環境変数のAPI経由更新、永続ディスク上のファイル等)に
    保存する仕組みの追加を強く推奨する。
    """
    now = time.time()
    if _freee_token_cache["access_token"] and now < _freee_token_cache["expires_at"] - 60:
        return _freee_token_cache["access_token"]

    resp = requests.post(
        FREEE_TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": FREEE_CLIENT_ID,
            "client_secret": FREEE_CLIENT_SECRET,
            "refresh_token": _freee_token_cache["refresh_token"],
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    _freee_token_cache["access_token"] = data["access_token"]
    _freee_token_cache["refresh_token"] = data["refresh_token"]
    _freee_token_cache["expires_at"] = now + data["expires_in"]
    logger.info("[freee] access token refreshed")
    return _freee_token_cache["access_token"]


def freee_headers() -> dict:
    return {"Authorization": f"Bearer {get_freee_access_token()}"}


def fetch_freee_sections() -> list:
    global _sections_cache
    if _sections_cache is not None:
        return _sections_cache
    resp = requests.get(
        f"{FREEE_API_BASE}/api/1/sections",
        headers=freee_headers(),
        params={"company_id": FREEE_COMPANY_ID},
        timeout=30,
    )
    resp.raise_for_status()
    _sections_cache = resp.json().get("sections", [])
    return _sections_cache


def find_section_id_by_name(name: str):
    for s in fetch_freee_sections():
        if s.get("name") == name:
            return s.get("id")
    return None


def upload_file_to_freee(file_bytes: bytes, filename: str) -> int:
    """freeeのファイルボックス(証憑)にアップロードし、receipt idを返す"""
    resp = requests.post(
        f"{FREEE_API_BASE}/api/1/receipts",
        headers=freee_headers(),
        data={"company_id": FREEE_COMPANY_ID, "description": filename},
        files={"receipt": (filename, file_bytes)},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["receipt"]["id"]


def create_nda_approval_request(
    *, title: str, counterparty: str, contract_date: str,
    receipt_id: int, section_id: int, method: str,
) -> dict:
    """NDA契約締結申請(freeeの汎用申請フォーム)を作成する。

    request_itemsの並び・typeは、既存の承認済みNDA申請(form_id=87137)を
    参考に組み立てている。freee側でフォーム定義(項目の追加/削除/並び替え)
    が変更された場合はここも合わせて調整すること。
    """
    body = {
        "company_id": FREEE_COMPANY_ID,
        "form_id": FREEE_NDA_FORM_ID,
        "title": title,
        "request_items": [
            {"type": "title", "value": title},
            {"type": "section", "value": str(section_id)},
            {"type": "single_line", "value": counterparty},
            {"type": "partner", "value": ""},
            {"type": "date", "value": contract_date},
            {"type": "receipt", "value": str(receipt_id)},
            {"type": "select", "value": method},
            {"type": "multi_line", "value": ""},
            {"type": "multi_line", "value": ""},
            {"type": "multi_line", "value": ""},
        ],
    }
    resp = requests.post(
        f"{FREEE_API_BASE}/api/1/approval_requests",
        headers=freee_headers(),
        json=body,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["approval_request"]


# =====================================================================
# Slack: 契約書ファイル受信 → Gemini解析 → NDA判定
# =====================================================================

def handle_contract_files(event: dict, say):
    files = event.get("files", [])
    target_files = [f for f in files if is_pdf_file(f) or is_docx_file(f)]

    if not target_files:
        return

    message_ts = event.get("ts")

    for f in target_files:
        file_id = f.get("id")
        filename = f.get("name", "unknown")

        if file_id in _processed_file_ids:
            logger.info(f"[SKIP] already processed: {filename} ({file_id})")
            continue
        _processed_file_ids.add(file_id)

        try:
            logger.info(f"[FILE] downloading: {filename} ({file_id})")
            raw_bytes = download_slack_file(f)

            if is_pdf_file(f):
                page_count = get_pdf_page_count(raw_bytes)
                logger.info(f"[PDF] sending to Gemini({GEMINI_MODEL}): {filename}")
                result = analyze_contract_with_gemini(pdf_bytes=raw_bytes)
                meta_line = f"ページ数: {page_count}"

            else:  # docx
                text = extract_docx_text(raw_bytes)
                if not text.strip():
                    logger.warning(f"[DOCX] no extractable text: {filename}")
                    say(f":warning: Wordファイルからテキストを抽出できませんでした: {filename}")
                    continue
                logger.info(f"[DOCX] sending to Gemini({GEMINI_MODEL}): {filename}")
                result = analyze_contract_with_gemini(text=text)
                meta_line = f"文字数: {len(text)}"

            logger.info(f"[FILE] gemini result: {result}")
            say(format_analysis_message(filename, meta_line, result))

            if result.get("is_nda"):
                thread_ts = message_ts
                _pending_nda[thread_ts] = {
                    "filename": filename,
                    "raw_bytes": raw_bytes,
                    "gemini_result": result,
                    "stage": "awaiting_fields",
                }
                say(
                    text=(
                        "freee申請作成のため、このスレッドで下記の形式で返信してください。\n"
                        f"`部門: <部門名> / 締結方法: <{' か '.join(CONTRACT_METHODS)}>`\n"
                        f"例: `部門: 事業開発部 / 締結方法: {CONTRACT_METHODS[0]}`"
                    ),
                    thread_ts=thread_ts,
                )

        except requests.HTTPError as e:
            logger.exception(f"[FILE] download failed: {filename}")
            say(f":warning: ファイルのダウンロードに失敗しました: {filename} ({e})")
        except json.JSONDecodeError as e:
            logger.exception(f"[FILE] gemini response was not valid JSON: {filename}")
            say(f":warning: Geminiの解析結果を読み取れませんでした: {filename} ({e})")
        except Exception as e:
            logger.exception(f"[FILE] processing failed: {filename}")
            say(f":warning: ファイルの処理に失敗しました: {filename} ({e})")


def handle_nda_field_reply(event: dict, say) -> bool:
    """NDA申請の「部門 / 締結方法」返信をスレッド内で処理する。
    処理した場合True、対象外ならFalseを返す。
    """
    thread_ts = event.get("thread_ts")
    pending = _pending_nda.get(thread_ts)
    if not pending or pending.get("stage") != "awaiting_fields":
        return False

    text = event.get("text", "")
    m = FIELD_REPLY_PATTERN.search(text)
    if not m:
        say(
            "形式が読み取れませんでした。次の形式で返信してください。\n"
            f"`部門: <部門名> / 締結方法: <{' か '.join(CONTRACT_METHODS)}>`",
            thread_ts=thread_ts,
        )
        return True

    section_name = m.group("section").strip()
    method = m.group("method").strip()

    if method not in CONTRACT_METHODS:
        say(
            f":warning: 締結方法は次のいずれかで指定してください: {', '.join(CONTRACT_METHODS)}",
            thread_ts=thread_ts,
        )
        return True

    try:
        section_id = find_section_id_by_name(section_name)
    except Exception as e:
        logger.exception("[freee] fetch sections failed")
        say(f":warning: freeeの部門一覧取得に失敗しました: {e}", thread_ts=thread_ts)
        return True

    if section_id is None:
        say(
            f":warning: 部門「{section_name}」がfreee上に見つかりません。"
            "freeeに登録されている部門名と完全に一致させて返信してください。",
            thread_ts=thread_ts,
        )
        return True

    pending["section_id"] = section_id
    pending["section_name"] = section_name
    pending["method"] = method
    pending["stage"] = "awaiting_confirm"

    result = pending["gemini_result"]
    confirm_text = (
        "以下の内容でfreeeにNDA契約締結申請を作成します。\n"
        "よろしければこのメッセージに :+1: で反応してください。\n"
        f"タイトル: {pending['filename']}\n"
        f"契約当事者: {'、'.join(result.get('parties') or []) or '不明'}\n"
        f"契約日: {result.get('contract_date') or '不明'}\n"
        f"部門: {section_name}\n"
        f"締結方法: {method}"
    )
    posted = say(confirm_text, thread_ts=thread_ts)
    pending["confirm_message_ts"] = posted["ts"]
    return True


@app.event("message")
def handle_message(event, say, logger):
    logger.info("========== EVENT RECEIVED ==========")
    logger.debug(json.dumps(event, ensure_ascii=False, indent=2))

    subtype = event.get("subtype")

    # bot自身の投稿は無視(無限ループ防止)
    if event.get("bot_id"):
        return

    # 指定チャンネル以外は無視(テスト/本番の切り分け)
    if event.get("channel") != CONTRACT_CHANNEL_ID:
        return

    # 通常のテキストメッセージ or ファイル添付(file_share)のみ処理
    # それ以外(message_changed / message_deleted 等)は無視
    if subtype not in (None, "file_share"):
        return

    if event.get("files"):
        handle_contract_files(event, say)
        return

    if event.get("thread_ts") and handle_nda_field_reply(event, say):
        return

    # ファイル無しの通常メッセージ(動作確認用)
    say("イベント受信成功")


@app.event("reaction_added")
def handle_reaction_added(event, say, logger):
    """確認メッセージへのリアクションでfreee申請を実行する"""
    if event.get("reaction") not in APPROVE_REACTIONS:
        return

    item = event.get("item", {})
    message_ts = item.get("ts")

    # 指定チャンネル以外は無視(テスト/本番の切り分け)
    if item.get("channel") != CONTRACT_CHANNEL_ID:
        return

    target_thread_ts = None
    for thread_ts, pending in _pending_nda.items():
        if pending.get("stage") == "awaiting_confirm" and pending.get("confirm_message_ts") == message_ts:
            target_thread_ts = thread_ts
            break

    if target_thread_ts is None:
        return

    pending = _pending_nda.pop(target_thread_ts)

    try:
        logger.info(f"[freee] uploading file: {pending['filename']}")
        receipt_id = upload_file_to_freee(pending["raw_bytes"], pending["filename"])

        result = pending["gemini_result"]
        logger.info(f"[freee] creating approval request: {pending['filename']}")
        approval = create_nda_approval_request(
            title=pending["filename"],
            counterparty="、".join(result.get("parties") or []) or "不明",
            contract_date=result.get("contract_date") or "",
            receipt_id=receipt_id,
            section_id=pending["section_id"],
            method=pending["method"],
        )
        say(
            f":white_check_mark: freeeへNDA契約締結申請を作成しました"
            f"(申請番号: {approval.get('application_number')})",
            thread_ts=target_thread_ts,
        )
    except Exception as e:
        logger.exception("[freee] approval request creation failed")
        say(f":warning: freee申請の作成に失敗しました: {e}", thread_ts=target_thread_ts)


if __name__ == "__main__":
    print("BOT START")

    handler = SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"],
    )
    handler.start()
