import os
import json
import logging

import requests
import fitz  # PyMuPDF

from google import genai
from google.genai import types

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# --- ログ設定 -------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

app = App(
    token=SLACK_BOT_TOKEN,
    logger=logger,
)

gemini_client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

# 二重処理防止(Slackのretry/イベント再送対策の簡易ガード)
_processed_file_ids = set()

NDA_ANALYSIS_PROMPT = """\
あなたは契約書を分析する専門家です。添付のPDF(スキャン画像のみで
テキスト層が無い場合も、画像として内容を読み取って判断してください)
を確認し、以下の項目をJSON形式のみで回答してください。説明文や
コードブロックのマークダウンは付けず、JSONオブジェクトのみを返すこと。

{
  "is_nda": true または false (秘密保持契約/NDAであればtrue),
  "document_type": "書類の種類(例: 秘密保持契約書, 業務委託契約書 等)",
  "parties": ["契約当事者1", "契約当事者2"],
  "contract_date": "YYYY-MM-DD形式の契約日。読み取れない場合はnull",
  "reason": "is_ndaと判定した理由の要約(1〜2文)"
}
"""


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


def analyze_contract_with_gemini(pdf_bytes: bytes) -> dict:
    """PDFバイナリをそのままGeminiに渡し、NDA判定を含む解析結果を得る。
    テキスト層の無いスキャンPDF(捺印済み契約書等)でもGeminiが
    画像として内容を読み取れるため、OCR処理を別途行う必要がない。
    """
    response = gemini_client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            NDA_ANALYSIS_PROMPT,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        ),
    )
    return json.loads(response.text)


def format_analysis_message(filename: str, page_count: int, result: dict) -> str:
    is_nda = result.get("is_nda")
    judgement = "✅ NDA(秘密保持契約)と判定" if is_nda else "❌ NDAではないと判定"
    parties = "、".join(result.get("parties") or []) or "不明"
    return (
        f"PDFを受信しました: *{filename}* (ページ数: {page_count})\n"
        f"{judgement}\n"
        f"書類種別: {result.get('document_type', '不明')}\n"
        f"契約当事者: {parties}\n"
        f"契約日: {result.get('contract_date') or '不明'}\n"
        f"判定理由: {result.get('reason', '-')}"
        + ("\n→ freeeへNDA契約締結申請を作成します(次ステップで実装)" if is_nda else "")
    )


def handle_pdf_files(event: dict, say):
    files = event.get("files", [])
    pdf_files = [
        f for f in files
        if f.get("mimetype") == "application/pdf"
        or f.get("filetype") == "pdf"
    ]

    if not pdf_files:
        return

    for f in pdf_files:
        file_id = f.get("id")
        filename = f.get("name", "unknown.pdf")

        if file_id in _processed_file_ids:
            logger.info(f"[SKIP] already processed: {filename} ({file_id})")
            continue
        _processed_file_ids.add(file_id)

        try:
            logger.info(f"[PDF] downloading: {filename} ({file_id})")
            pdf_bytes = download_slack_file(f)
            page_count = get_pdf_page_count(pdf_bytes)

            logger.info(f"[PDF] sending to Gemini({GEMINI_MODEL}): {filename}")
            result = analyze_contract_with_gemini(pdf_bytes)
            logger.info(f"[PDF] gemini result: {result}")

            say(format_analysis_message(filename, page_count, result))

        except requests.HTTPError as e:
            logger.exception(f"[PDF] download failed: {filename}")
            say(f":warning: PDFのダウンロードに失敗しました: {filename} ({e})")
        except json.JSONDecodeError as e:
            logger.exception(f"[PDF] gemini response was not valid JSON: {filename}")
            say(f":warning: Geminiの解析結果を読み取れませんでした: {filename} ({e})")
        except Exception as e:
            logger.exception(f"[PDF] processing failed: {filename}")
            say(f":warning: PDFの処理に失敗しました: {filename} ({e})")


@app.event("message")
def handle_message(event, say, logger):
    logger.info("========== EVENT RECEIVED ==========")
    logger.debug(json.dumps(event, ensure_ascii=False, indent=2))

    subtype = event.get("subtype")

    # bot自身の投稿は無視(無限ループ防止)
    if event.get("bot_id"):
        return

    # 通常のテキストメッセージ or ファイル添付(file_share)のみ処理
    # それ以外(message_changed / message_deleted 等)は無視
    if subtype not in (None, "file_share"):
        return

    if event.get("files"):
        handle_pdf_files(event, say)
        return

    # ファイル無しの通常メッセージ(動作確認用)
    say("イベント受信成功")


if __name__ == "__main__":
    print("BOT START")

    handler = SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"],
    )
    handler.start()
