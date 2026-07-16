import os
import io
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

app = App(
    token=SLACK_BOT_TOKEN,
    logger=logger,
)

gemini_client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

# 二重処理防止(Slackのretry/イベント再送対策の簡易ガード)
_processed_file_ids = set()

PDF_MIMETYPES = {"application/pdf"}
PDF_FILETYPES = {"pdf"}
DOCX_MIMETYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}
DOCX_FILETYPES = {"docx"}

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
        + ("\n→ freeeへNDA契約締結申請を作成します(次ステップで実装)" if is_nda else "")
    )


def handle_contract_files(event: dict, say):
    files = event.get("files", [])
    target_files = [f for f in files if is_pdf_file(f) or is_docx_file(f)]

    if not target_files:
        return

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

        except requests.HTTPError as e:
            logger.exception(f"[FILE] download failed: {filename}")
            say(f":warning: ファイルのダウンロードに失敗しました: {filename} ({e})")
        except json.JSONDecodeError as e:
            logger.exception(f"[FILE] gemini response was not valid JSON: {filename}")
            say(f":warning: Geminiの解析結果を読み取れませんでした: {filename} ({e})")
        except Exception as e:
            logger.exception(f"[FILE] processing failed: {filename}")
            say(f":warning: ファイルの処理に失敗しました: {filename} ({e})")


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
        handle_contract_files(event, say)
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
