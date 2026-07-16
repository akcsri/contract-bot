import os
import json
import logging
 
import requests
import fitz  # PyMuPDF
 
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
 
# --- デバッグログを有効化 -------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
 
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
 
app = App(
    token=SLACK_BOT_TOKEN,
    logger=logger,
)
 
# 二重処理防止(Slackのretry/イベント再送対策の簡易ガード)
_processed_file_ids = set()
 
 
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
 
 
def extract_pdf_text(pdf_bytes: bytes) -> str:
    """PyMuPDFでPDFからテキストを抽出する"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        pages_text = [page.get_text() for page in doc]
    finally:
        doc.close()
    return "\n".join(pages_text)
 
 
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
 
            logger.info(f"[PDF] extracting text: {filename}")
            text = extract_pdf_text(pdf_bytes)
 
            char_count = len(text)
            page_count = fitz.open(stream=pdf_bytes, filetype="pdf").page_count
 
            logger.info(
                f"[PDF] done: {filename} pages={page_count} chars={char_count}"
            )
            logger.debug(f"[PDF] extracted text preview:\n{text[:500]}")
 
            # TODO: 次のステップでこの text を Gemini解析へ渡す
            say(
                f"PDFを受信しました: *{filename}*\n"
                f"ページ数: {page_count} / 文字数: {char_count}\n"
                f"→ Gemini解析へ進みます(次ステップで実装)"
            )
 
        except requests.HTTPError as e:
            logger.exception(f"[PDF] download failed: {filename}")
            say(f":warning: PDFのダウンロードに失敗しました: {filename} ({e})")
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
 
