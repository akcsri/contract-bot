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
# 「NDA契約締結申請」フォームの「申請経路の選択」に対応する必須項目。
# NDA締結申請の経路(プロジェクトオーナー→リーガル(金子)→コーポレート(吉田))のID。
FREEE_APPROVAL_FLOW_ROUTE_ID = int(os.environ.get("FREEE_APPROVAL_FLOW_ROUTE_ID", "1431338"))
# 認可コード取得時に指定したコールバックURLと同じ値を指定する
# (ブラウザで手動取得した場合は "urn:ietf:wg:oauth:2.0:oob")
FREEE_REDIRECT_URI = os.environ.get("FREEE_REDIRECT_URI", "urn:ietf:wg:oauth:2.0:oob")

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

# 不足項目を聞き返すときに使う個別パターン(どちらか一方だけの返信でも拾える)
PROJECT_REPLY_PATTERN = re.compile(r"プロジェクト名?[:：]\s*(\S+)")
METHOD_REPLY_PATTERN = re.compile(r"締結方法[:：]\s*(.+)")


def build_nda_prompt(project_candidates: list) -> str:
    candidates_text = "、".join(project_candidates) if project_candidates else "(候補リスト取得不可)"
    return f"""\
あなたは契約書の申請処理を行う専門アシスタントです。
提供された契約書本体(PDFの場合はスキャン画像のみでテキスト層が無い
こともあるので、画像として内容を読み取って判断してください)に加えて、
Slack投稿の本文とファイル名も参考情報として渡されます。契約書の内容
だけでなく、投稿本文・ファイル名に書かれている情報もすべて確認し、
そこに既に書かれている情報は憶測せずそのまま採用してください。
どこにも情報が無い項目だけをnullにしてください。

以下の項目をJSON形式のみで回答してください。説明文やコードブロックの
マークダウンは付けず、JSONオブジェクトのみを返すこと。

{{
  "is_nda": true または false (秘密保持契約/NDAであればtrue),
  "document_type": "書類の種類(例: 秘密保持契約書, 業務委託契約書 等)",
  "parties": ["契約当事者1", "契約当事者2"],
  "contract_date": "YYYY-MM-DD形式の契約日。読み取れない場合はnull",
  "reason": "is_ndaと判定した理由の要約(1〜2文)",
  "project_name": "次の候補の中から一致するものを1つだけ選んで文字列で返す: [{candidates_text}]。
    ファイル名や投稿本文にプロジェクト名/顧客名らしき記載があれば最優先で使い、
    候補の中から最も一致するものを選ぶこと。プロジェクトに紐づかない場合や
    候補に一致するものが無い場合はnull",
  "method": "契約の締結方法。投稿本文に「捺印」「押印」とあれば\\"原本捺印\\"、
    「電子署名」とあれば\\"電子署名（CSRI発信）\\"。判断できなければnull",
  "physical_mail_address": "締結方法が原本捺印と判断できる場合、投稿本文に
    書かれている原本の送付先(郵便番号・住所・会社名・部署名・担当者名・
    電話番号など)をそのままの文字列でまとめて抽出する。記載が無い/
    該当しない場合はnull"
}}
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

def analyze_contract_with_gemini(
    *, pdf_bytes: bytes = None, text: str = None,
    message_text: str = "", filename: str = "",
    project_candidates: list = None,
) -> dict:
    """契約書(PDFバイナリ、またはWordから抽出済みのテキスト)に加えて、
    Slack投稿本文・ファイル名・freeeのプロジェクト候補一覧もあわせてGeminiに渡し、
    NDA判定と、freee申請に必要な項目(プロジェクト名/締結方法/原本送付先)の
    自動抽出を1回のリクエストでまとめて行う。
    PDFはスキャン画像(テキスト層無し)でもGeminiが画像として読み取れるため
    OCR処理を別途行う必要がない。Wordは事前にテキスト抽出したものを渡す。
    """
    prompt = build_nda_prompt(project_candidates or [])
    extra_context = (
        f"--- Slack投稿本文 ---\n{message_text or '(本文なし)'}\n\n"
        f"--- 添付ファイル名 ---\n{filename}\n"
    )

    if pdf_bytes is not None:
        contents = [
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt,
            extra_context,
        ]
    elif text is not None:
        contents = [
            prompt,
            extra_context,
            f"--- 契約書本文(Wordファイルから抽出) ---\n{text}",
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

    logger.info(
        f"[freee] refreshing access token (redirect_uri={FREEE_REDIRECT_URI}, "
        f"refresh_token末尾={_freee_token_cache['refresh_token'][-6:]})"
    )
    resp = requests.post(
        FREEE_TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": FREEE_CLIENT_ID,
            "client_secret": FREEE_CLIENT_SECRET,
            "refresh_token": _freee_token_cache["refresh_token"],
            "redirect_uri": FREEE_REDIRECT_URI,
        },
        timeout=30,
    )
    if not resp.ok:
        logger.error(f"[freee] token refresh failed: {resp.status_code} {resp.text}")
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
    """freeeの「部門」マスタを取得する。このフォーム上では部門IDが
    「プロジェクト名」として運用されている(プロジェクト名: CSRI = 不紐づけ用)。
    """
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
    """プロジェクト名(freee上は部門名)からIDを引く"""
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
    receipt_id: int, section_id: int, method: str, mail_address: str = "",
) -> dict:
    """NDA契約締結申請(freeeの汎用申請フォーム)を作成する。

    request_itemsの並び・typeは、既存の承認済みNDA申請(form_id=87137)を
    参考に組み立てている。freee側でフォーム定義(項目の追加/削除/並び替え)
    が変更された場合はここも合わせて調整すること。
    最初のmulti_lineは「原本送付先」欄(押印方法が原本捺印の場合のみ使用)。
    """
    body = {
        "company_id": FREEE_COMPANY_ID,
        "form_id": FREEE_NDA_FORM_ID,
        "approval_flow_route_id": FREEE_APPROVAL_FLOW_ROUTE_ID,
        "title": title,
        "request_items": [
            {"type": "title", "value": title},
            {"type": "section", "value": str(section_id)},
            {"type": "single_line", "value": counterparty},
            {"type": "partner", "value": ""},
            {"type": "date", "value": contract_date},
            {"type": "receipt", "value": str(receipt_id)},
            {"type": "select", "value": method},
            {"type": "multi_line", "value": mail_address},
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
# Slack: 契約書ファイル受信 → Gemini解析 → NDA判定 → freee申請準備
# =====================================================================

def post_confirmation(pending: dict, thread_ts: str, say):
    """freee申請の最終確認メッセージを投稿する"""
    result = pending["gemini_result"]
    lines = [
        "以下の内容でfreeeにNDA契約締結申請を作成します。",
        "よろしければこのメッセージに :+1: で反応してください。",
        f"タイトル: {pending['filename']}",
        f"契約当事者: {'、'.join(result.get('parties') or []) or '不明'}",
        f"契約日: {result.get('contract_date') or '不明'}",
        f"プロジェクト名: {pending['section_name']}",
        f"締結方法: {pending['method']}",
    ]
    if pending.get("mail_address"):
        lines.append(f"原本送付先: {pending['mail_address']}")
    posted = say("\n".join(lines), thread_ts=thread_ts)
    pending["confirm_message_ts"] = posted["ts"]
    pending["stage"] = "awaiting_confirm"


def prompt_for_missing_fields(missing_fields: list, thread_ts: str, say):
    prompts = []
    if "project_name" in missing_fields:
        prompts.append("プロジェクト名: <プロジェクト名>")
    if "method" in missing_fields:
        prompts.append(f"締結方法: <{' か '.join(CONTRACT_METHODS)}>")
    note = (
        "\n(プロジェクトに紐づかない契約書は「CSRI」を指定してください)"
        if "project_name" in missing_fields else ""
    )
    say(
        "投稿内容・ファイル名・契約書の中身から自動判定を試みましたが、"
        "次の項目が確認できませんでした。このスレッドで返信してください。\n"
        f"`{' / '.join(prompts)}`" + note,
        thread_ts=thread_ts,
    )


def handle_contract_files(event: dict, say):
    files = event.get("files", [])
    target_files = [f for f in files if is_pdf_file(f) or is_docx_file(f)]

    if not target_files:
        return

    message_ts = event.get("ts")
    message_text = event.get("text", "")

    # プロジェクト名の自動判定に使う候補一覧(取得できなくても処理は続行する)
    try:
        project_candidates = [s.get("name") for s in fetch_freee_sections() if s.get("name")]
    except Exception:
        logger.exception("[freee] fetch sections failed (project候補取得)")
        project_candidates = []

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
                result = analyze_contract_with_gemini(
                    pdf_bytes=raw_bytes,
                    message_text=message_text,
                    filename=filename,
                    project_candidates=project_candidates,
                )
                meta_line = f"ページ数: {page_count}"

            else:  # docx
                text = extract_docx_text(raw_bytes)
                if not text.strip():
                    logger.warning(f"[DOCX] no extractable text: {filename}")
                    say(f":warning: Wordファイルからテキストを抽出できませんでした: {filename}")
                    continue
                logger.info(f"[DOCX] sending to Gemini({GEMINI_MODEL}): {filename}")
                result = analyze_contract_with_gemini(
                    text=text,
                    message_text=message_text,
                    filename=filename,
                    project_candidates=project_candidates,
                )
                meta_line = f"文字数: {len(text)}"

            logger.info(f"[FILE] gemini result: {result}")
            say(format_analysis_message(filename, meta_line, result))

            if not result.get("is_nda"):
                continue

            thread_ts = message_ts

            project_name = (result.get("project_name") or "").strip()
            method = (result.get("method") or "").strip()
            mail_address = (result.get("physical_mail_address") or "").strip()

            section_id = None
            if project_name:
                try:
                    section_id = find_section_id_by_name(project_name)
                except Exception:
                    logger.exception("[freee] fetch sections failed")
                    section_id = None

            missing_fields = []
            if section_id is None:
                missing_fields.append("project_name")
                project_name = ""
            if method not in CONTRACT_METHODS:
                missing_fields.append("method")
                method = ""

            pending = {
                "filename": filename,
                "raw_bytes": raw_bytes,
                "gemini_result": result,
                "section_id": section_id,
                "section_name": project_name,
                "method": method,
                "mail_address": mail_address if method == "原本捺印" else "",
                "missing_fields": missing_fields,
                "stage": "awaiting_fields" if missing_fields else "awaiting_confirm",
            }
            _pending_nda[thread_ts] = pending

            if missing_fields:
                prompt_for_missing_fields(missing_fields, thread_ts, say)
            else:
                post_confirmation(pending, thread_ts, say)

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
    """NDA申請で不足している項目(プロジェクト名/締結方法)の返信を処理する。
    片方だけの返信でも受け付け、揃うまで不足分だけを聞き返す。
    処理した場合True、対象外ならFalseを返す。
    """
    thread_ts = event.get("thread_ts")
    pending = _pending_nda.get(thread_ts)
    if not pending or pending.get("stage") != "awaiting_fields":
        return False

    text = event.get("text", "")
    missing = list(pending.get("missing_fields", []))

    if "project_name" in missing:
        m = PROJECT_REPLY_PATTERN.search(text)
        if m:
            section_name = m.group(1).strip()
            try:
                section_id = find_section_id_by_name(section_name)
            except Exception as e:
                logger.exception("[freee] fetch sections failed")
                say(f":warning: freeeのプロジェクト一覧取得に失敗しました: {e}", thread_ts=thread_ts)
                return True

            if section_id is None:
                say(
                    f":warning: プロジェクト「{section_name}」がfreee上に見つかりません。"
                    "freeeに登録されているプロジェクト名(部門名)と完全に一致させて返信してください"
                    "(紐づかない場合は「CSRI」)。",
                    thread_ts=thread_ts,
                )
                return True

            pending["section_id"] = section_id
            pending["section_name"] = section_name
            missing.remove("project_name")

    if "method" in missing:
        m = METHOD_REPLY_PATTERN.search(text)
        if m:
            method = m.group(1).strip()
            if method not in CONTRACT_METHODS:
                say(
                    f":warning: 締結方法は次のいずれかで指定してください: {', '.join(CONTRACT_METHODS)}",
                    thread_ts=thread_ts,
                )
                return True
            pending["method"] = method
            missing.remove("method")

    pending["missing_fields"] = missing

    if missing:
        prompt_for_missing_fields(missing, thread_ts, say)
        return True

    post_confirmation(pending, thread_ts, say)
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
        contract_title = os.path.splitext(pending["filename"])[0]
        approval = create_nda_approval_request(
            title=contract_title,
            counterparty="、".join(result.get("parties") or []) or "不明",
            contract_date=result.get("contract_date") or "",
            receipt_id=receipt_id,
            section_id=pending["section_id"],
            method=pending["method"],
            mail_address=pending.get("mail_address", ""),
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
