import os

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_APP_TOKEN = os.environ["SLACK_APP_TOKEN"]

CONTRACT_CHANNEL_ID = os.environ["CONTRACT_CHANNEL_ID"]

app = App(
    token=SLACK_BOT_TOKEN
)


@app.event("message")
def handle_message(event, say):

    # Bot自身の投稿は無視
    if event.get("subtype"):
        return

    # テストチャンネル以外は無視
    if event.get("channel") != CONTRACT_CHANNEL_ID:
        return

    user = event.get("user", "")

    text = event.get("text", "")

    files = event.get("files", [])

    # PDF添付なし
    if not files:

        say(
            text=f"""
受信しました。

ユーザー:
{user}

本文:
{text}

※ PDFファイルが添付されていません。
"""
        )

        return

    # 添付ファイル情報
    first_file = files[0]

    file_name = first_file.get("name", "")

    file_type = first_file.get("filetype", "")

    say(
        text=f"""
契約依頼を受信しました。

ユーザー:
{user}

ファイル名:
{file_name}

ファイル種別:
{file_type}

次のステップ:
PDF解析
→ NDA判定
→ freee申請内容確認
"""
    )


if __name__ == "__main__":

    print("===================================")
    print("CONTRACT BOT START")
    print("===================================")

    print(
        f"Target channel: {CONTRACT_CHANNEL_ID}"
    )

    handler = SocketModeHandler(
        app,
        SLACK_APP_TOKEN
    )

    print("Socket Mode connecting...")

    handler.start()
