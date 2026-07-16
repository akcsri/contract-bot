import os
import requests

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

app = App(
    token=os.environ["SLACK_BOT_TOKEN"]
)

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]

@app.event("message")
def handle_message(event, say):

    if event.get("subtype"):
        return

    files = event.get("files", [])

    if not files:
        say("PDFが添付されていません")
        return

    file_info = files[0]

    file_name = file_info["name"]

    url = file_info["url_private_download"]

    say(
        f"受信しました\nファイル名: {file_name}"
    )

    headers = {
        "Authorization":
        f"Bearer {SLACK_BOT_TOKEN}"
    }

    response = requests.get(
        url,
        headers=headers
    )

    with open("/tmp/contract.pdf", "wb") as f:
        f.write(response.content)

    say("PDFダウンロード成功")


if __name__ == "__main__":

    SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"]
    ).start()
