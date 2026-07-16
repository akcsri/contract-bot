import os

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

app = App(
    token=os.environ["SLACK_BOT_TOKEN"]
)

@app.event("message")
def handle_message(event, say):
    if event.get("subtype"):
        return

    say("受信しました")

if __name__ == "__main__":
    handler = SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"]
    )
    handler.start()

@app.event("message")
def handle_message(event, say):

    if event.get("subtype"):
        return

    files = event.get("files", [])

    if not files:
        say("PDFが添付されていません")
        return

    pdf = files[0]

    say(
        f"受信しました\nファイル名: {pdf['name']}"
    )
