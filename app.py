import os
import json

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

app = App(
    token=os.environ["SLACK_BOT_TOKEN"]
)

@app.event("message")
def handle_message(event, say):

    print("========== EVENT RECEIVED ==========")
    print(json.dumps(event, ensure_ascii=False, indent=2))
    print("===================================")

    say("イベント受信成功")

if __name__ == "__main__":

    print("BOT START")

    SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"]
    ).start()
