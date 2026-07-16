import os

print("APP LOADED")
print("FILE=", __file__)

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

print("IMPORT OK")

app = App(
    token=os.environ["SLACK_BOT_TOKEN"]
)

print("APP CREATED")

@app.event("message")
def handle_message(event, say):
    say("受信しました")

print("EVENT REGISTERED")

if __name__ == "__main__":
    print("BOT START")

    handler = SocketModeHandler(
        app,
        os.environ["SLACK_APP_TOKEN"]
    )

    print("SOCKET START")

    handler.start()
