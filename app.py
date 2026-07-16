print("APP LOADED")

import os

print("IMPORT OS")

from slack_bolt import App

print("IMPORT APP")

from slack_bolt.adapter.socket_mode import SocketModeHandler

print("IMPORT SOCKET")

print("BOT TOKEN EXISTS",
      "SLACK_BOT_TOKEN" in os.environ)

print("APP TOKEN EXISTS",
      "SLACK_APP_TOKEN" in os.environ)

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
