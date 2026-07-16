@app.event("message")
def handle_message(event, say):

    print("EVENT RECEIVED")
    print(event)

    if event.get("subtype"):
        return

    say("イベント受信")
