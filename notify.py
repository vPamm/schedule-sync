import requests

def send_telegram_message(message, chat_id, bot_token):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = {
        "chat_id": chat_id,
        "text": message,
    }
    response = requests.post(url, data=data)
    return response.status_code, response.text

if __name__ == "__main__":
    # Example values, replace these with your actual values
    bot_token = "YOUR_BOT_TOKEN"
    chat_id = "YOUR_CHAT_ID"
    message = "Schedule sync script has completed successfully."

    status_code, response_text = send_telegram_message(message, chat_id, bot_token)
    if status_code == 200:
        print("Message sent successfully.")
    else:
        print(f"Failed to send message: {response_text}")
