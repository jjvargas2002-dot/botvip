import requests
from config import Config

PHONE_NUMBER_ID = Config.PHONE_NUMBER_ID


def enviar_mensaje(numero, texto):

    url = f"https://graph.facebook.com/v25.0/{PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {Config.WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "text",
        "text": {
            "body": texto
        }
    }

    response = requests.post(url, headers=headers, json=payload)

    print(f"WhatsApp -> {numero}: {response.status_code}")
    print(response.text)

    return response
