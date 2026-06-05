import os
import sys
import requests
from dotenv import load_dotenv

# Cargar variables locales
load_dotenv()

WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_NUMBER_ID = os.getenv("PHONE_NUMBER_ID", "1098169976722957")

def enviar_test(numero):
    if not WHATSAPP_TOKEN:
        print("❌ Error: WHATSAPP_TOKEN no está definido en tu archivo .env")
        return
        
    url = f"https://graph.facebook.com/v25.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "text",
        "text": {
            "body": "🤖 Mensaje de prueba de conexión de BotVip"
        }
    }
    
    print(f"Petición POST a: {url}")
    print(f"Destinatario: {numero}")
    print(f"Token (primeros 15 caracteres): {WHATSAPP_TOKEN[:15]}...")
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        print(f"\nCódigo de respuesta HTTP: {response.status_code}")
        print("Respuesta de Meta:")
        print(response.text)
    except Exception as e:
        print(f"❌ Error al enviar la petición: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python test_whatsapp.py <tu_numero_con_codigo_pais>")
        print("Ejemplo: python test_whatsapp.py 51900111222")
    else:
        enviar_test(sys.argv[1].strip())
