#!/bin/bash
cd "$(dirname "$0")"

# Récupère l'IP locale du Mac
LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")

echo ""
echo "========================================"
echo "  Meeting Notes — serveur démarré"
echo "========================================"
echo ""
echo "  Sur votre iPhone (même WiFi) :"
echo "  > https://$LOCAL_IP"
echo ""
echo "========================================"
echo ""

sudo /Users/victor/Library/Python/3.9/bin/uvicorn main:app \
  --host 0.0.0.0 --port 443 \
  --ssl-keyfile key.pem --ssl-certfile cert.pem
