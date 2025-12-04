# server.py
from __future__ import annotations
from dataclasses import asdict
from typing import Optional, Tuple, Dict, Any
from flask import Flask, request, jsonify
import controller
from model import Language, GreetingStyle, SignOffStyle

class OutlookAssistServer:
    def __init__(self):
        self.app = Flask(__name__)
        self._routes()

    def _routes(self):
        @self.app.post("/assist/reply")
        def assist_reply():
            # Expected JSON payload from the add-in
            data: Dict[str, Any] = request.get_json(force=True, silent=False)

            recipient_display_name: str = data.get("recipient_display_name") or "User"
            incoming_sender_name: str = data.get("incoming_sender_name") or ""
            incoming_sender_email: str = data.get("incoming_sender_email") or ""
            incoming_subject: str = data.get("incoming_subject") or ""
            incoming_body: str = data.get("incoming_body") or ""
            tone: str = data.get("tone") or ""
            extra: str = data.get("extra") or ""
            greeting_style = GreetingStyle(data.get("greeting_style", GreetingStyle.AUTO.value))
            signoff_style = SignOffStyle(data.get("signoff_style", SignOffStyle.BEST_REGARDS.value))
            language = Language(data.get("language", Language.EN.value))

            subj, body = controller.generate_reply_email(
                recipient_display_name=recipient_display_name,
                incoming_sender_name=incoming_sender_name,
                incoming_sender_email=incoming_sender_email,
                incoming_subject=incoming_subject,
                incoming_body=incoming_body,
                tone_instructions=tone,
                extra_instructions=extra,
                greeting_style=greeting_style,
                signoff_style=signoff_style,
                language=language
            )
            return jsonify({"subject": subj, "body": body})

        @self.app.get("/healthz")
        def health():
            return jsonify({"ok": True})

def main():
    srv = OutlookAssistServer()
    # For local dev with a browser add-in, HTTPS is strongly recommended.
    # Generate a dev cert or use a reverse-proxy like mkcert / OpenSSL.
    # app.run(ssl_context=('cert.pem','key.pem')) for HTTPS; plain HTTP shown for brevity.
    srv.app.run(host="127.0.0.1", port=5001, debug=True)

if __name__ == "__main__":
    main()
