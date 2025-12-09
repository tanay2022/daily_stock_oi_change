from http.server import BaseHTTPRequestHandler
import json
import sys
from pathlib import Path

# Ensure root module is importable when deployed on Vercel
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from stock_daily_OI_change import vercel_handler  # noqa: E402


class handler(BaseHTTPRequestHandler):
    """Vercel entry point (Python Serverless Function)."""

    def _json_response(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, indent=2)
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.end_headers()
        self.wfile.write(body.encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.end_headers()

    def do_GET(self):  # noqa: N802 (Vercel handler signature)
        try:
            response = vercel_handler()
            status = response.get("statusCode", 200)
            body = json.loads(response.get("body", "{}"))
            self._json_response(body, status=status)
        except Exception as exc:  # pragma: no cover - defensive
            self._json_response(
                {
                    "success": False,
                    "error": f"Unhandled server error: {exc}",
                },
                status=500,
            )
