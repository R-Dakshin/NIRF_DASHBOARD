import json
import os
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse

from sync_dashboard import EXCEL_FILE, build_data


HOST = "127.0.0.1"
PORT = 8000

_CACHE = {"mtime": None, "payload": None}


def get_payload():
    mtime = os.path.getmtime(EXCEL_FILE)
    if _CACHE["payload"] is None or _CACHE["mtime"] != mtime:
        print("Refreshing data from Excel...")
        _CACHE["payload"] = build_data(EXCEL_FILE)
        _CACHE["mtime"] = mtime
    return _CACHE["payload"]


class DashboardHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # Allow local file-origin dashboard to call API during development.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/data":
            try:
                payload = get_payload()
                body = json.dumps(payload).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as exc:
                err = json.dumps({"error": str(exc)}).encode("utf-8")
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(err)))
                self.end_headers()
                self.wfile.write(err)
            return
        return super().do_GET()


if __name__ == "__main__":
    print(f"Starting dashboard server at http://{HOST}:{PORT}")
    print("Open /dash.html in your browser.")
    print("API endpoint: /api/data")
    HTTPServer((HOST, PORT), DashboardHandler).serve_forever()
