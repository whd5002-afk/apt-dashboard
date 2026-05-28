#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apt dashboard local server
  http://localhost:8080  -> dashboard HTML
  GET  /api/data         -> apt_transactions.json summary
  POST /api/refresh      -> trigger data collection
  GET  /api/status       -> collection progress
"""

import http.server
import json
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

PORT = 8080
BASE_DIR = Path(__file__).parent
JSON_PATH = BASE_DIR / "apt_transactions.json"

_state = {"status": "idle", "message": ""}
_lock  = threading.Lock()


def run_refresh():
    with _lock:
        _state["status"] = "running"
        _state["message"] = "collecting"
    try:
        result = subprocess.run(
            [sys.executable, str(BASE_DIR / "apt_price_analysis.py")],
            capture_output=True, text=True, cwd=str(BASE_DIR),
            encoding="utf-8", errors="replace"
        )
        with _lock:
            if result.returncode == 0:
                _state["status"] = "done"
                _state["message"] = "done"
            else:
                _state["status"] = "error"
                _state["message"] = (result.stderr or "unknown error")[-300:]
    except Exception as e:
        with _lock:
            _state["status"] = "error"
            _state["message"] = str(e)


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_GET(self):
        if   self.path == "/api/data":   self._serve_data()
        elif self.path == "/api/status": self._serve_status()
        else: super().do_GET()

    def do_POST(self):
        if self.path == "/api/refresh":  self._trigger_refresh()
        else: self.send_error(404)

    def _serve_data(self):
        if not JSON_PATH.exists():
            return self._json(404, {"error": "no data"})
        with open(JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)
        # summary: apt_transactions.json은 한글 키, data.json은 영문 키
        raw = data.get("summary", {})
        first_val = next(iter(raw.values()), {}) if raw else {}
        if "거래건수" in first_val:
            # apt_transactions.json 형식 — 키 변환
            summary = {
                city: {
                    "count":  v.get("거래건수", 0),
                    "avg":    v.get("평균가_만원", 0),
                    "median": v.get("중위가_만원", 0),
                    "min":    v.get("최저가_만원", 0),
                    "max":    v.get("최고가_만원", 0),
                }
                for city, v in raw.items()
            }
        else:
            summary = raw  # data.json 형식 — 이미 영문 키
        out = {"meta": data.get("meta", {}), "summary": summary}
        for field in ("rent", "youth_housing", "yh_complexes"):
            if field in data:
                out[field] = data[field]
        # data.json에서 추가 필드 보완 (apt_transactions.json만 있을 경우)
        data_json = JSON_PATH.parent / "data.json"
        if data_json.exists() and data_json != JSON_PATH:
            try:
                with open(data_json, encoding="utf-8") as f2:
                    d2 = json.load(f2)
                for field in ("rent", "youth_housing", "yh_complexes"):
                    if field not in out and field in d2:
                        out[field] = d2[field]
            except Exception:
                pass
        self._json(200, out)

    def _serve_status(self):
        with _lock:
            self._json(200, dict(_state))

    def _trigger_refresh(self):
        with _lock:
            if _state["status"] == "running":
                return self._json(200, {"status": "running", "message": "already running"})
            _state["status"] = "running"
            _state["message"] = "starting"
        threading.Thread(target=run_refresh, daemon=True).start()
        self._json(200, {"status": "running", "message": "started"})

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def address_string(self):
        return self.client_address[0]  # skip reverse DNS lookup

    def log_message(self, *_):
        pass


def _open_browser():
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{PORT}/apt_dashboard.html")


if __name__ == "__main__":
    threading.Thread(target=_open_browser, daemon=True).start()
    print(f"\n  [OK] Dashboard server started -> http://localhost:{PORT}")
    print("  Stop: Ctrl+C\n")
    try:
        with http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  Server stopped.")
