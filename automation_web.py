import argparse
import base64
import json
import mimetypes
import os
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from automation_center import AutomationCenter
from kuaishou_client import KuaishouClient


ROOT_DIR = Path(__file__).resolve().parent
WEB_DIR = ROOT_DIR / "web_console"


def json_response(handler, payload, status=200):
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def error_response(handler, message, status=400):
    json_response(handler, {"ok": False, "error": message}, status=status)


def encode_state(value):
    raw = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_state(value):
    if not value:
        return {}
    try:
        padded = value + "=" * ((4 - len(value) % 4) % 4)
        return json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
    except Exception:
        return {}


def redact_token_response(value):
    if isinstance(value, list):
        return [redact_token_response(item) for item in value]
    if not isinstance(value, dict):
        return value
    result = {}
    for key, item in value.items():
        lower_key = key.lower()
        if "token" in lower_key or lower_key in {"secret", "app_id"}:
            result[key] = "<hidden>" if item else item
        else:
            result[key] = redact_token_response(item)
    return result


class AutomationHandler(BaseHTTPRequestHandler):
    server_version = "KuaishouAutomation/1.0"

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            return self._serve_static("index.html")
        if parsed.path == "/ksAuthCallback":
            return self._handle_auth_callback(parsed)
        if parsed.path.startswith("/assets/"):
            return self._serve_static(parsed.path.removeprefix("/assets/"))
        if parsed.path == "/api/auth/status":
            return self._handle_auth_status()
        if parsed.path == "/api/auth/authorize-url":
            return self._handle_authorize_url()
        if parsed.path == "/api/state":
            query = parse_qs(parsed.query)
            report_date = (query.get("date") or [""])[0]
            return self._with_center(lambda center: json_response(self, {"ok": True, **center.dashboard(report_date)}))
        if parsed.path == "/api/version":
            return self._with_center(lambda center: json_response(self, {"ok": True, **center.data_version()}))
        if parsed.path == "/api/logs":
            query = parse_qs(parsed.query)
            limit = int((query.get("limit") or ["100"])[0])
            return self._with_center(lambda center: json_response(self, {"ok": True, "logs": center.list_logs(limit)}))
        if parsed.path == "/api/candidates":
            query = parse_qs(parsed.query)
            limit = int((query.get("limit") or ["200"])[0])
            report_date = (query.get("date") or [""])[0]
            return self._with_center(
                lambda center: json_response(
                    self,
                    {"ok": True, "candidates": center.list_candidates(limit=limit, report_date=report_date)},
                )
            )
        error_response(self, "Not found", status=404)

    def do_POST(self):
        parsed = urlparse(self.path)
        payload = self._read_json()
        if parsed.path == "/api/settings":
            return self._with_center(
                lambda center: json_response(self, {"ok": True, "settings": center.update_settings(payload)})
            )
        if parsed.path == "/api/rules":
            return self._with_center(lambda center: json_response(self, {"ok": True, "rule": center.save_rule(payload)}))
        if parsed.path == "/api/run":
            return self._with_center(
                lambda center: json_response(
                    self,
                    {
                        "ok": True,
                        "result": center.run_rules(
                            source=payload.get("source") or "web",
                            dry_run=payload.get("dry_run"),
                            respect_enabled=bool(payload.get("respect_enabled", False)),
                        ),
                    },
                )
            )
        if parsed.path == "/api/auth/exchange":
            return self._handle_auth_exchange(payload)
        if parsed.path == "/api/auth/check":
            return self._handle_auth_status(force_check=True)
        error_response(self, "Not found", status=404)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/rules":
            query = parse_qs(parsed.query)
            rule_id = int((query.get("rule_id") or ["0"])[0])
            return self._with_center(
                lambda center: json_response(self, {"ok": True, "deleted": center.delete_rule(rule_id)})
            )
        error_response(self, "Not found", status=404)

    def log_message(self, format, *args):
        print("%s - %s" % (self.address_string(), format % args))

    def _serve_static(self, name):
        safe_name = name.strip("/").replace("\\", "/") or "index.html"
        path = (WEB_DIR / safe_name).resolve()
        if WEB_DIR.resolve() not in path.parents and path != WEB_DIR.resolve():
            return error_response(self, "Bad path", status=400)
        if not path.exists() or not path.is_file():
            return error_response(self, "Not found", status=404)
        content = path.read_bytes()
        content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        if path.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        data = self.rfile.read(length)
        try:
            return json.loads(data.decode("utf-8"))
        except json.JSONDecodeError:
            return {}

    def _base_url(self):
        host = self.headers.get("Host") or "127.0.0.1:8787"
        proto = self.headers.get("X-Forwarded-Proto") or "http"
        return f"{proto}://{host}"

    def _redirect_uri(self):
        configured = os.getenv("KUAISHOU_AUTH_REDIRECT_URI", "").strip()
        if configured:
            return configured
        callback_port = os.getenv("KUAISHOU_AUTH_CALLBACK_PORT", "8000").strip() or "8000"
        return f"http://127.0.0.1:{callback_port}/ksAuthCallback"

    def _handle_auth_status(self, force_check=False):
        client = KuaishouClient()
        status = client.check_token() if force_check else client.token_status()
        if not force_check:
            if not status["has_app_id"] or not status["has_secret"]:
                status = {**status, "ok": False, "state": "missing_config", "message": "缺少 APP_ID 或 SECRET"}
            elif not status["has_access_token"] or not status["has_refresh_token"]:
                status = {**status, "ok": False, "state": "missing_token", "message": "缺少 token"}
            else:
                status = {**status, "ok": None, "state": "unknown", "message": "未检测"}
        return json_response(self, {"ok": True, "auth": status})

    def _handle_authorize_url(self):
        client = KuaishouClient()
        state = encode_state({"returnTo": self._base_url() + "/?auth=success"})
        return json_response(self, {"ok": True, "url": client.authorize_url(self._redirect_uri(), state=state)})

    def _handle_auth_exchange(self, payload):
        auth_code = str(payload.get("auth_code") or payload.get("authCode") or "").strip()
        if not auth_code:
            return error_response(self, "授权码不能为空", status=400)
        client = KuaishouClient()
        result = client.exchange_access_token(auth_code)
        auth = {**client.token_status(), "ok": True, "state": "authorized", "message": "授权成功"}
        return json_response(self, {"ok": True, "result": redact_token_response(result), "auth": auth})

    def _handle_auth_callback(self, parsed):
        query = parse_qs(parsed.query)
        auth_code = (query.get("auth_code") or query.get("authCode") or query.get("code") or [""])[0]
        if not auth_code:
            self.send_response(400)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("<h2>快手授权回调缺少 auth_code</h2><p>请重新点击授权。</p>".encode("utf-8"))
            return
        state = decode_state((query.get("state") or [""])[0])
        return_to = state.get("returnTo") or self._base_url() + "/"
        error = ""
        try:
            KuaishouClient().exchange_access_token(auth_code)
        except Exception as exc:
            error = str(exc)
        self.send_response(200 if not error else 500)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if error:
            body = f"<h2>快手授权失败</h2><pre>{error}</pre>"
        else:
            body = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>快手授权成功</title></head>
<body><p>快手授权成功，正在返回中台...</p><script>location.replace({json.dumps(return_to, ensure_ascii=False)});</script></body></html>"""
        self.wfile.write(body.encode("utf-8"))

    def _with_center(self, callback):
        center = AutomationCenter(
            getattr(self.server, "db_path", None),
            init_schema=False,
            db_timeout=3,
            busy_timeout_ms=3000,
        )
        try:
            return callback(center)
        except sqlite3.OperationalError as error:
            message = str(error)
            if "locked" in message.lower() or "busy" in message.lower():
                return error_response(self, "数据库正在入库或写入，请稍后刷新", status=503)
            return error_response(self, message, status=500)
        except Exception as error:
            return error_response(self, str(error), status=500)
        finally:
            center.close()


def run_server(host="127.0.0.1", port=8787, db_path=None):
    server = ThreadingHTTPServer((host, int(port)), AutomationHandler)
    server.db_path = db_path
    print(f"快手自动投放中台已启动: http://{host}:{port}", flush=True)
    server.serve_forever()


def main(argv=None):
    parser = argparse.ArgumentParser(description="快手自动投放中台 Web 服务")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8787)
    parser.add_argument("--db", default=None)
    args = parser.parse_args(argv)
    run_server(host=args.host, port=args.port, db_path=args.db)


if __name__ == "__main__":
    main()
