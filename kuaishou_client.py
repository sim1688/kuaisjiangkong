import json
import msvcrt
import os
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin
from urllib.request import Request, urlopen


ROOT_DIR = Path(__file__).resolve().parent
ENV_PATH = ROOT_DIR / ".env"
SHARED_ENV_PATH = ROOT_DIR.parent / "2-快手自动创编" / ".env"
TOKEN_LOCK_PATH = ROOT_DIR / "logs" / "kuaishou_token_refresh.lock"

AUTH_SCOPE = [
    "ad_query",
    "ad_manage",
    "report_service",
    "account_service",
    "public_dmp_service",
    "public_agent_service",
    "public_account_service",
]


def _read_env_values(path):
    if not path.exists():
        return {}
    values = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        text = line.strip()
        if not text or text.startswith("#") or "=" not in text:
            continue
        key, value = text.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        values[key] = value
    return values


def load_env(path=ENV_PATH, override=False):
    for key, value in _read_env_values(path).items():
        if override:
            os.environ[key] = value
        else:
            os.environ.setdefault(key, value)


def load_runtime_env():
    load_env(ENV_PATH, override=True)


def write_env_values(updates, path=ENV_PATH):
    lines = []
    if path.exists():
        lines = path.read_text(encoding="utf-8").splitlines()

    seen = set()
    next_lines = []
    for line in lines:
        if "=" not in line or line.strip().startswith("#"):
            next_lines.append(line)
            continue
        key = line.split("=", 1)[0].strip()
        if key in updates:
            next_lines.append(f"{key}={updates[key]}")
            seen.add(key)
        else:
            next_lines.append(line)

    for key, value in updates.items():
        if key not in seen:
            next_lines.append(f"{key}={value}")

    path.write_text("\n".join(next_lines).rstrip() + "\n", encoding="utf-8")


def write_runtime_env_values(updates):
    write_env_values(updates, ENV_PATH)


def timestamp_seconds(value):
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return 0
    if number > 1000000000000:
        return number / 1000
    if number > 1000000000:
        return number
    return 0


def acquire_token_lock(timeout_seconds=20):
    TOKEN_LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    handle = TOKEN_LOCK_PATH.open("a+")
    deadline = time.time() + timeout_seconds
    while True:
        try:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            handle.seek(0)
            handle.truncate()
            handle.write(f"pid={os.getpid()} locked_at={int(time.time())}\n")
            handle.flush()
            return handle
        except OSError:
            if time.time() >= deadline:
                handle.close()
                raise KuaishouApiError("token refresh is busy, try again later")
            time.sleep(0.5)


def release_token_lock(handle):
    if not handle:
        return
    try:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        handle.close()


class KuaishouApiError(RuntimeError):
    def __init__(self, message, status=None, body=None):
        super().__init__(message)
        self.status = status
        self.body = body


class KuaishouClient:
    def __init__(self):
        load_runtime_env()
        self.base_url = os.getenv("KUAISHOU_API_BASE", "https://ad.e.kuaishou.com").rstrip("/")
        self.app_id = os.getenv("KUAISHOU_APP_ID", "")
        self.secret = os.getenv("KUAISHOU_SECRET", "")
        self.access_token = os.getenv("KUAISHOU_ACCESS_TOKEN", "")
        self.refresh_token = os.getenv("KUAISHOU_REFRESH_TOKEN", "")
        self.expires_at = timestamp_seconds(os.getenv("KUAISHOU_TOKEN_EXPIRES_AT", ""))
        self.refresh_token_expires_at = timestamp_seconds(os.getenv("KUAISHOU_REFRESH_TOKEN_EXPIRES_AT", ""))

    def token_status(self):
        return {
            "has_app_id": bool(self.app_id),
            "has_secret": bool(self.secret),
            "has_access_token": bool(self.access_token),
            "has_refresh_token": bool(self.refresh_token),
            "advertiser_id": os.getenv("KUAISHOU_ADVERTISER_ID", "") or None,
            "auth_user_id": os.getenv("KUAISHOU_AUTH_USER_ID", "") or None,
            "expires_at": self.expires_at or None,
            "refresh_token_expires_at": self.refresh_token_expires_at or None,
        }

    def reload_tokens_from_env(self):
        load_runtime_env()
        self.access_token = os.getenv("KUAISHOU_ACCESS_TOKEN", "")
        self.refresh_token = os.getenv("KUAISHOU_REFRESH_TOKEN", "")
        self.expires_at = timestamp_seconds(os.getenv("KUAISHOU_TOKEN_EXPIRES_AT", ""))
        self.refresh_token_expires_at = timestamp_seconds(os.getenv("KUAISHOU_REFRESH_TOKEN_EXPIRES_AT", ""))

    def needs_access_token_refresh(self, leeway_seconds=1800):
        if not self.access_token or not self.expires_at:
            return True
        return time.time() >= self.expires_at - leeway_seconds

    def authorize_url(self, redirect_uri, state=""):
        if not self.app_id:
            raise KuaishouApiError("Missing KUAISHOU_APP_ID")
        params = {
            "app_id": self.app_id,
            "scope": json.dumps(AUTH_SCOPE, ensure_ascii=False, separators=(",", ":")),
            "redirect_uri": redirect_uri,
            "oauth_type": "advertiser",
        }
        if state:
            params["state"] = state
        return "https://developers.e.kuaishou.com/tools/authorize?" + urlencode(params)

    def check_token(self):
        status = self.token_status()
        if not status["has_app_id"] or not status["has_secret"]:
            return {**status, "ok": False, "state": "missing_config", "message": "缺少 APP_ID 或 SECRET"}
        if not status["has_access_token"]:
            return {**status, "ok": False, "state": "missing_access_token", "message": "缺少 access token"}
        if not status["has_refresh_token"]:
            return {**status, "ok": False, "state": "missing_refresh_token", "message": "缺少 refresh token"}
        try:
            self.ensure_access_token(leeway_seconds=3600)
            return {**self.token_status(), "ok": True, "state": "valid", "message": "token 正常"}
        except Exception as error:
            body = getattr(error, "body", None)
            return {
                **self.token_status(),
                "ok": False,
                "state": "invalid",
                "message": str(error),
                "detail": body if isinstance(body, dict) else None,
            }

    def exchange_access_token(self, auth_code):
        if not self.app_id or not self.secret:
            raise KuaishouApiError("Missing KUAISHOU_APP_ID or KUAISHOU_SECRET")
        body = self._request_json(
            "/rest/openapi/oauth2/authorize/access_token",
            method="POST",
            body={
                "app_id": self.app_id,
                "secret": self.secret,
                "auth_code": auth_code,
            },
            auth=False,
        )
        self._assert_success(body)
        self._apply_token_response(body)
        self._persist_token_response(body)
        return body

    def ensure_access_token(self, leeway_seconds=1800):
        lock_handle = acquire_token_lock()
        try:
            self.reload_tokens_from_env()
            if not self.needs_access_token_refresh(leeway_seconds=leeway_seconds):
                return {"ok": True, "skipped": True, "message": "access token still valid"}
            return self.refresh_access_token(force=True, locked=True)
        finally:
            release_token_lock(lock_handle)

    def refresh_access_token(self, force=True, locked=False):
        if not self.app_id or not self.secret:
            raise KuaishouApiError("Missing KUAISHOU_APP_ID or KUAISHOU_SECRET")
        if not self.refresh_token:
            raise KuaishouApiError("Missing KUAISHOU_REFRESH_TOKEN")
        if not force and not self.needs_access_token_refresh():
            return {"ok": True, "skipped": True, "message": "access token still valid"}
        if not locked:
            lock_handle = acquire_token_lock()
            try:
                self.reload_tokens_from_env()
                return self.refresh_access_token(force=force, locked=True)
            finally:
                release_token_lock(lock_handle)
        body = self._request_json(
            "/rest/openapi/oauth2/authorize/refresh_token",
            method="POST",
            body={
                "app_id": self.app_id,
                "secret": self.secret,
                "refresh_token": self.refresh_token,
            },
            auth=False,
        )
        self._assert_success(body)
        self._apply_token_response(body)
        self._persist_token_response(body)
        return body

    def request(self, path, method="GET", body=None, query=None):
        return self._request_json(path, method=method, body=body, query=query, auth=True)

    def list_all(self, path, body, detail_key="details", page_size=None, raw_callback=None):
        page_size = int(page_size or os.getenv("KS_PAGE_SIZE", "100") or 100)
        page = int(body.get("page") or 1)
        details = []
        total = None

        while True:
            page_body = dict(body)
            page_body["page"] = page
            page_body["page_size"] = page_size
            response = self.request(path, method="POST", body=page_body)
            self._assert_success(response)
            if raw_callback:
                raw_callback(path, page_body, response)

            data = response.get("data") or {}
            page_details = data.get(detail_key) if isinstance(data, dict) else []
            if not isinstance(page_details, list):
                page_details = []
            details.extend(page_details)

            total_value = data.get("total_count") or data.get("total") or len(page_details)
            total = int(total_value or 0)
            if not page_details or len(details) >= total:
                break
            page += 1

        return {"total_count": total if total is not None else len(details), "details": details}

    def _request_json(self, path, method="GET", body=None, query=None, auth=True):
        url = path if path.startswith("http") else urljoin(self.base_url + "/", path.lstrip("/"))
        if query:
            items = {key: value for key, value in query.items() if value not in (None, "")}
            if items:
                separator = "&" if "?" in url else "?"
                url = f"{url}{separator}{urlencode(items)}"

        data = None
        headers = {"content-type": "application/json"}
        if auth:
            headers["Access-Token"] = self._get_access_token()
        if body is not None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")

        request = Request(url, data=data, method=method, headers=headers)
        try:
            with urlopen(request, timeout=30) as response:
                text = response.read().decode("utf-8")
                return json.loads(text) if text else {}
        except HTTPError as error:
            error_text = error.read().decode("utf-8", errors="replace")
            parsed = _parse_json_or_raw(error_text)
            raise KuaishouApiError(f"Kuaishou API HTTP {error.code}", status=error.code, body=parsed) from error
        except URLError as error:
            raise KuaishouApiError(f"Kuaishou API request failed: {error.reason}") from error
        except TimeoutError as error:
            raise KuaishouApiError("Kuaishou API request timeout") from error

    def _get_access_token(self):
        if not self.access_token:
            raise KuaishouApiError("Missing KUAISHOU_ACCESS_TOKEN. Exchange an auth code first.")
        if self.refresh_token and self.needs_access_token_refresh(leeway_seconds=120):
            self.ensure_access_token(leeway_seconds=120)
        return self.access_token

    def _apply_token_response(self, body):
        data = body.get("data") if isinstance(body, dict) else None
        if not isinstance(data, dict):
            data = body if isinstance(body, dict) else {}
        self.access_token = data.get("access_token") or self.access_token
        self.refresh_token = data.get("refresh_token") or self.refresh_token
        expires_in = int(data.get("access_token_expires_in") or data.get("expires_in") or data.get("expires") or 0)
        if expires_in > 0:
            self.expires_at = time.time() + expires_in
        refresh_expires_in = int(data.get("refresh_token_expires_in") or 0)
        if refresh_expires_in > 0:
            self.refresh_token_expires_at = time.time() + refresh_expires_in

    def _persist_token_response(self, body):
        data = body.get("data") if isinstance(body, dict) else None
        if not isinstance(data, dict):
            data = body if isinstance(body, dict) else {}
        updates = {}
        if data.get("access_token"):
            updates["KUAISHOU_ACCESS_TOKEN"] = data["access_token"]
        if data.get("refresh_token"):
            updates["KUAISHOU_REFRESH_TOKEN"] = data["refresh_token"]
        if self.expires_at:
            updates["KUAISHOU_TOKEN_EXPIRES_AT"] = str(int(self.expires_at * 1000))
        if self.refresh_token_expires_at:
            updates["KUAISHOU_REFRESH_TOKEN_EXPIRES_AT"] = str(int(self.refresh_token_expires_at * 1000))
        advertiser_id = data.get("advertiser_id") or data.get("advertiserId")
        if advertiser_id:
            updates["KUAISHOU_ADVERTISER_ID"] = str(advertiser_id)
        if updates:
            write_runtime_env_values(updates)

    @staticmethod
    def _assert_success(body):
        if not isinstance(body, dict):
            return
        code = body.get("code")
        if code in (None, 0):
            return
        raise KuaishouApiError(body.get("message") or f"Kuaishou API business error {code}", status=502, body=body)


def _parse_json_or_raw(text):
    try:
        return json.loads(text) if text else {}
    except json.JSONDecodeError:
        return {"raw": text}
