"""
asana_client.py — Minimal Asana REST client.

Deliberately not the official SDK: this needs three things the SDK makes
awkward — transparent 429/5xx retry, cursor pagination as a generator, and a
dry-run mode that logs writes without sending them.
"""
from __future__ import annotations

import logging
import time
from typing import Any, Iterator

import requests

log = logging.getLogger("rise.asana")

BASE = "https://app.asana.com/api/1.0"
RETRY_STATUS = {429, 500, 502, 503, 504}


class AsanaError(RuntimeError):
    pass


class AsanaClient:
    def __init__(self, token: str, *, dry_run: bool = False, timeout: int = 30,
                 max_retries: int = 5):
        if not token:
            raise AsanaError("No Asana token. Set ASANA_TOKEN in your .env file.")
        self.dry_run = dry_run
        self.timeout = timeout
        self.max_retries = max_retries
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        })
        self.writes = 0
        self.reads = 0

    # ------------------------------------------------------------------
    def _request(self, method: str, path: str, **kw) -> dict[str, Any]:
        url = path if path.startswith("http") else f"{BASE}{path}"
        is_write = method.upper() in {"POST", "PUT", "DELETE"}

        if is_write and self.dry_run:
            log.info("[dry-run] %s %s %s", method, path,
                     str(kw.get("json", {}))[:200])
            self.writes += 1
            return {"data": {"gid": f"dry-{self.writes}"}}

        delay = 1.0
        for attempt in range(1, self.max_retries + 1):
            resp = self.session.request(method, url, timeout=self.timeout, **kw)
            if resp.status_code in RETRY_STATUS:
                wait = float(resp.headers.get("Retry-After", delay))
                if attempt == self.max_retries:
                    raise AsanaError(
                        f"{method} {path} failed after {attempt} attempts: "
                        f"{resp.status_code} {resp.text[:300]}")
                log.warning("HTTP %s on %s — retrying in %.1fs (%d/%d)",
                            resp.status_code, path, wait, attempt, self.max_retries)
                time.sleep(wait)
                delay = min(delay * 2, 30)
                continue
            if not resp.ok:
                raise AsanaError(f"{method} {path} -> {resp.status_code} {resp.text[:500]}")

            if is_write:
                self.writes += 1
            else:
                self.reads += 1
            return resp.json() if resp.content else {}
        raise AsanaError("unreachable")

    # ------------------------------------------------------------------
    def get(self, path: str, params: dict | None = None) -> Any:
        return self._request("GET", path, params=params).get("data")

    def post(self, path: str, data: dict) -> Any:
        return self._request("POST", path, json={"data": data}).get("data")

    def put(self, path: str, data: dict) -> Any:
        return self._request("PUT", path, json={"data": data}).get("data")

    def paginate(self, path: str, params: dict | None = None,
                 page_size: int = 100) -> Iterator[dict]:
        params = dict(params or {})
        params["limit"] = page_size
        while True:
            payload = self._request("GET", path, params=params)
            self.reads += 1
            yield from payload.get("data", [])
            nxt = (payload.get("next_page") or {}).get("offset")
            if not nxt:
                return
            params["offset"] = nxt

    # ------------------------------------------------------------------
    def me(self) -> dict:
        return self.get("/users/me", {"opt_fields": "name,email,workspaces.name"})
