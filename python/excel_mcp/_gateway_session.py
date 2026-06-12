"""In-process JWT cache for MCP subprocesses authenticating via /api/chat/init."""
from __future__ import annotations

import hashlib
import ipaddress
import os
import threading
import time
from typing import Dict, Optional
from urllib.parse import urlparse

import httpx

_REFRESH_SKEW_SECONDS = 60
_INIT_TIMEOUT_SECONDS = 10
_CHANNEL = "mcp"
_LOOPBACK_HOSTS = {"localhost", "::1"}


class _CacheEntry:
  __slots__ = ("session_token", "session_id", "expires_at")

  def __init__(self, session_token: str, session_id: str, expires_at: int):
    self.session_token = session_token
    self.session_id = session_id
    self.expires_at = expires_at


_CACHE: Dict[str, _CacheEntry] = {}
_LOCK = threading.Lock()


def _cache_key(api_key: str, gateway_base_url: str) -> str:
  digest = hashlib.sha256(f"{api_key}|{gateway_base_url}".encode("utf-8")).hexdigest()
  return digest[:16]


def _is_loopback_host(host: Optional[str]) -> bool:
  if not host:
    return False
  if host.lower() in _LOOPBACK_HOSTS:
    return True
  try:
    return ipaddress.ip_address(host).is_loopback
  except ValueError:
    return False


def _normalize_gateway_url(raw: str) -> str:
  candidate = raw.strip()
  if "://" not in candidate:
    return f"https://{candidate}"
  return candidate


def default_tls_verify(gateway_base_url: str) -> bool:
  parsed = urlparse(_normalize_gateway_url(gateway_base_url))
  return not _is_loopback_host(parsed.hostname)


def _gateway_api_base_url(base_url: str) -> str:
  stripped = base_url.rstrip("/")
  return stripped if stripped.endswith("/api") else f"{stripped}/api"


def _init_url(gateway_base_url: str) -> str:
  return f"{_gateway_api_base_url(gateway_base_url)}/chat/init"


def _exchange(api_key: str, gateway_base_url: str, *, tls_verify: bool) -> _CacheEntry:
  body = {
    "api_key": api_key,
    "user_id": "mcp-subprocess",
    "context": {"channel": _CHANNEL},
  }
  with httpx.Client(verify=tls_verify, timeout=_INIT_TIMEOUT_SECONDS) as client:
    response = client.post(_init_url(gateway_base_url), json=body)
  if response.status_code >= 400:
    raise RuntimeError(f"Gateway /api/chat/init failed: HTTP {response.status_code}")
  payload = response.json()
  return _CacheEntry(
    session_token=str(payload["session_token"]),
    session_id=str(payload["session_id"]),
    expires_at=int(payload["expires_at"]),
  )


def get_session_token(
  api_key: str,
  *,
  gateway_base_url: str,
  tls_verify: Optional[bool] = None,
) -> str:
  """Return a valid JWT for the gateway, exchanging the API key if needed."""
  if tls_verify is None:
    tls_verify = default_tls_verify(gateway_base_url)
  key = _cache_key(api_key, gateway_base_url)
  now = int(time.time())
  with _LOCK:
    entry = _CACHE.get(key)
    if entry is not None and entry.expires_at - _REFRESH_SKEW_SECONDS > now:
      return entry.session_token
    entry = _exchange(api_key, gateway_base_url, tls_verify=tls_verify)
    _CACHE[key] = entry
    return entry.session_token


def invalidate(api_key: str, gateway_base_url: str) -> None:
  """Drop the cached entry for (api_key, gateway_base_url). Call on 401."""
  key = _cache_key(api_key, gateway_base_url)
  with _LOCK:
    _CACHE.pop(key, None)
