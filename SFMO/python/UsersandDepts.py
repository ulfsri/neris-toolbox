#!/usr/bin/env python3
"""
NERIS Monthly Incident Count Report

Standalone script version (no Jupyter/ipywidgets) — run from a terminal,
e.g. in VS Code's integrated terminal:

    python neris_report.py

You'll be prompted for your NERIS email/password (password entry is hidden)
and the report parameters (state code, optional entity ID, date range).
Produces an .xlsx pivot of incident counts / NAR months per department.
"""

import sys
import subprocess
import os
import csv
import threading
import time
import traceback
from datetime import datetime, timezone, timedelta
from urllib.parse import urlsplit, parse_qs
from collections import deque


def ensure_dependencies():
    def pip_install(*packages):
        subprocess.run(
            [sys.executable, "-m", "pip", "install", *packages, "--quiet"],
            check=True
        )

    try:
        import openpyxl  # noqa
        print("✓ openpyxl ready")
    except ImportError:
        print("Installing openpyxl...")
        pip_install("openpyxl")
        print("✓ openpyxl installed")

    try:
        from neris_api_client import NerisApiClient  # noqa
        print("✓ Dependencies ready")
    except ImportError:
        print("Installing neris-api-client...")
        pip_install(
            "https://github.com/ulfsri/neris-api-client/archive/refs/heads/main.zip"
        )
        print("✓ neris-api-client installed")


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS State Roster — Departments & Users")
    print("=" * 60)

    print("\n-- Credentials --")
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()

    print("\n-- Query Parameters --")
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()

    if not username or not password or not state_code:
        sys.exit("✗ Email, password, and state code are required.")

    return username, password, state_code


_auth_lock = threading.Lock()
_orig_call = None  
_last_reauth_at = {"t": 0.0}
_mfa_wait_state = {"waiting": False, "since": 0.0}
_MFA_INPUT_TIMEOUT_SECONDS = 600


def _input_with_timeout(prompt, timeout_seconds):
    """Like input(), but gives up after timeout_seconds instead of blocking
    forever. Implemented with a daemon thread + queue since input() itself
    has no cross-platform timeout. If nothing is typed in time, raises
    TimeoutError and the background thread is simply abandoned -- if a code
    is typed into it after the fact, it's read by that orphaned thread and
    discarded, never consumed by the next prompt, which is a fine trade for
    not hanging indefinitely."""
    import queue
    q = queue.Queue(maxsize=1)

    def _reader():
        try:
            q.put(input(prompt))
        except Exception:
            pass  

    threading.Thread(target=_reader, daemon=True).start()
    try:
        return q.get(timeout=timeout_seconds)
    except queue.Empty:
        raise TimeoutError(
            f"No MFA code entered within {timeout_seconds}s — requesting "
            f"a fresh code instead of waiting indefinitely"
        )


def _maybe_force_reauth(client_self, token_before, context_label):
    """Shared by every retry path that can trigger a forced re-auth (the
    patched client._call, and the hand-rolled GETs in get_vendor_names /
    get_users_for_entity, which bypass client._call() entirely). Must be
    called already holding _auth_lock. Forces a fresh re-authentication
    only if BOTH:
      (a) nobody has fixed the token since this caller detected the
          failure (token_before still matches the live token) -- if
          someone else already fixed it, just use that.
      (b) no re-auth (proactive or forced) completed in the last few
          seconds -- unconditional insurance on top of (a), making it
          structurally impossible to fire two full password+MFA logins
          within a few seconds of each other.

    Returns True if the CALLER should sleep briefly before retrying (the
    cooldown case) -- deliberately NOT sleeping in here, since this
    function runs while the caller is holding _auth_lock, and sleeping
    while holding that lock would stall every other thread's own
    (unrelated) auth check for the duration. The caller does the sleep
    itself, after releasing the lock.
    """
    if client_self.tokens.access_token != token_before:
        print(f"  (token already refreshed by another thread — "
              f"retrying with it instead of starting another login) "
              f"[{context_label}]")
        return False
    if time.monotonic() - _last_reauth_at["t"] < 5.0:
        print(f"  (a re-authentication just completed moments ago — "
              f"waiting a beat instead of starting a second one) "
              f"[{context_label}]")
        return True
    client_self.tokens.expires_at = datetime.min
    try:
        client_self._update_auth()
    except Exception as e:
        print(f"  ⚠ Forced re-auth attempt failed: {e} — will retry "
              f"[{context_label}]")
    return False

_RATE_LIMITS = [
    (20, 5.0),        
]

_rate_lock = threading.Lock()
_request_times = [deque() for _ in _RATE_LIMITS]  


def _looks_like_rate_limit(resp):
    return resp.status_code == 429


def _looks_like_auth_drop(res):
    """True when a raw Response looks like the token was rejected — a 401,
    or a 403 (including the bare nginx-branded error page: NERIS's gateway
    appears to reject an invalid/expired token right at the edge). `res` is
    the raw requests.Response client._call() returns on any HTTP error."""
    import requests
    if not isinstance(res, requests.Response):
        return False
    return res.status_code in (401, 403)


def _log_rate_limit_headers(resp):
    """Print any rate-limit-related response headers so the real numbers
    are visible instead of us continuing to guess."""
    interesting = {
        k: v for k, v in resp.headers.items()
        if "rate" in k.lower() or "retry" in k.lower() or "limit" in k.lower()
           or "quota" in k.lower()
    }
    if interesting:
        print(f"    Rate-limit-related response headers: {interesting}")
    else:
        print("    (no rate-limit-related headers present on this response)")

_request_log_lock = threading.Lock()
_request_log_path = None
_request_log_file = None
_request_log_writer = None

_last_activity_at = {"t": time.monotonic()}

_REQUEST_LOG_HEADER = [
    "timestamp_utc", "method", "path", "neris_id_entity",
    "status_code", "x_request_id", "attempt", "rate_limit_suspected",
]


def _init_request_log(state_code):
    """Open the trace-log CSV for this run. Call once, before any request
    is sent. Safe to call more than once (only the first call does anything)."""
    global _request_log_path, _request_log_file, _request_log_writer
    if _request_log_file is not None:
        return _request_log_path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    _request_log_path = f"neris_request_log_{state_code}_{timestamp}.csv"
    _request_log_file = open(_request_log_path, "w", newline="", encoding="utf-8")
    _request_log_writer = csv.writer(_request_log_file)
    _request_log_writer.writerow(_REQUEST_LOG_HEADER)
    _request_log_file.flush()
    return _request_log_path


def _watch_for_stalls(stop_event, get_progress):
    """Runs in a daemon thread for the duration of the fetch phase. Every
    30s, checks how long it's been since the LAST HTTP response of any
    kind. Distinguishes "waiting on a human for an MFA code" (via
    _mfa_wait_state) from a genuine idle stretch, and a short idle stretch
    (normal retry/backoff) from one that looks genuinely stuck."""
    while not stop_event.wait(30):
        idle = time.monotonic() - _last_activity_at["t"]
        completed, total = get_progress()

        if _mfa_wait_state["waiting"]:
            waited = time.monotonic() - _mfa_wait_state["since"]
            print(f"  ⏳ Waiting on YOU — an MFA code prompt is live above "
                  f"(asked {waited:.0f}s ago, times out and retries "
                  f"automatically at {_MFA_INPUT_TIMEOUT_SECONDS}s). Check "
                  f"your email and type the code into the terminal.")
            continue

        if idle < 90:
            continue
        if idle < 300:
            print(f"  ⏳ Still working — no request has completed in "
                  f"{idle:.0f}s ({completed}/{total} departments done so "
                  f"far). This is normal if a department is mid-retry.")
        else:
            print(f"  ⚠ No request has completed in {idle:.0f}s "
                  f"({completed}/{total} departments done so far) — this "
                  f"looks genuinely stuck, not just slow. It's likely safe "
                  f"to Ctrl+C and rerun.")


def _log_request(method, url, resp, attempt, rate_limit_suspected):
    """Append one row to the trace log for a single HTTP response.
    Never logs headers/body beyond the request-id and status — no tokens."""
    _last_activity_at["t"] = time.monotonic()
    if _request_log_writer is None:
        return  # log not initialized -- skip rather than crash
    try:
        parts = urlsplit(url)
        path = parts.path
        qs = parse_qs(parts.query)
        neris_id_entity = (qs.get("neris_id_entity") or [""])[0]
        request_id = resp.headers.get("x-request-id", "") if resp is not None else ""
        status_code = resp.status_code if resp is not None else ""
        row = [
            datetime.now(timezone.utc).isoformat(),
            method.upper(),
            path,
            neris_id_entity,
            status_code,
            request_id,
            attempt,
            "yes" if rate_limit_suspected else "no",
        ]
        with _request_log_lock:
            _request_log_writer.writerow(row)
            _request_log_file.flush()
    except Exception as e:
        # Logging must never take down the actual data fetch.
        print(f"  ⚠ (non-fatal) failed to write request-id log row: {e}")


def _install_thread_safety_patch():
    global _orig_call
    from neris_api_client.client import _NerisApiClient
    from neris_api_client.config import TokenSet, GrantType
    if _orig_call is not None:
        return  # already patched
    _orig_call = _NerisApiClient._call

    def _safe_update_auth(self):
        if self.tokens.expires_at > datetime.now() + timedelta(seconds=120):
            return 

        token_url = f"{self.config.base_url}/token"

        def _post(data, headers):
            return self._session.post(token_url, headers=headers, data=data, timeout=30)

        def _password_grant():
            return _post(
                {
                    "grant_type": GrantType.PASSWORD,
                    "username": self.config.username,
                    "password": self.config.password,
                },
                {"Content-Type": "application/x-www-form-urlencoded"},
            )

        def _refresh_grant():
            return _post(
                {"grant_type": "refresh_token", "refresh_token": self.tokens.refresh_token},
                {"Content-Type": "application/x-www-form-urlencoded"},
            )

        have_refresh = bool(self.tokens.refresh_token)
        res = _refresh_grant() if have_refresh else _password_grant()
        allow_fallback_to_password = have_refresh

        while True:
            try:
                got = res.json()
            except ValueError:
                got = {}

            if res.status_code == 200:
                self.tokens = TokenSet(
                    access_token=got["access_token"],
                    refresh_token=got["refresh_token"],
                    expires_at=datetime.now() + timedelta(seconds=got["expires_in"]),
                )
                _last_reauth_at["t"] = time.monotonic()
                return

            if res.status_code == 202:
                print("\n" + "=" * 60)
                print("  CHECK YOUR EMAIL FOR THE MFA CODE")
                print("  (mid-run re-authentication needs a code entered here —")
                print(f"   waits up to {_MFA_INPUT_TIMEOUT_SECONDS}s, then requests a fresh "
                      f"code automatically)")
                print("=" * 60)
                _mfa_wait_state["waiting"] = True
                _mfa_wait_state["since"] = time.monotonic()
                try:
                    code = _input_with_timeout(
                        f"Provide MFA code for {got.get('challenge_name', 'MFA')}: ",
                        _MFA_INPUT_TIMEOUT_SECONDS,
                    )
                finally:
                    _mfa_wait_state["waiting"] = False
                res = self._session.post(
                    token_url,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    data={
                        "grant_type": got["challenge_name"],
                        "username": self.config.username,
                        "session": got["session"],
                        got["challenge_name"]: code,
                    },
                    timeout=30,
                )
                continue

            if allow_fallback_to_password:
                print(f"  ⚠ refresh_token grant failed (HTTP {res.status_code}) — "
                      f"likely expired (NERIS refresh tokens live ~30 min) — "
                      f"falling back to a full password login")
                self.tokens = TokenSet(access_token="", refresh_token="", expires_at=datetime.min)
                allow_fallback_to_password = False
                res = _password_grant()
                continue

            raise RuntimeError(
                f"Authentication failed: token endpoint returned "
                f"HTTP {res.status_code} - {res.text[:300]!r}"
            )

    _NerisApiClient._update_auth = _safe_update_auth

    def _send_request(self, method, path, data=None, params=None, model=None):
        import requests
        if self.config.validate and model:
            if isinstance(data, str):
                data = model.model_validate_json(data).model_dump(mode="json", by_alias=True)
            if isinstance(data, dict):
                data = model.model_validate(data).model_dump(mode="json", by_alias=True)

        res = getattr(self._session, method)(
            f"{self.config.base_url}{path}", json=data, params=params
        )

        if self.config.debug:
            import json as _json
            print(_json.dumps({
                "request": {
                    "url": f"{self.config.base_url}{path}", "body": data,
                    "headers": dict(self._session.headers), "params": params,
                },
                "response": {
                    "status_code": res.status_code,
                    "headers": dict(res.headers), "content": res.text,
                },
            }, indent=4, default=str))

        try:
            res.raise_for_status()
        except requests.exceptions.HTTPError as e:
            try:
                e.args = e.args + (e.response.json(),)
            except requests.exceptions.JSONDecodeError:
                e.args = e.args + (e.response.text,)
            finally:
                print(str(e))
                return res

        try:
            return res.json()
        except requests.exceptions.JSONDecodeError:
            return res.text

    def _locked_call(self, *args, **kwargs):
        max_attempts = 4
        res = None
        for attempt in range(1, max_attempts + 1):
            token_before = self.tokens.access_token
            try:
                with _auth_lock:
                    self._update_auth()
                    self._session.headers.update(
                        {"Authorization": f"Bearer {self.tokens.access_token}"}
                    )
                res = _send_request(self, *args, **kwargs)
            except Exception as e:
                if attempt == max_attempts:
                    raise
                print(f"  ⚠ Auth error on attempt {attempt}/{max_attempts} "
                      f"for {args[:2] if args else kwargs.get('path', '?')}: "
                      f"{e} — retrying")
                time.sleep(min(15, 3 * attempt))
                continue

            if not _looks_like_auth_drop(res):
                if attempt > 1:
                    print(f"  ✓ Recovered on attempt {attempt}/{max_attempts}")
                return res

            if attempt == max_attempts:
                return res  

            print(f"  ⚠ Token rejected (HTTP {res.status_code}) on "
                  f"{args[:2] if args else kwargs.get('path', '?')} "
                  f"(attempt {attempt}/{max_attempts}) — token fingerprint "
                  f"was {_token_fingerprint(self)}")
            with _auth_lock:
                need_cooldown = _maybe_force_reauth(self, token_before, f"roster/{args[:2] if args else kwargs.get('path', '?')}")
            if need_cooldown:
                time.sleep(1.5)  
        return res

    _NerisApiClient._call = _locked_call

    orig_session_request = _NerisApiClient._session.request

    def _wait_for_rate_budget():
        """Block until sending one more request satisfies EVERY rule in
        _RATE_LIMITS simultaneously."""
        while True:
            with _rate_lock:
                now = time.monotonic()
                sleep_for = 0.0
                for (max_requests, window_seconds), times in zip(_RATE_LIMITS, _request_times):
                    while times and now - times[0] > window_seconds:
                        times.popleft()
                    if len(times) >= max_requests:
                        sleep_for = max(sleep_for, window_seconds - (now - times[0]) + 0.05)
                if sleep_for <= 0:
                    for times in _request_times:
                        times.append(now)
                    return
            time.sleep(max(sleep_for, 0.05))

    def _throttled_request(method, url, *args, **kwargs):
        kwargs.setdefault("timeout", 60)
        _wait_for_rate_budget()

        max_attempts = 9
        resp = None
        for attempt in range(1, max_attempts + 1):
            resp = orig_session_request(method, url, *args, **kwargs)
            suspected = _looks_like_rate_limit(resp)
            _log_request(method, url, resp, attempt, suspected)
            if suspected and attempt < max_attempts:
                wait = min(300, 5 * (2 ** (attempt - 1)))
                print(f"  ⚠ Possible rate limit (HTTP {resp.status_code}) on "
                      f"{method.upper()} {url.split('?')[0]} — backing off "
                      f"{wait}s (attempt {attempt}/{max_attempts})")
                if attempt == 1:
                    _log_rate_limit_headers(resp)
                time.sleep(wait)
                continue
            return resp
        return resp

    _NerisApiClient._session.request = _throttled_request
    print("✓ Applied thread-safety patch to NerisApiClient")
    limits_desc = ", ".join(f"{n}/{w:.0f}s" for n, w in _RATE_LIMITS)
    print(f"✓ Applied request throttle (budgets: {limits_desc}) and rate-limit backoff")


def _unwrap(res, context):
    """client._call() catches HTTP errors internally and returns the raw
    Response object instead of raising, so a 403/404/500/etc. would
    otherwise silently look like "no results" to the caller. Detect that
    case here and raise a clear error instead."""
    import requests
    if isinstance(res, requests.Response):
        raise RuntimeError(
            f"{context} failed: HTTP {res.status_code} - {res.text[:500]!r}"
        )
    return res


def _token_fingerprint(client):
    """A short, safe-to-log stand-in for the access token — never print the
    real token, it's a bearer credential."""
    tok = getattr(client, "tokens", None)
    if not tok or not getattr(tok, "access_token", None):
        return "<no token>"
    return f"...{tok.access_token[-8:]} (expires_at={tok.expires_at})"


def authenticate(username, password):
    from neris_api_client import NerisApiClient, Config
    _install_thread_safety_patch()

    print("\nConnecting to NERIS API...")
    client = NerisApiClient(Config(
        base_url="https://api.neris.fsri.org/v1",
        grant_type="password",
        username=username,
        password=password,
    ))

    print("\n" + "=" * 60)
    print("  CHECK YOUR EMAIL FOR THE MFA CODE")
    print("=" * 60)

    last_error = None
    for attempt in range(1, 4):
        try:
            _unwrap(client.health(), "auth check (health)")
            print("\n✓ Authentication successful!")
            return client
        except RuntimeError as e:
            last_error = e
            if "HTTP 50" in str(e) and attempt < 3:
                wait = attempt * 5
                print(f"⚠ Transient server error on attempt {attempt}: {e}")
                print(f"  Retrying in {wait}s...")
                time.sleep(wait)
                continue
            sys.exit(f"✗ Authentication failed: {e}")
        except Exception as e:
            sys.exit(f"✗ Authentication failed: {e}")

    sys.exit(f"✗ Authentication failed after retries: {last_error}")


def _hcell(ws, row, col, value):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    THIN = Side(style="thin")
    c = ws.cell(row=row, column=col, value=value)
    c.fill = PatternFill(start_color="262F68", end_color="262F68", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    return c


def _dcell(ws, row, col, value):
    from openpyxl.styles import Font, Alignment, Border, Side
    THIN = Side(style="thin")
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(size=11)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    return c


def autofit(ws, headers, min_width=14, max_width=40):
    from openpyxl.utils import get_column_letter
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(header))
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


def _parse_response(res):
    """Parse a Response object, raising a clear error on non-200 or empty
    body. By the time a response gets here it has already been through the
    patched, lock-protected, auto-retrying _call() (for client.list_entities
    etc.) or _robust_get() (for the hand-rolled endpoints below) -- so a
    403 surviving to this point means re-auth was already attempted and
    still failed, not that nothing tried."""
    if not hasattr(res, "status_code"):
        return res  
    if res.status_code == 403:
        raise PermissionError("403 Forbidden — re-authentication was already "
                               "attempted and the session still isn't valid.")
    if res.status_code != 200:
        raise ValueError(f"HTTP {res.status_code}: {res.text[:200]}")
    if not res.text or not res.text.strip():
        raise ValueError(f"Empty response body (HTTP {res.status_code})")
    return res.json()


def _call_with_retry(fn, *args, retries=3, delay=3, **kwargs):
    """Call fn(*args, **kwargs) with retry on transient failures. fn here is
    always a NerisApiClient method, so by the time it returns, the patched
    _call() has already retried auth drops internally -- this layer catches
    anything else (a parse hiccup, a brief network blip)."""
    for attempt in range(1, retries + 1):
        try:
            res = fn(*args, **kwargs)
            return _parse_response(res)
        except PermissionError:
            raise
        except Exception as e:
            if attempt < retries:
                print(f"\n  ⚠ Attempt {attempt} failed ({e}) — retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise


def _token_locked_get(client, url, params):
    """Make sure auth is current and capture the token fingerprint under
    _auth_lock (fast -- _update_auth() is a no-op whenever the token still
    looks valid), then fire the actual GET WITHOUT holding the lock. This
    is for endpoints the library has no official method for (vendor
    enrollments, user_entity_membership below), which bypass
    client._call() entirely and so get none of its protection unless it's
    added here explicitly."""
    with _auth_lock:
        client._update_auth()
        client._session.headers.update(
            {"Authorization": f"Bearer {client.tokens.access_token}"}
        )
        fingerprint = _token_fingerprint(client)
    r = client._session.get(url, params=params)
    return r, fingerprint


def _token_locked_get_retrying(client, url, params, max_attempts=4):
    """_token_locked_get, but retries on a hard exception from the auth
    check itself (an MFA timeout, or NERIS flatly rejecting a refresh/
    login attempt)."""
    last_exc = None
    for attempt in range(1, max_attempts + 1):
        try:
            return _token_locked_get(client, url, params)
        except Exception as e:
            last_exc = e
            if attempt == max_attempts:
                raise
            print(f"  ⚠ Auth error on {url.rsplit('/', 1)[-1]} "
                  f"(attempt {attempt}/{max_attempts}): {e} — retrying")
            time.sleep(min(15, 3 * attempt))
    raise last_exc


def _robust_get(client, url, params, context_label, max_attempts=4):
    """GET a URL directly via client._session (bypassing client._call(),
    the same way the no_activity_report endpoint does in
    neris_monthly_counts.py) with the same lock-protected auth-check +
    forced-reauth retry loop _locked_call gives every other call. Returns
    the parsed JSON dict on success; raises RuntimeError if every attempt
    still looks like an auth failure."""
    r, fingerprint = _token_locked_get_retrying(client, url, params)

    for attempt in range(1, max_attempts + 1):
        if not _looks_like_auth_drop(r):
            break
        if attempt == max_attempts:
            break  

        token_before = client.tokens.access_token
        print(f"  ⚠ HTTP {r.status_code} on {context_label} "
              f"[token {fingerprint}] (attempt {attempt}/{max_attempts}) — "
              f"checking re-auth")
        with _auth_lock:
            need_cooldown = _maybe_force_reauth(client, token_before, context_label)
        if need_cooldown:
            time.sleep(1.5)  

        r, fingerprint = _token_locked_get_retrying(client, url, params)

    if not r.ok:
        raise RuntimeError(
            f"{context_label} failed: HTTP {r.status_code} - "
            f"{r.text[:500]!r} [token used: {fingerprint}]"
        )
    try:
        return r.json()
    except ValueError as e:
        raise RuntimeError(
            f"{context_label} returned non-JSON (HTTP {r.status_code}): "
            f"{r.text[:500]!r} [token used: {fingerprint}]"
        ) from e


def get_entity_departments(client, state_code):
    """Page through all departments for the given state code."""
    departments = []
    page_number = 1
    page_count  = None

    print(f"\nFetching departments for state: {state_code}")
    while True:
        print(f"  Page {page_number}... ", end="", flush=True)
        try:
            res = _call_with_retry(
                client.list_entities,
                page_size=100,
                page_number=page_number,
                state=state_code
            )

            if page_number == 1:
                page_count  = res.get("page_count") or 1
                total_count = res.get("total_count") or 0
                print(f"\n  [info] total_count={total_count}, page_count={page_count}")

            entities = res.get("entities", [])
            if not entities:
                print("empty — done")
                break

            departments.extend(entities)
            print(f"{len(entities)} retrieved (total so far: {len(departments)})")

            if page_number >= page_count:
                print("  All pages retrieved.")
                break

            page_number += 1

        except Exception as e:
            print(f"\n✗ Error on page {page_number}: {e}")
            traceback.print_exc()
            break

    print(f"✓ Total departments fetched: {len(departments)}")
    return departments


def get_vendor_names(client, neris_id):
    """
    Fetch integration_title values for a department via
    GET /account/enrollment/{neris_id} — paginated, routed through
    _robust_get() so a dropped token here gets the same forced-reauth
    retry every other call gets instead of silently giving up.
    Returns a comma-separated string of vendor/integration names.
    """
    base_url = client.config.base_url
    titles   = []
    page     = 1

    while True:
        try:
            data = _robust_get(
                client,
                f"{base_url}/account/enrollment/{neris_id}",
                {"page_size": 100, "page_number": page},
                f"account/enrollment({neris_id})",
            )
        except Exception as e:
            print(f"  ⚠ Could not fetch enrollments for {neris_id}: {e}")
            break

        enrollments = data.get("enrollments", [])
        page_count  = data.get("page_count", 1)

        for e in enrollments:
            title = e.get("integration_title", "")
            if title and title not in titles:
                titles.append(title)

        if page >= page_count:
            break
        page += 1

    return ", ".join(titles)


def get_users_for_entity(client, neris_id):
    """
    Fetch all users and their roles for a department via
    GET /entity/{neris_id}/user_entity_membership — paginated, routed
    through _robust_get() for the same reason as get_vendor_names above.
    Returns a list of row dicts ready for the USERS sheet.
    """
    base_url = client.config.base_url
    rows     = []
    page     = 1

    while True:
        try:
            data = _robust_get(
                client,
                f"{base_url}/entity/{neris_id}/user_entity_membership",
                {"page_size": 100, "page_number": page},
                f"user_entity_membership({neris_id})",
            )
        except Exception as e:
            print(f"  ⚠ Could not fetch users for {neris_id}: {e}")
            break

        users      = data.get("users", [])
        page_count = data.get("page_count", 1)

        for u in users:
            roles      = u.get("roles", [])
            role_names = ", ".join(r.get("name", "") for r in roles if r.get("name"))
            status     = u.get("status", "")
            rows.append({
                "sub":           u.get("sub", ""),
                "given_name":    u.get("given_name", ""),
                "family_name":   u.get("family_name", ""),
                "email":         u.get("email", ""),
                "active":        u.get("active"),
                "logged_in":     "Yes" if status == "CONFIRMED" else "No",
                "dept_neris_id": neris_id,
                "dept_name":     "",
                "role":          role_names,
            })

        if page >= page_count:
            break
        page += 1

    return rows


def get_reporting_dept_ids(client, dept_ids, max_workers=20):
    """
    Check which departments have at least one incident using parallel
    single-record API calls via ThreadPoolExecutor. client.list_incidents()
    goes through the now-patched, lock-protected client._call(), so this
    concurrent loop gets the same auth-check serialization + forced-reauth
    protection as neris_monthly_counts.py's department fetch, instead of
    the unprotected ThreadPoolExecutor this had before.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    reporting = set()
    errors    = 0
    print(f"\nChecking reporting status for {len(dept_ids)} departments...")

    def check_one(neris_id):
        try:
            res = client.list_incidents(neris_id_entity=neris_id, page_size=1)
            res = _unwrap(res, f"list_incidents({neris_id})")
            if not isinstance(res, dict):
                res = res.json()
            return neris_id, len(res.get("incidents", [])) > 0, None
        except Exception as e:
            return neris_id, False, str(e)

    completed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(check_one, d): d for d in dept_ids}
        for future in as_completed(futures):
            neris_id, has, err = future.result()
            if has:
                reporting.add(neris_id)
            if err:
                errors += 1
            completed += 1
            if completed % 50 == 0:
                print(f"  {completed}/{len(dept_ids)} checked...")

    print(f"✓ {len(reporting)} departments have at least one incident")
    if errors:
        print(f"  ⚠ {errors} department(s) could not be checked and were "
              f"counted as 'no incidents' — see any ⚠ lines above for detail")
    return reporting


USER_HEADERS = [
    "NERIS ID",
    "Department Name",
    "First Name",
    "Last Name",
    "Email",
    "Active?",
    "Logged in?",
    "Role",
]

DEPT_HEADERS_BASE = [
    "NERIS ID",
    "Department Name",
    "Department Type",
    "Onboarding Status",
    "Direct Reporting Active?",
    "No Activity Report Active?",
    "Reporting",
    "Vendor / Integration",  # conditionally dropped if all blank
]


def write_users_sheet(ws, user_rows):
    ws.freeze_panes = "A2"
    for col, h in enumerate(USER_HEADERS, start=1):
        _hcell(ws, 1, col, h)
    for r, row in enumerate(user_rows, start=2):
        _dcell(ws, r, 1, row.get("dept_neris_id", ""))
        _dcell(ws, r, 2, row.get("dept_name", ""))
        _dcell(ws, r, 3, row.get("given_name", ""))
        _dcell(ws, r, 4, row.get("family_name", ""))
        _dcell(ws, r, 5, row.get("email", ""))
        _dcell(ws, r, 6, "Yes" if row.get("active") else "No")
        _dcell(ws, r, 7, row.get("logged_in", "No"))
        _dcell(ws, r, 8, row.get("role", ""))
    autofit(ws, USER_HEADERS)


def write_departments_sheet(ws, dept_rows, include_vendor):
    headers = [h for h in DEPT_HEADERS_BASE if h != "Vendor / Integration" or include_vendor]
    ws.freeze_panes = "A2"
    for col, h in enumerate(headers, start=1):
        _hcell(ws, 1, col, h)
    for r, row in enumerate(dept_rows, start=2):
        flags = row.get("feature_flags", {}) or {}
        _dcell(ws, r, 1, row.get("neris_id", ""))
        _dcell(ws, r, 2, row.get("name", ""))
        _dcell(ws, r, 3, row.get("department_type", ""))
        _dcell(ws, r, 4, row.get("onboarding_status", ""))
        _dcell(ws, r, 5, "Yes" if flags.get("allow_ui_incident_creation") else "No")
        _dcell(ws, r, 6, "Yes" if flags.get("allow_ui_no_activity_report_creation") else "No")
        _dcell(ws, r, 7, row.get("_has_incidents", ""))
        if include_vendor:
            _dcell(ws, r, 8, row.get("_vendor_names", ""))
    autofit(ws, headers)


def main():
    ensure_dependencies()

    username, password, state_code = prompt_config()
    log_path = _init_request_log(state_code)
    print(f"\n✓ Logging every request's x-request-id to: {log_path}")

    client = authenticate(username, password)

    departments = get_entity_departments(client, state_code)

    if not departments:
        print("\n⚠ No departments found.")
        if _request_log_file:
            _request_log_file.close()
        return

    print(f"\nFetching details for {len(departments)} departments...")
    dept_rows = []
    user_rows = []

    all_dept_ids = [
        d.get("neris_id") or d.get("id", "") for d in departments
    ]

    _progress = {"completed": 0, "total": len(departments)}
    _stall_watch_stop = threading.Event()
    _stall_watch_thread = threading.Thread(
        target=_watch_for_stalls,
        args=(_stall_watch_stop, lambda: (_progress["completed"], _progress["total"])),
        daemon=True,
    )
    _stall_watch_thread.start()

    try:
        reporting_dept_ids = get_reporting_dept_ids(client, all_dept_ids)

        for i, dept in enumerate(departments, start=1):
            neris_id  = dept.get("neris_id") or dept.get("id", "")
            dept_name = dept.get("name", "")
            print(f"  [{i}/{len(departments)}] {dept_name} ({neris_id})")

            try:
                detail = _unwrap(client.get_entity(neris_id), f"get_entity({neris_id})")
                if not isinstance(detail, dict):
                    detail = dept
            except Exception as e:
                print(f"  ⚠ Could not fetch entity detail for {neris_id}: {e}")
                detail = dept

            detail["_vendor_names"]  = get_vendor_names(client, neris_id)
            detail["_has_incidents"] = "Yes" if neris_id in reporting_dept_ids else "No"
            dept_rows.append(detail)

            rows = get_users_for_entity(client, neris_id)
            for row in rows:
                row["dept_name"] = dept_name
            user_rows.extend(rows)

            _progress["completed"] = i
    finally:
        _stall_watch_stop.set()

    print(f"\n✓ {len(dept_rows)} departments | {len(user_rows)} users collected")

    include_vendor = any(d.get("_vendor_names", "") for d in dept_rows)

    from openpyxl import Workbook
    wb       = Workbook()
    ws_users = wb.active
    ws_users.title = "USERS"
    write_users_sheet(ws_users, user_rows)

    ws_depts = wb.create_sheet("DEPARTMENTS")
    write_departments_sheet(ws_depts, dept_rows, include_vendor)

    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"Departments and Users {state_code} {date_str}.xlsx"

    wb.save(filename)
    print(f"\n✓ Report saved: {filename}")

    if _request_log_path:
        print(f"\n  Request trace log : {_request_log_path}")
        print("  (every call's x-request-id, status, department, and attempt —")
        print("   send this to NERIS's team along with failing timestamps)")
        if _request_log_file:
            _request_log_file.close()

    print("\n" + "=" * 60)
    print("✓ PROCESS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
