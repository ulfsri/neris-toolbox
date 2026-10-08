#!/usr/bin/env python3
"""
NERIS Monthly Incident Count Report
------------------------------------
Fetches all incidents and no-activity reports for every department in a
state and exports a pivot table (departments x months) to Excel.
"""

import sys
import subprocess
import os
import csv
import threading
import time
from datetime import datetime, timezone, timedelta
from urllib.parse import urlsplit, parse_qs
from collections import defaultdict, deque
from concurrent.futures import ThreadPoolExecutor, as_completed



def ensure_dependencies():
    def pip_install(*packages):
        subprocess.run(
            [sys.executable, "-m", "pip", "install", *packages, "--quiet"],
            check=True
        )

    try:
        import pandas   
        import openpyxl 
        import requests 
    except ImportError:
        print("Installing pandas, openpyxl, and requests...")
        pip_install("pandas", "openpyxl", "requests")
        print("✓ pandas, openpyxl, and requests installed")

    try:
        from neris_api_client import NerisApiClient  
        print("✓ Dependencies ready")
    except ImportError:
        print("Installing neris-api-client...")
        pip_install("neris-api-client")
        print("✓ neris-api-client installed")


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Monthly Incident Count Report")
    print("=" * 60)
    print("\nNote: This report covers all data from 1 Jan 2025 onward.")
    print("      NAR = No-Activity Report (confirmed zero incidents).")
    print("      Blank = no data of either kind submitted.\n")

    print("-- Credentials --")
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()

    print("\n-- Query Parameters --")
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()
    entity_id  = input("NERIS Entity ID (optional — leave blank for all departments): ").strip() or None

    if not username or not password or not state_code:
        sys.exit("✗ Email, password, and state code are required.")

    return username, password, state_code, entity_id


_auth_lock = threading.Lock()
_orig_call = None  

_last_reauth_at = {"t": 0.0}

_mfa_wait_state = {"waiting": False, "since": 0.0}

_MFA_INPUT_TIMEOUT_SECONDS = 240


def _input_with_timeout(prompt, timeout_seconds):
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
    import requests
    if not isinstance(res, requests.Response):
        return False
    return res.status_code in (401, 403)


def _log_rate_limit_headers(resp):
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
        return  
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
                need_cooldown = _maybe_force_reauth(self, token_before, f"incident/{args[:2] if args else kwargs.get('path', '?')}")
            if need_cooldown:
                time.sleep(1.5)  
        return res

    _NerisApiClient._call = _locked_call

    orig_session_request = _NerisApiClient._session.request

    def _wait_for_rate_budget():
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
    import requests
    if isinstance(res, requests.Response):
        raise RuntimeError(
            f"{context} failed: HTTP {res.status_code} - {res.text[:500]!r}"
        )
    return res


def _token_fingerprint(client):
    tok = getattr(client, "tokens", None)
    if not tok or not getattr(tok, "access_token", None):
        return "<no token>"
    return f"...{tok.access_token[-8:]} (expires_at={tok.expires_at})"


def authenticate(username, password):
    os.environ["NERIS_BASE_URL"]   = "https://api.neris.fsri.org/v1"
    os.environ["NERIS_GRANT_TYPE"] = "password"
    os.environ["NERIS_USERNAME"]   = username
    os.environ["NERIS_PASSWORD"]   = password

    from neris_api_client import NerisApiClient
    _install_thread_safety_patch()

    print("\nConnecting to NERIS API...")
    client = NerisApiClient()

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

REPORT_START = datetime(2025, 1, 1, tzinfo=timezone.utc)
REPORT_END = datetime.now(timezone.utc)


def fetch_all_entities(client, state_code, neris_id_entity=None, page_size=100):
    if neris_id_entity:
        try:
            entity = _unwrap(client.get_entity(neris_id_entity), f"get_entity({neris_id_entity})")
            name   = entity.get("name", "") if isinstance(entity, dict) else ""
            print(f"✓ Single entity lookup: {neris_id_entity} → {name}")
            return {neris_id_entity: name}
        except Exception as e:
            print(f"⚠ Could not fetch entity {neris_id_entity}: {e}")
            return {neris_id_entity: ""}

    all_entities = {}
    page_number  = 1

    print(f"\nFetching all registered entities for state: {state_code}")
    while True:
        print(f"  Page {page_number}... ", end="", flush=True)
        try:
            res = _unwrap(
                client.list_entities(state=state_code, page_size=page_size,
                                      page_number=page_number),
                f"list_entities(page={page_number})",
            )
            if not isinstance(res, dict):
                res = res.json()
        except Exception as e:
            print(f"\n⚠ list_entities failed (page {page_number}): {e}")
            break

        batch = res.get("entities", [])
        if not batch:
            print("empty — done.")
            break

        for ent in batch:
            eid  = ent.get("neris_id", "")
            name = ent.get("name", "")
            if eid:
                all_entities[eid] = name

        print(f"retrieved {len(batch)} (total so far: {len(all_entities)})")

        total_count = res.get("total_count", 0)
        if len(all_entities) >= total_count or len(batch) < page_size:
            print("  ✓ All pages retrieved")
            break

        page_number += 1

    print(f"\n{'='*50}")
    print(f"Total registered departments: {len(all_entities)}")
    print(f"{'='*50}")
    return all_entities


def _month_label(dt_or_str):
    """Convert a datetime or 'MM/YYYY' / ISO string to 'Mon-YYYY'. Returns None on failure."""
    if dt_or_str is None:
        return None
    if isinstance(dt_or_str, str):
        if "/" in dt_or_str and len(dt_or_str) <= 7:
            try:
                m, y = dt_or_str.split("/")
                return datetime(int(y), int(m), 1).strftime("%b-%Y")
            except Exception:
                pass
        try:
            dt_or_str = datetime.fromisoformat(dt_or_str.replace("Z", "+00:00"))
        except Exception:
            return None
    try:
        return dt_or_str.strftime("%b-%Y")
    except Exception:
        return None


def fetch_incidents_for_entity(client, state_code, neris_id_entity, page_size=100):
    """
    Fetch incidents for a single department within REPORT_START..REPORT_END.
    Returns a dict: { 'Mon-YYYY': count }
    """
    month_counts = defaultdict(int)
    next_cursor  = None

    while True:
        kwargs = dict(
            state=state_code,
            neris_id_entity=neris_id_entity,
            page_size=page_size,
            call_create_start=REPORT_START,
            call_create_end=REPORT_END,
        )
        if next_cursor:
            kwargs["cursor"] = next_cursor

        res = _unwrap(client.list_incidents(**kwargs), f"list_incidents({neris_id_entity})")
        if not isinstance(res, dict):
            res = res.json()

        batch = res.get("incidents", [])
        if not batch:
            break

        for inc in batch:
            disp = inc.get("dispatch") or {}
            ts   = disp.get("call_create") or disp.get("call_create_start")
            lbl  = _month_label(ts)
            if lbl:
                month_counts[lbl] += 1

        next_cursor = res.get("next_cursor")
        if not next_cursor:
            break

    return dict(month_counts)


def _token_locked_get(client, url, params):
    with _auth_lock:
        client._update_auth()
        client._session.headers.update(
            {"Authorization": f"Bearer {client.tokens.access_token}"}
        )
        fingerprint = _token_fingerprint(client)
    r = client._session.get(url, params=params)
    return r, fingerprint


def _token_locked_get_retrying(client, url, params, max_attempts=4):
    last_exc = None
    for attempt in range(1, max_attempts + 1):
        try:
            return _token_locked_get(client, url, params)
        except Exception as e:
            last_exc = e
            if attempt == max_attempts:
                raise
            eid = params.get("neris_id_entity", "?")
            print(f"  ⚠ Auth error on NAR check for {eid} "
                  f"(attempt {attempt}/{max_attempts}): {e} — retrying")
            time.sleep(min(15, 3 * attempt))
    raise last_exc


def fetch_nars_for_entity(client, state_code, neris_id_entity, page_size=100):
    """
    Fetch no-activity reports for a single department within
    REPORT_START..REPORT_END. Returns a set of 'Mon-YYYY' labels.
    """
    nar_months  = set()
    next_cursor = None
    base_url    = os.environ.get("NERIS_BASE_URL", "https://api.neris.fsri.org/v1")
    url = f"{base_url}/no_activity_report"

    while True:
        params = dict(
            state=state_code,
            neris_id_entity=neris_id_entity,
            page_size=page_size,
            call_create_start=REPORT_START.isoformat(),
            call_create_end=REPORT_END.isoformat(),
        )
        if next_cursor:
            params["cursor"] = next_cursor

        r, fingerprint = _token_locked_get_retrying(client, url, params)

        max_attempts = 4
        for attempt in range(1, max_attempts + 1):
            if not _looks_like_auth_drop(r):
                break
            if attempt == max_attempts:
                break  

            token_before = client.tokens.access_token
            print(f"  ⚠ HTTP {r.status_code} on no_activity_report({neris_id_entity}) "
                  f"[token {fingerprint}] (attempt {attempt}/{max_attempts}) — "
                  f"checking re-auth")
            with _auth_lock:
                need_cooldown = _maybe_force_reauth(client, token_before, f"NAR/{neris_id_entity}")
            if need_cooldown:
                time.sleep(1.5)  

            r, fingerprint = _token_locked_get_retrying(client, url, params)

        if not r.ok:
            raise RuntimeError(
                f"no_activity_report({neris_id_entity}) failed: "
                f"HTTP {r.status_code} - {r.text[:500]!r} [token used: {fingerprint}]"
            )

        try:
            res = r.json()
        except ValueError as e:
            raise RuntimeError(
                f"no_activity_report({neris_id_entity}) returned non-JSON "
                f"(HTTP {r.status_code}): {r.text[:500]!r} [token used: {fingerprint}]"
            ) from e

        batch = res.get("reports", [])
        if not batch:
            break

        for report in batch:
            lbl = _month_label(report.get("month_year", ""))
            if lbl:
                nar_months.add(lbl)

        next_cursor = res.get("next_cursor")
        if not next_cursor:
            break

    return nar_months


def generate_month_columns():
    """All month labels from Jan-2025 through the current month."""
    cols, d = [], datetime(2025, 1, 1)
    now = datetime.now()
    while d <= now:
        cols.append(d.strftime("%b-%Y"))
        m = d.month + 1
        d = datetime(d.year + (m // 13), ((m - 1) % 12) + 1, 1)
    return cols


def export_to_excel(df_pivot, month_cols, state_code, inc_counts, nar_flags):
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename  = f"neris_monthly_counts_{state_code}_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Monthly Incident Counts"

    header_cols = ["NERIS Entity ID", "Department Name"] + month_cols
    ws.append(header_cols)

    hdr_fill     = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    nar_fill     = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_fill   = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    no_data_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    hdr_font     = Font(color="FFFFFF", bold=True, size=11)
    total_font   = Font(bold=True, size=11)
    hdr_aln      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ctr_aln      = Alignment(horizontal="center", vertical="center")
    left_aln     = Alignment(horizontal="left",   vertical="center")
    thin_side    = Side(style="thin", color="BFBFBF")
    thin_border  = Border(left=thin_side, right=thin_side,
                          top=thin_side,  bottom=thin_side)

    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = hdr_aln

    total_row_idx = len(df_pivot) + 1   
    rows = df_pivot.values.tolist()
    for r_idx, row_data in enumerate(rows, start=2):
        is_total = (r_idx == total_row_idx)
        for c_idx, val in enumerate(row_data, start=1):
            cell        = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border

            if is_total:
                cell.fill      = total_fill
                cell.font      = total_font
                cell.alignment = ctr_aln
            elif c_idx <= 2:
                cell.alignment = left_aln
            elif val == "NAR":
                cell.fill      = nar_fill
                cell.font      = Font(italic=True, color="7F6000")
                cell.alignment = ctr_aln
            elif val == "":
                cell.fill      = no_data_fill
                cell.alignment = ctr_aln
            else:
                cell.alignment = ctr_aln

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 35
    for i in range(3, len(header_cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11

    ws.freeze_panes = "C2"

    # Legend sheet
    ws_legend  = wb.create_sheet("Legend")
    legend_data = [
        ("Symbol",       "Meaning"),
        ("(number)",     "Count of incidents reported for that department in that month"),
        ("NAR",          "No-Activity Report filed — department confirmed zero incidents"),
        ("(grey blank)", "No incidents and no no-activity report submitted for that month"),
        ("TOTAL row",    "Sum of numeric incident counts across all departments per month"),
    ]
    for lr in legend_data:
        ws_legend.append(lr)
    for cell in ws_legend[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
    ws_legend.column_dimensions["A"].width = 14
    ws_legend.column_dimensions["B"].width = 75

    wb.save(filename)
    return filename

def main():
    ensure_dependencies()

    import pandas as pd

    username, password, state_code, entity_id = prompt_config()

    log_path = _init_request_log(state_code)
    print(f"\n✓ Logging every request's x-request-id to: {log_path}")

    client = authenticate(username, password)

    all_entities = fetch_all_entities(client, state_code, entity_id)
    all_eids     = list(all_entities.keys())

    print(f"\nFetching incidents and NARs for {len(all_eids)} departments...")
    print("(progress updates every 50 departments, more often near the end;")
    print(" a '⏳ still working' line means it's alive but slow, not stuck)\n")

    inc_counts = {}   
    nar_flags  = {}   
    completed  = 0
    errors     = []

    _stall_watch_stop = threading.Event()
    _stall_watch_thread = threading.Thread(
        target=_watch_for_stalls,
        args=(_stall_watch_stop, lambda: (completed, len(all_eids))),
        daemon=True,
    )
    _stall_watch_thread.start()

    _first_dept_done = {"done": False}

    def fetch_dept_data(eid):
        counts = fetch_incidents_for_entity(client, state_code, eid)
        nars   = fetch_nars_for_entity(client, state_code, eid)
        if not _first_dept_done["done"]:
            _first_dept_done["done"] = True
            print(f"\n  [Diagnostic] First dept {eid}: "
                  f"{sum(counts.values())} incidents, {len(nars)} NAR months")
        return eid, counts, nars

    try:
        with ThreadPoolExecutor(max_workers=20) as executor:
            futures = {executor.submit(fetch_dept_data, eid): eid for eid in all_eids}

            first_errors_shown = 0
            for future in as_completed(futures):
                eid = futures[future]
                try:
                    eid, counts, nars = future.result()
                    inc_counts[eid]   = counts
                    nar_flags[eid]    = nars
                except Exception as e:
                    errors.append((eid, str(e)))
                    inc_counts[eid] = {}
                    nar_flags[eid]  = set()
                    if first_errors_shown < 3:
                        print(f"\n  ⚠ ERROR for {eid}: {e}")
                        first_errors_shown += 1

                completed += 1
                remaining = len(all_eids) - completed
                if completed % 50 == 0 or completed == len(all_eids) or remaining < 20:
                    total_incidents = sum(sum(v.values()) for v in inc_counts.values())
                    total_nars      = sum(len(v) for v in nar_flags.values())
                    print(f"  {completed}/{len(all_eids)} departments complete "
                          f"| {total_incidents} incidents | {total_nars} NAR months")
    finally:
        _stall_watch_stop.set()

    if errors:
        print(f"\n⚠ {len(errors)} department(s) had fetch errors "
              f"(these will show as BLANK in the export, not silently as 0):")
        for eid, err in errors[:10]:
            print(f"    {eid}: {err}")
        if len(errors) > 10:
            print(f"    ... and {len(errors) - 10} more")

    total_incidents = sum(sum(v.values()) for v in inc_counts.values())
    total_nars      = sum(len(v) for v in nar_flags.values())
    print(f"\n{'='*50}")
    print(f"Total incidents retrieved : {total_incidents}")
    print(f"Total NAR months on record: {total_nars}")
    print(f"{'='*50}")

    month_cols = generate_month_columns()
    rows = []

    failed_eids = {eid for eid, _ in errors}

    for eid in sorted(all_eids):
        row = {
            "NERIS Entity ID": eid,
            "Department Name": all_entities.get(eid, ""),
        }
        dept_counts = inc_counts.get(eid, {})
        dept_nars   = nar_flags.get(eid, set())
        dept_failed = eid in failed_eids

        for mc in month_cols:
            if dept_failed:
                row[mc] = "ERROR"
            elif mc in dept_counts:
                row[mc] = dept_counts[mc]
            elif mc in dept_nars:
                row[mc] = "NAR"
            else:
                row[mc] = ""
        rows.append(row)

    df_pivot = pd.DataFrame(rows)

    summary = {"NERIS Entity ID": "TOTAL", "Department Name": ""}
    for mc in month_cols:
        nums = pd.to_numeric(df_pivot[mc], errors="coerce").dropna()
        summary[mc] = int(nums.sum()) if not nums.empty else ""
    df_pivot = pd.concat([df_pivot, pd.DataFrame([summary])], ignore_index=True)

    print(f"\nPivot table: {len(df_pivot)-1} departments x {len(month_cols)} months")

    filename = export_to_excel(df_pivot, month_cols, state_code, inc_counts, nar_flags)

    print(f"\n✓ Exported: {filename}")
    print(f"  Departments : {len(df_pivot)-1}")
    print(f"  Months      : {len(month_cols)}  ({month_cols[0]} - {month_cols[-1]})")
    print(f"  Incidents   : {total_incidents}")
    print(f"  NAR months  : {total_nars}")
    if errors:
        print(f"  Failed      : {len(errors)} department(s) — see ERROR cells in the export")
    print("\nKey:")
    print("  (number)     = incident count")
    print("  NAR          = no-activity report filed (confirmed 0 incidents)")
    print("  (grey blank) = no data submitted for that month")
    print("  ERROR        = fetch failed for this department — data is INCOMPLETE, not zero")

    if _request_log_path:
        print(f"\n  Request trace log : {_request_log_path}")
        print("  (every call's x-request-id, status, department, and attempt —")
        print("   send this to NERIS's team along with failing timestamps)")
        if _request_log_file:
            _request_log_file.close()

    print("\n" + "=" * 60)
    print("  PROCESS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
