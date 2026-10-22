"""
NERIS Incident Export Tool — States .py
------------------------------------
Fetches incidents for all departments in a state from the NERIS API and
exports them to a formatted Excel file with Overview, Raw Data, and (if
applicable) Casualty & Rescue sheets.

Note that if a column has no data throughout the entire dataset, it will be dropped from the Excel file output.

RUNNING IN VS CODE: the script pauses for you to type your NERIS email,
password and the emailed MFA code, so it must run in the TERMINAL panel:
  - Use "Run Python File in Terminal" (the play button's default), or
  - When debugging (F5), make sure launch.json has "console": "integratedTerminal".
The Code Runner extension's OUTPUT tab and the DEBUG CONSOLE can't accept
typed input, so the script can't pause for the code there.
"""

import sys
import subprocess
import os
import csv
import json
import threading
import time
import traceback
from datetime import datetime, timezone, timedelta
from urllib.parse import urlsplit, parse_qs
from collections import deque

pd = None


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
        from neris_api_client import NerisApiClient  # noqa
        print("✓ Dependencies ready")
    except ImportError:
        print("Installing neris-api-client...")
        pip_install(
            "https://github.com/ulfsri/neris-api-client/archive/refs/heads/main.zip"
        )
        print("✓ neris-api-client installed")


INCIDENT_TYPE_MAP = {
    "FIRE – Outside Fire": [
        "FIRE||OUTSIDE_FIRE||CONSTRUCTION_WASTE",
        "FIRE||OUTSIDE_FIRE||DUMPSTER_OUTDOOR_CONTAINER_FIRE",
        "FIRE||OUTSIDE_FIRE||OTHER_OUTSIDE_FIRE",
        "FIRE||OUTSIDE_FIRE||OUTSIDE_TANK_FIRE",
        "FIRE||OUTSIDE_FIRE||TRASH_RUBBISH_FIRE",
        "FIRE||OUTSIDE_FIRE||UTILITY_INFRASTRUCTURE_FIRE",
        "FIRE||OUTSIDE_FIRE||VEGETATION_GRASS_FIRE",
        "FIRE||OUTSIDE_FIRE||WILDFIRE_URBAN_INTERFACE",
        "FIRE||OUTSIDE_FIRE||WILDFIRE_WILDLAND",
    ],
    "FIRE – Special Fire": [
        "FIRE||SPECIAL_FIRE||ESS_FIRE",
        "FIRE||SPECIAL_FIRE||EXPLOSION",
        "FIRE||SPECIAL_FIRE||INFRASTRUCTURE_FIRE",
    ],
    "FIRE – Structure Fire": [
        "FIRE||STRUCTURE_FIRE||CHIMNEY_FIRE",
        "FIRE||STRUCTURE_FIRE||CONFINED_COOKING_APPLIANCE_FIRE",
        "FIRE||STRUCTURE_FIRE||ROOM_AND_CONTENTS_FIRE",
        "FIRE||STRUCTURE_FIRE||STRUCTURAL_INVOLVEMENT_FIRE",
    ],
    "FIRE – Transportation Fire": [
        "FIRE||TRANSPORTATION_FIRE||AIRCRAFT_FIRE",
        "FIRE||TRANSPORTATION_FIRE||BOAT_PERSONAL_WATERCRAFT_BARGE_FIRE",
        "FIRE||TRANSPORTATION_FIRE||POWERED_MOBILITY_DEVICE_FIRE",
        "FIRE||TRANSPORTATION_FIRE||TRAIN_RAIL_FIRE",
        "FIRE||TRANSPORTATION_FIRE||VEHICLE_FIRE_COMMERCIAL",
        "FIRE||TRANSPORTATION_FIRE||VEHICLE_FIRE_FOOD_TRUCK",
        "FIRE||TRANSPORTATION_FIRE||VEHICLE_FIRE_PASSENGER",
        "FIRE||TRANSPORTATION_FIRE||VEHICLE_FIRE_RV",
    ],
    "HAZSIT – Hazardous Materials": [
        "HAZSIT||HAZARDOUS_MATERIALS||BIOLOGICAL_RELEASE_INCIDENT",
        "HAZSIT||HAZARDOUS_MATERIALS||CARBON_MONOXIDE_RELEASE",
        "HAZSIT||HAZARDOUS_MATERIALS||FUEL_SPILL_ODOR",
        "HAZSIT||HAZARDOUS_MATERIALS||GAS_LEAK_ODOR",
        "HAZSIT||HAZARDOUS_MATERIALS||HAZMAT_RELEASE_FACILITY",
        "HAZSIT||HAZARDOUS_MATERIALS||HAZMAT_RELEASE_TRANSPORT",
        "HAZSIT||HAZARDOUS_MATERIALS||RADIOACTIVE_RELEASE_INCIDENT",
    ],
    "HAZSIT – Non-Chemical Hazard": [
        "HAZSIT||HAZARD_NONCHEM||BOMB_THREAT_RESPONSE_SUSPICIOUS_PACKAGE",
        "HAZSIT||HAZARD_NONCHEM||ELEC_HAZARD_SHORT_CIRCUIT",
        "HAZSIT||HAZARD_NONCHEM||ELEC_POWER_LINE_DOWN_ARCHING_MALFUNC",
        "HAZSIT||HAZARD_NONCHEM||MOTOR_VEHICLE_COLLISION",
    ],
    "HAZSIT – Investigation": [
        "HAZSIT||INVESTIGATION||ODOR",
        "HAZSIT||INVESTIGATION||SMOKE_INVESTIGATION",
    ],
    "HAZSIT – Overpressure": [
        "HAZSIT||OVERPRESSURE||NO_RUPTURE",
        "HAZSIT||OVERPRESSURE||RUPTURE_WITHOUT_FIRE",
    ],
    "LAWENFORCE": [
        "LAWENFORCE",
    ],
    "MEDICAL – Illness": [
        "MEDICAL||ILLNESS",
        "MEDICAL||ILLNESS||ABDOMINAL_PAIN",
        "MEDICAL||ILLNESS||ALLERGIC_REACTION_STINGS",
        "MEDICAL||ILLNESS||ALTERED_MENTAL_STATUS",
        "MEDICAL||ILLNESS||BACK_PAIN_NON_TRAUMA",
        "MEDICAL||ILLNESS||BREATHING_PROBLEMS",
        "MEDICAL||ILLNESS||CARDIAC_ARREST",
        "MEDICAL||ILLNESS||CHEST_PAIN_NON_TRAUMA",
        "MEDICAL||ILLNESS||CONVULSIONS_SEIZURES",
        "MEDICAL||ILLNESS||DIABETIC_PROBLEMS",
        "MEDICAL||ILLNESS||HEADACHE",
        "MEDICAL||ILLNESS||HEART_PROBLEMS",
        "MEDICAL||ILLNESS||NAUSEA_VOMITING",
        "MEDICAL||ILLNESS||NO_APPROPRIATE_CHOICE",
        "MEDICAL||ILLNESS||OVERDOSE",
        "MEDICAL||ILLNESS||PANDEMIC_EPIDEMIC_OUTBREAK",
        "MEDICAL||ILLNESS||PREGNANCY_CHILDBIRTH",
        "MEDICAL||ILLNESS||PSYCHOLOGICAL_BEHAVIOR_ISSUES",
        "MEDICAL||ILLNESS||SICK_CASE",
        "MEDICAL||ILLNESS||STROKE_CVA",
        "MEDICAL||ILLNESS||UNCONSCIOUS_VICTIM",
        "MEDICAL||ILLNESS||UNKNOWN_PROBLEM",
        "MEDICAL||ILLNESS||WELL_PERSON_CHECK",
    ],
    "MEDICAL – Injury": [
        "MEDICAL||INJURY",
        "MEDICAL||INJURY||ANIMAL_BITES",
        "MEDICAL||INJURY||ASSAULT",
        "MEDICAL||INJURY||BURNS_EXPLOSION",
        "MEDICAL||INJURY||CARBON_MONOXIDE_OTHER_INHALATION_INJURY",
        "MEDICAL||INJURY||CHOKING",
        "MEDICAL||INJURY||DROWNING_DIVING_SCUBA_ACCIDENT",
        "MEDICAL||INJURY||ELECTROCUTION",
        "MEDICAL||INJURY||EYE_TRAUMA",
        "MEDICAL||INJURY||FALL",
        "MEDICAL||INJURY||GUNSHOT_WOUND",
        "MEDICAL||INJURY||HEAT_COLD_EXPOSURE",
        "MEDICAL||INJURY||HEMORRHAGE_LACERATION",
        "MEDICAL||INJURY||INDUSTRIAL_INACCESSIBLE_ENTRAPMENT",
        "MEDICAL||INJURY||MOTOR_VEHICLE_COLLISION",
        "MEDICAL||INJURY||OTHER_TRAUMATIC_INJURY",
        "MEDICAL||INJURY||POISONING",
        "MEDICAL||INJURY||STAB_PENETRATING_TRAUMA",
    ],
    "MEDICAL – Other": [
        "MEDICAL||OTHER||AIRMEDICAL_TRANSPORT",
        "MEDICAL||OTHER||COMMUNITY_PUBLIC_HEALTH",
        "MEDICAL||OTHER||HEALTHCARE_PROFESSIONAL_ADMISSION",
        "MEDICAL||OTHER||INTERCEPT_OTHER_UNIT",
        "MEDICAL||OTHER||MEDICAL_ALARM",
        "MEDICAL||OTHER||STANDBY_REQUEST",
        "MEDICAL||OTHER||TRANSFER_INTERFACILITY",
    ],
    "NOEMERG – Cancelled": [
        "NOEMERG||CANCELLED",
    ],
    "NOEMERG – False Alarm": [
        "NOEMERG||FALSE_ALARM||ACCIDENTAL_ALARM",
        "NOEMERG||FALSE_ALARM||BOMB_SCARE",
        "NOEMERG||FALSE_ALARM||INTENTIONAL_FALSE_ALARM",
        "NOEMERG||FALSE_ALARM||MALFUNCTIONING_ALARM",
        "NOEMERG||FALSE_ALARM||OTHER_FALSE_CALL",
    ],
    "NOEMERG – Good Intent": [
        "NOEMERG||GOOD_INTENT||CONTROLLED_BURNING_AUTHORIZED",
        "NOEMERG||GOOD_INTENT||INVESTIGATE_HAZARDOUS_RELEASE",
        "NOEMERG||GOOD_INTENT||NO_INCIDENT_FOUND_LOCATION_ERROR",
        "NOEMERG||GOOD_INTENT||SMOKE_FROM_NONHOSTILE_SOURCE",
    ],
    "PUBSERV – Alarms (Non-Medical)": [
        "PUBSERV||ALARMS_NONMED||CO_ALARM",
        "PUBSERV||ALARMS_NONMED||FIRE_ALARM",
        "PUBSERV||ALARMS_NONMED||GAS_ALARM",
        "PUBSERV||ALARMS_NONMED||OTHER_ALARM",
    ],
    "PUBSERV – Citizen Assist": [
        "PUBSERV||CITIZEN_ASSIST||CITIZEN_ASSIST_SERVICE_CALL",
        "PUBSERV||CITIZEN_ASSIST||LIFT_ASSIST",
        "PUBSERV||CITIZEN_ASSIST||LOST_PERSON",
        "PUBSERV||CITIZEN_ASSIST||PERSON_IN_DISTRESS",
    ],
    "PUBSERV – Disaster / Weather": [
        "PUBSERV||DISASTER_WEATHER||DAMAGE_ASSESSMENT",
        "PUBSERV||DISASTER_WEATHER||WEATHER_RESPONSE",
    ],
    "PUBSERV – Other": [
        "PUBSERV||OTHER||DAMAGED_HYDRANT",
        "PUBSERV||OTHER||MOVE_UP",
        "PUBSERV||OTHER||STANDBY",
    ],
    "RESCUE – Outside": [
        "RESCUE||OUTSIDE||BACKCOUNTRY_RESCUE",
        "RESCUE||OUTSIDE||CONFINED_SPACE_RESCUE",
        "RESCUE||OUTSIDE||EXTRICATION_ENTRAPPED",
        "RESCUE||OUTSIDE||HIGH_ANGLE_RESCUE",
        "RESCUE||OUTSIDE||LIMITED_NO_ACCESS",
        "RESCUE||OUTSIDE||LOW_ANGLE_RESCUE",
        "RESCUE||OUTSIDE||STEEP_ANGLE_RESCUE",
        "RESCUE||OUTSIDE||TRENCH",
    ],
    "RESCUE – Structure": [
        "RESCUE||STRUCTURE||BUILDING_STRUCTURE_COLLAPSE",
        "RESCUE||STRUCTURE||CONFINED_SPACE_RESCUE",
        "RESCUE||STRUCTURE||ELEVATOR_ESCALATOR_RESCUE",
        "RESCUE||STRUCTURE||EXTRICATION_ENTRAPPED",
    ],
    "RESCUE – Transportation": [
        "RESCUE||TRANSPORTATION||AVIATION_COLLISION_CRASH",
        "RESCUE||TRANSPORTATION||AVIATION_STANDBY",
        "RESCUE||TRANSPORTATION||MOTOR_VEHICLE_EXTRICATION_ENTRAPPED",
        "RESCUE||TRANSPORTATION||TRAIN_RAIL_COLLISION_DERAILMENT",
    ],
    "RESCUE – Water": [
        "RESCUE||WATER||PERSON_IN_WATER_STANDING",
        "RESCUE||WATER||PERSON_IN_WATER_SWIFTWATER",
        "RESCUE||WATER||WATERCRAFT_IN_DISTRESS",
    ],
}


DATE_RANGE_OPTIONS = [
    'All Records', 'Occurred Today', 'Occurred Yesterday',
    'Occurred Last 7 Days', 'Occurred Last 30 Days', 'Occurred in 2025',
    'Custom Date Range',
]


_NO_INPUT_MESSAGE = (
    "\n✗ This program can't read what you type, so it can't pause for your "
    "login details or MFA code.\n"
    "  In VS Code, run it in the TERMINAL panel instead:\n"
    "    • Use 'Run Python File in Terminal' (not Code Runner's OUTPUT tab), or\n"
    "    • When debugging (F5), set \"console\": \"integratedTerminal\" in launch.json\n"
    "      (not \"internalConsole\", which is the Debug Console)."
)


def _parse_incident_type_selection(raw):
    cats  = list(INCIDENT_TYPE_MAP.keys())
    valid = {t for types in INCIDENT_TYPE_MAP.values() for t in types}
    selected = []
    for tok in [t.strip() for t in raw.split(",") if t.strip()]:
        if tok.isdigit() and 1 <= int(tok) <= len(cats):
            for t in INCIDENT_TYPE_MAP[cats[int(tok) - 1]]:
                if t not in selected:
                    selected.append(t)
        elif tok.upper() in valid:
            if tok.upper() not in selected:
                selected.append(tok.upper())
        else:
            print(f"  ⚠ Ignoring unrecognized category/type: {tok!r}")
    return selected or None


def _prompt_date(label, default):
    while True:
        raw = input(f"{label} (YYYY-MM-DD) [{default.isoformat()}]: ").strip()
        if not raw:
            return default
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            print("  ✗ Use the format YYYY-MM-DD, e.g. 2025-03-01")


def prompt_config():
    print("\n" + "=" * 60)
    print("  NERIS Incident Export")
    print("=" * 60)

    print("\n-- Credentials --")
    username = input("NERIS Email: ").strip()
    print("NERIS Password (note: characters will be visible):")
    password = input("> ").strip()

    print("\n-- Query Parameters --")
    state_code = input("State Code (e.g. VA, MI, CA): ").strip().upper()
    entity_id  = input("NERIS Entity ID (optional — leave blank for all departments): ").strip() or None

    if not username or not password or not state_code:
        sys.exit("✗ Email, password, and state code are required.")

    print("\n-- Incident Type Filter (optional) --")
    for i, cat in enumerate(INCIDENT_TYPE_MAP, 1):
        print(f"  {i:>2}. {cat} ({len(INCIDENT_TYPE_MAP[cat])} types)")
    raw_types = input("Category numbers and/or exact type strings, comma-separated "
                      "(blank = all types): ").strip()
    incident_types = _parse_incident_type_selection(raw_types) if raw_types else None

    print("\n-- Date Range (based on when incidents occurred — dispatch call_create) --")
    for i, opt in enumerate(DATE_RANGE_OPTIONS, 1):
        print(f"  {i}. {opt}")
    while True:
        choice = input(f"Choose 1-{len(DATE_RANGE_OPTIONS)} [1]: ").strip() or "1"
        if choice.isdigit() and 1 <= int(choice) <= len(DATE_RANGE_OPTIONS):
            range_type = DATE_RANGE_OPTIONS[int(choice) - 1]
            break
        print("  ✗ Enter one of the numbers above.")

    start_date = end_date = None
    if range_type == 'Custom Date Range':
        today = datetime.now().date()
        while True:
            start_date = _prompt_date("Start Date", today - timedelta(days=365))
            end_date   = _prompt_date("End Date", today)
            if start_date <= end_date:
                break
            print("  ✗ Start Date must be on or before End Date.")

    return {
        "username": username, "password": password,
        "state_code": state_code, "entity_id": entity_id,
        "incident_types": incident_types, "range_type": range_type,
        "start_date": start_date, "end_date": end_date,
    }

_auth_lock = threading.Lock()
_orig_call = None
_last_reauth_at = {"t": 0.0}

_mfa_wait_state = {"waiting": False, "since": 0.0}

# How many times a typed MFA code may be rejected before giving up on this
# sign-in attempt (each rejection automatically emails a fresh code).
_MFA_MAX_REJECTED_CODES = 3


def _prompt_mfa_code(challenge_name):
    """Pause the script until the user types the emailed MFA code.

    input() is called directly on the calling thread (not in a background
    thread), so it truly blocks and works in VS Code's integrated terminal,
    a normal terminal, and the Jupyter/Interactive window alike. There is
    deliberately no timeout: the run can't continue without a code anyway,
    and a background reader thread would swallow the next thing typed. If a
    code expires while you're away, just press Enter on a blank line and a
    fresh one is emailed.
    """
    print("\a", end="")  # terminal bell, to get attention if you've tabbed away
    print("\n" + "=" * 60)
    print("  ⏸  PAUSED — CHECK YOUR EMAIL FOR THE NERIS MFA CODE")
    print("  Type the code below and press Enter to continue.")
    print("  (Press Enter on a blank line to have a NEW code emailed.)")
    print("=" * 60, flush=True)
    _mfa_wait_state["waiting"] = True
    _mfa_wait_state["since"] = time.monotonic()
    try:
        return input(f"MFA code ({challenge_name}): ").strip()
    except EOFError:
        # SystemExit is not caught by the retry loops' `except Exception`,
        # so this stops cleanly instead of emailing code after code.
        sys.exit(_NO_INPUT_MESSAGE)
    finally:
        _mfa_wait_state["waiting"] = False
        # Don't let the stall watcher count the time spent typing as "stuck".
        _last_activity_at["t"] = time.monotonic()


def _maybe_force_reauth(client_self, token_before, context_label):
    """Shared by every retry path that can trigger a forced re-auth (the
    patched client._call, plus any hand-rolled GET that bypasses
    client._call() entirely). Must be
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
    kind. Stays silent while an MFA prompt is waiting (so it never prints
    over the line you're typing the code into), and distinguishes a short
    idle stretch (normal retry/backoff) from one that looks genuinely stuck."""
    while not stop_event.wait(30):
        if _mfa_wait_state["waiting"]:
            continue

        idle = time.monotonic() - _last_activity_at["t"]
        progress_desc = get_progress()

        if idle < 90:
            continue
        if idle < 300:
            print(f"  ⏳ Still working — no request has completed in "
                  f"{idle:.0f}s ({progress_desc} so far). This is normal "
                  f"if a request is mid-retry.")
        else:
            print(f"  ⚠ No request has completed in {idle:.0f}s "
                  f"({progress_desc} so far) — this "
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
        rejected_codes = 0

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
                print("  ✓ Signed in to NERIS", flush=True)
                return

            if res.status_code == 202:
                challenge = got.get("challenge_name", "MFA")
                code = _prompt_mfa_code(challenge)

                if not code:
                    print("  Requesting a new MFA code — check your email again...")
                    res = _password_grant()
                    continue

                res = self._session.post(
                    token_url,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    data={
                        "grant_type": challenge,
                        "username": self.config.username,
                        "session": got["session"],
                        challenge: code,
                    },
                    timeout=30,
                )

                if res.status_code not in (200, 202):
                    rejected_codes += 1
                    if rejected_codes < _MFA_MAX_REJECTED_CODES:
                        print(f"  ✗ That code wasn't accepted (HTTP {res.status_code}). "
                              f"A new code is being emailed — please try again.")
                        res = _password_grant()
                        continue
                    raise RuntimeError(
                        f"MFA code rejected {rejected_codes} times "
                        f"(last: HTTP {res.status_code} - {res.text[:300]!r})"
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
                need_cooldown = _maybe_force_reauth(self, token_before, f"api/{args[:2] if args else kwargs.get('path', '?')}")
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
    print("(NERIS will email you an MFA code — the script will pause and ask for it.)",
          flush=True)
    client = NerisApiClient(Config(
        base_url="https://api.neris.fsri.org/v1",
        grant_type="password",
        username=username,
        password=password,
    ))

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

_progress     = {"desc": "0 incidents retrieved"}
_fetch_status = {"incomplete": False, "error": None}


def calculate_date_range(range_type, start_date=None, end_date=None):
    """Calculate date filters for the query.
    Returns a tuple of (call_create_start, call_create_end).
    Both are sent to the API as server-side filters:
    - call_create_start: API returns incidents where call_create >= this value
    - call_create_end: API returns incidents where call_create <= this value
    All datetimes are timezone-aware (UTC)."""

    from datetime import timezone

    now = datetime.now(tz=timezone.utc)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if range_type == 'All Records':
        return None, None
    elif range_type == 'Occurred Today':
        return today, None
    elif range_type == 'Occurred Yesterday':
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday.replace(hour=23, minute=59, second=59)
    elif range_type == 'Occurred Last 7 Days':
        return today - timedelta(days=7), None
    elif range_type == 'Occurred Last 30 Days':
        return today - timedelta(days=30), None
    elif range_type == 'Occurred Last 90 Days':
        return today - timedelta(days=90), None
    elif range_type.startswith('Occurred in '):
        year = int(range_type.split()[-1])
        return (datetime(year, 1, 1, 0, 0, 0, tzinfo=timezone.utc),
                datetime(year, 12, 31, 23, 59, 59, tzinfo=timezone.utc))
    elif range_type == 'Custom Date Range':
        start = datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc)
        end = datetime.combine(end_date, datetime.max.time().replace(microsecond=0), tzinfo=timezone.utc)
        return start, end

    return None, None


def get_state_incidents(client, state_code, call_create_start=None, call_create_end=None,
                        neris_id_entity=None, incident_types=None, max_incidents=None, page_size=100):
    """Retrieve incidents using the NERIS client library with server-side filtering.

    incident_types: list of NERIS type strings (e.g. ['FIRE||STRUCTURE_FIRE||CHIMNEY_FIRE']).
                    When provided, each value is passed as incident_types to the API,
                    which ORs them together. Pass None or [] to skip type filtering.

    Every page goes through the patched, lock-protected client._call()
    (auth check, proactive refresh, forced re-auth on a 401/403), and is
    additionally retried up to 3 times on any other error (e.g. a 5xx)
    before giving up. If a page still fails, whatever was retrieved so far
    is returned and _fetch_status["incomplete"] is set, so the run reports
    PARTIAL results loudly instead of silently treating them as the
    complete set.
    """
    all_incidents = []
    next_cursor = None
    incident_count = 0
    page_count = 0
    _fetch_status["incomplete"] = False
    _fetch_status["error"] = None

    print(f"Retrieving incidents for state: {state_code}")
    if neris_id_entity:
        print(f"Filter: neris_id_entity = {neris_id_entity}")
    if call_create_start:
        print(f"Filter: call_create >= {call_create_start.strftime('%Y-%m-%d %H:%M:%S')}")
    if call_create_end:
        print(f"Filter: call_create <= {call_create_end.strftime('%Y-%m-%d %H:%M:%S')}")
    if incident_types:
        print(f"Filter: incident_types = {incident_types}  ({len(incident_types)} type(s))")
    else:
        print("Filter: All incident types (no type filter)")
    if not call_create_start and not call_create_end:
        print("Filter: All records (no date filter)")

    records_remain = True
    max_page_attempts = 3

    while records_remain:
        page_count += 1
        print(f"  Page {page_count}... ", end='', flush=True)

        kwargs = {
            'state': state_code,
            'page_size': page_size,
        }
        if call_create_start:
            kwargs['call_create_start'] = call_create_start
        if call_create_end:
            kwargs['call_create_end'] = call_create_end
        if neris_id_entity:
            kwargs['neris_id_entity'] = neris_id_entity
        if incident_types:
            kwargs['incident_types'] = list(incident_types)
        if next_cursor:
            kwargs['cursor'] = next_cursor

        res = None
        for attempt in range(1, max_page_attempts + 1):
            try:
                res = _unwrap(client.list_incidents(**kwargs),
                              f"list_incidents(page {page_count})")
                if not isinstance(res, dict):
                    res = res.json()
                break
            except Exception as e:
                _fetch_status["error"] = str(e)
                if attempt < max_page_attempts:
                    wait = 5 * attempt
                    print(f"\n  ⚠ Page {page_count} attempt {attempt}/{max_page_attempts} "
                          f"failed: {e} — retrying in {wait}s...")
                    time.sleep(wait)
                else:
                    print(f"\n✗ Error on page {page_count} after {max_page_attempts} attempts: {e}")
                    traceback.print_exc()
                    res = None

        if res is None:
            _fetch_status["incomplete"] = True
            print(f"⚠ STOPPING EARLY — only {incident_count} incident(s) from "
                  f"{page_count - 1} page(s) were retrieved. Results are INCOMPLETE.")
            break

        incidents = res.get('incidents', [])

        if not incidents:
            print("empty page")
            break

        all_incidents.extend(incidents)
        incident_count += len(incidents)
        _progress["desc"] = f"{incident_count} incidents retrieved over {page_count} page(s)"
        print(f"retrieved {len(incidents)} incidents (Total: {incident_count})")

        if max_incidents and incident_count >= max_incidents:
            all_incidents = all_incidents[:max_incidents]
            print(f"✓ Reached limit: {max_incidents}")
            break

        next_cursor = res.get('next_cursor')

        if not next_cursor:
            records_remain = False
            print("✓ No more pages")

    print(f"\n{'='*50}")
    print(f"Total incidents retrieved: {len(all_incidents)}")
    if _fetch_status["incomplete"]:
        print("⚠ INCOMPLETE — the fetch stopped early (see errors above)")
    print(f"{'='*50}")
    return all_incidents


def incidents_to_dataframe(incidents):
    """Convert incident list to pandas DataFrame."""
    if not incidents:
        print("⚠ No incidents to convert")
        return pd.DataFrame()

    for incident in incidents:
        if 'base' in incident and isinstance(incident['base'], dict):
            if 'neris_uid' in incident['base']:
                incident['neris_uid'] = incident['base']['neris_uid']
            if 'incident_number' in incident['base']:
                incident['incident_number'] = incident['base']['incident_number']

    df = pd.json_normalize(incidents, sep='_')

    import warnings
    for col in df.columns:
        if any(field in col.lower() for field in [
            'date', 'time', 'created', 'updated', 'modified',
            'clear', 'arrival', 'answered', 'create', 'enroute',
            'on_scene', 'staging', 'canceled', 'hospital',
            'patient', 'transferred'
        ]):
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    df[col] = pd.to_datetime(df[col], errors='coerce', utc=True)
            except:
                pass

    print(f"✓ Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
    return df

def export_to_excel(incidents, client, filename=None, state_code=None):
    """Export incidents to formatted Excel with Overview and Raw Data sheets.

    - Overview: curated columns with clean extracted values
    - Raw Data: full json_normalize with array fields flattened to strings
    - Both sheets: blank columns removed, original dark blue header formatting
    """
    if not incidents:
        print("⚠ No data to export")
        return None

    if filename is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"neris_incidents_{state_code}_{timestamp}.xlsx"

    dept_ids = set()
    for inc in incidents:
        base = inc.get('base', {}) or {}
        did = base.get('department_neris_id')
        if did:
            dept_ids.add(did)

    print(f"Looking up {len(dept_ids)} department names...")
    dept_names = {}
    for i, did in enumerate(dept_ids, 1):
        try:
            entity = _unwrap(client.get_entity(did), f"get_entity({did})")
            dept_names[did] = entity.get('name', '') if isinstance(entity, dict) else ''
        except Exception as e:
            print(f"  ⚠ Could not look up department name for {did}: {e}")
            dept_names[did] = ''
        _progress["desc"] = f"{i}/{len(dept_ids)} department names looked up"
        if i % 25 == 0:
            print(f"  {i}/{len(dept_ids)} departments looked up...")
    print(f"✓ Department lookup complete")

    # Helpers
    def extract_list_field(lst, key):
        if not lst or not isinstance(lst, list):
            return ''
        vals = [str(item.get(key, '')) for item in lst if isinstance(item, dict) and item.get(key)]
        return '||'.join(vals)

    def join_list(lst):
        if not lst or not isinstance(lst, list):
            return ''
        return '||'.join(str(x) for x in lst if x)

    overview_rows = []
    for inc in incidents:
        base = inc.get('base', {}) or {}
        dispatch = inc.get('dispatch', {}) or {}
        status_obj = inc.get('incident_status', {}) or {}
        dept_id = base.get('department_neris_id', '')

        inc_types = inc.get('incident_types', []) or []
        incident_type_list = []
        for it in inc_types:
            if isinstance(it, dict):
                t = it.get('type', '')
                if t:
                    incident_type_list.append(t)

        first_type = incident_type_list[0] if len(incident_type_list) > 0 else ''
        second_type = incident_type_list[1] if len(incident_type_list) > 1 else ''
        third_type = incident_type_list[2] if len(incident_type_list) > 2 else ''

        overview_rows.append({
            'Incident NERIS ID': inc.get('neris_id', ''),
            'Department/Entity ID': dept_id,
            'Department Name': dept_names.get(dept_id, ''),
            'Incident Number': base.get('incident_number', ''),
            'Call Date': dispatch.get('call_create', ''),
            'Last Modified': inc.get('last_modified', ''),
            'Incident Status': status_obj.get('status', ''),
            'First Incident Type': first_type,
            'Second Incident Type': second_type,
            'Third Incident Type': third_type,
            'People Present': base.get('people_present', ''),
            'Animals Rescued': base.get('animals_rescued', ''),
            'Outcome Narrative': base.get('outcome_narrative', ''),
            'Displacement Count': base.get('displacement_count', ''),
            'Submitter Type': inc.get('submitter_account_type', ''),
        })

    import copy
    processed = []
    casualty_detail_rows = []
    for inc in copy.deepcopy(incidents):
        base = inc.get('base', {}) or {}
        dispatch = inc.get('dispatch', {}) or {}
        dept_id = base.get('department_neris_id', '')

        inc['department_name'] = dept_names.get(dept_id, '')

        if isinstance(base, dict):
            inc['department_neris_id'] = base.get('department_neris_id', '')
            inc['incident_number'] = base.get('incident_number', '')
        dept_obj = inc.get('department', {}) or {}
        inc['department_time_zone'] = dept_obj.get('time_zone', '')

        inc_types = inc.get('incident_types', []) or []
        primary = [it.get('type', '') for it in inc_types if isinstance(it, dict) and it.get('primary')]
        other = [it.get('type', '') for it in inc_types if isinstance(it, dict) and not it.get('primary')]
        inc['primary_incident_types'] = '||'.join(primary)
        inc['other_incident_types'] = '||'.join(other)
        inc.pop('incident_types', None)

        inc['special_modifiers'] = extract_list_field(inc.get('special_modifiers', []), 'type')

        aids = inc.get('aids', []) or []
        inc['aid_types'] = extract_list_field(aids, 'aid_type')
        inc['aid_directions'] = extract_list_field(aids, 'aid_direction')
        inc['aid_department_ids'] = extract_list_field(aids, 'department_neris_id')
        inc.pop('aids', None)

        inc['nonfd_aid_types'] = extract_list_field(inc.get('nonfd_aids', []), 'type')
        inc.pop('nonfd_aids', None)

        units = inc.get('unit_responses', []) or []
        inc['unit_ids'] = extract_list_field(units, 'reported_unit_id')
        inc['unit_neris_ids'] = extract_list_field(units, 'unit_neris_id')
        inc['total_units'] = len(units)
        inc['total_staffing'] = sum(u.get('staffing', 0) or 0 for u in units if isinstance(u, dict))
        inc['unit_response_modes'] = extract_list_field(units, 'response_mode')
        inc.pop('unit_responses', None)

        disp_units = dispatch.get('unit_responses', []) or []
        dispatch['disp_unit_ids'] = extract_list_field(disp_units, 'reported_unit_id')
        dispatch['disp_total_units'] = len(disp_units)
        dispatch['disp_total_staffing'] = sum(u.get('staffing', 0) or 0 for u in disp_units if isinstance(u, dict))
        dispatch.pop('unit_responses', None)

        comments = dispatch.get('comments', []) or []
        dispatch['comments_text'] = extract_list_field(comments, 'comment')
        dispatch.pop('comments', None)

        base['displacement_causes'] = join_list(base.get('displacement_causes', []))

        def get_coordinates_from_point(point_obj):
            if not point_obj or not isinstance(point_obj, dict):
                return None, None
            geometry = point_obj.get('geometry', {})
            if geometry and isinstance(geometry, dict):
                coords = geometry.get('coordinates', [])
                if coords and isinstance(coords, list) and len(coords) >= 2:
                    return coords[1], coords[0]
            return None, None

        def get_location_data(loc_dict, point_obj=None, neris_uid=None):
            if not loc_dict or not isinstance(loc_dict, dict):
                return {}
            lat, lon = get_coordinates_from_point(point_obj)
            return {
                'street_number': loc_dict.get('number', ''),
                'street_prefix': loc_dict.get('street_prefix_direction', ''),
                'street_name': loc_dict.get('street', ''),
                'street_type': loc_dict.get('street_postfix', ''),
                'street_suffix': loc_dict.get('street_postfix_direction', ''),
                'apartment': loc_dict.get('unit_value', ''),
                'city': loc_dict.get('incorporated_municipality', ''),
                'state': loc_dict.get('state', ''),
                'zip_code': loc_dict.get('postal_code', ''),
                'county': loc_dict.get('county', ''),
                'latitude': lat if lat is not None else '',
                'longitude': lon if lon is not None else '',
                'census_tract': '',
                'census_block': '',
                'parcel_id': '',
            }

        base_loc = base.get('location', {}) or {}
        base_loc_geocoded = base_loc.get('geocoded_location', {}) or {}
        base_loc_geocoded_point = base_loc_geocoded.get('point', {}) if base_loc_geocoded else {}
        base_loc_geocoded_uid = base_loc_geocoded.get('neris_uid') if base_loc_geocoded else None

        disp_loc = dispatch.get('location', {}) or {}
        disp_loc_geocoded = disp_loc.get('geocoded_location', {}) or {}
        disp_loc_geocoded_point = disp_loc_geocoded.get('point', {}) if disp_loc_geocoded else {}
        disp_loc_geocoded_uid = disp_loc_geocoded.get('neris_uid') if disp_loc_geocoded else None

        location_source = ''
        final_location = {}

        if base_loc_geocoded and any(base_loc_geocoded.values()):
            final_location = get_location_data(base_loc_geocoded, base_loc_geocoded_point, base_loc_geocoded_uid)
            location_source = 'base_location_geocoded'
        elif base_loc and any(base_loc.values()):
            final_location = get_location_data(base_loc)
            location_source = 'base_location'
        elif disp_loc_geocoded and any(disp_loc_geocoded.values()):
            final_location = get_location_data(disp_loc_geocoded, disp_loc_geocoded_point, disp_loc_geocoded_uid)
            location_source = 'dispatch_location_geocoded'
        elif disp_loc and any(disp_loc.values()):
            final_location = get_location_data(disp_loc)
            location_source = 'dispatch_location'

        inc['location_source'] = location_source
        for field, value in final_location.items():
            inc[f'location_{field}'] = value

        base.pop('location', None)
        dispatch.pop('location', None)

        parcel = inc.get('parcel', {}) or {}
        parcel_assessment = (parcel.get('summary', {}) or {}).get('assessment', {}) or {}
        inc['parcel_primary_year_built'] = parcel_assessment.get('primary_year_built', '')
        inc['parcel_primary_zoning'] = parcel_assessment.get('primary_zoning', '')
        inc['parcel_primary_use'] = parcel_assessment.get('primary_use', '')
        inc['parcel_primary_owner_occupied'] = parcel_assessment.get('primary_owner_occupied', '')
        inc['parcel_primary_assessment_land'] = parcel_assessment.get('primary_assessment_land', '')
        inc['parcel_computed_loss'] = parcel_assessment.get('computed_loss', '')

        exposures = inc.get('exposures', []) or []
        inc['exposure_count'] = len(exposures)
        inc['exposure_damage_types'] = extract_list_field(exposures, 'damage_type')
        inc.pop('exposures', None)

        casualties = inc.get('casualty_rescues', []) or []
        inc['casualty_count'] = len(casualties)
        inc['casualty_types'] = extract_list_field(casualties, 'type')
        inc['casualty_ranks'] = extract_list_field(casualties, 'rank')
        inc['casualty_genders'] = extract_list_field(casualties, 'gender')

        neris_id = inc.get('neris_id', '')
        for cas in casualties:
            if not isinstance(cas, dict):
                continue
            casualty_obj = cas.get('casualty', {}) or {}
            injury = casualty_obj.get('injury_or_noninjury', {}) or {}
            ff_details = injury.get('ff_injury_details', {}) or {}
            rescue_obj = cas.get('rescue', {}) or {}
            ffrescue = rescue_obj.get('ffrescue_or_nonffrescue', {}) or {}
            removal = ffrescue.get('removal_or_nonremoval', {}) or {}
            fire_removal = removal.get('fire_removal', {}) or {}
            presence_known = rescue_obj.get('presence_known', {}) or {}
            mayday_obj = rescue_obj.get('mayday', {}) or {}

            casualty_detail_rows.append({
                'Incident NERIS ID': neris_id,
                'Department/Entity ID': dept_id,
                'Department Name': dept_names.get(dept_id, ''),
                'Call Date': dispatch.get('call_create', ''),
                'County': (base.get('location', {}) or {}).get('county', ''),
                'Type': cas.get('type', ''),
                'Rank': cas.get('rank', ''),
                'Years of Service': cas.get('years_of_service', ''),
                'Birth Month Year': cas.get('birth_month_year', ''),
                'Gender': cas.get('gender', ''),
                'Race': cas.get('race', ''),
                'Injury Type': injury.get('type', ''),
                'Injury Cause': injury.get('cause', ''),
                'FF Unit NERIS ID': ff_details.get('unit_neris_id', ''),
                'FF Reported Unit ID': ff_details.get('reported_unit_id', ''),
                'FF Unit Continuity': ff_details.get('unit_continuity', ''),
                'FF Incident Command': ff_details.get('incident_command', ''),
                'FF Job Classification': ff_details.get('job_classification', ''),
                'FF Duty Type': ff_details.get('duty_type', ''),
                'FF Action Type': ff_details.get('action_type', ''),
                'FF Incident Stage': ff_details.get('incident_stage', ''),
                'FF PPE Items': join_list(ff_details.get('ppe_items', [])),
                'Rescue Type': ffrescue.get('type', ''),
                'Rescue Actions': join_list(ffrescue.get('actions', [])),
                'Rescue Impediments': join_list(ffrescue.get('impediments', [])),
                'Removal Type': removal.get('type', ''),
                'Removal Room Type': removal.get('room_type', ''),
                'Removal Elevation Type': removal.get('elevation_type', ''),
                'Rescue Path Type': removal.get('rescue_path_type', ''),
                'Gas Isolation': removal.get('gas_isolation', ''),
                'Fire Removal Relative Suppression Time': fire_removal.get('relative_suppression_time', ''),
                'Presence Known Type': presence_known.get('presence_known_type', ''),
                'Mayday': mayday_obj.get('mayday', ''),
                'RIT Activated': mayday_obj.get('rit_activated', ''),
                'Mayday Relative Suppression Time': mayday_obj.get('relative_suppression_time', ''),
            })

        inc.pop('casualty_rescues', None)

        med = inc.get('medical_details', []) or []
        inc['med_report_ids'] = extract_list_field(med, 'patient_care_report_id')
        inc['med_evaluations'] = extract_list_field(med, 'patient_care_evaluation')
        inc['med_patient_statuses'] = extract_list_field(med, 'patient_status')
        inc['med_transport_dispositions'] = extract_list_field(med, 'transport_disposition')
        inc.pop('medical_details', None)

        elec = inc.get('electric_hazards', []) or []
        inc['electric_hazard_types'] = extract_list_field(elec, 'type')
        inc['electric_hazard_source_target'] = extract_list_field(elec, 'source_or_target')
        inc.pop('electric_hazards', None)

        pgen = inc.get('powergen_hazards', []) or []
        inc['powergen_hazard_count'] = len(pgen)
        inc.pop('powergen_hazards', None)

        fire = inc.get('fire_detail', {}) or {}
        fire['investigation_types'] = join_list(fire.get('investigation_types', []))
        fire['suppression_appliances'] = join_list(fire.get('suppression_appliances', []))

        hazsit = inc.get('hazsit_detail', {}) or {}
        chems = hazsit.get('chemicals', []) or []
        hazsit['chemical_names'] = extract_list_field(chems, 'name')
        hazsit['chemical_dot_classes'] = extract_list_field(chems, 'dot_class')
        hazsit.pop('chemicals', None)

        sa_presence = (inc.get('smoke_alarm', {}) or {}).get('presence', {}) or {}
        sa_presence['alarm_types'] = join_list(sa_presence.get('alarm_types', []))

        fa_presence = (inc.get('fire_alarm', {}) or {}).get('presence', {}) or {}
        fa_presence['alarm_types'] = join_list(fa_presence.get('alarm_types', []))

        oa_presence = (inc.get('other_alarm', {}) or {}).get('presence', {}) or {}
        oa_presence['alarm_types'] = join_list(oa_presence.get('alarm_types', []))

        fs_presence = (inc.get('fire_suppression', {}) or {}).get('presence', {}) or {}
        supp_types = fs_presence.get('suppression_types', []) or []
        fs_presence['suppression_types'] = extract_list_field(supp_types, 'type') if supp_types and isinstance(supp_types[0], dict) else join_list(supp_types)

        cfs_presence = (inc.get('cooking_fire_suppression', {}) or {}).get('presence', {}) or {}
        cfs_types = cfs_presence.get('suppression_types', []) or []
        cfs_presence['suppression_types'] = join_list(cfs_types)

        at = inc.get('actions_tactics', {}) or {}
        an = at.get('action_noaction', {}) or {}
        an['actions'] = join_list(an.get('actions', []))

        processed.append(inc)

    df_raw = pd.json_normalize(processed, sep='_')

    exclude_cols = {
        'actions_tactics_last_modified', 'actions_tactics_neris_uid',
        'base_location_geocoded_location_neris_uid', 'base_location_neris_uid',
        'base_location_use_in_use_last_modified', 'base_location_use_in_use_neris_uid',
        'base_location_use_neris_uid',
        'census_tract_incident_neris_id', 'census_tract_last_modified', 'census_tract_neris_uid',
        'cooking_fire_suppression_last_modified', 'cooking_fire_suppression_neris_uid',
        'cooking_fire_suppression_presence_last_modified', 'cooking_fire_suppression_presence_neris_uid',
        'csst_hazard_last_modified', 'csst_hazard_neris_uid',
        'dispatch_last_modified', 'dispatch_location_geocoded_location_neris_uid', 'dispatch_neris_uid',
        'dispatch_tactic_timestamps_last_modified', 'dispatch_tactic_timestamps_neris_uid',
        'fire_alarm_last_modified', 'fire_alarm_neris_uid',
        'fire_alarm_presence_last_modified', 'fire_alarm_presence_neris_uid',
        'fire_detail_last_modified', 'fire_detail_neris_uid',
        'fire_suppression_last_modified', 'fire_suppression_neris_uid',
        'fire_suppression_presence_last_modified', 'fire_suppression_presence_neris_uid',
        'fire_suppression_presence_operation_type_effectiveness_last_modified',
        'fire_suppression_presence_operation_type_effectiveness_neris_uid',
        'fire_suppression_presence_operation_type_last_modified',
        'fire_suppression_presence_operation_type_neris_uid',
        'hazsit_detail_last_modified', 'hazsit_detail_neris_uid',
        'incident_status_last_modified',
        'other_alarm_last_modified', 'other_alarm_neris_uid',
        'other_alarm_presence_last_modified', 'other_alarm_presence_neris_uid',
        'smoke_alarm_last_modified', 'smoke_alarm_neris_uid',
        'smoke_alarm_presence_operation_alerted_failed_other_last_modified',
        'smoke_alarm_presence_operation_alerted_failed_other_neris_uid',
        'smoke_alarm_presence_operation_last_modified', 'smoke_alarm_presence_operation_neris_uid',
        'tactic_timestamps_last_modified', 'tactic_timestamps_neris_uid',
        'weather_last_modified', 'weather_neris_uid',
        'department_time_zone',
        'base_neris_uid',
        'base_last_modified',
        'base_department_neris_id',
        'incident_status_created_by',
        'dispatch_location_neris_uid',
        'tactic_timestamps_command_established',
        'tactic_timestamps_completed_sizeup',
        'tactic_timestamps_suppression_complete',
        'tactic_timestamps_primary_search_begin',
        'tactic_timestamps_primary_search_complete',
        'tactic_timestamps_water_on_fire',
        'tactic_timestamps_fire_under_control',
        'tactic_timestamps_fire_knocked_down',
        'weather_incident_neris_id',
        'dispatch_tactic_timestamps_command_established',
        'dispatch_tactic_timestamps_completed_sizeup',
        'dispatch_tactic_timestamps_suppression_complete',
        'dispatch_tactic_timestamps_fire_under_control',
    }
    df_raw = df_raw.drop(columns=[c for c in exclude_cols if c in df_raw.columns])

    def reorder_columns(df):
        first_cols = [
            'department_neris_id',
            'department_name',
            'neris_id',
            'neris_uid',
            'incident_status_status',
        ]
        location_cols = [
            'location_source',
            'location_street_number',
            'location_street_prefix',
            'location_street_name',
            'location_street_type',
            'location_street_suffix',
            'location_apartment',
            'location_city',
            'location_state',
            'location_zip_code',
            'location_county',
            'location_latitude',
            'location_longitude',
            'location_census_tract',
            'location_census_block',
            'location_parcel_id',
        ]
        hazard_fields = [
            'electric_hazard_types',
            'electric_hazard_source_target',
            'powergen_hazard_count',
        ]
        existing_first = [c for c in first_cols if c in df.columns]
        existing_location = [c for c in location_cols if c in df.columns]
        existing_hazard = [c for c in hazard_fields if c in df.columns]
        all_cols = df.columns.tolist()
        middle_cols = [c for c in all_cols if c not in existing_first + existing_location + existing_hazard + ['last_modified']]
        final_order = existing_first + existing_location + middle_cols + existing_hazard
        if 'last_modified' in df.columns:
            final_order.append('last_modified')
        return df[final_order]

    df_raw = reorder_columns(df_raw)

    def clean_df(df):
        cols_to_drop = []
        for col in df.columns:
            vals = df[col].dropna()
            if vals.empty:
                cols_to_drop.append(col)
                continue
            if vals.apply(lambda x: x == '' or x == [] or (isinstance(x, str) and x == '[]')).all():
                cols_to_drop.append(col)
        return df.drop(columns=cols_to_drop)

    df_overview = clean_df(pd.DataFrame(overview_rows).fillna(''))
    df_raw = clean_df(df_raw)
    df_casualties = clean_df(pd.DataFrame(casualty_detail_rows).fillna('')) if casualty_detail_rows else pd.DataFrame()

    print(f"  Overview: {len(df_overview.columns)} columns")
    print(f"  Raw Data: {len(df_raw.columns)} columns")
    print(f"  Casualty & Rescue: {len(df_casualties)} rows, {len(df_casualties.columns)} columns")

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell import WriteOnlyCell
    try:
        from openpyxl.cell.cell import KNOWN_TYPES as _XL_TYPES
    except ImportError:  # very old openpyxl
        _XL_TYPES = (str, int, float, bool, datetime, type(None))

    import re
    ILLEGAL_CHARS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
    EXCEL_CELL_MAX = 32767  # Excel's per-cell character limit

    def sanitize(val):
        """Make any value safe to put in an Excel cell.

        Some NERIS fields are still lists or dicts after the flattening
        above (e.g. an empty [] array on a field that isn't specifically
        handled), and openpyxl refuses those outright ("Cannot convert []
        to Excel"), which killed a 96k-incident export at the very end.
        Lists of plain values are joined with '||' (same convention as the
        rest of this export); anything nested is written as JSON text.
        """
        if isinstance(val, (list, tuple, set)):
            if not val:
                return None
            if all(not isinstance(x, (dict, list, tuple, set)) for x in val):
                val = '||'.join(str(x) for x in val)
            else:
                val = json.dumps(list(val), default=str)
        elif isinstance(val, dict):
            if not val:
                return None
            val = json.dumps(val, default=str)
        elif isinstance(val, datetime) and val.tzinfo is not None:
            # Excel has no time zones; keep the UTC wall-clock time.
            val = val.astimezone(timezone.utc).replace(tzinfo=None)
        try:
            if pd.isna(val):
                return None
        except (TypeError, ValueError):
            pass
        if isinstance(val, str):
            val = ILLEGAL_CHARS_RE.sub('', val)
            if len(val) > EXCEL_CELL_MAX:
                suffix = ' ...[truncated]'
                val = val[:EXCEL_CELL_MAX - len(suffix)] + suffix
        elif not isinstance(val, _XL_TYPES):
            val = str(val)
        return val

    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def write_formatted_sheet(wb, df, sheet_name):
        """Stream one DataFrame into a sheet of a write-only workbook.

        Write-only mode streams rows to disk instead of holding every cell
        in memory: a statewide export (~96k incidents x ~170 columns, about
        16 million cells) needs roughly 6 GB in openpyxl's normal mode but
        well under 100 MB this way. Write-only sheets must have their
        column widths and frozen header set BEFORE any row is written, so
        widths are sized from the first 50 data rows up front.
        """
        sheet = wb.create_sheet(sheet_name)
        headers = list(df.columns)

        sample = df.head(50).itertuples(index=False, name=None)
        widths = [len(str(h)) for h in headers]
        for row in sample:
            for i, v in enumerate(row):
                v = sanitize(v)
                if v:
                    widths[i] = max(widths[i], len(str(v)))
        for col_idx, w in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(w + 2, 12), 60)
        sheet.freeze_panes = 'A2'

        header_row = []
        for h in headers:
            c = WriteOnlyCell(sheet, value=sanitize(h))
            c.fill = header_fill
            c.font = header_font
            c.alignment = header_align
            header_row.append(c)
        sheet.append(header_row)

        for i, row in enumerate(df.itertuples(index=False, name=None), 1):
            sheet.append([sanitize(v) for v in row])
            if i % 20000 == 0:
                print(f"    {sheet_name}: {i:,}/{len(df):,} rows written...")

    wb = Workbook(write_only=True)
    write_formatted_sheet(wb, df_overview, 'Overview')
    write_formatted_sheet(wb, df_raw, 'Raw Data')
    if not df_casualties.empty:
        write_formatted_sheet(wb, df_casualties, 'Casualty & Rescue')
    wb.save(filename)

    print(f"\n✓ Excel exported: {filename}")
    print(f"  Incidents: {len(incidents)}")
    sheets = "Overview, Raw Data" + (", Casualty & Rescue" if not df_casualties.empty else "")
    print(f"  Sheets: {sheets}")
    return filename

def main():
    global pd

    # Show every printed line immediately (VS Code's panels can otherwise
    # hold output back, so the "check your email" banner appears late).
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass

    ensure_dependencies()
    import pandas as pd

    try:
        cfg = prompt_config()
    except EOFError:
        sys.exit(_NO_INPUT_MESSAGE)
    state_code = cfg["state_code"]

    log_path = _init_request_log(state_code)
    print(f"\n✓ Logging every request's x-request-id to: {log_path}")

    client = authenticate(cfg["username"], cfg["password"])

    print(f"\nState: {state_code}")
    if cfg["entity_id"]:
        print(f"Entity ID: {cfg['entity_id']}")
    print(f"Date Range: {cfg['range_type']}")
    if cfg["incident_types"]:
        print(f"Incident Types: {cfg['incident_types']}")
    else:
        print("Incident Types: All (no filter)")

    start_date_filter, end_date_filter = calculate_date_range(
        cfg["range_type"], cfg["start_date"], cfg["end_date"]
    )
    print("\n" + "=" * 70)

    _stall_watch_stop = threading.Event()
    threading.Thread(
        target=_watch_for_stalls,
        args=(_stall_watch_stop, lambda: _progress["desc"]),
        daemon=True,
    ).start()

    excel_file = None
    try:
        incidents = get_state_incidents(
            client,
            state_code,
            call_create_start=start_date_filter,
            call_create_end=end_date_filter,
            neris_id_entity=cfg["entity_id"],
            incident_types=cfg["incident_types"],
            max_incidents=None,
            page_size=100,
        )

        if not incidents:
            print("\n⚠ No incidents found")
        else:
            print("\nConverting to DataFrame...")
            df = incidents_to_dataframe(incidents)
            print(f"\nTotal Incidents: {len(df)}")
            print(df.head(3))

            print("\nExporting to Excel...")
            excel_file = export_to_excel(incidents, client, state_code=state_code)
            if excel_file:
                print(f"\n✓ Complete! File: {excel_file}")
    finally:
        _stall_watch_stop.set()

    if _fetch_status["incomplete"]:
        print("\n" + "!" * 70)
        print("  ⚠ WARNING: the incident fetch stopped early — the export contains")
        print("    PARTIAL results, not the full set. Last error:")
        print(f"    {_fetch_status['error']}")
        print("    Re-run to get the complete data.")
        print("!" * 70)

    if _request_log_path:
        print(f"\n  Request trace log : {_request_log_path}")
        print("  (every call's x-request-id, status, and attempt —")
        print("   send this to NERIS's team along with failing timestamps)")
        if _request_log_file:
            _request_log_file.close()

    print("\n" + "=" * 70)
    print("✓ PROCESS COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
