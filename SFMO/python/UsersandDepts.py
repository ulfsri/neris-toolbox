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
import threading
import getpass
from datetime import datetime, timezone, timedelta
from collections import defaultdict


def ensure_neris_client_installed():
    print("Installing NERIS API client...")
    try:
        result = subprocess.run([
            sys.executable, '-m', 'pip', 'install',
            'https://github.com/ulfsri/neris-api-client/archive/refs/heads/main.zip',
            '--quiet'
        ], capture_output=True, text=True)

        if result.returncode == 0:
            print("✓ NERIS API client installed successfully")
        else:
            print(f"Installation output: {result.stdout}")
            print(f"Installation errors: {result.stderr}")
    except Exception as e:
        print(f"Installation error: {e}")


ensure_neris_client_installed()

try:
    from neris_api_client import NerisApiClient
    from neris_api_client.client import _NerisApiClient
    print("✓ NERIS API Client loaded")
except ImportError:
    print("⚠ NERIS API Client not found. Try running: pip install --break-system-packages "
          "https://github.com/ulfsri/neris-api-client/archive/refs/heads/main.zip")
    sys.exit(1)

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("⚠ Missing dependency. Run: pip install pandas openpyxl")
    sys.exit(1)


_auth_lock = threading.Lock()
_orig_call = _NerisApiClient._call


def _locked_call(self, *args, **kwargs):
    with _auth_lock:
        return _orig_call(self, *args, **kwargs)


_NerisApiClient._call = _locked_call
print("✓ Applied thread-safety patch to NerisApiClient")


MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']


def prompt_credentials():
    print("\n" + "=" * 60)
    print("NERIS Login")
    print("=" * 60)
    username = input("NERIS Email: ").strip()
    password = getpass.getpass("NERIS Password: ")
    return username, password


def authenticate(username, password):
    os.environ.update({
        'NERIS_BASE_URL':   'https://api.neris.fsri.org/v1',
        'NERIS_GRANT_TYPE': 'password',
        'NERIS_USERNAME':   username,
        'NERIS_PASSWORD':   password,
    })

    print("\nCreating NERIS API client…")
    client = NerisApiClient()

    print("\n" + "=" * 60)
    print("📧  CHECK YOUR EMAIL FOR THE MFA CODE")
    print("   (you'll be prompted for it right here in the terminal)")
    print("=" * 60)

    try:
        client.list_incidents(page_size=1)
        print("✓ Authentication successful")
        return client
    except Exception as e:
        print(f"✗ Authentication failed: {e}")
        return None


def prompt_parameters():
    print("\n" + "=" * 60)
    print("Query Parameters")
    print("=" * 60)

    state_code = input("State Code (e.g. MI, CA, NY): ").strip().upper()
    while not state_code:
        state_code = input("State Code is required. Enter it: ").strip().upper()

    entity_id = input("NERIS Entity ID (optional, e.g. FD26163151 — leave blank for all): ").strip() or None

    now = datetime.now()
    years = list(range(2025, now.year + 1))

    print(f"\nMonths: {', '.join(MONTHS)}")
    start_month = _prompt_choice("Start Month", MONTHS, default='January')
    start_year = _prompt_year("Start Year", years, default=2025)
    end_month = _prompt_choice("End Month", MONTHS, default=MONTHS[now.month - 1])
    end_year = _prompt_year("End Year", years, default=now.year)

    return state_code, entity_id, start_month, start_year, end_month, end_year


def _prompt_choice(label, options, default):
    raw = input(f"{label} [{default}]: ").strip()
    if not raw:
        return default
    # allow either the full name or a case-insensitive prefix match
    matches = [o for o in options if o.lower().startswith(raw.lower())]
    if len(matches) == 1:
        return matches[0]
    if raw in options:
        return raw
    print(f"  Didn't recognize '{raw}', using default: {default}")
    return default


def _prompt_year(label, options, default):
    raw = input(f"{label} [{default}]: ").strip()
    if not raw:
        return default
    try:
        year = int(raw)
        if year in options:
            return year
        print(f"  {year} not in expected range {options}, using default: {default}")
        return default
    except ValueError:
        print(f"  Didn't recognize '{raw}', using default: {default}")
        return default


def fetch_all_entities(client, state_code, neris_id_entity=None, page_size=100):
    """
    Fetch every registered department/entity for the state.
    Returns a dict: { neris_id: name }
    """
    if neris_id_entity:
        try:
            entity = client.get_entity(neris_id_entity)
            name = entity.get('name', '') if isinstance(entity, dict) else ''
            print(f"✓ Single entity lookup: {neris_id_entity} → {name}")
            return {neris_id_entity: name}
        except Exception as e:
            print(f"⚠ Could not fetch entity {neris_id_entity}: {e}")
            return {neris_id_entity: ''}

    all_entities = {}
    page_number = 1

    print(f"Fetching all registered entities for state: {state_code}")

    while True:
        print(f"  Page {page_number}... ", end='', flush=True)
        try:
            res = client.list_entities(state=state_code, page_size=page_size,
                                        page_number=page_number)
            if not isinstance(res, dict):
                res = res.json()
        except Exception as e:
            print(f"\n⚠ list_entities failed (page {page_number}): {e}")
            break

        batch = res.get('entities', [])
        if not batch:
            print("empty — done")
            break

        for ent in batch:
            eid = ent.get('neris_id', '')
            name = ent.get('name', '')
            if eid:
                all_entities[eid] = name

        print(f"retrieved {len(batch)} (total so far: {len(all_entities)})")

        total_count = res.get('total_count', 0)
        if len(all_entities) >= total_count or len(batch) < page_size:
            print("  ✓ All pages retrieved")
            break

        page_number += 1

    print(f"\n{'=' * 50}")
    print(f"Total registered departments: {len(all_entities)}")
    print(f"{'=' * 50}")
    return all_entities


def fetch_incidents_for_entity(client, state_code, neris_id_entity,
                                start_dt, end_dt, page_size=100):
    """
    Fetch incidents for a single department within the date range.
    Passes call_create_start/end to the API — no client-side filtering.
    Returns a dict: { 'Mon-YYYY': count }
    """
    month_counts = defaultdict(int)
    next_cursor = None

    while True:
        kwargs = dict(
            state=state_code,
            neris_id_entity=neris_id_entity,
            page_size=page_size,
            call_create_start=start_dt,
            call_create_end=end_dt,
        )
        if next_cursor:
            kwargs['cursor'] = next_cursor

        res = client.list_incidents(**kwargs)
        if not isinstance(res, dict):
            res = res.json()

        batch = res.get('incidents', [])
        if not batch:
            break

        for inc in batch:
            disp = inc.get('dispatch') or {}
            ts = disp.get('call_create') or disp.get('call_create_start')
            lbl = _month_label(ts)
            if lbl:
                month_counts[lbl] += 1

        next_cursor = res.get('next_cursor')
        if not next_cursor:
            break

    return dict(month_counts)


def fetch_nars_for_entity(client, state_code, neris_id_entity,
                           start_dt, end_dt, page_size=100):
    """
    Fetch no-activity reports for a single department within the date range.
    Returns a set of 'Mon-YYYY' labels.
    """
    nar_months = set()
    next_cursor = None
    base_url = os.environ.get('NERIS_BASE_URL', 'https://api.neris.fsri.org/v1')

    while True:
        params = dict(
            state=state_code,
            neris_id_entity=neris_id_entity,
            page_size=page_size,
            call_create_start=start_dt.isoformat(),
            call_create_end=end_dt.isoformat(),
        )
        if next_cursor:
            params['cursor'] = next_cursor


        with _auth_lock:
            r = client._session.get(f"{base_url}/no_activity_report", params=params)
        res = r.json()

        batch = res.get('reports', [])
        if not batch:
            break

        for report in batch:
            lbl = _month_label(report.get('month_year', ''))
            if lbl:
                nar_months.add(lbl)

        next_cursor = res.get('next_cursor')
        if not next_cursor:
            break

    return nar_months


def _month_label(dt_or_str):
    """Convert a datetime or 'MM/YYYY' / ISO string to 'Mon-YYYY'. Returns None on failure."""
    if dt_or_str is None:
        return None
    if isinstance(dt_or_str, str):
        if '/' in dt_or_str and len(dt_or_str) <= 7:
            try:
                m, y = dt_or_str.split('/')
                return datetime(int(y), int(m), 1).strftime('%b-%Y')
            except Exception:
                pass
        try:
            dt_or_str = datetime.fromisoformat(dt_or_str.replace('Z', '+00:00'))
        except Exception:
            return None
    try:
        return dt_or_str.strftime('%b-%Y')
    except Exception:
        return None


def generate_month_columns(start_dt, end_dt):
    """Month labels covering start_dt through end_dt."""
    cols = []
    d = datetime(start_dt.year, start_dt.month, 1)
    end = datetime(end_dt.year, end_dt.month, 1)
    while d <= end:
        cols.append(d.strftime('%b-%Y'))
        m = d.month + 1
        d = datetime(d.year + (m // 13), ((m - 1) % 12) + 1, 1)
    return cols


def build_report(client, state_code, entity_id, start_month, start_year, end_month, end_year):
    from concurrent.futures import ThreadPoolExecutor, as_completed

    start_month_num = MONTHS.index(start_month) + 1
    end_month_num = MONTHS.index(end_month) + 1
    start_dt = datetime(start_year, start_month_num, 1, tzinfo=timezone.utc)

    next_m = end_month_num + 1
    end_dt = datetime(
        end_year + (next_m // 13),
        ((next_m - 1) % 12) + 1,
        1, tzinfo=timezone.utc
    ) - timedelta(seconds=1)

    if start_dt > end_dt:
        print("✗ ERROR: Start date must be before end date.")
        return

    print(f"\nState: {state_code}")
    if entity_id:
        print(f"Entity ID filter: {entity_id}")
    print(f"Date Range: {start_month} {start_year} — {end_month} {end_year}")
    print("\n" + "=" * 70)

    try:
        all_entities = fetch_all_entities(client, state_code, entity_id)
        all_eids = list(all_entities.keys())

        print(f"\nFetching incidents and NARs for {len(all_eids)} departments in parallel...")
        print("(progress updates every 50 departments)\n")

        inc_counts = {}
        nar_flags = {}
        completed = 0
        errors = []

        _first_dept_done = {'done': False}

        def fetch_dept_data(eid):
            counts = fetch_incidents_for_entity(
                client, state_code, eid, start_dt, end_dt)
            nars = fetch_nars_for_entity(
                client, state_code, eid, start_dt, end_dt)
            if not _first_dept_done['done']:
                _first_dept_done['done'] = True
                print(f"\n  [Diagnostic] First dept {eid}: "
                      f"{sum(counts.values())} incidents, {len(nars)} NAR months")
            return eid, counts, nars

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(fetch_dept_data, eid): eid
                       for eid in all_eids}

            first_errors_shown = 0
            for future in as_completed(futures):
                eid = futures[future]
                try:
                    eid, counts, nars = future.result()
                    inc_counts[eid] = counts
                    nar_flags[eid] = nars
                except Exception as e:
                    errors.append((eid, str(e)))
                    inc_counts[eid] = {}
                    nar_flags[eid] = set()
                    if first_errors_shown < 3:
                        print(f"\n  ⚠ ERROR for {eid}: {e}")
                        print(f"    Token state at failure: {NerisApiClient.tokens}")
                        first_errors_shown += 1

                completed += 1
                if completed % 50 == 0 or completed == len(all_eids):
                    total_incidents = sum(sum(v.values()) for v in inc_counts.values())
                    total_nars = sum(len(v) for v in nar_flags.values())
                    print(f"  {completed}/{len(all_eids)} departments complete "
                          f"| {total_incidents} incidents | {total_nars} NAR months")

        if errors:
            print(f"\n⚠ {len(errors)} department(s) had fetch errors:")
            for eid, err in errors[:10]:
                print(f"    {eid}: {err}")
            if len(errors) > 10:
                print(f"    ... and {len(errors) - 10} more")

        total_incidents = sum(sum(v.values()) for v in inc_counts.values())
        print(f"\n{'=' * 50}")
        print(f"Total incidents retrieved : {total_incidents}")
        print(f"Total NAR months on record: {sum(len(v) for v in nar_flags.values())}")
        print(f"{'=' * 50}")

        month_cols = generate_month_columns(start_dt, end_dt)
        rows = []

        for eid in sorted(all_eids):
            row = {
                'NERIS Entity ID': eid,
                'Department Name': all_entities.get(eid, ''),
            }
            dept_counts = inc_counts.get(eid, {})
            dept_nars = nar_flags.get(eid, set())

            for mc in month_cols:
                if mc in dept_counts:
                    row[mc] = dept_counts[mc]
                elif mc in dept_nars:
                    row[mc] = 'NAR'
                else:
                    row[mc] = ''
            rows.append(row)

        df_pivot = pd.DataFrame(rows)

        summary = {'NERIS Entity ID': 'TOTAL', 'Department Name': ''}
        for mc in month_cols:
            nums = pd.to_numeric(df_pivot[mc], errors='coerce').dropna()
            summary[mc] = int(nums.sum()) if not nums.empty else ''
        df_pivot = pd.concat([df_pivot, pd.DataFrame([summary])],
                              ignore_index=True)

        print(f"\nPivot table: {len(df_pivot) - 1} departments  x  {len(month_cols)} months")

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"neris_monthly_counts_{state_code}_{timestamp}.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = 'Monthly Incident Counts'

        header_cols = ['NERIS Entity ID', 'Department Name'] + month_cols
        ws.append(header_cols)

        DARK_BLUE = '1F4E78'
        NAR_COLOR = 'FFF2CC'
        TOTAL_COLOR = 'D9E1F2'
        NO_DATA_COLOR = 'F5F5F5'

        hdr_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type='solid')
        nar_fill = PatternFill(start_color=NAR_COLOR, end_color=NAR_COLOR, fill_type='solid')
        total_fill = PatternFill(start_color=TOTAL_COLOR, end_color=TOTAL_COLOR, fill_type='solid')
        no_data_fill = PatternFill(start_color=NO_DATA_COLOR, end_color=NO_DATA_COLOR, fill_type='solid')
        hdr_font = Font(color='FFFFFF', bold=True, size=11)
        total_font = Font(bold=True, size=11)
        hdr_aln = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ctr_aln = Alignment(horizontal='center', vertical='center')
        left_aln = Alignment(horizontal='left', vertical='center')
        thin_side = Side(style='thin', color='BFBFBF')
        thin_border = Border(left=thin_side, right=thin_side,
                              top=thin_side, bottom=thin_side)

        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = hdr_aln

        total_row_idx = len(df_pivot) + 1
        for r_idx, row_data in enumerate(df_pivot.itertuples(index=False), start=2):
            is_total = (r_idx == total_row_idx)
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border

                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                    cell.alignment = ctr_aln
                elif c_idx <= 2:
                    cell.alignment = left_aln
                elif val == 'NAR':
                    cell.fill = nar_fill
                    cell.font = Font(italic=True, color='7F6000')
                    cell.alignment = ctr_aln
                elif val == '':
                    cell.fill = no_data_fill
                    cell.alignment = ctr_aln
                else:
                    cell.alignment = ctr_aln

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 35
        for i in range(3, len(header_cols) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 11

        ws.freeze_panes = 'C2'

        ws_legend = wb.create_sheet('Legend')
        legend_data = [
            ('Symbol', 'Meaning'),
            ('(number)', 'Count of incidents reported for that department in that month'),
            ('NAR', 'No-Activity Report filed — department confirmed zero incidents'),
            ('(grey blank)', 'No incidents and no no-activity report submitted for that month'),
            ('TOTAL row', 'Sum of numeric incident counts across all departments per month'),
        ]
        for lr in legend_data:
            ws_legend.append(lr)
        for cell in ws_legend[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
        ws_legend.column_dimensions['A'].width = 14
        ws_legend.column_dimensions['B'].width = 75

        wb.save(filename)

        print(f"\n✓ Excel exported: {filename}")
        print(f"  Departments : {len(df_pivot) - 1}")
        print(f"  Months      : {len(month_cols)}  ({month_cols[0]} — {month_cols[-1]})")
        print(f"  Incidents   : {total_incidents}")
        print(f"  NAR months  : {sum(len(v) for v in nar_flags.values())}")
        print("\nKey:")
        print("  (number)     = incident count")
        print("  NAR          = no-activity report filed (confirmed 0 incidents)")
        print("  (grey blank) = no data submitted for that month")

        return df_pivot, filename

    except Exception:
        print("\n✗ Error building report:")
        import traceback
        traceback.print_exc()
        return None

def main():
    username, password = prompt_credentials()
    if not username or not password:
        print("✗ Email and password are required.")
        sys.exit(1)

    client = authenticate(username, password)
    if client is None:
        sys.exit(1)

    state_code, entity_id, start_month, start_year, end_month, end_year = prompt_parameters()

    build_report(client, state_code, entity_id, start_month, start_year, end_month, end_year)

    print("\n" + "=" * 70)
    print("✓ PROCESS COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
