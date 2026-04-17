"""
database.py - persistence helpers for insurer mappings and legacy admin data.
"""

import json
import os
import re
import sqlite3
from io import StringIO
from urllib.parse import urlencode

import pandas as pd
import requests
import streamlit as st


BASE_DIR = os.path.dirname(__file__)
SQLITE_PATH = os.path.join(BASE_DIR, "insurance.db")
BASEROW_API_URL = "https://api.baserow.io"


CONSUMPTION_MASTER_FIELDS = [
    ("member_name", "Member Name"),
    ("service_date", "Service Date"),
    ("net_amount", "Net Payable Amount"),
    ("diagnosis", "Diagnosis"),
    ("provider", "Billing Provider"),
    ("category", "Category / Service Type"),
]


OPTIONAL_CONSUMPTION_FIELDS = [
    ("card_no", "Card Number"),
    ("provider_type", "Provider Type"),
    ("incident_type", "Incident Type"),
    ("chronic_flag", "Chronic Flag"),
    ("relation", "Relation"),
    ("plan", "Plan"),
    ("submission", "Submission Method"),
]


MASTER_COLUMNS = [
    "Full Name",
    "First Name",
    "Second Name",
    "Third Name",
    "DOB",
    "Gender",
    "Relation",
    "Plan",
    "Phone",
    "National ID",
    "Staff ID",
    "Empty",
]


DASHBOARD_FIELDS = CONSUMPTION_MASTER_FIELDS + OPTIONAL_CONSUMPTION_FIELDS


def _safe_secret(key, default=""):
    try:
        return st.secrets[key]
    except Exception:
        return default


def _safe_secret_dict(key):
    try:
        value = st.secrets[key]
        return dict(value)
    except Exception:
        return None


def _consumption_sheet_url():
    return _safe_secret("CONSUMPTION_MAPPINGS_SHEET_URL", "").strip()


def _baserow_api_url():
    return _safe_secret("BASEROW_API_URL", BASEROW_API_URL).strip().rstrip("/")


def _baserow_api_token():
    return _safe_secret("BASEROW_API_TOKEN", "").strip()


def _baserow_consumption_table_id():
    return str(_safe_secret("BASEROW_CONSUMPTION_TABLE_ID", "")).strip()


def _baserow_headers():
    token = _baserow_api_token()
    if not token:
        return None
    return {
        "Authorization": f"Token {token}",
        "Content-Type": "application/json",
    }


def _baserow_rows_url(table_id=None):
    table_id = table_id or _baserow_consumption_table_id()
    if not table_id:
        return None
    return f"{_baserow_api_url()}/api/database/rows/table/{table_id}/"


def _baserow_row_url(row_id, table_id=None):
    base_url = _baserow_rows_url(table_id)
    if not base_url or not row_id:
        return None
    return f"{base_url}{row_id}/"


def _normalize_remote_consumption_row(raw_row):
    insurer_name = str(raw_row.get("insurer_name", "")).strip()
    if not insurer_name:
        return None

    item = {"insurer_name": insurer_name}
    for field_key, _ in DASHBOARD_FIELDS:
        item[field_key] = str(raw_row.get(field_key, "")).strip()
    if "id" in raw_row:
        item["id"] = raw_row["id"]
    return item


def _load_baserow_consumption_rows():
    headers = _baserow_headers()
    rows_url = _baserow_rows_url()
    if not headers or not rows_url:
        return []

    params = {
        "user_field_names": "true",
        "size": "200",
    }
    response = requests.get(f"{rows_url}?{urlencode(params)}", headers=headers, timeout=20)
    response.raise_for_status()
    payload = response.json() or {}

    rows = []
    for raw_row in payload.get("results", []):
        normalized = _normalize_remote_consumption_row(raw_row)
        if normalized:
            rows.append(normalized)

    rows.sort(key=lambda item: item["insurer_name"].lower())
    return rows


def _find_baserow_consumption_row(insurer_name):
    for row in _load_baserow_consumption_rows():
        if row["insurer_name"].strip().lower() == insurer_name.strip().lower():
            return row
    return None


def _save_baserow_consumption_mapping(insurer_name, field_vals, original_insurer_name=None):
    headers = _baserow_headers()
    rows_url = _baserow_rows_url()
    if not headers or not rows_url:
        raise RuntimeError("Baserow write credentials are missing.")

    normalized = normalize_field_mapping(field_vals)
    payload = {"insurer_name": insurer_name.strip(), **normalized}
    lookup_name = original_insurer_name or insurer_name
    existing_row = _find_baserow_consumption_row(lookup_name)

    params = {"user_field_names": "true"}
    if existing_row and existing_row.get("id"):
        response = requests.patch(
            f"{_baserow_row_url(existing_row['id'])}?{urlencode(params)}",
            headers=headers,
            data=json.dumps(payload),
            timeout=20,
        )
    else:
        response = requests.post(
            f"{rows_url}?{urlencode(params)}",
            headers=headers,
            data=json.dumps(payload),
            timeout=20,
        )
    response.raise_for_status()


def _delete_baserow_consumption_mapping(insurer_name):
    headers = _baserow_headers()
    existing_row = _find_baserow_consumption_row(insurer_name)
    if not headers or not existing_row or not existing_row.get("id"):
        return

    response = requests.delete(_baserow_row_url(existing_row["id"]), headers=headers, timeout=20)
    response.raise_for_status()


def _get_gspread_client():
    creds_dict = _safe_secret_dict("GOOGLE_SERVICE_ACCOUNT")
    if not creds_dict:
        creds_json = _safe_secret("GOOGLE_SERVICE_ACCOUNT_JSON", "")
        if creds_json:
            try:
                creds_dict = json.loads(creds_json)
            except Exception:
                creds_dict = None

    if not creds_dict:
        return None

    import gspread

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return gspread.service_account_from_dict(creds_dict, scopes=scopes)


def _get_consumption_worksheet():
    client = _get_gspread_client()
    sheet_url = _consumption_sheet_url()
    if not client or not sheet_url:
        return None

    workbook = client.open_by_url(sheet_url)
    try:
        return workbook.worksheet("consumption_mappings")
    except Exception:
        worksheet = workbook.add_worksheet(title="consumption_mappings", rows=200, cols=30)
        headers = ["insurer_name"] + [field_key for field_key, _ in DASHBOARD_FIELDS]
        worksheet.update("A1", [headers])
        return worksheet


def _get_conn():
    db_url = os.getenv("DATABASE_URL", "")
    if not db_url:
        db_url = _safe_secret("DATABASE_URL", "")

    if db_url and db_url.startswith("postgres"):
        import psycopg2

        return psycopg2.connect(db_url), "pg"

    conn = sqlite3.connect(SQLITE_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn, "sqlite"


def _exec(sql, params=(), fetch=None):
    conn, _ = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        conn.commit()
        if fetch == "all":
            return [dict(r) for r in cur.fetchall()]
        if fetch in {"one", "one2"}:
            row = cur.fetchone()
            return dict(row) if row else None
        return None
    finally:
        conn.close()


def init_db():
    conn, mode = _get_conn()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS companies (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT UNIQUE NOT NULL,
            template_name TEXT,
            start_row     INTEGER DEFAULT 2,
            date_format   TEXT DEFAULT '%Y-%m-%d',
            notes         TEXT
        );

        CREATE TABLE IF NOT EXISTS column_mappings (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id    INTEGER NOT NULL,
            target_order  INTEGER NOT NULL,
            master_col    TEXT NOT NULL,
            label         TEXT DEFAULT '',
            FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS value_mappings (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id    INTEGER NOT NULL,
            master_col    TEXT NOT NULL,
            source_value  TEXT NOT NULL,
            target_value  TEXT NOT NULL,
            FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS consumption_mappings (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            insurer_name  TEXT UNIQUE NOT NULL,
            mapping_json  TEXT NOT NULL
        );
        """
    )

    if mode == "sqlite":
        existing_cols = {
            row["name"]
            for row in cur.execute("PRAGMA table_info(column_mappings)").fetchall()
        }
        if "label" not in existing_cols:
            cur.execute("ALTER TABLE column_mappings ADD COLUMN label TEXT DEFAULT ''")

    conn.commit()
    conn.close()


init_db()


def normalize_field_mapping(field_vals):
    cleaned = {}
    for field_key, _ in DASHBOARD_FIELDS:
        value = field_vals.get(field_key, "") if field_vals else ""
        cleaned[field_key] = str(value).strip()
    return cleaned


def using_google_consumption_backend():
    return bool(_consumption_sheet_url())


def using_baserow_consumption_backend():
    return bool(_baserow_api_token() and _baserow_consumption_table_id())


def using_remote_consumption_backend():
    return using_baserow_consumption_backend() or using_google_consumption_backend()


def consumption_remote_backend_label():
    if using_baserow_consumption_backend():
        return "Baserow"
    if using_google_consumption_backend():
        return "Server"
    return ""


def consumption_backend_is_writable():
    if using_baserow_consumption_backend():
        return True
    return _get_gspread_client() is not None


def current_consumption_backend():
    if using_baserow_consumption_backend():
        return "Baserow"
    if using_google_consumption_backend():
        return "Google Sheet"
    return "Local SQLite"


def _sheet_export_csv_url(sheet_url):
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_url)
    if not match:
        return None
    gid_match = re.search(r"[?&]gid=(\d+)", sheet_url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv&gid={gid}"


def _load_google_consumption_rows():
    if not using_google_consumption_backend():
        return []

    export_url = _sheet_export_csv_url(_consumption_sheet_url())
    if not export_url:
        return []

    response = requests.get(export_url, timeout=15)
    response.raise_for_status()
    df = pd.read_csv(StringIO(response.text)).fillna("")

    expected_cols = ["insurer_name"] + [field_key for field_key, _ in DASHBOARD_FIELDS]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    rows = []
    for _, row in df.iterrows():
        insurer_name = str(row.get("insurer_name", "")).strip()
        if not insurer_name:
            continue
        item = {"insurer_name": insurer_name}
        for field_key, _ in DASHBOARD_FIELDS:
            item[field_key] = str(row.get(field_key, "")).strip()
        rows.append(item)
    return rows


def get_all_companies():
    return _exec("SELECT * FROM companies ORDER BY name", fetch="all") or []


def get_company_by_id(cid):
    return _exec("SELECT * FROM companies WHERE id=?", (cid,), fetch="one")


def add_company(name, template_name, start_row, date_format, notes=""):
    _exec(
        """
        INSERT INTO companies (name,template_name,start_row,date_format,notes)
        VALUES (?,?,?,?,?)
        """,
        (name, template_name, start_row, date_format, notes),
    )


def update_company(cid, name, template_name, start_row, date_format, notes=""):
    _exec(
        """
        UPDATE companies
        SET name=?,template_name=?,start_row=?,date_format=?,notes=?
        WHERE id=?
        """,
        (name, template_name, start_row, date_format, notes, cid),
    )


def delete_company(cid):
    _exec("DELETE FROM companies WHERE id=?", (cid,))


def get_column_mappings(company_id):
    return (
        _exec(
            """
            SELECT * FROM column_mappings
            WHERE company_id=?
            ORDER BY target_order
            """,
            (company_id,),
            fetch="all",
        )
        or []
    )


def save_column_mappings(company_id, mappings):
    conn, _ = _get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM column_mappings WHERE company_id=?", (company_id,))
    for mapping in mappings:
        cur.execute(
            """
            INSERT INTO column_mappings (company_id,target_order,master_col,label)
            VALUES (?,?,?,?)
            """,
            (
                company_id,
                mapping["target_order"],
                mapping["master_col"],
                mapping.get("label", ""),
            ),
        )
    conn.commit()
    conn.close()


def get_value_mappings(company_id):
    return (
        _exec(
            """
            SELECT * FROM value_mappings
            WHERE company_id=?
            ORDER BY id
            """,
            (company_id,),
            fetch="all",
        )
        or []
    )


def save_value_mappings(company_id, mappings):
    conn, _ = _get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM value_mappings WHERE company_id=?", (company_id,))
    for mapping in mappings:
        cur.execute(
            """
            INSERT INTO value_mappings (company_id,master_col,source_value,target_value)
            VALUES (?,?,?,?)
            """,
            (
                company_id,
                mapping["master_col"],
                mapping["source_value"],
                mapping["target_value"],
            ),
        )
    conn.commit()
    conn.close()


def get_all_consumption_mappings():
    if using_baserow_consumption_backend():
        try:
            return _load_baserow_consumption_rows()
        except Exception:
            pass

    if using_google_consumption_backend():
        try:
            return _load_google_consumption_rows()
        except Exception:
            pass

    rows = (
        _exec(
            "SELECT * FROM consumption_mappings ORDER BY insurer_name",
            fetch="all",
        )
        or []
    )
    result = []
    for row in rows:
        item = dict(row)
        item.update(normalize_field_mapping(json.loads(item.get("mapping_json", "{}"))))
        result.append(item)
    return result


def get_consumption_mapping(insurer_name):
    if using_baserow_consumption_backend():
        try:
            return _find_baserow_consumption_row(insurer_name)
        except Exception:
            pass

    if using_google_consumption_backend():
        try:
            for row in _load_google_consumption_rows():
                if row["insurer_name"] == insurer_name:
                    return row
        except Exception:
            pass

    row = _exec(
        "SELECT * FROM consumption_mappings WHERE insurer_name=?",
        (insurer_name,),
        fetch="one",
    )
    if not row:
        return None
    row.update(normalize_field_mapping(json.loads(row.get("mapping_json", "{}"))))
    return row


def save_consumption_mapping(insurer_name, field_vals, original_insurer_name=None):
    if using_baserow_consumption_backend():
        _save_baserow_consumption_mapping(insurer_name, field_vals, original_insurer_name=original_insurer_name)
        return

    if using_google_consumption_backend() and consumption_backend_is_writable():
        worksheet = _get_consumption_worksheet()
        if worksheet:
            headers = worksheet.row_values(1)
            expected_headers = ["insurer_name"] + [field_key for field_key, _ in DASHBOARD_FIELDS]
            if not headers:
                worksheet.update("A1", [expected_headers])
                headers = expected_headers

            normalized = normalize_field_mapping(field_vals)
            row_map = {"insurer_name": insurer_name.strip(), **normalized}
            row_values = [row_map.get(header, "") for header in headers]

            existing_col = worksheet.col_values(1)
            existing_row_idx = None
            lookup_name = (original_insurer_name or insurer_name).strip()
            for idx, value in enumerate(existing_col[1:], start=2):
                if str(value).strip() == lookup_name:
                    existing_row_idx = idx
                    break

            if existing_row_idx:
                end_col = chr(64 + len(headers))
                worksheet.update(f"A{existing_row_idx}:{end_col}{existing_row_idx}", [row_values])
            else:
                worksheet.append_row(row_values, value_input_option="USER_ENTERED")
            return

    json_str = json.dumps(normalize_field_mapping(field_vals), ensure_ascii=False)
    conn, _ = _get_conn()
    cur = conn.cursor()
    lookup_name = (original_insurer_name or insurer_name).strip()
    cur.execute(
        "SELECT id FROM consumption_mappings WHERE insurer_name=?",
        (lookup_name,),
    )
    if cur.fetchone():
        cur.execute(
            """
            UPDATE consumption_mappings
            SET insurer_name=?, mapping_json=?
            WHERE insurer_name=?
            """,
            (insurer_name, json_str, lookup_name),
        )
    else:
        cur.execute(
            """
            INSERT INTO consumption_mappings (insurer_name,mapping_json)
            VALUES (?,?)
            """,
            (insurer_name, json_str),
        )
    conn.commit()
    conn.close()


def delete_consumption_mapping(insurer_name):
    if using_baserow_consumption_backend():
        _delete_baserow_consumption_mapping(insurer_name)
        return

    if using_google_consumption_backend() and consumption_backend_is_writable():
        worksheet = _get_consumption_worksheet()
        if worksheet:
            existing_col = worksheet.col_values(1)
            for idx, value in enumerate(existing_col[1:], start=2):
                if str(value).strip() == insurer_name.strip():
                    worksheet.delete_rows(idx)
                    return

    _exec("DELETE FROM consumption_mappings WHERE insurer_name=?", (insurer_name,))
