#!/usr/bin/env python3
"""Weekly Retail Business Dashboard refresh.

Workflow:
1. Run the monthly Finance Actuals query and compare its newest PERIOD with the
   local Marketing_Spend_2026YTD.csv. Replace the CSV only when Snowflake has a
   strictly newer month.
2. Run the weekly Retail Dashboard query and save its result as a dated CSV.
3. Call the existing build_outputs.py to generate the dated Excel and HTML.

The Snowflake identity only needs SELECT access to objects referenced by the
two SQL files, plus USAGE on its warehouse/database/schemas.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import subprocess
import sys
import tempfile
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Sequence


ROOT = Path(__file__).resolve().parent
SQL_DIR = ROOT / "sql"
WEEKLY_SQL = SQL_DIR / "weekly_retail_actuals.sql"
FINANCE_SQL = SQL_DIR / "finance_actuals.sql"
BUILDER = ROOT / "scripts" / "build_outputs.py"
FINANCE_CSV = ROOT / "inputs" / "finance" / "Marketing_Spend_2026YTD.csv"
ENV_FILE = ROOT / ".env"

FINANCE_COLUMNS = {"PERIOD", "CATEGORY", "CHANNEL", "MDF_SPLIT", "SPEND"}
WEEKLY_REQUIRED_COLUMNS = {"WEEK_START", "DASHBOARD_CHANNEL"}


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def load_dotenv(path: Path = ENV_FILE) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"Snowflake configuration not found: {path}. Copy .env.example to .env first."
        )
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = re.split(r"\s+#", value.strip(), maxsplit=1)[0].strip().strip("'\"")
        if key:
            os.environ.setdefault(key, value)


def load_select_sql(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"SQL file not found: {path}")
    sql = path.read_text(encoding="utf-8").strip()
    sql_without_comments = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
    sql_without_comments = re.sub(r"--[^\n]*", " ", sql_without_comments).strip()
    first_word = re.match(r"([A-Za-z]+)", sql_without_comments)
    if not first_word or first_word.group(1).upper() not in {"SELECT", "WITH"}:
        raise ValueError(f"Only read-only SELECT/WITH SQL is allowed: {path.name}")
    return sql.rstrip(";\n ")


def _private_key_bytes():
    from cryptography.hazmat.primitives import serialization

    raw_key = os.environ.get("SNOWFLAKE_PRIVATE_KEY")
    if not raw_key:
        raise ValueError("SNOWFLAKE_PRIVATE_KEY is required for keypair auth.")

    # Keep the PEM on one line in .env, with each original line break written
    # as the two characters \n. Environment variables supplied by another
    # mechanism may already contain real line breaks, which also work.
    pem_bytes = raw_key.replace("\\n", "\n").strip().encode("utf-8")
    passphrase = (
        os.environ.get("SNOWFLAKE_PRIVATE_KEY_PASSPHRASE")
        or os.environ.get("SNOWFLAKE_PRIVATE_KEY_PWD")
    )
    private_key = serialization.load_pem_private_key(
        pem_bytes, password=passphrase.encode() if passphrase else None
    )
    return private_key.private_bytes(
        encoding=serialization.Encoding.DER,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )


def connect():
    try:
        import snowflake.connector
    except ImportError as exc:
        raise RuntimeError(
            "Missing Snowflake connector. Run .venv/bin/pip install -r requirements.txt"
        ) from exc

    load_dotenv()
    required = ["SNOWFLAKE_ACCOUNT", "SNOWFLAKE_USER", "SNOWFLAKE_WAREHOUSE"]
    missing = [key for key in required if not os.environ.get(key)]
    if missing:
        raise ValueError(f"Missing Snowflake settings: {', '.join(missing)}")

    params = {
        "account": os.environ["SNOWFLAKE_ACCOUNT"],
        "user": os.environ["SNOWFLAKE_USER"],
        "role": os.environ.get("SNOWFLAKE_ROLE"),
        "warehouse": os.environ["SNOWFLAKE_WAREHOUSE"],
        "database": os.environ.get("SNOWFLAKE_DATABASE", "DATA_MART"),
        "schema": os.environ.get("SNOWFLAKE_SCHEMA", "FINANCE"),
        "client_session_keep_alive": False,
        "session_parameters": {"QUERY_TAG": "weekly_retail_dashboard_refresh"},
    }
    params = {key: value for key, value in params.items() if value is not None}

    auth = os.environ.get("SNOWFLAKE_AUTH", "keypair").lower()
    if auth == "keypair":
        params["private_key"] = _private_key_bytes()
    elif auth == "password":
        password = os.environ.get("SNOWFLAKE_PASSWORD")
        if not password:
            raise ValueError("SNOWFLAKE_PASSWORD is required for password auth.")
        params["password"] = password
    elif auth == "externalbrowser":
        params["authenticator"] = "externalbrowser"
    else:
        raise ValueError(
            f"Unsupported SNOWFLAKE_AUTH={auth!r}; use keypair, password, or externalbrowser."
        )

    log(
        "Connecting to Snowflake "
        f"as {params['user']} / role {params.get('role') or '(default)'}..."
    )
    return snowflake.connector.connect(**params)


def execute_query(conn, sql_path: Path) -> tuple[list[str], list[Sequence]]:
    sql = load_select_sql(sql_path)
    log(f"Running {sql_path.name}...")
    cur = conn.cursor()
    try:
        cur.execute(sql)
        columns = []
        for col in cur.description:
            name = getattr(col, "name", None)
            columns.append(name if name is not None else col[0])
        rows = cur.fetchall()
    finally:
        cur.close()
    log(f"  returned {len(rows):,} rows.")
    return columns, rows


def parse_period_month(value) -> date:
    if isinstance(value, datetime):
        return value.date().replace(day=1)
    if isinstance(value, date):
        return value.replace(day=1)
    text = str(value).strip()

    # Finance exports use English month labels such as "JUN 2026". Parse
    # those explicitly so the result is chronological and independent of the
    # machine's locale; never compare or sort the raw labels as strings.
    month_numbers = {
        "JAN": 1,
        "JANUARY": 1,
        "FEB": 2,
        "FEBRUARY": 2,
        "MAR": 3,
        "MARCH": 3,
        "APR": 4,
        "APRIL": 4,
        "MAY": 5,
        "JUN": 6,
        "JUNE": 6,
        "JUL": 7,
        "JULY": 7,
        "AUG": 8,
        "AUGUST": 8,
        "SEP": 9,
        "SEPT": 9,
        "SEPTEMBER": 9,
        "OCT": 10,
        "OCTOBER": 10,
        "NOV": 11,
        "NOVEMBER": 11,
        "DEC": 12,
        "DECEMBER": 12,
    }
    english_month = re.fullmatch(r"([A-Za-z]+)\s+(\d{4})", text)
    if english_month:
        month_number = month_numbers.get(english_month.group(1).upper())
        if month_number is not None:
            return date(int(english_month.group(2)), month_number, 1)

    for fmt in ("%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(text, fmt).date().replace(day=1)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized PERIOD value: {value!r}")


def newest_local_finance_month(path: Path = FINANCE_CSV) -> date | None:
    if not path.exists() or path.stat().st_size == 0:
        return None
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or "PERIOD" not in {name.upper() for name in reader.fieldnames}:
            raise ValueError(f"{path.name} is missing PERIOD.")
        period_field = next(name for name in reader.fieldnames if name.upper() == "PERIOD")
        months = [parse_period_month(row[period_field]) for row in reader if row.get(period_field)]
    return max(months) if months else None


def _csv_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, "f")
    return value


def write_csv_atomic(path: Path, columns: Sequence[str], rows: Iterable[Sequence]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8-sig",
            newline="",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            writer = csv.writer(handle)
            writer.writerow(columns)
            writer.writerows([_csv_value(value) for value in row] for row in rows)
        os.replace(temp_path, path)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def refresh_finance_if_new_month(
    columns: Sequence[str], rows: Sequence[Sequence], path: Path = FINANCE_CSV
) -> bool:
    upper_columns = [column.upper() for column in columns]
    missing = FINANCE_COLUMNS - set(upper_columns)
    if missing:
        raise ValueError(f"Finance query is missing columns: {sorted(missing)}")
    if not rows:
        raise ValueError("Finance query returned no rows; local CSV was not changed.")

    period_index = upper_columns.index("PERIOD")
    remote_latest = max(parse_period_month(row[period_index]) for row in rows)
    local_latest = newest_local_finance_month(path)
    local_label = local_latest.strftime("%b %Y") if local_latest else "none"
    remote_label = remote_latest.strftime("%b %Y")
    log(f"Finance month check: local={local_label}, Snowflake={remote_label}.")

    if local_latest is not None and remote_latest <= local_latest:
        log("  no newer month; Marketing_Spend_2026YTD.csv was left untouched.")
        return False

    write_csv_atomic(path, upper_columns, rows)
    log(f"  updated {path.name} through {remote_label}. ✅")
    return True


def validate_weekly_result(columns: Sequence[str], rows: Sequence[Sequence]) -> None:
    upper_columns = {column.upper() for column in columns}
    missing = WEEKLY_REQUIRED_COLUMNS - upper_columns
    if missing:
        raise ValueError(f"Weekly query is missing columns: {sorted(missing)}")
    if not rows:
        raise ValueError("Weekly dashboard query returned no rows.")


def build_dashboard(csv_path: Path, snapshot_date: str) -> None:
    command = [
        sys.executable,
        str(BUILDER),
        str(csv_path),
        "--snapshot-date",
        snapshot_date,
    ]
    log("Building the Excel and HTML dashboard...")
    subprocess.run(command, cwd=ROOT, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Refresh monthly Finance data when needed, then build the weekly Retail Dashboard."
    )
    parser.add_argument(
        "--finance-check-only",
        action="store_true",
        help="Check/update Marketing_Spend_2026YTD.csv, but do not run the weekly dashboard.",
    )
    parser.add_argument(
        "--snapshot-date",
        help="Dashboard date YYYY-MM-DD (default: today).",
    )
    args = parser.parse_args()

    snapshot_date = args.snapshot_date or date.today().isoformat()
    date.fromisoformat(snapshot_date)  # validate before connecting

    conn = connect()
    try:
        finance_columns, finance_rows = execute_query(conn, FINANCE_SQL)
        refresh_finance_if_new_month(finance_columns, finance_rows)

        if args.finance_check_only:
            log("Finance-only check completed. ✅")
            return 0

        weekly_columns, weekly_rows = execute_query(conn, WEEKLY_SQL)
    finally:
        conn.close()

    validate_weekly_result(weekly_columns, weekly_rows)
    stamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    with tempfile.TemporaryDirectory(prefix="retail-dashboard-") as temp_dir:
        weekly_csv = Path(temp_dir) / f"Retail Business Dashboard_{stamp}.csv"
        write_csv_atomic(
            weekly_csv,
            [column.upper() for column in weekly_columns],
            weekly_rows,
        )
        log("Prepared temporary weekly Snowflake result.")
        build_dashboard(weekly_csv, snapshot_date)

    output_dir = ROOT / "outputs" / snapshot_date
    log(f"Weekly Retail Dashboard completed: {output_dir} ✅")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as exc:
        log(f"FAILED: dashboard builder exited with status {exc.returncode}.")
        raise SystemExit(exc.returncode)
    except Exception as exc:
        log(f"FAILED: {exc}")
        raise SystemExit(1)
