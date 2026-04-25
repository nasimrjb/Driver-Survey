"""
Load Driver Survey processed CSVs into Cab_Studies database.

Creates 6 separate tables, one per CSV file:
  - Cab.DriverSurvey_ShortMain
  - Cab.DriverSurvey_ShortRare
  - Cab.DriverSurvey_WideMain
  - Cab.DriverSurvey_WideRare
  - Cab.DriverSurvey_LongMain
  - Cab.DriverSurvey_LongRare

INCREMENTAL MODE (default):
  - Tables are created only if they don't exist.
  - Existing recordIDs are fetched from the database.
  - Only NEW rows (recordIDs not yet in the table) are inserted.
  - Safe to re-run after adding new survey weeks to the CSVs.

FULL RELOAD MODE (--full-reload):
  - Drops and recreates all 6 tables, then loads everything.

Column types are determined from Sources/column_rename_mapping.json
and auto-detected for computed/derived columns not in the JSON.

Usage:
    python load_to_database.py              # Incremental (only new rows)
    python load_to_database.py --full-reload   # Drop & reload everything
"""

import json
import sys
import pyodbc
import pandas as pd
import numpy as np

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = r"D:\Work\Driver Survey"
MAPPING_PATH = rf"{BASE_DIR}\Sources\column_rename_mapping.json"
SERVER = "192.168.18.37"
DATABASE = "Cab_Studies"
SCHEMA = "Cab"
SQL_USER = "nasim.rajabi"
SQL_PASS = "ISLRv2_corrected_June_2023"
BATCH_SIZE = 1000

# CSV file → table name mapping
TABLES = {
    "short_main": {
        "csv": rf"{BASE_DIR}\processed\short_survey_main.csv",
        "table": "DriverSurvey_ShortMain",
    },
    "short_rare": {
        "csv": rf"{BASE_DIR}\processed\short_survey_rare.csv",
        "table": "DriverSurvey_ShortRare",
    },
    "wide_main": {
        "csv": rf"{BASE_DIR}\processed\wide_survey_main.csv",
        "table": "DriverSurvey_WideMain",
    },
    "wide_rare": {
        "csv": rf"{BASE_DIR}\processed\wide_survey_rare.csv",
        "table": "DriverSurvey_WideRare",
    },
    "long_main": {
        "csv": rf"{BASE_DIR}\processed\long_survey_main.csv",
        "table": "DriverSurvey_LongMain",
    },
    "long_rare": {
        "csv": rf"{BASE_DIR}\processed\long_survey_rare.csv",
        "table": "DriverSurvey_LongRare",
    },
}


# ============================================================
# COLUMN TYPE DETECTION
# ============================================================

# Computed/derived columns not in the JSON mapping
EXTRA_FLOAT = {
    "snapp_ride", "tapsi_ride", "tapsi_commfree_disc_ride",
    "snapp_commfree_disc_ride", "snapp_diff_commfree", "tapsi_diff_commfree",
    "snapp_commfree", "tapsi_commfree", "snapp_incentive", "tapsi_incentive",
    "wheel", "snapp_LOC", "tapsi_LOC", "edu", "marr_stat",
}
EXTRA_INT = {"weeknumber", "joint_by_signup", "active_joint", "yearweek"}
EXTRA_STR = {
    "_source_file", "age_group", "cooperation_type", "question", "answer",
    "question_type", "snapp_incentive_category", "tapsi_incentive_category",
    "gotbonus_snapp", "gotbonus_tapsi", "city_clean",
}


def load_json_dtypes():
    """Load column dtype mapping from the JSON file."""
    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        mapping = json.load(f)
    dtypes = {}
    for col, info in mapping.items():
        if not col.startswith("ignore"):
            dtypes[col] = info.get("dtype", "str")
    return dtypes


def get_col_type(col_name, json_dtypes, sample_values=None):
    """Determine the logical type for a column.

    Returns one of: 'int', 'float', 'datetime', 'str'.
    """
    if col_name == "recordID":
        return "int"
    if col_name == "datetime":
        return "datetime"
    if col_name in json_dtypes:
        return json_dtypes[col_name]
    if col_name in EXTRA_FLOAT:
        return "float"
    if col_name in EXTRA_INT:
        return "int"
    if col_name in EXTRA_STR:
        return "str"

    # Auto-detect from sample values
    if sample_values:
        non_empty = {str(s).strip() for s in sample_values
                     if pd.notna(s) and str(s).strip() and str(s).strip() != "nan"}
        if not non_empty:
            return "str"
        if non_empty <= {"0", "1"}:
            return "int"
        if all(s.lstrip("-").isdigit() for s in non_empty):
            return "int"
        try:
            [float(s) for s in non_empty]
            return "float"
        except ValueError:
            pass
    return "str"


def detect_column_types(df, json_dtypes):
    """Return list of (col_name, col_type) for every column in the DataFrame."""
    col_types = []
    for c in df.columns:
        samples = df[c].dropna().astype(str).unique()[:50].tolist()
        col_types.append((c, get_col_type(c, json_dtypes, samples)))
    return col_types


# ============================================================
# SQL TYPE MAPPING
# ============================================================
SQL_TYPE_MAP = {
    "int": "INT NULL",
    "float": "FLOAT NULL",
    "datetime": "DATETIME NULL",
    "str": "NVARCHAR(500) NULL",
}


def get_sql_type(col_name, col_type):
    """Return the SQL Server column definition."""
    if col_name == "recordID":
        return "INT NOT NULL"
    return SQL_TYPE_MAP.get(col_type, "NVARCHAR(500) NULL")


# ============================================================
# DATABASE OPERATIONS
# ============================================================

def connect_db():
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={SQL_USER};"
        f"PWD={SQL_PASS};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


def table_exists(cursor, table_name):
    """Check if a table exists. Returns True/False."""
    cursor.execute(f"""
        SELECT CASE WHEN OBJECT_ID(N'[{SCHEMA}].[{table_name}]', N'U') IS NOT NULL
                    THEN 1 ELSE 0 END
    """)
    return cursor.fetchone()[0] == 1


def get_existing_record_ids(cursor, table_name):
    """Fetch ALL existing recordIDs from a table as a Python set."""
    cursor.execute(f"SELECT DISTINCT [recordID] FROM [{SCHEMA}].[{table_name}]")
    return {row[0] for row in cursor.fetchall()}


def create_table(cursor, conn, table_name, col_types):
    """Create a table with properly typed columns (IF NOT EXISTS)."""
    col_defs = []
    for col_name, col_type in col_types:
        sql_type = get_sql_type(col_name, col_type)
        col_defs.append(f"    [{col_name}] {sql_type}")

    create_sql = f"""
    IF OBJECT_ID(N'[{SCHEMA}].[{table_name}]', N'U') IS NULL
    BEGIN
        CREATE TABLE [{SCHEMA}].[{table_name}] (
    {',\n'.join(col_defs)}
        );
    END
    """
    cursor.execute(create_sql)
    conn.commit()

    # Print type summary
    type_counts = {}
    for _, ct in col_types:
        type_counts[ct] = type_counts.get(ct, 0) + 1
    print(f"  Created [{SCHEMA}].[{table_name}] — "
          f"{len(col_types)} columns: {type_counts}")


def drop_table(cursor, conn, table_name):
    """Drop a table if it exists."""
    cursor.execute(f"""
        IF OBJECT_ID(N'[{SCHEMA}].[{table_name}]', N'U') IS NOT NULL
            DROP TABLE [{SCHEMA}].[{table_name}];
    """)
    conn.commit()
    print(f"  Dropped [{SCHEMA}].[{table_name}]")


def prepare_dataframe(df, col_types):
    """Prepare DataFrame for SQL insert.

    Data is sent as strings — SQL Server handles implicit conversion
    to the proper column types (INT, FLOAT, DATETIME, NVARCHAR).
    This avoids Python/pyodbc type-casting edge cases (NaN→int overflow, etc.).

    Datetime columns are normalized to 'YYYY-MM-DD HH:MM:SS' format
    (SQL Server DATETIME cannot handle nanosecond-precision strings).

    Returns df_ready with empty strings replaced by None (NULL).
    """
    columns = [c for c, _ in col_types]
    df_ready = df[columns].copy()

    # Normalize datetime columns: truncate nanoseconds
    for col_name, col_type in col_types:
        if col_type == "datetime":
            dt = pd.to_datetime(df_ready[col_name], errors="coerce")
            df_ready[col_name] = dt.dt.strftime("%Y-%m-%d %H:%M:%S")
            # strftime turns NaT into 'NaT' string — fix that
            df_ready[col_name] = df_ready[col_name].replace("NaT", "")

    # Replace empty strings with None (SQL NULL)
    # Note: df.where() can leave float NaN in mixed-type columns,
    # so we convert via applymap to ensure proper None values.
    df_ready = df_ready.replace("", None)
    df_ready = df_ready.where(df_ready.notna(), None)
    # Belt-and-suspenders: ensure no numpy NaN survives
    df_ready = df_ready.map(lambda x: None if x is np.nan or (isinstance(x, float) and np.isnan(x)) else x)
    return df_ready


COMMIT_EVERY = 20  # commit every N batches so progress isn't lost on disconnect
MAX_RETRIES = 3    # retry on network failure


def insert_data(conn, df, col_types, table_name):
    """Bulk insert DataFrame into the table using fast_executemany.
    Commits periodically and retries on network failure."""
    columns = [c for c, _ in col_types]

    if len(df) == 0:
        print(f"  Nothing to insert into [{SCHEMA}].[{table_name}].")
        return conn

    print(f"  Inserting {len(df):,} rows into [{SCHEMA}].[{table_name}]...")

    # Prepare data: empty → NULL, send as strings (SQL Server converts implicitly)
    print(f"  Preparing data...")
    df_ready = prepare_dataframe(df, col_types)

    col_list = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join(["?"] * len(columns))
    sql = f"INSERT INTO [{SCHEMA}].[{table_name}] ({col_list}) VALUES ({placeholders})"

    total = len(df_ready)
    rows_inserted = 0
    batch_num = 0

    cursor = conn.cursor()
    cursor.fast_executemany = True
    cursor.setinputsizes([(pyodbc.SQL_WVARCHAR, 500, 0)] * len(columns))

    for start in range(0, total, BATCH_SIZE):
        batch_num += 1
        batch = df_ready.iloc[start:start + BATCH_SIZE]
        rows = batch.values.tolist()

        # Retry logic for network drops
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                cursor.executemany(sql, rows)
                break
            except (pyodbc.OperationalError, pyodbc.Error) as e:
                err_msg = str(e)
                if "Communication link failure" in err_msg or "08S01" in err_msg:
                    print(f"    [!] Network error at batch {batch_num}, "
                          f"retry {attempt}/{MAX_RETRIES}...", flush=True)
                    import time
                    time.sleep(3)
                    try:
                        conn = connect_db()
                        cursor = conn.cursor()
                        cursor.fast_executemany = True
                        cursor.setinputsizes(
                            [(pyodbc.SQL_WVARCHAR, 500, 0)] * len(columns))
                        print(f"    [!] Reconnected.", flush=True)
                    except Exception:
                        if attempt == MAX_RETRIES:
                            raise
                else:
                    raise

        rows_inserted += len(rows)
        pct = rows_inserted / total * 100
        if batch_num % 50 == 0 or rows_inserted == total:
            print(f"    {rows_inserted:,} / {total:,} rows  ({pct:.1f}%)",
              flush=True)

        # Commit periodically to save progress
        if batch_num % COMMIT_EVERY == 0:
            conn.commit()

    conn.commit()
    cursor.close()
    print(f"  Done: {rows_inserted:,} rows inserted.")
    return conn  # return conn in case it was reconnected


# ============================================================
# MAIN
# ============================================================

def main():
    full_reload = "--full-reload" in sys.argv

    print("Loading column type mapping...")
    json_dtypes = load_json_dtypes()
    print(f"  {len(json_dtypes)} columns mapped from JSON.\n")

    if full_reload:
        print("*** FULL RELOAD MODE — all tables will be dropped and recreated ***\n")
    else:
        print("*** INCREMENTAL MODE — only new recordIDs will be inserted ***\n")

    print("Connecting to database...")
    conn = connect_db()
    print("  Connected.\n")

    for key, info in TABLES.items():
        csv_path = info["csv"]
        table_name = info["table"]
        print(f"=== {table_name} ({key}) ===")

        # Load CSV
        print(f"  Loading {csv_path}...")
        df = pd.read_csv(csv_path, encoding="utf-8-sig",
                         dtype=str, keep_default_na=False)
        print(f"  CSV shape: {df.shape}")

        # Detect column types
        col_types = detect_column_types(df, json_dtypes)

        cursor = conn.cursor()

        if full_reload:
            # Drop and recreate
            drop_table(cursor, conn, table_name)
            create_table(cursor, conn, table_name, col_types)
            df_new = df  # insert everything

        else:
            # Incremental: create if not exists, then find new rows
            exists = table_exists(cursor, table_name)

            if not exists:
                print(f"  Table does not exist — creating it.")
                create_table(cursor, conn, table_name, col_types)
                df_new = df  # insert everything (fresh table)
            else:
                # Fetch existing recordIDs
                print(f"  Fetching existing recordIDs...")
                existing_ids = get_existing_record_ids(cursor, table_name)
                print(f"  Found {len(existing_ids):,} existing recordIDs.")

                # Convert CSV recordID to int for comparison
                csv_ids = df["recordID"].astype(int)
                mask = ~csv_ids.isin(existing_ids)
                df_new = df[mask].copy()

                new_ids = csv_ids[mask].nunique()
                print(f"  New recordIDs to insert: {new_ids:,} "
                      f"({len(df_new):,} rows)")

                if len(df_new) == 0:
                    print(f"  UP TO DATE — no new data.\n")
                    cursor.close()
                    continue

        cursor.close()

        # Insert new data
        conn = insert_data(conn, df_new, col_types, table_name)

        # Compute yearweek for newly inserted rows that don't have it yet.
        # Priority:
        #   1. _source_file matches 'Data Raw Driver NNNN.xlsx' → 4-digit yearweek at chars 17-20
        #   2. ISO week of [datetime] column → CAST(RIGHT(YEAR(dt)*100+ISO_WK,4) …) via a formula
        #   3. Fallback: 2500 + weeknumber
        if table_name in ("DriverSurvey_ShortMain", "DriverSurvey_ShortRare",
                          "DriverSurvey_WideMain", "DriverSurvey_WideRare",
                          "DriverSurvey_LongMain", "DriverSurvey_LongRare"):
            print(f"  Computing yearweek for new rows...")
            cursor = conn.cursor()
            cursor.execute(f"""
                UPDATE [{SCHEMA}].[{table_name}]
                SET [yearweek] =
                    CASE
                        -- Pattern: 'Data Raw Driver NNNN.xlsx'  (new-format files from 2026 onward)
                        WHEN [_source_file] LIKE 'Data Raw Driver [0-9][0-9][0-9][0-9].xlsx'
                            THEN CAST(SUBSTRING([_source_file], 17, 4) AS INT)
                        -- Datetime-based ISO yearweek (handles year boundaries correctly)
                        WHEN TRY_CAST([datetime] AS DATETIME) IS NOT NULL
                            THEN (
                                YEAR(DATEADD(dd, 26 - DATEPART(iso_week, TRY_CAST([datetime] AS DATETIME)), TRY_CAST([datetime] AS DATETIME))) % 100
                            ) * 100 + DATEPART(iso_week, TRY_CAST([datetime] AS DATETIME))
                        -- Fallback: assume 2025
                        ELSE 2500 + TRY_CAST([weeknumber] AS INT)
                    END
                WHERE [yearweek] IS NULL
            """)
            rows_updated = cursor.rowcount
            conn.commit()
            cursor.close()
            print(f"  yearweek computed for {rows_updated:,} rows.")

        # Verify final count
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM [{SCHEMA}].[{table_name}]")
        total = cursor.fetchone()[0]
        cursor.close()
        print(f"  Total rows in table: {total:,}\n")

    conn.close()
    mode = "FULL RELOAD" if full_reload else "INCREMENTAL"
    print(f"All done! ({mode} mode)")


if __name__ == "__main__":
    main()
