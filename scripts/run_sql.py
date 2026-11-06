"""Load the competitor tracker CSVs into SQLite and run every query in sql/.

Usage:  python scripts/run_sql.py
Output: outputs/tracker.db and one CSV per query in outputs/sql_results/
"""
import csv
import sqlite3
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SQL = ROOT / "sql"
OUT = ROOT / "outputs"
DB_PATH = OUT / "tracker.db"

TABLES = ["companies", "industry_events"]


def to_value(text):
    """Empty cells become NULL, numbers become floats, everything else stays text."""
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return text


def load(conn):
    conn.executescript((SQL / "schema.sql").read_text())
    for table in TABLES:
        with open(DATA / f"{table}.csv", newline="") as f:
            reader = csv.reader(f)
            header = next(reader)
            rows = [[to_value(v) for v in row] for row in reader]
        marks = ", ".join("?" * len(header))
        conn.executemany(f"INSERT INTO {table} ({', '.join(header)}) VALUES ({marks})", rows)
        print(f"loaded {len(rows):>3} rows into {table}")
    conn.commit()


def run_queries(conn):
    results_dir = OUT / "sql_results"
    results_dir.mkdir(parents=True, exist_ok=True)
    for path in sorted(SQL.glob("0*.sql")):
        cur = conn.execute(path.read_text())
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
        with open(results_dir / f"{path.stem}.csv", "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(cols)
            w.writerows(rows)
        print(f"\n== {path.name} ==")
        print(" | ".join(cols))
        for r in rows:
            print(" | ".join("" if v is None else str(v) for v in r))


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    if DB_PATH.exists():
        DB_PATH.unlink()
    with sqlite3.connect(DB_PATH) as conn:
        load(conn)
        run_queries(conn)
