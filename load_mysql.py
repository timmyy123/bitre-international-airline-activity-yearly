"""
Load the yearly workbook straight into MySQL with Python (no generated SQL file needed).

Reads the 7 data sheets of output/International_airline_activity_yearly_ALL.xlsx and inserts
them into the tables defined in sql/create_tables.sql, in one transaction.

    pip install -r requirements.txt
    python load_mysql.py --create                       # create database/tables if missing, then load
    python load_mysql.py                                # reload data into existing tables
    python load_mysql.py --host db.example --user me    # another server
    python load_mysql.py --dry-run                      # read and check the workbook, don't connect

Login: --host/--port/--user/--database, or the environment variables MYSQL_HOST, MYSQL_PORT,
MYSQL_USER, MYSQL_DATABASE. The password comes from MYSQL_PASSWORD or is asked for at the prompt;
it is never stored in this file.

Re-runnable: each table is emptied (DELETE) and reloaded inside the same transaction, so a failure
leaves the previous data in place.
"""
import argparse
import getpass
import os
import re
import sys

import export_sql as ex

HERE = os.path.dirname(os.path.abspath(__file__))


def schema_statements(database):
    """Statements from sql/create_tables.sql, made safe to re-run (CREATE TABLE IF NOT EXISTS)."""
    text = open(ex.SCHEMA, encoding="utf-8").read()
    text = "\n".join(line for line in text.splitlines() if not line.lstrip().startswith("--"))
    text = text.replace("airline_activity", database)
    text = re.sub(r"CREATE TABLE (?!IF NOT EXISTS)", "CREATE TABLE IF NOT EXISTS ", text)
    # split only on a ';' that ends a line — column COMMENTs may contain ';'
    return [s.strip() for s in re.split(r";[ \t]*(?:\n|$)", text) if s.strip()]


def prepared_tables(workbook):
    """[(table, columns, rows)] with each value converted to its column's type."""
    out = []
    for sheet, table, cols, types, recs in ex.table_rows(workbook):
        rows = [tuple(ex.py_value(v, *t) for v, t in zip(rec, types)) for rec in recs]
        out.append((table, cols, rows))
    return out


def load(conn, tables, database, create=False):
    cur = conn.cursor()
    if create:
        for stmt in schema_statements(database):
            cur.execute(stmt)
        print(f"schema ready: database {database}, {sum(1 for s in schema_statements(database) if s.startswith('CREATE TABLE'))} tables")
    cur.execute(f"USE `{database}`")
    try:
        for table, cols, rows in tables:
            cur.execute(f"DELETE FROM {table}")
            sql = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join(['%s'] * len(cols))})"
            for i in range(0, len(rows), 1000):
                cur.executemany(sql, rows[i:i + 1000])
            print(f"  {table}: {len(rows)} rows")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    # check what landed
    problems = 0
    for table, cols, rows in tables:
        cur.execute(f"SELECT COUNT(*) FROM {table}")
        n = cur.fetchone()[0]
        if n != len(rows):
            problems += 1
            print(f"  CHECK FAILED {table}: {n} rows in MySQL, {len(rows)} in workbook")
    cur.execute("SELECT year, passengers_total FROM summary_by_year ORDER BY year DESC LIMIT 1")
    y, p = cur.fetchone()
    cur.execute("SELECT SUM(total_passengers) FROM t1_traffic_by_operator WHERE row_type = 'detail' AND year = %s", (y,))
    p1 = cur.fetchone()[0]
    if int(p1) != int(p):
        problems += 1
        print(f"  CHECK FAILED {y}: Table 1 detail rows sum to {p1}, summary says {p}")
    print(f"loaded and checked: {problems} problem(s); {y} passengers = {int(p):,}")
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--host", default=os.environ.get("MYSQL_HOST", "localhost"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("MYSQL_PORT", 3306)))
    ap.add_argument("--user", default=os.environ.get("MYSQL_USER", "root"))
    ap.add_argument("--database", default=os.environ.get("MYSQL_DATABASE", "airline_activity"))
    ap.add_argument("--workbook", default=ex.WORKBOOK)
    ap.add_argument("--create", action="store_true", help="create the database and tables first if they don't exist")
    ap.add_argument("--dry-run", action="store_true", help="read and convert the workbook only; don't connect")
    args = ap.parse_args()

    if not re.fullmatch(r"\w+", args.database):
        sys.exit("--database may only contain letters, digits and _")
    tables = prepared_tables(args.workbook)
    print("read workbook: " + ", ".join(f"{t} {len(r)}" for t, _, r in tables))
    if args.dry_run:
        return

    try:
        import pymysql
    except ImportError:
        sys.exit("PyMySQL is not installed: pip install -r requirements.txt")
    password = os.environ.get("MYSQL_PASSWORD")
    if password is None:
        password = getpass.getpass(f"MySQL password for {args.user}@{args.host}: ")
    conn = pymysql.connect(host=args.host, port=args.port, user=args.user, password=password,
                           charset="utf8mb4", autocommit=False)
    try:
        problems = load(conn, tables, args.database, create=args.create)
    finally:
        conn.close()
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
