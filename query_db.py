import argparse
import os
import sqlite3
from pathlib import Path

from kuaishou_client import load_runtime_env


ROOT_DIR = Path(__file__).resolve().parent


def resolve_db_path(value):
    load_runtime_env()
    db_path = Path(value or os.getenv("KS_DB_PATH", "data/kuaishou.db"))
    if not db_path.is_absolute():
        db_path = ROOT_DIR / db_path
    return db_path


def print_rows(cursor):
    columns = [item[0] for item in cursor.description or []]
    if columns:
        print("\t".join(columns))
    for row in cursor.fetchall():
        print("\t".join("" if value is None else str(value) for value in row))


def main():
    parser = argparse.ArgumentParser(description="查询快手 SQLite 落库数据")
    parser.add_argument("sql", nargs="*", help="要执行的 SQL；不传则执行 --file")
    parser.add_argument("--file", default="", help="SQL 文件路径")
    parser.add_argument("--db", default="", help="SQLite 数据库路径")
    args = parser.parse_args()

    db_path = resolve_db_path(args.db)
    sql = " ".join(args.sql) if args.sql else ""
    if args.file:
        file_path = Path(args.file)
        if not file_path.is_absolute():
            file_path = ROOT_DIR / file_path
        sql = file_path.read_text(encoding="utf-8-sig")
    if not sql.strip():
        raise SystemExit("请传入 SQL，或使用 --file 指定 SQL 文件")

    conn = sqlite3.connect(str(db_path))
    try:
        for statement in [item.strip().lstrip("\ufeff") for item in sql.split(";") if item.strip()]:
            cursor = conn.execute(statement)
            if cursor.description:
                print_rows(cursor)
            else:
                conn.commit()
                print(f"ok: {cursor.rowcount} rows affected")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
