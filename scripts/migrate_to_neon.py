from __future__ import annotations

from argparse import ArgumentParser
from pathlib import Path
import os
import sqlite3
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


TABLES = (
    "import_batches",
    "users",
    "sales_records",
    "import_issues",
    "audit_logs",
    "hometax_records",
)


def source_columns(connection: sqlite3.Connection, table: str) -> list[str]:
    return [row["name"] for row in connection.execute(f"PRAGMA table_info({table})")]


def target_columns(connection, table: str) -> set[str]:
    return {
        row["column_name"]
        for row in connection.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = ?
            """,
            (table,),
        ).fetchall()
    }


def main() -> None:
    parser = ArgumentParser(description="SQLite 데이터를 Neon PostgreSQL로 이전합니다.")
    parser.add_argument(
        "--source",
        type=Path,
        default=PROJECT_ROOT / "data" / "sales.db",
    )
    parser.add_argument(
        "--uploads",
        type=Path,
        default=PROJECT_ROOT / "uploads",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="대상 DB가 비어 있는지 확인한 뒤 이전을 실행합니다.",
    )
    arguments = parser.parse_args()

    database_url = os.getenv("DATABASE_URL", "").strip()
    if not database_url.startswith(("postgresql://", "postgres://")):
        raise SystemExit("DATABASE_URL에 Neon PostgreSQL 연결 문자열을 설정해 주세요.")
    if not arguments.yes:
        raise SystemExit("확인 후 --yes 옵션을 붙여 다시 실행해 주세요.")
    if not arguments.source.exists():
        raise SystemExit(f"SQLite DB를 찾을 수 없습니다: {arguments.source}")

    from app.database import get_connection, initialize_database

    initialize_database()
    source = sqlite3.connect(arguments.source)
    source.row_factory = sqlite3.Row

    try:
        with get_connection() as target:
            occupied = {
                table: target.execute(f"SELECT COUNT(*) AS count FROM {table}")
                .fetchone()["count"]
                for table in TABLES
            }
            if any(occupied.values()):
                details = ", ".join(
                    f"{table}={count}" for table, count in occupied.items() if count
                )
                raise SystemExit(
                    f"대상 Neon DB가 비어 있지 않아 중단했습니다: {details}"
                )

            for table in TABLES:
                available = target_columns(target, table)
                sqlite_columns = source_columns(source, table)
                columns = [
                    column
                    for column in sqlite_columns
                    if column in available
                ]
                if table == "import_batches" and "source_content" not in columns:
                    columns.append("source_content")

                query_columns = [column for column in columns if column in sqlite_columns]
                rows = source.execute(
                    f"SELECT {', '.join(query_columns)} FROM {table} ORDER BY id"
                ).fetchall()
                placeholders = ", ".join("?" for _ in columns)
                insert_sql = (
                    f"INSERT INTO {table} ({', '.join(columns)}) "
                    f"VALUES ({placeholders})"
                )

                values = []
                for row in rows:
                    item = []
                    for column in columns:
                        if column != "source_content":
                            item.append(row[column])
                            continue
                        content = (
                            row["source_content"]
                            if "source_content" in sqlite_columns
                            else None
                        )
                        if content is None:
                            file_path = arguments.uploads / row["stored_filename"]
                            content = file_path.read_bytes() if file_path.exists() else None
                        item.append(content)
                    values.append(tuple(item))

                if values:
                    target.executemany(insert_sql, values)
                print(f"{table}: {len(values):,}행 이전")

            for table in TABLES:
                target.execute(
                    f"""SELECT setval(
                        pg_get_serial_sequence('{table}', 'id'),
                        COALESCE(MAX(id), 1),
                        MAX(id) IS NOT NULL
                    ) FROM {table}"""
                )
    finally:
        source.close()

    print("Neon 데이터 이전이 완료되었습니다.")


if __name__ == "__main__":
    main()
