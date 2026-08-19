from pathlib import Path

from app.database import get_connection, initialize_database
from app.services.excel_importer import import_workbook


def main() -> None:
    initialize_database()
    repaired = []

    with get_connection() as connection:
        batches = connection.execute(
            """
            SELECT * FROM import_batches
            WHERE status = 'completed'
              AND imported_rows = 0
              AND duplicate_rows > 0
              AND NOT EXISTS (
                SELECT 1 FROM sales_records
                WHERE import_batch_id = import_batches.id
              )
            """
        ).fetchall()

        for batch in batches:
            connection.execute(
                "DELETE FROM import_issues WHERE import_batch_id = ?",
                (batch["id"],),
            )
            result = import_workbook(
                connection,
                Path("uploads") / batch["stored_filename"],
                batch["id"],
            )
            connection.execute(
                """
                UPDATE import_batches
                SET total_rows = ?, imported_rows = ?,
                    duplicate_rows = ?, issue_rows = ?
                WHERE id = ?
                """,
                (
                    result.total_rows,
                    result.imported_rows,
                    result.duplicate_rows,
                    result.issue_rows,
                    batch["id"],
                ),
            )
            connection.execute(
                """
                INSERT INTO audit_logs (action, target_type, target_id, description)
                VALUES ('repair_import', 'import_batch', ?, ?)
                """,
                (
                    batch["id"],
                    f"취소 배치 중복 판정 오류 복구: {result.imported_rows}건 반영",
                ),
            )
            repaired.append((batch["id"], result.imported_rows))

    print({"repaired": repaired})


if __name__ == "__main__":
    main()
