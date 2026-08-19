from pathlib import Path

from app.database import get_connection, initialize_database
from app.services.excel_importer import import_workbook


def main() -> None:
    initialize_database()
    connection = get_connection()
    connection.execute("BEGIN")

    try:
        cancelled = connection.execute(
            """
            SELECT stored_filename, total_rows
            FROM import_batches
            WHERE status = 'cancelled'
            ORDER BY id DESC LIMIT 1
            """
        ).fetchone()
        assert cancelled is not None
        connection.execute("UPDATE sales_records SET batch_is_active = 0")
        cursor = connection.execute(
            """
            INSERT INTO import_batches (original_filename, stored_filename, status)
            VALUES ('cancelled-reupload-test.xlsx', 'test.xlsx', 'processing')
            """
        )
        result = import_workbook(
            connection,
            Path("uploads") / cancelled["stored_filename"],
            cursor.lastrowid,
        )
        assert result.total_rows == cancelled["total_rows"]
        assert result.imported_rows == cancelled["total_rows"]
        assert result.duplicate_rows == 0
        print(result)
    finally:
        connection.rollback()
        connection.close()


if __name__ == "__main__":
    main()
