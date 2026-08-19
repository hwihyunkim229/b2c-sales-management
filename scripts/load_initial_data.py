from pathlib import Path

from app.database import get_connection, initialize_database
from app.services.excel_importer import import_workbook


def main() -> None:
    source_path = Path("uploads/initial-source.xlsx")

    initialize_database()

    with get_connection() as connection:
        cursor = connection.execute(
            """
            INSERT INTO import_batches (
                original_filename,
                stored_filename,
                status
            ) VALUES (?, ?, 'processing')
            """,
            (
                "B2C_판매내역_자동템플릿_2607_v2.xlsx",
                source_path.name,
            ),
        )
        batch_id = cursor.lastrowid
        result = import_workbook(
            connection=connection,
            file_path=source_path,
            batch_id=batch_id,
        )
        connection.execute(
            """
            UPDATE import_batches
            SET
                status = 'completed',
                total_rows = ?,
                imported_rows = ?,
                duplicate_rows = ?
            WHERE id = ?
            """,
            (
                result.total_rows,
                result.imported_rows,
                result.duplicate_rows,
                batch_id,
            ),
        )
        record_count = connection.execute(
            "SELECT COUNT(*) FROM sales_records"
        ).fetchone()[0]

    print(result)
    print(f"Stored sales records: {record_count}")


if __name__ == "__main__":
    main()

