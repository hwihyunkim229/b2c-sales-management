from pathlib import Path
from tempfile import TemporaryDirectory

import app.database as database
import app.main as main_app


def main() -> None:
    original_database_path = database.DATABASE_PATH
    original_upload_directory = main_app.UPLOAD_DIRECTORY

    with TemporaryDirectory(ignore_cleanup_errors=True) as temporary_directory:
        root = Path(temporary_directory)
        database.DATABASE_PATH = root / "sales.db"
        main_app.UPLOAD_DIRECTORY = root / "uploads"
        main_app.UPLOAD_DIRECTORY.mkdir()

        try:
            database.initialize_database()
            with database.get_connection() as connection:
                source_batch_id = connection.execute(
                    """INSERT INTO import_batches
                    (original_filename, stored_filename, status)
                    VALUES ('source.xlsx', 'source.xlsx', 'completed')
                    RETURNING id"""
                ).fetchone()[0]
                source_record_id = connection.execute(
                    """INSERT INTO sales_records
                    (import_batch_id, platform, source_row, order_number,
                     item_type, quantity, row_fingerprint)
                    VALUES (?, '테스트', 2, 'SOURCE', '제품확정', 1, 'source-row')
                    RETURNING id""",
                    (source_batch_id,),
                ).fetchone()[0]

                deleted_batch_id = connection.execute(
                    """INSERT INTO import_batches
                    (original_filename, stored_filename, status)
                    VALUES ('delete.xlsx', 'delete.xlsx', 'completed')
                    RETURNING id"""
                ).fetchone()[0]
                connection.execute(
                    """INSERT INTO import_record_updates
                    (import_batch_id, sales_record_id, old_item_type, new_item_type)
                    VALUES (?, ?, '제품', '제품확정')""",
                    (deleted_batch_id, source_record_id),
                )
                connection.execute(
                    """INSERT INTO sales_records
                    (import_batch_id, platform, source_row, order_number,
                     item_type, quantity, row_fingerprint)
                    VALUES (?, '테스트', 3, 'DELETE', '제품', 1, 'delete-row')""",
                    (deleted_batch_id,),
                )
                connection.execute(
                    """INSERT INTO import_issues
                    (import_batch_id, platform, source_row, field_name,
                     issue_code, message)
                    VALUES (?, '테스트', 3, 'quantity', 'test', 'test')""",
                    (deleted_batch_id,),
                )

            stored_path = main_app.UPLOAD_DIRECTORY / "delete.xlsx"
            stored_path.write_bytes(b"test")
            response = main_app.delete_import(deleted_batch_id)
            assert response.status_code == 303

            with database.get_connection() as connection:
                assert connection.execute(
                    "SELECT COUNT(*) FROM import_batches WHERE id = ?",
                    (deleted_batch_id,),
                ).fetchone()[0] == 0
                assert connection.execute(
                    "SELECT COUNT(*) FROM sales_records WHERE import_batch_id = ?",
                    (deleted_batch_id,),
                ).fetchone()[0] == 0
                assert connection.execute(
                    "SELECT item_type FROM sales_records WHERE id = ?",
                    (source_record_id,),
                ).fetchone()[0] == "제품"
            assert not stored_path.exists()
        finally:
            database.DATABASE_PATH = original_database_path
            main_app.UPLOAD_DIRECTORY = original_upload_directory


if __name__ == "__main__":
    main()
    print("permanent import deletion ok")
