from pathlib import Path

from fastapi.testclient import TestClient

from app.database import get_connection, initialize_database
from app.main import UPLOAD_DIRECTORY, app


def main() -> None:
    initialize_database()

    with get_connection() as connection:
        before_count = connection.execute(
            "SELECT COUNT(*) FROM sales_records"
        ).fetchone()[0]

    batch_id = None
    stored_filename = None

    try:
        with TestClient(app) as client:
            with Path("uploads/initial-source.xlsx").open("rb") as workbook:
                response = client.post(
                    "/imports",
                    files={
                        "file": (
                            "preview-test.xlsx",
                            workbook,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    },
                    follow_redirects=False,
                )

            assert response.status_code == 303

        with get_connection() as connection:
            batch = connection.execute(
                """
                SELECT *
                FROM import_batches
                WHERE original_filename = 'preview-test.xlsx'
                ORDER BY id DESC
                LIMIT 1
                """
            ).fetchone()
            after_count = connection.execute(
                "SELECT COUNT(*) FROM sales_records"
            ).fetchone()[0]

        assert batch is not None
        assert batch["status"] == "preview"
        assert batch["total_rows"] == 671
        assert before_count == after_count

        batch_id = batch["id"]
        stored_filename = batch["stored_filename"]

        print(
            {
                "batch_id": batch_id,
                "status": batch["status"],
                "total_rows": batch["total_rows"],
                "imported_rows": batch["imported_rows"],
                "duplicate_rows": batch["duplicate_rows"],
                "sales_count_unchanged": before_count == after_count,
            }
        )
    finally:
        if batch_id is not None:
            with get_connection() as connection:
                connection.execute(
                    "DELETE FROM import_batches WHERE id = ?",
                    (batch_id,),
                )

        if stored_filename is not None:
            test_file = UPLOAD_DIRECTORY / stored_filename

            if test_file.exists():
                test_file.unlink()


if __name__ == "__main__":
    main()
