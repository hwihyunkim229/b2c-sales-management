from openpyxl import load_workbook

from app.database import get_connection, initialize_database
from app.services.excel_importer import _import_hometax_sheet


def main() -> None:
    initialize_database()
    workbook = load_workbook(
        "uploads/initial-source.xlsx",
        data_only=True,
        read_only=True,
    )

    with get_connection() as connection:
        batch = connection.execute(
            "SELECT id FROM import_batches ORDER BY id LIMIT 1"
        ).fetchone()

        if batch is None:
            raise RuntimeError("먼저 판매자료를 업로드해주세요.")

        _import_hometax_sheet(
            connection=connection,
            workbook=workbook,
            batch_id=batch["id"],
        )
        count = connection.execute(
            "SELECT COUNT(*) FROM hometax_records"
        ).fetchone()[0]

    print(f"Stored HomeTax records: {count}")


if __name__ == "__main__":
    main()
