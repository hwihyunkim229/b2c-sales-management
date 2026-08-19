from fastapi.testclient import TestClient

from app.database import get_connection, initialize_database
from app.main import app


def active_summary(connection):
    return connection.execute(
        """
        SELECT COUNT(*) AS count,
            COALESCE(SUM(credit_card_amount + cash_receipt_amount
                + tax_invoice_amount + other_amount), 0) AS sales_amount
        FROM sales_records
        WHERE is_active = 1 AND batch_is_active = 1
        """
    ).fetchone()


def main() -> None:
    initialize_database()
    with get_connection() as connection:
        batch = connection.execute(
            "SELECT id FROM import_batches WHERE status = 'completed' ORDER BY id LIMIT 1"
        ).fetchone()
        assert batch is not None
        batch_id = int(batch["id"])
        before = active_summary(connection)
        excluded_before = connection.execute(
            "SELECT COUNT(*) AS count FROM sales_records WHERE is_active = 0"
        ).fetchone()["count"]

    client = TestClient(app)
    assert client.post(f"/imports/{batch_id}/cancel", follow_redirects=False).status_code == 303
    with get_connection() as connection:
        cancelled = active_summary(connection)
        assert connection.execute(
            "SELECT status FROM import_batches WHERE id = ?", (batch_id,)
        ).fetchone()["status"] == "cancelled"
        assert cancelled["count"] == 0

    assert client.post(f"/imports/{batch_id}/restore", follow_redirects=False).status_code == 303
    with get_connection() as connection:
        after = active_summary(connection)
        excluded_after = connection.execute(
            "SELECT COUNT(*) AS count FROM sales_records WHERE is_active = 0"
        ).fetchone()["count"]
        assert connection.execute(
            "SELECT status FROM import_batches WHERE id = ?", (batch_id,)
        ).fetchone()["status"] == "completed"

    assert before["count"] == after["count"]
    assert before["sales_amount"] == after["sales_amount"]
    assert excluded_before == excluded_after
    print({
        "batch_id": batch_id,
        "before_count": before["count"],
        "cancelled_count": cancelled["count"],
        "restored_count": after["count"],
        "sales_amount": after["sales_amount"],
        "excluded_preserved": excluded_after,
    })


if __name__ == "__main__":
    main()
