from app.database import get_connection


def main() -> None:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT
                id,
                platform,
                source_row,
                order_number,
                tax_date,
                item_type,
                product_option,
                quantity,
                credit_card_amount,
                cash_receipt_amount,
                tax_invoice_amount,
                other_amount,
                buyer_name,
                import_batch_id,
                row_fingerprint
            FROM sales_records
            WHERE TRIM(order_number) IN ('8', '9')
            ORDER BY order_number, platform, source_row
            """
        ).fetchall()

    print(f"count: {len(rows)}")

    for row in rows:
        print(dict(row))


if __name__ == "__main__":
    main()
