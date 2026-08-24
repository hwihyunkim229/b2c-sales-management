from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from pathlib import Path
import os
import sqlite3

try:
    import psycopg
    from psycopg import IntegrityError as PostgresIntegrityError
except ImportError:
    psycopg = None
    PostgresIntegrityError = type("PostgresIntegrityError", (Exception,), {})


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIRECTORY = PROJECT_ROOT / "data"
DATABASE_PATH = DATA_DIRECTORY / "sales.db"
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
USING_POSTGRES = DATABASE_URL.startswith(("postgresql://", "postgres://"))
DB_INTEGRITY_ERRORS = (sqlite3.IntegrityError, PostgresIntegrityError)


class DatabaseRow(Mapping):
    def __init__(self, columns: Sequence[str], values: Sequence[object]):
        self._columns = tuple(columns)
        self._values = tuple(values)
        self._mapping = dict(zip(self._columns, self._values))

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        return self._mapping[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._columns)

    def __len__(self) -> int:
        return len(self._columns)

    def __getattr__(self, name: str):
        try:
            return self._mapping[name]
        except KeyError as error:
            raise AttributeError(name) from error


class PostgresResult:
    def __init__(self, cursor):
        self.cursor = cursor
        self.columns = tuple(
            description.name for description in (cursor.description or ())
        )

    def _row(self, values):
        if values is None:
            return None
        return DatabaseRow(self.columns, values)

    def fetchone(self):
        values = self.cursor.fetchone()
        self.cursor.close()
        return self._row(values)

    def fetchall(self):
        values = self.cursor.fetchall()
        self.cursor.close()
        return [self._row(row) for row in values]


class EmptyPostgresResult:
    def fetchone(self):
        return None

    def fetchall(self):
        return []


class PostgresConnection:
    def __init__(self):
        if psycopg is None:
            raise RuntimeError("PostgreSQL을 사용하려면 psycopg 패키지가 필요합니다.")
        self.raw_connection = psycopg.connect(
            DATABASE_URL,
            options="-c timezone=Asia/Seoul",
        )

    @staticmethod
    def _sql(sql: str) -> str:
        # psycopg uses percent-style placeholders. Escape literal percent
        # characters first, then translate the app's portable qmark syntax.
        return sql.replace("%", "%%").replace("?", "%s")

    def execute(self, sql: str, parameters: Sequence[object] = ()):
        cursor = self.raw_connection.cursor()
        cursor.execute(self._sql(sql), tuple(parameters))
        if cursor.description is None:
            cursor.close()
            return EmptyPostgresResult()
        return PostgresResult(cursor)

    def executemany(self, sql: str, parameters):
        cursor = self.raw_connection.cursor()
        cursor.executemany(self._sql(sql), list(parameters))
        cursor.close()
        return EmptyPostgresResult()

    def commit(self) -> None:
        self.raw_connection.commit()

    def rollback(self) -> None:
        self.raw_connection.rollback()

    def close(self) -> None:
        self.raw_connection.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        self.close()


def get_connection():
    if USING_POSTGRES:
        return PostgresConnection()

    DATA_DIRECTORY.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS import_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    source_content BLOB,
    status TEXT NOT NULL,
    total_rows INTEGER NOT NULL DEFAULT 0,
    imported_rows INTEGER NOT NULL DEFAULT 0,
    duplicate_rows INTEGER NOT NULL DEFAULT 0,
    issue_rows INTEGER NOT NULL DEFAULT 0,
    error_message TEXT,
    cancelled_at TEXT,
    restored_at TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS sales_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER NOT NULL,
    platform TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    order_number TEXT NOT NULL,
    sales_month INTEGER,
    tax_date TEXT,
    item_type TEXT NOT NULL DEFAULT '상품',
    product_option TEXT,
    quantity REAL NOT NULL DEFAULT 0,
    credit_card_amount REAL NOT NULL DEFAULT 0,
    cash_receipt_amount REAL NOT NULL DEFAULT 0,
    tax_invoice_amount REAL NOT NULL DEFAULT 0,
    other_amount REAL NOT NULL DEFAULT 0,
    fee_amount_1 REAL NOT NULL DEFAULT 0,
    fee_amount_2 REAL NOT NULL DEFAULT 0,
    settlement_amount REAL NOT NULL DEFAULT 0,
    advance_payment REAL NOT NULL DEFAULT 0,
    buyer_name TEXT,
    note TEXT,
    row_fingerprint TEXT NOT NULL UNIQUE,
    review_status TEXT NOT NULL DEFAULT 'normal',
    review_note TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    batch_is_active INTEGER NOT NULL DEFAULT 1,
    updated_at TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (import_batch_id) REFERENCES import_batches(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_sales_platform ON sales_records(platform);
CREATE INDEX IF NOT EXISTS idx_sales_month ON sales_records(sales_month);
CREATE INDEX IF NOT EXISTS idx_sales_order ON sales_records(platform, order_number);

CREATE TABLE IF NOT EXISTS import_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER NOT NULL,
    platform TEXT NOT NULL,
    source_row INTEGER NOT NULL,
    order_number TEXT,
    field_name TEXT NOT NULL,
    issue_code TEXT NOT NULL,
    message TEXT NOT NULL,
    severity TEXT NOT NULL DEFAULT 'warning',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (import_batch_id) REFERENCES import_batches(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_import_issues_batch ON import_issues(import_batch_id);

CREATE TABLE IF NOT EXISTS import_record_updates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER NOT NULL,
    sales_record_id INTEGER NOT NULL,
    old_item_type TEXT NOT NULL,
    new_item_type TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(import_batch_id, sales_record_id),
    FOREIGN KEY (import_batch_id) REFERENCES import_batches(id) ON DELETE CASCADE,
    FOREIGN KEY (sales_record_id) REFERENCES sales_records(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_import_record_updates_batch
    ON import_record_updates(import_batch_id);

CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action TEXT NOT NULL,
    target_type TEXT NOT NULL,
    target_id INTEGER,
    description TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'viewer',
    password_hash TEXT NOT NULL,
    password_salt TEXT NOT NULL,
    password_iterations INTEGER NOT NULL,
    must_change_password INTEGER NOT NULL DEFAULT 1,
    is_active INTEGER NOT NULL DEFAULT 1,
    last_login_at TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS auth_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    token_hash TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL,
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_auth_sessions_expires ON auth_sessions(expires_at);

CREATE TABLE IF NOT EXISTS hometax_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER NOT NULL,
    record_type TEXT NOT NULL,
    platform TEXT,
    source_platform TEXT,
    platform_status TEXT NOT NULL DEFAULT 'unassigned',
    sales_month INTEGER,
    sales_date TEXT,
    transaction_count INTEGER NOT NULL DEFAULT 0,
    supply_amount REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    service_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    credit_card_amount REAL NOT NULL DEFAULT 0,
    other_payment_amount REAL NOT NULL DEFAULT 0,
    approval_number TEXT,
    merchant_name TEXT,
    transaction_type TEXT,
    note TEXT,
    batch_is_active INTEGER NOT NULL DEFAULT 1,
    row_fingerprint TEXT NOT NULL UNIQUE,
    FOREIGN KEY (import_batch_id) REFERENCES import_batches(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_hometax_month_platform
    ON hometax_records(sales_month, platform);
"""


POSTGRES_SCHEMA_STATEMENTS = (
    """CREATE TABLE IF NOT EXISTS import_batches (
        id BIGSERIAL PRIMARY KEY, original_filename TEXT NOT NULL,
        stored_filename TEXT NOT NULL, source_content BYTEA, status TEXT NOT NULL,
        total_rows INTEGER NOT NULL DEFAULT 0, imported_rows INTEGER NOT NULL DEFAULT 0,
        duplicate_rows INTEGER NOT NULL DEFAULT 0, issue_rows INTEGER NOT NULL DEFAULT 0,
        error_message TEXT, cancelled_at TIMESTAMPTZ, restored_at TIMESTAMPTZ,
        created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
    )""",
    """CREATE TABLE IF NOT EXISTS sales_records (
        id BIGSERIAL PRIMARY KEY,
        import_batch_id BIGINT NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
        platform TEXT NOT NULL, source_row INTEGER NOT NULL, order_number TEXT NOT NULL,
        sales_month INTEGER, tax_date TEXT, item_type TEXT NOT NULL DEFAULT '상품',
        product_option TEXT, quantity DOUBLE PRECISION NOT NULL DEFAULT 0,
        credit_card_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        cash_receipt_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        tax_invoice_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        other_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        fee_amount_1 DOUBLE PRECISION NOT NULL DEFAULT 0,
        fee_amount_2 DOUBLE PRECISION NOT NULL DEFAULT 0,
        settlement_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        advance_payment DOUBLE PRECISION NOT NULL DEFAULT 0,
        buyer_name TEXT, note TEXT, row_fingerprint TEXT NOT NULL UNIQUE,
        review_status TEXT NOT NULL DEFAULT 'normal', review_note TEXT,
        is_active INTEGER NOT NULL DEFAULT 1, batch_is_active INTEGER NOT NULL DEFAULT 1,
        updated_at TIMESTAMPTZ, created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
    )""",
    "CREATE INDEX IF NOT EXISTS idx_sales_platform ON sales_records(platform)",
    "CREATE INDEX IF NOT EXISTS idx_sales_month ON sales_records(sales_month)",
    "CREATE INDEX IF NOT EXISTS idx_sales_order ON sales_records(platform, order_number)",
    """CREATE TABLE IF NOT EXISTS import_issues (
        id BIGSERIAL PRIMARY KEY,
        import_batch_id BIGINT NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
        platform TEXT NOT NULL, source_row INTEGER NOT NULL, order_number TEXT,
        field_name TEXT NOT NULL, issue_code TEXT NOT NULL, message TEXT NOT NULL,
        severity TEXT NOT NULL DEFAULT 'warning',
        created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
    )""",
    "CREATE INDEX IF NOT EXISTS idx_import_issues_batch ON import_issues(import_batch_id)",
    """CREATE TABLE IF NOT EXISTS import_record_updates (
        id BIGSERIAL PRIMARY KEY,
        import_batch_id BIGINT NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
        sales_record_id BIGINT NOT NULL REFERENCES sales_records(id) ON DELETE CASCADE,
        old_item_type TEXT NOT NULL, new_item_type TEXT NOT NULL,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(import_batch_id, sales_record_id)
    )""",
    "CREATE INDEX IF NOT EXISTS idx_import_record_updates_batch ON import_record_updates(import_batch_id)",
    """CREATE TABLE IF NOT EXISTS audit_logs (
        id BIGSERIAL PRIMARY KEY, action TEXT NOT NULL, target_type TEXT NOT NULL,
        target_id BIGINT, description TEXT NOT NULL,
        created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
    )""",
    """CREATE TABLE IF NOT EXISTS users (
        id BIGSERIAL PRIMARY KEY, username TEXT NOT NULL UNIQUE, display_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'viewer', password_hash TEXT NOT NULL,
        password_salt TEXT NOT NULL, password_iterations INTEGER NOT NULL,
        must_change_password INTEGER NOT NULL DEFAULT 1,
        is_active INTEGER NOT NULL DEFAULT 1, last_login_at TIMESTAMPTZ,
        created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMPTZ
    )""",
    """CREATE TABLE IF NOT EXISTS auth_sessions (
        id BIGSERIAL PRIMARY KEY, user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        token_hash TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL, last_seen_at TEXT NOT NULL
    )""",
    "CREATE INDEX IF NOT EXISTS idx_auth_sessions_expires ON auth_sessions(expires_at)",
    """CREATE TABLE IF NOT EXISTS hometax_records (
        id BIGSERIAL PRIMARY KEY,
        import_batch_id BIGINT NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
        record_type TEXT NOT NULL, platform TEXT, source_platform TEXT,
        platform_status TEXT NOT NULL DEFAULT 'unassigned', sales_month INTEGER,
        sales_date TEXT, transaction_count INTEGER NOT NULL DEFAULT 0,
        supply_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        vat_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        service_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        total_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        credit_card_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        other_payment_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
        approval_number TEXT, merchant_name TEXT, transaction_type TEXT, note TEXT,
        batch_is_active INTEGER NOT NULL DEFAULT 1, row_fingerprint TEXT NOT NULL UNIQUE
    )""",
    "CREATE INDEX IF NOT EXISTS idx_hometax_month_platform ON hometax_records(sales_month, platform)",
)


def _sqlite_columns(connection, table: str) -> set[str]:
    return {row["name"] for row in connection.execute(f"PRAGMA table_info({table})")}


def _normalize_existing_data(connection) -> None:
    connection.execute(
        """UPDATE sales_records SET review_status = CASE
            WHEN EXISTS (SELECT 1 FROM import_issues
                WHERE import_issues.import_batch_id = sales_records.import_batch_id
                  AND import_issues.platform = sales_records.platform
                  AND import_issues.source_row = sales_records.source_row)
            THEN 'issue' ELSE 'normal' END WHERE review_status = 'pending'"""
    )
    connection.execute(
        """UPDATE hometax_records
        SET source_platform = COALESCE(source_platform, platform),
            platform_status = CASE WHEN platform IS NULL OR platform = ''
                THEN 'unassigned' ELSE 'manual' END
        WHERE source_platform IS NULL"""
    )
    connection.execute(
        """UPDATE sales_records SET item_type = '배송비'
        WHERE REPLACE(COALESCE(product_option, ''), ' ', '') LIKE '%배송비%'"""
    )
    connection.execute(
        """UPDATE sales_records SET item_type = '부가제품'
        WHERE REPLACE(COALESCE(product_option, ''), ' ', '') LIKE '%오므론%'"""
    )
    connection.execute(
        """UPDATE sales_records SET item_type = '사이즈미정'
        WHERE REPLACE(COALESCE(product_option, ''), ' ', '') LIKE '%사이즈미정%'"""
    )
    connection.execute(
        """UPDATE sales_records SET item_type = '제품취소'
        WHERE quantity < 0
          AND REPLACE(COALESCE(product_option, ''), ' ', '') NOT LIKE '%배송비%'
          AND REPLACE(COALESCE(product_option, ''), ' ', '') NOT LIKE '%오므론%'
          AND REPLACE(COALESCE(product_option, ''), ' ', '') NOT LIKE '%사이즈미정%'"""
    )


def _initialize_sqlite() -> None:
    with get_connection() as connection:
        connection.executescript(SQLITE_SCHEMA)
        additions = {
            "hometax_records": (
                ("source_platform", "TEXT"),
                ("platform_status", "TEXT NOT NULL DEFAULT 'unassigned'"),
                ("batch_is_active", "INTEGER NOT NULL DEFAULT 1"),
            ),
            "sales_records": (
                ("review_status", "TEXT NOT NULL DEFAULT 'normal'"),
                ("review_note", "TEXT"),
                ("is_active", "INTEGER NOT NULL DEFAULT 1"),
                ("batch_is_active", "INTEGER NOT NULL DEFAULT 1"),
                ("updated_at", "TEXT"),
            ),
            "import_batches": (
                ("issue_rows", "INTEGER NOT NULL DEFAULT 0"),
                ("cancelled_at", "TEXT"),
                ("restored_at", "TEXT"),
                ("source_content", "BLOB"),
            ),
            "users": (("role", "TEXT NOT NULL DEFAULT 'viewer'"),),
        }
        for table, columns in additions.items():
            existing = _sqlite_columns(connection, table)
            for name, definition in columns:
                if name not in existing:
                    connection.execute(
                        f"ALTER TABLE {table} ADD COLUMN {name} {definition}"
                    )
        connection.execute("UPDATE users SET role = 'admin' WHERE username = 'admin'")
        _normalize_existing_data(connection)


def _initialize_postgres() -> None:
    with get_connection() as connection:
        for statement in POSTGRES_SCHEMA_STATEMENTS:
            connection.execute(statement)
        connection.execute(
            "ALTER TABLE import_batches ADD COLUMN IF NOT EXISTS source_content BYTEA"
        )
        connection.execute(
            "ALTER TABLE users ADD COLUMN IF NOT EXISTS role TEXT NOT NULL DEFAULT 'viewer'"
        )
        connection.execute("UPDATE users SET role = 'admin' WHERE username = 'admin'")
        _normalize_existing_data(connection)


def initialize_database() -> None:
    if USING_POSTGRES:
        _initialize_postgres()
    else:
        _initialize_sqlite()
