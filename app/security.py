from __future__ import annotations

from datetime import datetime, timedelta, timezone
from hashlib import pbkdf2_hmac, sha256
from pathlib import Path
import hmac
import os
import secrets
import sqlite3


PROJECT_ROOT = Path(__file__).resolve().parent.parent
INITIAL_CREDENTIALS_PATH = PROJECT_ROOT / "data" / "initial_admin_credentials.txt"
SESSION_COOKIE_NAME = "b2c_session"
SESSION_DURATION_SECONDS = 2 * 60 * 60
PASSWORD_ITERATIONS = 600_000


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_utc(value: datetime) -> str:
    return value.isoformat(timespec="seconds")


def hash_password(
    password: str,
    salt_hex: str | None = None,
    iterations: int = PASSWORD_ITERATIONS,
) -> tuple[str, str, int]:
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    password_hash = pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        iterations,
    )
    return password_hash.hex(), salt.hex(), iterations


def verify_password(
    password: str,
    expected_hash: str,
    salt_hex: str,
    iterations: int,
) -> bool:
    actual_hash, _, _ = hash_password(password, salt_hex, iterations)
    return hmac.compare_digest(actual_hash, expected_hash)


def password_error(password: str) -> str | None:
    if len(password) < 10:
        return "비밀번호는 10자 이상이어야 합니다."
    if not any(character.isalpha() for character in password):
        return "비밀번호에 영문자를 포함해 주세요."
    if not any(character.isdigit() for character in password):
        return "비밀번호에 숫자를 포함해 주세요."
    if not any(not character.isalnum() for character in password):
        return "비밀번호에 특수문자를 포함해 주세요."
    return None


def ensure_initial_admin(connection: sqlite3.Connection) -> bool:
    existing_user = connection.execute("SELECT id FROM users LIMIT 1").fetchone()
    if existing_user:
        return False

    username = os.getenv("INITIAL_ADMIN_USERNAME", "admin").strip().lower()
    configured_password = os.getenv("INITIAL_ADMIN_PASSWORD", "")
    temporary_password = configured_password or secrets.token_urlsafe(14)
    if configured_password and password_error(configured_password):
        raise RuntimeError(
            "INITIAL_ADMIN_PASSWORD는 10자 이상이며 영문자·숫자·특수문자를 포함해야 합니다."
        )
    password_hash, password_salt, iterations = hash_password(temporary_password)
    connection.execute(
        """
        INSERT INTO users (
            username, display_name, role, password_hash, password_salt,
            password_iterations, must_change_password
        ) VALUES (?, ?, 'admin', ?, ?, ?, 1)
        """,
        (username, "관리자", password_hash, password_salt, iterations),
    )
    connection.commit()

    if not configured_password:
        INITIAL_CREDENTIALS_PATH.parent.mkdir(parents=True, exist_ok=True)
        INITIAL_CREDENTIALS_PATH.write_text(
            "B2C Sales 최초 관리자 계정\n"
            f"아이디: {username}\n"
            f"임시 비밀번호: {temporary_password}\n\n"
            "최초 로그인 후 비밀번호를 변경하면 이 파일은 자동 삭제됩니다.\n",
            encoding="utf-8",
        )
    return True


def session_token_hash(token: str) -> str:
    return sha256(token.encode("utf-8")).hexdigest()


def create_session(connection: sqlite3.Connection, user_id: int) -> str:
    token = secrets.token_urlsafe(32)
    now = utc_now()
    connection.execute(
        """
        INSERT INTO auth_sessions (
            user_id, token_hash, created_at, expires_at, last_seen_at
        ) VALUES (?, ?, ?, ?, ?)
        """,
        (
            user_id,
            session_token_hash(token),
            iso_utc(now),
            iso_utc(now + timedelta(seconds=SESSION_DURATION_SECONDS)),
            iso_utc(now),
        ),
    )
    return token


def get_session_user(connection: sqlite3.Connection, token: str | None):
    if not token:
        return None

    now = iso_utc(utc_now())
    row = connection.execute(
        """
        SELECT users.*, auth_sessions.id AS session_id
        FROM auth_sessions
        JOIN users ON users.id = auth_sessions.user_id
        WHERE auth_sessions.token_hash = ?
          AND auth_sessions.expires_at > ?
          AND users.is_active = 1
        """,
        (session_token_hash(token), now),
    ).fetchone()
    if row:
        connection.execute(
            "UPDATE auth_sessions SET last_seen_at = ? WHERE id = ?",
            (now, row["session_id"]),
        )
    return row


def delete_session(connection: sqlite3.Connection, token: str | None) -> None:
    if token:
        connection.execute(
            "DELETE FROM auth_sessions WHERE token_hash = ?",
            (session_token_hash(token),),
        )


def remove_initial_credentials() -> None:
    INITIAL_CREDENTIALS_PATH.unlink(missing_ok=True)
