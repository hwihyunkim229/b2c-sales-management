from pathlib import Path

from fastapi.testclient import TestClient

import app.database as database
import app.security as security
from app.main import app


def initial_password(path: Path) -> str:
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("임시 비밀번호:"):
            return line.split(": ", 1)[1]
    raise AssertionError("initial password missing")


def test_account_creation_and_role_permissions(tmp_path, monkeypatch):
    database_path = tmp_path / "sales.db"
    credentials_path = tmp_path / "initial_admin_credentials.txt"
    monkeypatch.setattr(database, "DATABASE_PATH", database_path)
    monkeypatch.setattr(security, "INITIAL_CREDENTIALS_PATH", credentials_path)

    with TestClient(app) as admin_client:
        temporary_admin_password = initial_password(credentials_path)
        response = admin_client.post(
            "/login",
            data={
                "username": "admin",
                "password": temporary_admin_password,
                "next": "/",
            },
            follow_redirects=False,
        )
        assert response.headers["location"] == "/change-password"

        response = admin_client.post(
            "/change-password",
            data={
                "current_password": temporary_admin_password,
                "new_password": "NewAdmin1!x",
                "confirm_password": "NewAdmin1!x",
            },
            follow_redirects=False,
        )
        assert response.headers["location"] == "/"
        assert not credentials_path.exists()

        response = admin_client.post(
            "/users",
            data={
                "username": "viewer01",
                "display_name": "조회 담당",
                "role": "viewer",
                "temporary_password": "ViewerPass1!",
                "confirm_password": "ViewerPass1!",
            },
            follow_redirects=False,
        )
        assert response.headers["location"] == "/users?message=created"

        with database.get_connection() as connection:
            user = connection.execute(
                "SELECT * FROM users WHERE username = 'viewer01'"
            ).fetchone()
            assert user["role"] == "viewer"
            assert user["must_change_password"] == 1
            assert user["password_hash"] != "ViewerPass1!"

    with TestClient(app) as viewer_client:
        viewer_client.post(
            "/login",
            data={"username": "viewer01", "password": "ViewerPass1!", "next": "/"},
        )
        viewer_client.post(
            "/change-password",
            data={
                "current_password": "ViewerPass1!",
                "new_password": "ViewerChanged1!",
                "confirm_password": "ViewerChanged1!",
            },
        )
        assert viewer_client.get("/users").status_code == 403
        assert viewer_client.get("/sales/1/edit").status_code == 403
        assert viewer_client.post("/imports/1/cancel").status_code == 403
