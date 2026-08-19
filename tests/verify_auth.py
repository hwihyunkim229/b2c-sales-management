from pathlib import Path
from tempfile import TemporaryDirectory

from fastapi.testclient import TestClient

import app.database as database
import app.security as security
from app.main import app


PROJECT_ROOT = Path.cwd()


with TemporaryDirectory(ignore_cleanup_errors=True) as directory:
    temporary = Path(directory)
    original_database_path = database.DATABASE_PATH
    original_credentials_path = security.INITIAL_CREDENTIALS_PATH
    database.DATABASE_PATH = temporary / "sales.db"
    security.INITIAL_CREDENTIALS_PATH = temporary / "credentials.txt"
    uploaded_test_file = None
    try:
        with TestClient(app) as admin:
            password = next(
                line.split(": ", 1)[1]
                for line in security.INITIAL_CREDENTIALS_PATH.read_text(
                    encoding="utf-8"
                ).splitlines()
                if line.startswith("임시 비밀번호:")
            )
            response = admin.post(
                "/login",
                data={"username": "admin", "password": password, "next": "/"},
                follow_redirects=False,
            )
            assert response.headers["location"] == "/change-password"
            response = admin.post(
                "/change-password",
                data={
                    "current_password": password,
                    "new_password": "NewAdmin1!x",
                    "confirm_password": "NewAdmin1!x",
                },
                follow_redirects=False,
            )
            assert response.headers["location"] == "/"
            assert not security.INITIAL_CREDENTIALS_PATH.exists()
            response = admin.post(
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

            with (PROJECT_ROOT / "uploads" / "initial-source.xlsx").open("rb") as workbook:
                response = admin.post(
                    "/imports",
                    files={
                        "file": (
                            "database-source-test.xlsx",
                            workbook,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    },
                    follow_redirects=False,
                )
            assert response.status_code == 303
            with database.get_connection() as connection:
                batch = connection.execute(
                    """SELECT * FROM import_batches
                    WHERE original_filename = 'database-source-test.xlsx'"""
                ).fetchone()
                assert batch is not None
                assert batch["source_content"]
                uploaded_test_file = PROJECT_ROOT / "uploads" / batch["stored_filename"]
                uploaded_batch_id = batch["id"]
            uploaded_test_file.unlink()
            response = admin.get(f"/imports/{uploaded_batch_id}")
            assert response.status_code == 200
            assert uploaded_test_file.exists()

        with TestClient(app) as viewer:
            viewer.post(
                "/login",
                data={
                    "username": "viewer01",
                    "password": "ViewerPass1!",
                    "next": "/",
                },
            )
            viewer.post(
                "/change-password",
                data={
                    "current_password": "ViewerPass1!",
                    "new_password": "ViewerChanged1!",
                    "confirm_password": "ViewerChanged1!",
                },
            )
            assert viewer.get("/users").status_code == 403
            assert viewer.get("/sales/1/edit").status_code == 403
            assert viewer.post("/imports/1/cancel").status_code == 403

        with database.get_connection() as connection:
            viewer_user = connection.execute(
                "SELECT * FROM users WHERE username = 'viewer01'"
            ).fetchone()
            assert viewer_user["role"] == "viewer"
            assert viewer_user["password_hash"] != "ViewerChanged1!"
    finally:
        database.DATABASE_PATH = original_database_path
        security.INITIAL_CREDENTIALS_PATH = original_credentials_path
        if uploaded_test_file and uploaded_test_file.exists():
            uploaded_test_file.unlink()

print("authentication verification ok")
