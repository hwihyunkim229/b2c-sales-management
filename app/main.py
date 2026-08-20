from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4
from urllib.parse import quote
import io
import math
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from fastapi import FastAPI, File, Form, Query, Request, UploadFile
from fastapi.responses import JSONResponse, RedirectResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.database import get_connection, initialize_database
from app.services.excel_importer import import_workbook
from app.services.excel_importer import SUPPORTED_PLATFORMS
from app.services.data_validator import build_validation_report
from app.security import (
    INITIAL_CREDENTIALS_PATH,
    SESSION_COOKIE_NAME,
    SESSION_DURATION_SECONDS,
    create_session,
    delete_session,
    ensure_initial_admin,
    get_session_user,
    hash_password,
    password_error,
    remove_initial_credentials,
    verify_password,
)


APP_DIRECTORY = Path(__file__).resolve().parent
PROJECT_ROOT = APP_DIRECTORY.parent
UPLOAD_DIRECTORY = PROJECT_ROOT / "uploads"


def ensure_batch_source_file(connection, batch) -> Path:
    stored_path = UPLOAD_DIRECTORY / batch["stored_filename"]
    if stored_path.exists():
        return stored_path

    source_content = batch["source_content"]
    if source_content is None:
        source_row = connection.execute(
            "SELECT source_content FROM import_batches WHERE id = ?",
            (batch["id"],),
        ).fetchone()
        source_content = source_row["source_content"] if source_row else None

    if source_content is None:
        raise FileNotFoundError(
            f"업로드 원본을 찾을 수 없습니다: {batch['original_filename']}"
        )

    UPLOAD_DIRECTORY.mkdir(parents=True, exist_ok=True)
    stored_path.write_bytes(bytes(source_content))
    return stored_path


def restore_completed_source_files(connection) -> None:
    batches = connection.execute(
        """
        SELECT id, original_filename, stored_filename, source_content
        FROM import_batches
        WHERE status = 'completed'
        """
    ).fetchall()
    for batch in batches:
        ensure_batch_source_file(connection, batch)


def backfill_source_content(connection) -> int:
    batches = connection.execute(
        """
        SELECT id, stored_filename
        FROM import_batches
        WHERE source_content IS NULL
        """
    ).fetchall()
    updated = 0
    for batch in batches:
        file_path = UPLOAD_DIRECTORY / batch["stored_filename"]
        if not file_path.exists():
            continue
        connection.execute(
            "UPDATE import_batches SET source_content = ? WHERE id = ?",
            (file_path.read_bytes(), batch["id"]),
        )
        updated += 1
    return updated


@asynccontextmanager
async def lifespan(app: FastAPI):
    initialize_database()
    UPLOAD_DIRECTORY.mkdir(parents=True, exist_ok=True)
    with get_connection() as connection:
        created = ensure_initial_admin(connection)
        restored_sources = backfill_source_content(connection)
    if created:
        if INITIAL_CREDENTIALS_PATH.exists():
            print(f"[B2C Sales] 최초 관리자 자격증명: {INITIAL_CREDENTIALS_PATH}")
        else:
            print("[B2C Sales] 환경변수로 최초 관리자 계정을 생성했습니다.")
    if restored_sources:
        print(f"[B2C Sales] 기존 업로드 원본 {restored_sources}개를 DB에 저장했습니다.")
    yield


app = FastAPI(
    title="B2C Sales Management",
    lifespan=lifespan,
)
app.mount(
    "/static",
    StaticFiles(directory=APP_DIRECTORY / "static"),
    name="static",
)

templates = Jinja2Templates(
    directory=APP_DIRECTORY / "templates",
)


PUBLIC_PATHS = {"/login", "/health", "/favicon.ico"}


@app.middleware("http")
async def authentication_middleware(request: Request, call_next):
    path = request.url.path
    token = request.cookies.get(SESSION_COOKIE_NAME)
    user = None
    if token:
        with get_connection() as connection:
            user = get_session_user(connection, token)

    request.state.user = user

    if path.startswith("/static/") or path in PUBLIC_PATHS:
        return await call_next(request)

    if user is None:
        next_path = path if request.method == "GET" else "/"
        return RedirectResponse(
            f"/login?next={quote(next_path, safe='/')}",
            status_code=303,
        )

    if user["must_change_password"] and path not in {
        "/change-password",
        "/logout",
    }:
        return RedirectResponse("/change-password", status_code=303)

    if path.startswith("/users") and user["role"] != "admin":
        return templates.TemplateResponse(
            request=request,
            name="permission_denied.html",
            context={},
            status_code=403,
        )

    if (
        user["role"] == "viewer"
        and path.startswith("/sales/")
        and path.endswith("/edit")
    ):
        return templates.TemplateResponse(
            request=request,
            name="permission_denied.html",
            context={},
            status_code=403,
        )

    if request.method not in {"GET", "HEAD", "OPTIONS"} and path not in {
        "/change-password",
        "/logout",
    }:
        role = user["role"]
        reviewer_restricted = path.startswith("/imports") or path.startswith("/hometax")
        if role == "viewer" or (role == "reviewer" and reviewer_restricted):
            return templates.TemplateResponse(
                request=request,
                name="permission_denied.html",
                context={},
                status_code=403,
            )

    return await call_next(request)


def set_session_cookie(response: RedirectResponse, token: str, request: Request) -> None:
    response.set_cookie(
        SESSION_COOKIE_NAME,
        token,
        max_age=SESSION_DURATION_SECONDS,
        httponly=True,
        secure=request.url.scheme == "https",
        samesite="lax",
        path="/",
    )


def safe_next_path(value: str) -> str:
    if value.startswith("/") and not value.startswith("//"):
        return value
    return "/"


@app.api_route("/health", methods=["GET", "HEAD"])
async def health(request: Request):
    if request.method == "HEAD":
        return Response(status_code=200)
    return JSONResponse({"status": "ok"})


@app.get("/login")
async def login_page(request: Request, next: str = "/", error: str = ""):
    if request.state.user:
        destination = (
            "/change-password"
            if request.state.user["must_change_password"]
            else safe_next_path(next)
        )
        return RedirectResponse(destination, status_code=303)
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"next": safe_next_path(next), "error": error},
    )


@app.post("/login")
async def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next: str = Form("/"),
):
    with get_connection() as connection:
        connection.execute(
            "DELETE FROM auth_sessions WHERE expires_at <= ?",
            (datetime.now(timezone.utc).isoformat(timespec="seconds"),),
        )
        user = connection.execute(
            "SELECT * FROM users WHERE username = ? AND is_active = 1",
            (username.strip(),),
        ).fetchone()
        valid = bool(
            user
            and verify_password(
                password,
                user["password_hash"],
                user["password_salt"],
                user["password_iterations"],
            )
        )
        if not valid:
            connection.execute(
                """
                INSERT INTO audit_logs (action, target_type, target_id, description)
                VALUES ('login_failed', 'user', ?, ?)
                """,
                (user["id"] if user else None, f"로그인 실패: {username.strip()}"),
            )
            return RedirectResponse(
                f"/login?error=invalid&next={quote(safe_next_path(next), safe='/')}",
                status_code=303,
            )

        token = create_session(connection, user["id"])
        connection.execute(
            "UPDATE users SET last_login_at = CURRENT_TIMESTAMP WHERE id = ?",
            (user["id"],),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('login', 'user', ?, ?)
            """,
            (user["id"], f"{user['username']} 로그인"),
        )

    destination = "/change-password" if user["must_change_password"] else safe_next_path(next)
    response = RedirectResponse(destination, status_code=303)
    set_session_cookie(response, token, request)
    return response


@app.get("/change-password")
async def change_password_page(request: Request, error: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="change_password.html",
        context={"error": error},
    )


@app.post("/change-password")
async def change_password(
    request: Request,
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...),
):
    user = request.state.user
    if new_password != confirm_password:
        return RedirectResponse("/change-password?error=confirm", status_code=303)
    validation_error = password_error(new_password)
    if validation_error:
        return templates.TemplateResponse(
            request=request,
            name="change_password.html",
            context={"error": "policy", "policy_error": validation_error},
            status_code=400,
        )
    if verify_password(
        new_password,
        user["password_hash"],
        user["password_salt"],
        user["password_iterations"],
    ):
        return RedirectResponse("/change-password?error=same", status_code=303)
    if not verify_password(
        current_password,
        user["password_hash"],
        user["password_salt"],
        user["password_iterations"],
    ):
        return RedirectResponse("/change-password?error=current", status_code=303)

    password_hash, password_salt, iterations = hash_password(new_password)
    with get_connection() as connection:
        connection.execute(
            """
            UPDATE users
            SET password_hash = ?, password_salt = ?, password_iterations = ?,
                must_change_password = 0, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (password_hash, password_salt, iterations, user["id"]),
        )
        connection.execute("DELETE FROM auth_sessions WHERE user_id = ?", (user["id"],))
        token = create_session(connection, user["id"])
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('password_changed', 'user', ?, ?)
            """,
            (user["id"], f"{user['username']} 비밀번호 변경"),
        )

    remove_initial_credentials()
    response = RedirectResponse("/", status_code=303)
    set_session_cookie(response, token, request)
    return response


@app.post("/logout")
async def logout(request: Request):
    token = request.cookies.get(SESSION_COOKIE_NAME)
    user = request.state.user
    with get_connection() as connection:
        delete_session(connection, token)
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('logout', 'user', ?, ?)
            """,
            (user["id"], f"{user['username']} 로그아웃"),
        )
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(SESSION_COOKIE_NAME, path="/")
    return response


ROLE_LABELS = {
    "admin": "관리자",
    "reviewer": "검수자",
    "viewer": "조회자",
}


@app.get("/users")
async def users_page(request: Request, message: str = "", error: str = ""):
    with get_connection() as connection:
        users = connection.execute(
            """
            SELECT id, username, display_name, role, must_change_password,
                   is_active, last_login_at, created_at
            FROM users
            ORDER BY is_active DESC, id
            """
        ).fetchall()
    return templates.TemplateResponse(
        request=request,
        name="users.html",
        context={
            "users": users,
            "role_labels": ROLE_LABELS,
            "message": message,
            "error": error,
        },
    )


@app.post("/users")
async def create_user(
    request: Request,
    username: str = Form(...),
    display_name: str = Form(...),
    role: str = Form(...),
    temporary_password: str = Form(...),
    confirm_password: str = Form(...),
):
    username = username.strip().lower()
    display_name = display_name.strip()
    if not re.fullmatch(r"[a-z0-9._-]{3,30}", username):
        return RedirectResponse("/users?error=username", status_code=303)
    if not display_name or len(display_name) > 30:
        return RedirectResponse("/users?error=name", status_code=303)
    if role not in ROLE_LABELS:
        return RedirectResponse("/users?error=role", status_code=303)
    if temporary_password != confirm_password:
        return RedirectResponse("/users?error=confirm", status_code=303)
    if password_error(temporary_password):
        return RedirectResponse("/users?error=password", status_code=303)

    password_hash, password_salt, iterations = hash_password(temporary_password)
    with get_connection() as connection:
        if connection.execute(
            "SELECT 1 FROM users WHERE username = ?", (username,)
        ).fetchone():
            return RedirectResponse("/users?error=duplicate", status_code=303)
        cursor = connection.execute(
            """
            INSERT INTO users (
                username, display_name, role, password_hash, password_salt,
                password_iterations, must_change_password, is_active
            ) VALUES (?, ?, ?, ?, ?, ?, 1, 1)
            RETURNING id
            """,
            (
                username,
                display_name,
                role,
                password_hash,
                password_salt,
                iterations,
            ),
        )
        created_user = cursor.fetchone()
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('user_created', 'user', ?, ?)
            """,
            (
                created_user["id"],
                f"{request.state.user['username']} 계정 생성: {username} ({ROLE_LABELS[role]})",
            ),
        )
    return RedirectResponse("/users?message=created", status_code=303)


@app.post("/users/{user_id}/role")
async def change_user_role(request: Request, user_id: int, role: str = Form(...)):
    if role not in ROLE_LABELS:
        return RedirectResponse("/users?error=role", status_code=303)
    if user_id == request.state.user["id"]:
        return RedirectResponse("/users?error=self_role", status_code=303)
    with get_connection() as connection:
        target = connection.execute(
            "SELECT username FROM users WHERE id = ?", (user_id,)
        ).fetchone()
        if not target:
            return RedirectResponse("/users?error=missing", status_code=303)
        connection.execute(
            "UPDATE users SET role = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (role, user_id),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('user_role_changed', 'user', ?, ?)
            """,
            (user_id, f"{target['username']} 권한을 {ROLE_LABELS[role]}로 변경"),
        )
    return RedirectResponse("/users?message=role", status_code=303)


@app.post("/users/{user_id}/status")
async def toggle_user_status(request: Request, user_id: int):
    if user_id == request.state.user["id"]:
        return RedirectResponse("/users?error=self_status", status_code=303)
    with get_connection() as connection:
        target = connection.execute(
            "SELECT username, is_active FROM users WHERE id = ?", (user_id,)
        ).fetchone()
        if not target:
            return RedirectResponse("/users?error=missing", status_code=303)
        new_status = 0 if target["is_active"] else 1
        connection.execute(
            "UPDATE users SET is_active = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (new_status, user_id),
        )
        if not new_status:
            connection.execute("DELETE FROM auth_sessions WHERE user_id = ?", (user_id,))
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('user_status_changed', 'user', ?, ?)
            """,
            (user_id, f"{target['username']} 계정 {'활성화' if new_status else '비활성화'}"),
        )
    return RedirectResponse("/users?message=status", status_code=303)


@app.post("/users/{user_id}/reset-password")
async def reset_user_password(
    request: Request,
    user_id: int,
    temporary_password: str = Form(...),
    confirm_password: str = Form(...),
):
    if temporary_password != confirm_password:
        return RedirectResponse("/users?error=reset_confirm", status_code=303)
    if password_error(temporary_password):
        return RedirectResponse("/users?error=reset_password", status_code=303)
    password_hash, password_salt, iterations = hash_password(temporary_password)
    with get_connection() as connection:
        target = connection.execute(
            "SELECT username FROM users WHERE id = ?", (user_id,)
        ).fetchone()
        if not target:
            return RedirectResponse("/users?error=missing", status_code=303)
        connection.execute(
            """
            UPDATE users
            SET password_hash = ?, password_salt = ?, password_iterations = ?,
                must_change_password = 1, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (password_hash, password_salt, iterations, user_id),
        )
        connection.execute("DELETE FROM auth_sessions WHERE user_id = ?", (user_id,))
        connection.execute(
            """
            INSERT INTO audit_logs (action, target_type, target_id, description)
            VALUES ('password_reset', 'user', ?, ?)
            """,
            (user_id, f"{target['username']} 임시 비밀번호 재설정"),
        )
    return RedirectResponse("/users?message=reset", status_code=303)


def year_month(value: object) -> str:
    if value in (None, ""):
        return "-"

    text_value = str(value).strip()

    if len(text_value) >= 7:
        return text_value[:7]

    return text_value


def date_only(value: object) -> str:
    if value in (None, ""):
        return "-"
    if isinstance(value, (date, datetime)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return str(value).strip().split("T", 1)[0].split(" ", 1)[0]


KOREA_TIMEZONE = timezone(timedelta(hours=9))


def korean_datetime(value: object) -> str:
    if value in (None, ""):
        return "-"

    if isinstance(value, datetime):
        parsed_value = value
    else:
        try:
            parsed_value = datetime.fromisoformat(str(value).strip().replace("Z", "+00:00"))
        except ValueError:
            return str(value)

    if parsed_value.tzinfo is None:
        parsed_value = parsed_value.replace(tzinfo=timezone.utc)

    return parsed_value.astimezone(KOREA_TIMEZONE).strftime("%Y-%m-%d %H:%M")


def optional_number(
    value: str,
    minimum: int,
    maximum: int,
) -> int | None:
    if not value.strip():
        return None

    try:
        parsed_value = int(value)
    except ValueError:
        return None

    if minimum <= parsed_value <= maximum:
        return parsed_value

    return None


def excel_date_value(value: object) -> object:
    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    try:
        return datetime.fromisoformat(str(value).strip()).date()
    except ValueError:
        return value


def add_export_sheet(
    workbook: Workbook,
    title: str,
    headers: list[str],
    rows: list[list[object]],
    currency_columns: tuple[int, ...] = (),
    date_columns: tuple[int, ...] = (),
) -> None:
    worksheet = workbook.create_sheet(title)
    worksheet.append(headers)

    for row in rows:
        worksheet.append(row)

    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="176B5B")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30

    for column in currency_columns:
        for cell in worksheet.iter_cols(
            min_col=column, max_col=column, min_row=2
        ):
            for item in cell:
                item.number_format = "#,##0"

    for column in date_columns:
        for cell in worksheet.iter_cols(
            min_col=column, max_col=column, min_row=2
        ):
            for item in cell:
                item.value = excel_date_value(item.value)
                item.number_format = "yyyy-mm-dd"

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False

    for column in worksheet.columns:
        width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 36)
        worksheet.column_dimensions[column[0].column_letter].width = width


def dated_excel_filename(korean_name: str) -> str:
    today = datetime.now().strftime("%Y%m%d")
    return f"{korean_name}_{today}.xlsx"


def workbook_response(workbook: Workbook, filename: str) -> StreamingResponse:
    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        del workbook["Sheet"]

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": (
                f"attachment; filename=download.xlsx; "
                f"filename*=UTF-8''{quote(filename)}"
            )
        },
    )

templates.env.filters["year_month"] = year_month
templates.env.filters["date_only"] = date_only
templates.env.filters["korean_datetime"] = korean_datetime


SALES_EXPRESSION = """
    credit_card_amount
    + cash_receipt_amount
    + tax_invoice_amount
    + other_amount
"""

SALE_QUANTITY_EXPRESSION = """
    CASE
        WHEN COALESCE(item_type, '') NOT LIKE '%배송비%'
         AND COALESCE(item_type, '') NOT LIKE '%오므론%'
         AND COALESCE(product_option, '') NOT LIKE '%배송비%'
         AND COALESCE(product_option, '') NOT LIKE '%오므론%'
         AND (
             COALESCE(item_type, '') != '취소'
             OR REPLACE(COALESCE(product_option, ''), ' ', '') = '본품취소'
         )
        THEN COALESCE(quantity, 0)
        ELSE 0
    END
"""

PRODUCT_CODES = {
    "rev1_7": "SL-P-PD-00033",
    "rev1_8": "SL-P-PD-00034",
    "rev1_9": "SL-P-PD-00035",
    "rev1_10": "SL-P-PD-00036",
    "rev1_11": "SL-P-PD-00037",
    "rev1_12": "SL-P-PD-00038",
    "rev1_13": "SL-P-PD-00039",
    "rev2_7": "SL-P-PD-00054",
    "rev2_8": "SL-P-PD-00055",
    "rev2_9": "SL-P-PD-00056",
    "rev2_10": "SL-P-PD-00057",
    "rev2_11": "SL-P-PD-00058",
    "rev2_12": "SL-P-PD-00059",
    "rev2_13": "SL-P-PD-00060",
}


def normalized_size_key(value: object) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        numeric_value = float(value)
        if numeric_value.is_integer():
            return f"rev1_{int(numeric_value)}"

    label = str(value or "").strip().lower().replace(" ", "").replace(".", "")
    label = label.replace("-", "_")
    numeric_match = re.fullmatch(r"(\d+)\.0+", str(value or "").strip())
    if numeric_match:
        return f"rev1_{numeric_match.group(1)}"
    if re.fullmatch(r"\d+", label):
        return f"rev1_{label}"
    match = re.fullmatch(r"rev([12])_?(\d+)", label)
    if match:
        return f"rev{match.group(1)}_{match.group(2)}"
    return label


def product_code_for_size(value: object) -> str:
    original_value = str(value or "").strip()
    return PRODUCT_CODES.get(normalized_size_key(value), original_value or "-")


def product_type_label(value: object) -> str:
    return {
        "상품": "제품",
        "부가상품": "부가제품",
    }.get(str(value or ""), str(value or "-").replace("상품", "제품"))


def dashboard_size_label(value: object) -> str:
    label = str(value or "미지정").strip()
    normalized_label = normalized_size_key(value)
    rev1_match = re.fullmatch(r"rev1_(\d+)", normalized_label)
    if rev1_match:
        return f"Rev.1_{rev1_match.group(1)}"
    rev2_match = re.fullmatch(r"rev2_(\d+)", normalized_label)
    if rev2_match:
        return f"Rev.2_{rev2_match.group(1)}"
    return label


def dashboard_size_sort_key(value: object) -> tuple[object, ...]:
    label = str(value or "미지정").strip()
    normalized_label = normalized_size_key(value)
    rev1_match = re.fullmatch(r"rev1_(\d+)", normalized_label)
    if rev1_match:
        size = rev1_match.group(1)
        return (0, float(size), size)
    rev2_match = re.fullmatch(r"rev2_(\d+)", normalized_label)
    if rev2_match:
        size = rev2_match.group(1)
        return (1, float(size), size)
    return (2, math.inf, label)


templates.env.filters["product_code"] = product_code_for_size
templates.env.filters["product_type"] = product_type_label
templates.env.filters["product_size"] = dashboard_size_label


@app.get("/")
def dashboard(
    request: Request,
    size_month: str = Query(default=""),
):
    selected_size_month = optional_number(size_month, 1, 12)
    with get_connection() as connection:
        summary = connection.execute(
            """
            SELECT
                COUNT(*) AS record_count,
                COALESCE(SUM({quantity_expression}), 0) AS quantity,
                COALESCE(
                    SUM({sales_expression}),
                    0
                ) AS sales_amount,
                COALESCE(SUM(settlement_amount), 0) AS settlement_amount,
                COALESCE(SUM(fee_amount_1 + fee_amount_2), 0) AS fee_amount
            FROM sales_records
            WHERE is_active = 1 AND batch_is_active = 1
            """.format(
                quantity_expression=SALE_QUANTITY_EXPRESSION,
                sales_expression=SALES_EXPRESSION,
            )
        ).fetchone()

        platforms = connection.execute(
            """
            SELECT
                platform,
                COUNT(*) AS record_count,
                COALESCE(SUM({quantity_expression}), 0) AS quantity,
                COALESCE(
                    SUM({sales_expression}),
                    0
                ) AS sales_amount,
                COALESCE(SUM(settlement_amount), 0) AS settlement_amount
            FROM sales_records
            WHERE is_active = 1 AND batch_is_active = 1
            GROUP BY platform
            ORDER BY sales_amount DESC
            """.format(
                quantity_expression=SALE_QUANTITY_EXPRESSION,
                sales_expression=SALES_EXPRESSION,
            )
        ).fetchall()

        monthly_sales = connection.execute(
            """
            SELECT
                SUBSTR(tax_date, 1, 7) AS period,
                COALESCE(
                    SUM({sales_expression}),
                    0
                ) AS sales_amount
            FROM sales_records
            WHERE tax_date IS NOT NULL AND is_active = 1 AND batch_is_active = 1
            GROUP BY period
            ORDER BY period
            """.format(
                sales_expression=SALES_EXPRESSION,
            )
        ).fetchall()

        size_conditions = ["is_active = 1", "batch_is_active = 1"]
        size_parameters = []
        if selected_size_month is not None:
            size_conditions.append("sales_month = ?")
            size_parameters.append(selected_size_month)
        size_where_clause = " AND ".join(size_conditions)

        platform_sizes = connection.execute(
            f"""
            SELECT
                platform,
                COALESCE(product_option, '미지정') AS product_option,
                COALESCE(SUM({SALE_QUANTITY_EXPRESSION}), 0) AS quantity
            FROM sales_records
            WHERE {size_where_clause}
            GROUP BY platform, product_option
            HAVING ABS(SUM({SALE_QUANTITY_EXPRESSION})) > 0.001
            ORDER BY platform, quantity DESC, product_option
            """,
            size_parameters,
        ).fetchall()

        size_months = connection.execute(
            """
            SELECT DISTINCT sales_month
            FROM sales_records
            WHERE is_active = 1 AND batch_is_active = 1
              AND sales_month IS NOT NULL
            ORDER BY sales_month
            """
        ).fetchall()

    platform_size_groups = []
    platform_size_lookup = {}
    for row in platform_sizes:
        platform_name = row["platform"]
        if platform_name not in platform_size_lookup:
            group = {
                "platform": platform_name,
                "items": [],
                "total_quantity": 0,
                "maximum_quantity": 0,
            }
            platform_size_lookup[platform_name] = group
            platform_size_groups.append(group)
        group = platform_size_lookup[platform_name]
        quantity = row["quantity"] or 0
        group["items"].append(
            {
                "product_option": row["product_option"],
                "display_label": product_code_for_size(row["product_option"]),
                "quantity": quantity,
                "absolute_quantity": abs(quantity),
            }
        )
        group["total_quantity"] += quantity
        group["maximum_quantity"] = max(
            group["maximum_quantity"], abs(quantity)
        )

    for group in platform_size_groups:
        group["items"].sort(
            key=lambda item: dashboard_size_sort_key(item["product_option"])
        )
        absolute_total = sum(
            item["absolute_quantity"] for item in group["items"]
        )
        for item in group["items"]:
            item["share"] = (
                item["absolute_quantity"] / absolute_total * 100
                if absolute_total else 0
            )

    maximum_monthly_sales = max(
        (row["sales_amount"] for row in monthly_sales),
        default=0,
    )
    maximum_platform_sales = max(
        (row["sales_amount"] for row in platforms),
        default=0,
    )

    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "summary": summary,
            "platforms": platforms,
            "monthly_sales": monthly_sales,
            "platform_sizes": platform_sizes,
            "platform_size_groups": platform_size_groups,
            "size_months": size_months,
            "selected_size_month": selected_size_month,
            "maximum_monthly_sales": maximum_monthly_sales,
            "maximum_platform_sales": maximum_platform_sales,
        },
    )


@app.get("/validation")
def validation_center(request: Request):
    with get_connection() as connection:
        restore_completed_source_files(connection)
        report = build_validation_report(
            connection=connection,
            upload_directory=UPLOAD_DIRECTORY,
        )

    return templates.TemplateResponse(
        request=request,
        name="validation.html",
        context={"report": report},
    )


@app.get("/exports/validation.xlsx")
def export_validation():
    with get_connection() as connection:
        restore_completed_source_files(connection)
        report = build_validation_report(connection, UPLOAD_DIRECTORY)

    workbook = Workbook()
    add_export_sheet(workbook, "플랫폼 대사",
        ["플랫폼", "원본 건수", "웹 건수", "건수 차이", "원본 수량", "웹 수량", "수량 차이", "원본 매출", "웹 매출", "매출 차이"],
        [[r["platform"], r["source_count"], r["database_count"], r["count_difference"], r["source_quantity"], r["database_quantity"], r["quantity_difference"], r["source_sales"], r["database_sales"], r["sales_difference"]] for r in report["reconciliation"]],
        currency_columns=(8, 9, 10))
    add_export_sheet(workbook, "원본 미반영",
        ["플랫폼", "원본 행", "주문번호", "수량", "매출"],
        [[r["platform"], r["source_row"], r["order_number"], r["quantity"], r["sales_amount"]] for r in report["missing_source_rows"]], currency_columns=(5,))
    add_export_sheet(workbook, "필수값 누락",
        ["플랫폼", "원본 행", "주문번호", "누락 항목", "제품사이즈", "수량", "매출"],
        [[r["platform"], r["source_row"], r["order_number"], r["issue_detail"], dashboard_size_label(r["product_option"]), r["quantity"], r["sales_amount"]] for r in report["missing_fields"]], currency_columns=(7,))
    add_export_sheet(workbook, "0원 거래",
        ["플랫폼", "원본 행", "주문번호", "제품사이즈", "수량", "매출"],
        [[r["platform"], r["source_row"], r["order_number"], dashboard_size_label(r["product_option"]), r["quantity"], r["sales_amount"]] for r in report["zero_amount_rows"]], currency_columns=(6,))
    add_export_sheet(workbook, "확인 필요",
        ["플랫폼", "원본 행", "주문번호", "제품사이즈", "수량", "매출"],
        [[r["platform"], r["source_row"], r["order_number"], dashboard_size_label(r["product_option"]), r["quantity"], r["sales_amount"]] for r in report["review_issue_rows"]], currency_columns=(6,))
    add_export_sheet(workbook, "중복 의심",
        ["플랫폼", "주문번호", "구매자", "제품사이즈", "수량", "매출", "행 수", "원본 행"],
        [[r["platform"], r["order_number"], r["buyer_name"], dashboard_size_label(r["product_option"]), r["quantity"], r["sales_amount"], r["record_count"], r["source_rows"]] for r in report["duplicate_suspects"]], currency_columns=(6,))
    return workbook_response(workbook, dated_excel_filename("매출대사"))


@app.get("/sales")
def sales_list(
    request: Request,
    platform: str = Query(default=""),
    keyword: str = Query(default=""),
    sales_month: str = Query(default=""),
    item_type: str = Query(default=""),
    needs_review: bool = Query(default=False),
    sort_by: str = Query(default="tax_date"),
    sort_dir: str = Query(default="desc"),
    page: int = Query(default=1, ge=1),
):
    conditions = ["is_active = 1", "batch_is_active = 1"]
    parameters: list[object] = []
    selected_month = optional_number(sales_month, 1, 12)

    if platform:
        conditions.append("platform = ?")
        parameters.append(platform)

    if keyword:
        conditions.append(
            "(order_number LIKE ? OR buyer_name LIKE ? OR product_option LIKE ?)"
        )
        search_value = f"%{keyword}%"
        parameters.extend(
            [
                search_value,
                search_value,
                search_value,
            ]
        )

    if selected_month is not None:
        conditions.append("sales_month = ?")
        parameters.append(selected_month)

    if item_type:
        conditions.append("item_type = ?")
        parameters.append(item_type)

    if needs_review:
        conditions.append("review_status = 'issue'")

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    sortable_columns = {
        "platform": "platform",
        "order_number": "order_number",
        "tax_date": "tax_date",
        "product_option": "product_option",
        "quantity": "quantity",
        "sales_amount": (
            "credit_card_amount + cash_receipt_amount "
            "+ tax_invoice_amount + other_amount"
        ),
        "fee_amount": "fee_amount_1 + fee_amount_2",
        "settlement_amount": "settlement_amount",
        "buyer_name": "buyer_name",
    }
    selected_sort = sort_by if sort_by in sortable_columns else "tax_date"
    selected_direction = "asc" if sort_dir.lower() == "asc" else "desc"
    order_expression = sortable_columns[selected_sort]

    if selected_sort == "order_number":
        order_clause = f"""
            CASE
                WHEN TRIM(order_number) GLOB '[0-9]*' THEN 0
                ELSE 1
            END ASC,
            CAST(TRIM(order_number) AS INTEGER) {selected_direction},
            TRIM(order_number) {selected_direction},
            id DESC
        """
    else:
        order_clause = (
            f"{order_expression} {selected_direction}, id DESC"
        )

    with get_connection() as connection:
        total_count = connection.execute(
            f"SELECT COUNT(*) FROM sales_records {where_clause}",
            parameters,
        ).fetchone()[0]
        summary = connection.execute(
            f"""
            SELECT
                COALESCE(SUM({SALE_QUANTITY_EXPRESSION}), 0) AS quantity,
                COALESCE(SUM({SALES_EXPRESSION}), 0) AS sales_amount,
                COALESCE(SUM(credit_card_amount), 0) AS credit_card_amount,
                COALESCE(SUM(cash_receipt_amount), 0) AS cash_receipt_amount,
                COALESCE(SUM(tax_invoice_amount), 0) AS tax_invoice_amount,
                COALESCE(SUM(other_amount), 0) AS other_amount
            FROM sales_records
            {where_clause}
            """,
            parameters,
        ).fetchone()
        page_size = 50
        total_pages = max(1, math.ceil(total_count / page_size))
        current_page = min(page, total_pages)

        records = connection.execute(
            f"""
            SELECT
                *,
                (
                    credit_card_amount
                    + cash_receipt_amount
                    + tax_invoice_amount
                    + other_amount
                ) AS sales_amount
            FROM sales_records
            {where_clause}
            ORDER BY {order_clause}
            LIMIT ? OFFSET ?
            """,
            [
                *parameters,
                page_size,
                (current_page - 1) * page_size,
            ],
        ).fetchall()

        platforms = connection.execute(
            """
            SELECT DISTINCT platform
            FROM sales_records
            ORDER BY platform
            """
        ).fetchall()

    return templates.TemplateResponse(
        request=request,
        name="sales.html",
        context={
            "records": records,
            "platforms": platforms,
            "selected_platform": platform,
            "keyword": keyword,
            "selected_month": selected_month,
            "selected_item_type": item_type,
            "needs_review": needs_review,
            "selected_sort": selected_sort,
            "selected_direction": selected_direction,
            "current_page": current_page,
            "total_pages": total_pages,
            "total_count": total_count,
            "summary": summary,
        },
    )


@app.get("/sales/{record_id}")
def sales_detail(request: Request, record_id: int):
    with get_connection() as connection:
        record = connection.execute(
            "SELECT * FROM sales_records WHERE id = ?",
            (record_id,),
        ).fetchone()

    if record is None:
        return RedirectResponse(url="/sales", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="sales_detail.html",
        context={"record": record},
    )


@app.get("/sales/{record_id}/edit")
def sales_edit(request: Request, record_id: int):
    with get_connection() as connection:
        record = connection.execute(
            "SELECT * FROM sales_records WHERE id = ?",
            (record_id,),
        ).fetchone()

    if record is None:
        return RedirectResponse(url="/sales", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="sales_edit.html",
        context={"record": record},
    )


@app.post("/sales/{record_id}/edit")
def update_sales_record(
    record_id: int,
    product_option: str = Form(default=""),
    quantity: float = Form(default=0),
    credit_card_amount: float = Form(default=0),
    cash_receipt_amount: float = Form(default=0),
    tax_invoice_amount: float = Form(default=0),
    other_amount: float = Form(default=0),
    fee_amount_1: float = Form(default=0),
    fee_amount_2: float = Form(default=0),
    settlement_amount: float = Form(default=0),
    buyer_name: str = Form(default=""),
    note: str = Form(default=""),
    review_status: str = Form(default="normal"),
    review_note: str = Form(default=""),
):
    if review_status not in ("normal", "reviewed", "issue"):
        review_status = "normal"

    with get_connection() as connection:
        connection.execute(
            """
            UPDATE sales_records
            SET
                product_option = ?, quantity = ?,
                credit_card_amount = ?, cash_receipt_amount = ?,
                tax_invoice_amount = ?, other_amount = ?,
                fee_amount_1 = ?, fee_amount_2 = ?,
                settlement_amount = ?, buyer_name = ?, note = ?,
                review_status = ?, review_note = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (
                product_option.strip() or None,
                quantity,
                credit_card_amount,
                cash_receipt_amount,
                tax_invoice_amount,
                other_amount,
                fee_amount_1,
                fee_amount_2,
                settlement_amount,
                buyer_name.strip() or None,
                note.strip() or None,
                review_status,
                review_note.strip() or None,
                record_id,
            ),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (
                action, target_type, target_id, description
            ) VALUES ('update', 'sales_record', ?, '판매내역 수정 및 검수')
            """,
            (record_id,),
        )

    return RedirectResponse(url=f"/sales/{record_id}", status_code=303)


@app.post("/sales/{record_id}/exclude")
def exclude_sales_record(record_id: int):
    with get_connection() as connection:
        connection.execute(
            """
            UPDATE sales_records
            SET is_active = 0, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (record_id,),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (
                action, target_type, target_id, description
            ) VALUES ('exclude', 'sales_record', ?, '판매내역 집계 제외')
            """,
            (record_id,),
        )

    return RedirectResponse(url="/sales", status_code=303)


@app.get("/exports/sales.xlsx")
def export_sales(
    platform: str = Query(default=""),
    sales_month: str = Query(default=""),
    item_type: str = Query(default=""),
):
    conditions = ["is_active = 1", "batch_is_active = 1"]
    parameters: list[object] = []
    selected_month = optional_number(sales_month, 1, 12)

    for column, value in (
        ("platform", platform),
        ("sales_month", selected_month),
        ("item_type", item_type),
    ):
        if value not in (None, ""):
            conditions.append(f"{column} = ?")
            parameters.append(value)

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    with get_connection() as connection:
        rows = connection.execute(
            f"SELECT * FROM sales_records {where_clause} ORDER BY id",
            parameters,
        ).fetchall()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "통합 판매내역"
    headers = [
        "플랫폼", "주문번호", "매출월", "신고일", "제품유형",
        "제품코드", "수량", "신용카드", "현금영수증",
        "세금계산서", "기타", "수수료1", "수수료2",
        "정산금액", "구매자", "제품사이즈", "비고",
    ]
    worksheet.append(headers)

    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="176B5B")

    for row in rows:
        worksheet.append(
            [
                row["platform"], row["order_number"], row["sales_month"],
                excel_date_value(row["tax_date"]),
                product_type_label(row["item_type"]),
                product_code_for_size(row["product_option"]), row["quantity"],
                row["credit_card_amount"],
                row["cash_receipt_amount"], row["tax_invoice_amount"],
                row["other_amount"], row["fee_amount_1"],
                row["fee_amount_2"], row["settlement_amount"],
                row["buyer_name"], dashboard_size_label(row["product_option"]), row["note"],
            ]
        )

        worksheet.cell(worksheet.max_row, 4).number_format = "yyyy-mm-dd"

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column in worksheet.columns:
        width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            35,
        )
        worksheet.column_dimensions[column[0].column_letter].width = width

    return workbook_response(workbook, dated_excel_filename("통합판매내역"))


@app.get("/imports")
def import_history(request: Request):
    with get_connection() as connection:
        batches = connection.execute(
            """
            SELECT
                import_batches.*,
                COUNT(import_issues.id) AS issue_count,
                (SELECT COUNT(*) FROM sales_records WHERE import_batch_id = import_batches.id) AS affected_sales_count,
                COALESCE((SELECT SUM(credit_card_amount + cash_receipt_amount + tax_invoice_amount + other_amount) FROM sales_records WHERE import_batch_id = import_batches.id), 0) AS affected_sales_amount,
                (SELECT COUNT(*) FROM hometax_records WHERE import_batch_id = import_batches.id) AS affected_hometax_count
            FROM import_batches
            LEFT JOIN import_issues
                ON import_issues.import_batch_id = import_batches.id
            GROUP BY import_batches.id
            ORDER BY import_batches.id DESC
            """
        ).fetchall()

    return templates.TemplateResponse(
        request=request,
        name="imports.html",
        context={
            "batches": batches,
        },
    )


@app.get("/reports")
def reports(request: Request):
    with get_connection() as connection:
        monthly = connection.execute(
            f"""
            SELECT
                SUBSTR(tax_date, 1, 7) AS period,
                SUM(quantity) AS quantity,
                SUM({SALES_EXPRESSION}) AS sales_amount
            FROM sales_records
            WHERE tax_date IS NOT NULL AND is_active = 1 AND batch_is_active = 1
            GROUP BY period
            ORDER BY period
            """
        ).fetchall()
        products = connection.execute(
            f"""
            SELECT
                item_type,
                COALESCE(product_option, '미지정') AS product_option,
                SUM(quantity) AS quantity,
                SUM({SALES_EXPRESSION}) AS sales_amount
            FROM sales_records
            WHERE is_active = 1 AND batch_is_active = 1
            GROUP BY item_type, product_option
            ORDER BY quantity DESC
            LIMIT 100
            """
        ).fetchall()
        taxes = connection.execute(
            """
            SELECT
                SUBSTR(tax_date, 1, 7) AS period,
                SUM(credit_card_amount) AS credit_card_amount,
                SUM(cash_receipt_amount) AS cash_receipt_amount,
                SUM(tax_invoice_amount) AS tax_invoice_amount,
                SUM(other_amount) AS other_amount,
                ROUND(SUM(
                    credit_card_amount + cash_receipt_amount
                    + tax_invoice_amount + other_amount
                ) / 11.0) AS vat_amount
            FROM sales_records
            WHERE tax_date IS NOT NULL AND is_active = 1 AND batch_is_active = 1
            GROUP BY period
            ORDER BY period
            """
        ).fetchall()
        accounting = connection.execute(
            f"""
            SELECT
                platform,
                SUM({SALES_EXPRESSION}) AS sales_amount,
                ROUND(SUM({SALES_EXPRESSION}) / 1.1) AS supply_amount,
                ROUND(SUM({SALES_EXPRESSION}) / 11.0) AS vat_amount,
                SUM(fee_amount_1 + fee_amount_2) AS fee_amount
            FROM sales_records
            WHERE is_active = 1 AND batch_is_active = 1
            GROUP BY platform
            ORDER BY sales_amount DESC
            """
        ).fetchall()
        hometax_reconciliation = connection.execute(
            f"""
            WITH sales AS (
                SELECT platform, SUBSTR(tax_date, 1, 7) AS period,
                    SUM({SALES_EXPRESSION}) AS sales_amount
                FROM sales_records
                WHERE tax_date IS NOT NULL AND is_active = 1 AND batch_is_active = 1
                GROUP BY platform, period
            ),
            hometax AS (
                SELECT platform, SUBSTR(sales_date, 1, 7) AS period,
                    SUM(total_amount) AS hometax_amount
                FROM hometax_records
                WHERE platform IS NOT NULL
                    AND batch_is_active = 1
                    AND platform != '취소'
                    AND sales_date IS NOT NULL
                GROUP BY platform, period
            ),
            keys AS (
                SELECT platform, period FROM sales
                UNION
                SELECT platform, period FROM hometax
            )
            SELECT keys.period, keys.platform,
                COALESCE(sales.sales_amount, 0) AS sales_amount,
                COALESCE(hometax.hometax_amount, 0) AS hometax_amount,
                COALESCE(sales.sales_amount, 0)
                    - COALESCE(hometax.hometax_amount, 0) AS difference_amount
            FROM keys
            LEFT JOIN sales
                ON sales.platform = keys.platform
                AND sales.period = keys.period
            LEFT JOIN hometax
                ON hometax.platform = keys.platform
                AND hometax.period = keys.period
            ORDER BY keys.period, keys.platform
            """
        ).fetchall()

    return templates.TemplateResponse(
        request=request,
        name="reports.html",
        context={
            "monthly": monthly,
            "products": products,
            "taxes": taxes,
            "accounting": accounting,
            "hometax_reconciliation": hometax_reconciliation,
        },
    )


@app.get("/exports/reports.xlsx")
def export_reports():
    with get_connection() as connection:
        monthly = connection.execute(f"SELECT SUBSTR(tax_date,1,7) period, SUM(quantity) quantity, SUM({SALES_EXPRESSION}) sales_amount FROM sales_records WHERE tax_date IS NOT NULL AND is_active=1 AND batch_is_active=1 GROUP BY period ORDER BY period").fetchall()
        products = connection.execute(f"SELECT item_type, COALESCE(product_option,'미지정') product_option, SUM(quantity) quantity, SUM({SALES_EXPRESSION}) sales_amount FROM sales_records WHERE is_active=1 AND batch_is_active=1 GROUP BY item_type,product_option ORDER BY quantity DESC LIMIT 100").fetchall()
        taxes = connection.execute("SELECT SUBSTR(tax_date,1,7) period, SUM(credit_card_amount) credit_card_amount, SUM(cash_receipt_amount) cash_receipt_amount, SUM(tax_invoice_amount) tax_invoice_amount, SUM(other_amount) other_amount, ROUND(SUM(credit_card_amount+cash_receipt_amount+tax_invoice_amount+other_amount)/11.0) vat_amount FROM sales_records WHERE tax_date IS NOT NULL AND is_active=1 AND batch_is_active=1 GROUP BY period ORDER BY period").fetchall()
        accounting = connection.execute(f"SELECT platform, SUM({SALES_EXPRESSION}) sales_amount, ROUND(SUM({SALES_EXPRESSION})/1.1) supply_amount, ROUND(SUM({SALES_EXPRESSION})/11.0) vat_amount, SUM(fee_amount_1+fee_amount_2) fee_amount FROM sales_records WHERE is_active=1 AND batch_is_active=1 GROUP BY platform ORDER BY sales_amount DESC").fetchall()
        reconciliation = connection.execute(f"""WITH sales AS (SELECT platform,SUBSTR(tax_date,1,7) period,SUM({SALES_EXPRESSION}) sales_amount FROM sales_records WHERE tax_date IS NOT NULL AND is_active=1 AND batch_is_active=1 GROUP BY platform,period), hometax AS (SELECT platform,SUBSTR(sales_date,1,7) period,SUM(total_amount) hometax_amount FROM hometax_records WHERE platform IS NOT NULL AND batch_is_active=1 AND platform!='취소' AND sales_date IS NOT NULL GROUP BY platform,period), keys AS (SELECT platform,period FROM sales UNION SELECT platform,period FROM hometax) SELECT keys.period,keys.platform,COALESCE(sales.sales_amount,0) sales_amount,COALESCE(hometax.hometax_amount,0) hometax_amount,COALESCE(sales.sales_amount,0)-COALESCE(hometax.hometax_amount,0) difference_amount FROM keys LEFT JOIN sales ON sales.platform=keys.platform AND sales.period=keys.period LEFT JOIN hometax ON hometax.platform=keys.platform AND hometax.period=keys.period ORDER BY keys.period,keys.platform""").fetchall()

    workbook = Workbook()
    add_export_sheet(workbook, "월별 매출", ["연월", "수량", "매출"], [[r["period"], r["quantity"], r["sales_amount"]] for r in monthly], currency_columns=(3,))
    add_export_sheet(workbook, "제품 사이즈", ["유형", "제품사이즈", "수량", "매출"], [[product_type_label(r["item_type"]), dashboard_size_label(r["product_option"]), r["quantity"], r["sales_amount"]] for r in products], currency_columns=(4,))
    add_export_sheet(workbook, "부가세 기초자료", ["연월", "신용카드", "현금영수증", "세금계산서", "기타", "예상 부가세"], [[r["period"], r["credit_card_amount"], r["cash_receipt_amount"], r["tax_invoice_amount"], r["other_amount"], r["vat_amount"]] for r in taxes], currency_columns=(2,3,4,5,6))
    add_export_sheet(workbook, "홈택스 대사", ["연월", "플랫폼", "웹 판매자료", "홈택스 자료", "차이"], [[r["period"], r["platform"], r["sales_amount"], r["hometax_amount"], r["difference_amount"]] for r in reconciliation], currency_columns=(3,4,5))
    add_export_sheet(workbook, "회계처리 집계", ["플랫폼", "매출", "공급가액", "부가세", "수수료"], [[r["platform"], r["sales_amount"], r["supply_amount"], r["vat_amount"], r["fee_amount"]] for r in accounting], currency_columns=(2,3,4,5))
    return workbook_response(workbook, dated_excel_filename("분석세무회계"))


@app.get("/hometax")
def hometax_management(
    request: Request,
    quarter: str = Query(default=""),
):
    conditions = ["batch_is_active = 1"]
    parameters: list[object] = []
    selected_quarter = optional_number(quarter, 1, 4)

    if selected_quarter is not None:
        start_month = (selected_quarter - 1) * 3 + 1
        end_month = start_month + 2
        conditions.append("sales_month BETWEEN ? AND ?")
        parameters.extend([start_month, end_month])

    where_clause = ""

    if conditions:
        where_clause = "WHERE " + " AND ".join(conditions)

    with get_connection() as connection:
        records = connection.execute(
            f"""
            SELECT *
            FROM hometax_records
            {where_clause}
            ORDER BY sales_date DESC, id DESC
            """,
            parameters,
        ).fetchall()

    return templates.TemplateResponse(
        request=request,
        name="hometax.html",
        context={
            "records": records,
            "platforms": SUPPORTED_PLATFORMS,
            "selected_quarter": selected_quarter,
        },
    )


@app.get("/exports/hometax.xlsx")
def export_hometax(quarter: str = Query(default="")):
    conditions = ["batch_is_active = 1"]
    parameters: list[object] = []
    selected_quarter = optional_number(quarter, 1, 4)

    if selected_quarter is not None:
        start_month = (selected_quarter - 1) * 3 + 1
        conditions.append("sales_month BETWEEN ? AND ?")
        parameters.extend([start_month, start_month + 2])

    with get_connection() as connection:
        records = connection.execute(
            f"SELECT * FROM hometax_records WHERE {' AND '.join(conditions)} ORDER BY sales_date DESC,id DESC",
            parameters,
        ).fetchall()

    workbook = Workbook()
    add_export_sheet(
        workbook, "홈택스 플랫폼 지정",
        ["구분", "매출월", "매출일", "거래건수", "공급가액", "부가세", "봉사료", "합계금액", "신용카드", "기타결제", "승인번호", "판매대행사", "원본 플랫폼", "지정 플랫폼", "지정 상태", "비고"],
        [["현금과세" if r["record_type"] == "cash_taxable" else "신용카드·건별", r["sales_month"], r["sales_date"], r["transaction_count"], r["supply_amount"], r["vat_amount"], r["service_amount"], r["total_amount"], r["credit_card_amount"], r["other_payment_amount"], r["approval_number"], r["merchant_name"], r["source_platform"], r["platform"], {"manual": "직접 지정", "unassigned": "미지정"}.get(r["platform_status"], r["platform_status"]), r["note"]] for r in records],
        currency_columns=(5,6,7,8,9,10), date_columns=(3,),
    )
    label = f"홈택스플랫폼지정_{selected_quarter}분기" if selected_quarter else "홈택스플랫폼지정_전체"
    return workbook_response(workbook, dated_excel_filename(label))


@app.post("/hometax/{record_id}/platform")
def update_hometax_platform(
    record_id: int,
    platform: str = Form(...),
    quarter: int | None = Form(default=None),
):
    normalized_platform = platform.strip()

    if normalized_platform not in SUPPORTED_PLATFORMS:
        return RedirectResponse(url="/hometax", status_code=303)

    with get_connection() as connection:
        connection.execute(
            """
            UPDATE hometax_records
            SET platform = ?, platform_status = 'manual'
            WHERE id = ? AND batch_is_active = 1
            """,
            (normalized_platform, record_id),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (
                action, target_type, target_id, description
            ) VALUES ('update', 'hometax_record', ?, ?)
            """,
            (record_id, f"홈택스 플랫폼 지정: {normalized_platform}"),
        )

    redirect_url = "/hometax"

    if quarter is not None and 1 <= quarter <= 4:
        redirect_url = f"/hometax?quarter={quarter}"

    return RedirectResponse(url=redirect_url, status_code=303)


@app.post("/imports/{batch_id}/cancel")
def cancel_import(batch_id: int):
    with get_connection() as connection:
        batch = connection.execute(
            "SELECT * FROM import_batches WHERE id = ?",
            (batch_id,),
        ).fetchone()
        if batch is None or batch["status"] != "completed":
            return RedirectResponse(url="/imports", status_code=303)
        connection.execute(
            "UPDATE sales_records SET batch_is_active = 0 WHERE import_batch_id = ?",
            (batch_id,),
        )
        connection.execute(
            "UPDATE hometax_records SET batch_is_active = 0 WHERE import_batch_id = ?",
            (batch_id,),
        )
        connection.execute(
            "UPDATE import_batches SET status = 'cancelled', cancelled_at = CURRENT_TIMESTAMP WHERE id = ?",
            (batch_id,),
        )
        connection.execute(
            """
            INSERT INTO audit_logs (
                action, target_type, target_id, description
            ) VALUES ('cancel', 'import_batch', ?, ?)
            """,
            (
                batch_id,
                f"업로드 반영 취소: {batch['original_filename']}",
            ),
        )

    return RedirectResponse(url="/imports", status_code=303)


@app.post("/imports/{batch_id}/restore")
def restore_import(batch_id: int):
    with get_connection() as connection:
        batch = connection.execute(
            "SELECT * FROM import_batches WHERE id = ?", (batch_id,)
        ).fetchone()
        if batch is None or batch["status"] != "cancelled":
            return RedirectResponse(url="/imports", status_code=303)
        connection.execute(
            """
            UPDATE sales_records AS target
            SET batch_is_active = 1
            WHERE import_batch_id = ?
              AND NOT EXISTS (
                SELECT 1 FROM sales_records AS other
                WHERE other.id != target.id
                  AND other.batch_is_active = 1
                  AND SUBSTR(other.row_fingerprint, 1, 64)
                      = SUBSTR(target.row_fingerprint, 1, 64)
              )
            """,
            (batch_id,),
        )
        connection.execute(
            """
            UPDATE hometax_records AS target
            SET batch_is_active = 1
            WHERE import_batch_id = ?
              AND NOT EXISTS (
                SELECT 1 FROM hometax_records AS other
                WHERE other.id != target.id
                  AND other.batch_is_active = 1
                  AND SUBSTR(other.row_fingerprint, 1, 64)
                      = SUBSTR(target.row_fingerprint, 1, 64)
              )
            """,
            (batch_id,),
        )
        connection.execute(
            "UPDATE import_batches SET status = 'completed', restored_at = CURRENT_TIMESTAMP WHERE id = ?",
            (batch_id,),
        )
        connection.execute(
            "INSERT INTO audit_logs (action, target_type, target_id, description) VALUES ('restore', 'import_batch', ?, ?)",
            (batch_id, f"업로드 다시 복원: {batch['original_filename']}"),
        )
    return RedirectResponse(url="/imports", status_code=303)


@app.get("/imports/{batch_id}")
def import_detail(
    request: Request,
    batch_id: int,
):
    with get_connection() as connection:
        batch = connection.execute(
            """
            SELECT *
            FROM import_batches
            WHERE id = ?
            """,
            (batch_id,),
        ).fetchone()

        issues = connection.execute(
            """
            SELECT *
            FROM import_issues
            WHERE import_batch_id = ?
            ORDER BY platform, source_row, id
            """,
            (batch_id,),
        ).fetchall()
        blocking_issue_count = sum(
            1 for issue in issues if issue["severity"] == "error"
        )
        preview_rows = []

        if batch is not None and batch["status"] == "preview":
            stored_path = ensure_batch_source_file(connection, batch)
            connection.execute("SAVEPOINT detail_preview")
            try:
                preview_result = import_workbook(
                    connection,
                    stored_path,
                    batch_id,
                )
                preview_rows = preview_result.preview_rows
            finally:
                connection.execute("ROLLBACK TO detail_preview")
                connection.execute("RELEASE detail_preview")

    if batch is None:
        return RedirectResponse(
            url="/imports",
            status_code=303,
        )

    return templates.TemplateResponse(
        request=request,
        name="import_detail.html",
        context={
            "batch": batch,
            "issues": issues,
            "blocking_issue_count": blocking_issue_count,
            "preview_rows": preview_rows,
        },
    )


@app.post("/imports")
def upload_workbook(file: UploadFile = File(...)):
    original_filename = file.filename or "upload.xlsx"

    if Path(original_filename).suffix.lower() != ".xlsx":
        return RedirectResponse(
            url="/imports?error=xlsx",
            status_code=303,
        )

    stored_filename = f"{uuid4().hex}.xlsx"
    stored_path = UPLOAD_DIRECTORY / stored_filename

    source_content = file.file.read()
    stored_path.write_bytes(source_content)

    with get_connection() as connection:
        cursor = connection.execute(
            """
            INSERT INTO import_batches (
                original_filename,
                stored_filename,
                source_content,
                status
            ) VALUES (?, ?, ?, 'preview')
            RETURNING id
            """,
            (
                original_filename,
                stored_filename,
                source_content,
            ),
        )
        batch_id = cursor.fetchone()["id"]

        preview_active = False

        try:
            connection.execute("SAVEPOINT import_preview")
            preview_active = True
            result = import_workbook(
                connection=connection,
                file_path=stored_path,
                batch_id=batch_id,
            )
            preview_issues = [
                dict(row)
                for row in connection.execute(
                    "SELECT * FROM import_issues WHERE import_batch_id = ?",
                    (batch_id,),
                ).fetchall()
            ]
            connection.execute("ROLLBACK TO import_preview")
            connection.execute("RELEASE import_preview")
            preview_active = False
            connection.executemany(
                """
                INSERT INTO import_issues (
                    import_batch_id, platform, source_row, order_number,
                    field_name, issue_code, message, severity
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        issue["import_batch_id"], issue["platform"],
                        issue["source_row"], issue["order_number"],
                        issue["field_name"], issue["issue_code"],
                        issue["message"], issue["severity"],
                    )
                    for issue in preview_issues
                ],
            )
            connection.execute(
                """
                UPDATE import_batches
                SET
                    status = 'preview',
                    total_rows = ?,
                    imported_rows = ?,
                    duplicate_rows = ?,
                    issue_rows = ?
                WHERE id = ?
                """,
                (
                    result.total_rows,
                    result.imported_rows,
                    result.duplicate_rows,
                    result.issue_rows,
                    batch_id,
                ),
            )
        except Exception as error:
            if preview_active:
                connection.execute("ROLLBACK TO import_preview")
                connection.execute("RELEASE import_preview")

            connection.execute(
                """
                UPDATE import_batches
                SET
                    status = 'failed',
                    error_message = ?
                WHERE id = ?
                """,
                (
                    str(error),
                    batch_id,
                ),
            )

    return RedirectResponse(
        url=f"/imports/{batch_id}",
        status_code=303,
    )


@app.post("/imports/{batch_id}/confirm")
def confirm_import(batch_id: int):
    with get_connection() as connection:
        batch = connection.execute(
            "SELECT * FROM import_batches WHERE id = ? AND status = 'preview'",
            (batch_id,),
        ).fetchone()

        if batch is None:
            return RedirectResponse(url="/imports", status_code=303)

        blocking_issue = connection.execute(
            """
            SELECT 1 FROM import_issues
            WHERE import_batch_id = ? AND severity = 'error'
            LIMIT 1
            """,
            (batch_id,),
        ).fetchone()

        if blocking_issue is not None:
            return RedirectResponse(
                url=f"/imports/{batch_id}?error=blocking",
                status_code=303,
            )

        stored_path = ensure_batch_source_file(connection, batch)

        try:
            connection.execute(
                "DELETE FROM import_issues WHERE import_batch_id = ?",
                (batch_id,),
            )
            result = import_workbook(
                connection=connection,
                file_path=stored_path,
                batch_id=batch_id,
            )
            connection.execute(
                """
                UPDATE import_batches
                SET status = 'completed', total_rows = ?,
                    imported_rows = ?, duplicate_rows = ?, issue_rows = ?
                WHERE id = ?
                """,
                (
                    result.total_rows,
                    result.imported_rows,
                    result.duplicate_rows,
                    result.issue_rows,
                    batch_id,
                ),
            )
            connection.execute(
                """
                INSERT INTO audit_logs (
                    action, target_type, target_id, description
                ) VALUES ('import', 'import_batch', ?, ?)
                """,
                (
                    batch_id,
                    f"엑셀 최종 반영: {batch['original_filename']} ({result.imported_rows}행)",
                ),
            )
        except Exception as error:
            connection.execute(
                """
                UPDATE import_batches
                SET status = 'failed', error_message = ?
                WHERE id = ?
                """,
                (str(error), batch_id),
            )

    return RedirectResponse(url=f"/imports/{batch_id}", status_code=303)
