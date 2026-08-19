from pathlib import Path

from fastapi.testclient import TestClient

from app.main import app


def main() -> None:
    output = Path("work/export-qa")
    output.mkdir(parents=True, exist_ok=True)
    client = TestClient(app)

    for name, path in {
        "validation.xlsx": "/exports/validation.xlsx",
        "reports.xlsx": "/exports/reports.xlsx",
        "hometax.xlsx": "/exports/hometax.xlsx?quarter=3",
    }.items():
        response = client.get(path)
        response.raise_for_status()
        (output / name).write_bytes(response.content)


if __name__ == "__main__":
    main()
