# B2C Sales Management

판매채널별 엑셀 데이터를 통합하여 매출, 정산, 수수료와 부가세를 관리하는 웹 프로그램입니다.

## 실행 방법

최초 환경 구성:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

서버 실행:

```powershell
.\run.ps1
```

브라우저에서 `http://127.0.0.1:8002`로 접속합니다.

## 현재 지원 기능

- 판매 현황 대시보드
- 통합 판매내역 조회 및 검색
- 기존 자동 템플릿 형식의 `.xlsx` 업로드
- 업로드 이력과 중복 행 방지
- 로컬 SQLite / 운영 PostgreSQL 자동 전환
- 업로드 원본 엑셀 DB 영구 저장

## 엑셀 업로드 규칙

기존 템플릿의 플랫폼 시트 중 다음 시트를 읽습니다.

- 네이버
- 쿠팡
- 카페24
- 11번가
- SSG
- 지마켓
- 옥션
- 베네피아
- 현대홈쇼핑
- 현대백화점

각 시트의 3행을 헤더로, 4행부터 판매 데이터로 처리합니다.

## Render + Neon 배포

`DATABASE_URL`이 없으면 로컬 `data/sales.db`를 사용하고, PostgreSQL 연결
문자열이 있으면 Neon을 사용합니다. Render Blueprint 설정은 `render.yaml`에
포함되어 있습니다.

Render 환경변수:

```env
DATABASE_URL=Neon pooled connection string
INITIAL_ADMIN_USERNAME=admin
INITIAL_ADMIN_PASSWORD=10자 이상 영문자·숫자·특수문자 포함
```

헬스체크 주소는 `/health`이며 DB에 접근하지 않습니다. Uptime 모니터도 이
주소를 사용합니다.

기존 로컬 데이터를 빈 Neon DB로 이전할 때는 PowerShell에서 `DATABASE_URL`을
현재 세션에 설정한 뒤 실행합니다.

```powershell
$env:DATABASE_URL = "postgresql://..."
.\.venv\Scripts\python.exe scripts\migrate_to_neon.py --yes
```

이전 도구는 대상 DB에 기존 데이터가 하나라도 있으면 중단합니다. 사용자,
비밀번호 해시, 판매내역, 홈택스 자료, 작업 이력과 업로드 원본 엑셀을 함께
복사합니다. 로그인 세션은 보안상 이전하지 않습니다.
