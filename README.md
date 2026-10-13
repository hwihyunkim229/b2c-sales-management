# B2C 판매 관리

네이버, 쿠팡, 카페24 등 판매채널의 주문 자료를 한데 모아 보는 내부용 웹 서비스입니다. 업로드한 엑셀을 기준으로 매출과 정산 내역을 확인하고, 수수료·부가세 자료를 정리할 수 있습니다.

## 로컬에서 실행하기

처음 한 번만 가상환경을 만들고 필요한 패키지를 설치합니다.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

이후에는 `run.ps1`을 실행하면 됩니다.

```powershell
.\run.ps1
```

서버가 뜨면 [http://127.0.0.1:8002](http://127.0.0.1:8002)로 접속합니다. 동일한 네트워크의 다른 PC에서 접속할 때 사용할 주소는 실행 창에 같이 표시됩니다.

## 할 수 있는 일

- 월별·판매채널별 매출 현황 확인
- 판매내역 검색, 수정 및 엑셀 내보내기
- 회계, 부가세, 홈택스 자료 조회
- `.xlsx` 파일 업로드와 중복 데이터 확인
- 업로드 이력 취소, 복구 및 삭제
- 사용자별 권한 관리

## 엑셀 업로드 규칙

업로드 파일은 기존 판매 템플릿의 `.xlsx` 형식을 사용합니다. 아래 이름의 시트만 가져옵니다.

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

각 시트의 3번째 행이 헤더이며, 실제 판매 데이터는 4번째 행부터 시작해야 합니다.

## 데이터베이스와 배포

`DATABASE_URL`을 지정하지 않으면 `data/sales.db` SQLite 파일을 사용합니다. PostgreSQL 연결 문자열을 지정하면 해당 DB로 연결합니다. 현재 운영 환경은 Render와 Neon을 기준으로 하며, Render Blueprint는 `render.yaml`에 있습니다.

배포할 때는 다음 환경변수를 설정합니다.

```env
DATABASE_URL=Neon pooled connection string
INITIAL_ADMIN_USERNAME=admin
INITIAL_ADMIN_PASSWORD=10자 이상 영문자·숫자·특수문자 포함
```

헬스체크는 `/health`에서 확인할 수 있습니다. 이 경로는 DB에 접근하지 않으므로 Uptime 모니터의 확인 주소로도 사용합니다.

### 로컬 데이터 이전

로컬 SQLite 데이터를 빈 Neon DB로 옮길 때는 PowerShell 세션에 `DATABASE_URL`을 설정한 뒤 이전 스크립트를 실행합니다.

```powershell
$env:DATABASE_URL = "postgresql://..."
.\.venv\Scripts\python.exe scripts\migrate_to_neon.py --yes
```

대상 DB에 데이터가 하나라도 있으면 이전을 중단합니다. 사용자와 비밀번호 해시, 판매내역, 홈택스 자료, 작업 이력, 업로드한 원본 엑셀이 함께 복사됩니다. 로그인 세션은 옮기지 않습니다.
