# Global Operation Index

Flask 기반 데이터 분석 및 대시보드 애플리케이션입니다. HR Index, LMS Learning, HONG 데이터 등 다양한 Excel 파일을 전처리하고 분석 결과를 시각화합니다.

## 주요 기능

- **데이터 전처리**: Excel 파일(HR Index, LMS Learning, HONG 데이터) 자동 처리
- **대시보드**: Bootstrap 5 기반 반응형 웹 인터페이스
- **데이터 분석**: 법인별, 지역별, 월별 데이터 분석 및 리포트 생성
- **Dynamic DataTables**: 코딩 없이 데이터 관리
- **인증**: Session 기반, GitHub, Google OAuth 지원
- **데이터베이스**: SQLite (기본), MySQL/MariaDB, PostgreSQL 전환 가능

---

## 🚀 빠른 시작 (Quick Start)

### 시스템 요구사항

- **운영체제**: macOS, Windows, Linux
- **Python**: 3.8 이상 (3.9 ~ 3.11 권장)
- **메모리**: 최소 4GB RAM 권장

### Mac (M1/M2/M3 칩 포함)

```bash
# 1. Python 버전 확인
python3 --version

# 2. 가상환경 생성
python3 -m venv venv

# 3. 가상환경 활성화
source venv/bin/activate

# 4. pip 업그레이드
pip install --upgrade pip

# 5. 의존성 설치
pip install -r requirements.txt --timeout 1000

# 6. Flask 웹서버 실행
python run.py
```

### Windows

```bash
# 1. Python 버전 확인
python --version

# 2. 가상환경 생성
python -m venv venv

# 3. 가상환경 활성화
venv\Scripts\activate

# 4. pip 업그레이드
pip install --upgrade pip

# 5. 의존성 설치
pip install -r requirements.txt --timeout 1000

# 6. Flask 웹서버 실행
python run.py
```

### 웹 브라우저에서 접속

```
http://localhost:8080
```

---

## 📖 상세 설치 가이드

### 1. Python 설치

#### macOS
```bash
# Homebrew로 Python 설치 (권장)
brew install python@3.11

# 또는 python.org에서 다운로드
# https://www.python.org/downloads/
```

#### Windows
```bash
# python.org에서 설치 프로그램 다운로드
# https://www.python.org/downloads/
# 설치 시 "Add Python to PATH" 체크 필수!
```

### 2. 프로젝트 클론

```bash
git clone <repository-url>
cd global_operation_index
```

### 3. 가상환경 설정

가상환경을 사용하면 프로젝트별로 독립적인 Python 패키지를 관리할 수 있습니다.

```bash
# 가상환경 생성
python3 -m venv venv

# 가상환경 활성화
# macOS/Linux:
source venv/bin/activate

# Windows:
venv\Scripts\activate

# 활성화 확인: 터미널 앞에 (venv)가 표시됨
```

### 4. 의존성 설치

```bash
# pip 최신 버전으로 업그레이드
pip install --upgrade pip

# 패키지 설치 (타임아웃 시간 설정)
pip install -r requirements.txt --timeout 1000
```

### 5. 데이터 전처리 (선택사항)

Excel 데이터 파일이 있는 경우에만 실행합니다.

```bash
python main.py
```

**필요한 데이터 파일 위치:**
- `data/{월}/index_management.xlsx`
- `data/{월}/prev_hr_index.xlsx`
- `data/{월}/hr_index.xlsx`
- `data/{월}/lms_learning.xlsx`
- `data/{월}/hong_data.xlsx`

> **주의**: 데이터 파일이 없어도 Flask 웹서버는 정상 실행됩니다.

### 6. Flask 웹서버 실행

```bash
python run.py
```

서버가 성공적으로 시작되면 다음과 같은 메시지가 표시됩니다:
```
* Running on http://0.0.0.0:8080
```

웹 브라우저에서 `http://localhost:8080` 접속

---

## 🔧 문제 해결 (Troubleshooting)

### 1. `python` 명령어를 찾을 수 없음 (macOS)

**문제:**
```bash
zsh: command not found: python
```

**해결:**
macOS에서는 `python3` 명령어를 사용합니다.
```bash
python3 --version
python3 run.py
```

### 2. pip 설치 시 타임아웃 에러

**문제:**
```bash
ReadTimeoutError: HTTPSConnectionPool... Read timed out.
```

**해결 방법 1: 타임아웃 시간 늘리기**
```bash
pip install -r requirements.txt --timeout 1000
```

**해결 방법 2: 한국 미러 서버 사용**
```bash
pip install -r requirements.txt --timeout 1000 --index-url https://mirror.kakao.com/pypi/simple
```

**해결 방법 3: 패키지를 나눠서 설치**
```bash
# 핵심 패키지만 먼저 설치
pip install flask==3.1.0 --timeout 1000
pip install pandas --timeout 1000
pip install openpyxl --timeout 1000

# 나머지 패키지 설치
pip install -r requirements.txt --timeout 1000
```

### 3. Permission 에러 (Windows)

**문제:**
```bash
PermissionError: [Errno 13] Permission denied
```

**해결:**
PowerShell을 관리자 권한으로 실행하거나 `--user` 플래그 사용
```bash
pip install --user -r requirements.txt
```

### 4. 가상환경 활성화가 안 됨 (Windows)

**문제:**
```bash
... cannot be loaded because running scripts is disabled on this system
```

**해결:**
PowerShell에서 실행 정책 변경
```bash
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 5. 포트 8080이 이미 사용 중

**문제:**
```bash
OSError: [Errno 48] Address already in use
```

**해결:**
`run.py` 파일에서 포트 번호 변경
```python
# run.py 파일 60번째 줄
app.run(host='0.0.0.0', port=8081)  # 8080 → 8081
```

---

## 📁 프로젝트 구조

```
global_operation_index/
├── apps/                      # Flask 애플리케이션 모듈
├── data/                      # 데이터 파일 디렉토리 (월별)
│   └── {월}/                 # 월별 데이터 폴더 (예: 9/)
├── templates/                 # HTML 템플릿
├── static/                    # CSS, JS, 이미지 파일
├── main.py                    # 데이터 전처리 스크립트
├── run.py                     # Flask 웹서버 실행 파일
├── requirements.txt           # Python 패키지 의존성
├── excel_preprocess_hr.py     # HR 데이터 전처리 모듈
├── excel_preprocess_lms.py    # LMS 데이터 전처리 모듈
├── excel_preprocess_hong.py   # HONG 데이터 전처리 모듈
├── make_logic.py              # 분석 로직 생성 모듈
├── logger_config.py           # 로깅 설정
└── README.md                  # 이 파일
```

---

## 🛠️ 기술 스택

### Backend
- **Flask 3.1.0**: Python 웹 프레임워크
- **SQLAlchemy 2.0**: ORM (Object-Relational Mapping)
- **Pandas**: 데이터 분석 및 처리
- **openpyxl**: Excel 파일 읽기/쓰기

### Frontend
- **Bootstrap 5**: 반응형 UI 프레임워크
- **Gradient Able Dashboard**: 오픈소스 대시보드 템플릿

### Database
- **SQLite**: 기본 데이터베이스 (개발/테스트)
- **MySQL/PostgreSQL**: 프로덕션 환경 (선택사항)

### 배포
- **Gunicorn**: WSGI HTTP 서버
- **Docker**: 컨테이너화 (선택사항)
- **Render/AWS**: 클라우드 배포 (선택사항)

---

## 📚 추가 문서

- **[INSTALLATION.md](./INSTALLATION.md)**: 상세 설치 가이드 및 Python 버전 호환성
- **[CHANGELOG.md](./CHANGELOG.md)**: 버전별 변경 사항

---

## 🔐 환경 변수 설정 (선택사항)

프로덕션 환경에서는 `.env` 파일을 생성하여 환경 변수를 설정합니다.

```bash
# .env 파일 생성
cp env.sample .env
```

`.env` 파일 예시:
```bash
DEBUG=False
SECRET_KEY=your-secret-key-here
DATABASE_URL=sqlite:///db.sqlite3
```

---

## 📝 데이터 분석 설정

`main.py` 파일에서 분석 기준 월을 설정할 수 있습니다:

```python
# main.py 18-20번째 줄
ANALYSIS_YEAR = 2025   # 분석 기준 년도
ANALYSIS_MONTH = 9     # 분석 기준 월 (1~12)
```

---

## 🤝 기여

이슈 및 풀 리퀘스트는 언제나 환영합니다!

---

## 📄 라이선스

이 프로젝트는 MIT 라이선스를 따릅니다. 자세한 내용은 [LICENSE.md](./LICENSE.md)를 참조하세요.

---

## 📞 지원

문제가 발생하면 다음을 확인해주세요:

1. [INSTALLATION.md](./INSTALLATION.md) - 상세 설치 가이드
2. [문제 해결](#-문제-해결-troubleshooting) 섹션
3. GitHub Issues에 문의

---

**개발 환경**: Python 3.9+ | Flask 3.1.0 | Bootstrap 5
**마지막 업데이트**: 2025년 11월
