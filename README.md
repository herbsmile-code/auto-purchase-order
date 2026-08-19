# 카카오톡 주문 원버튼 발주서 및 ERP 연동 시스템

카카오톡으로 수신된 주문 메시지를 자동으로 파싱하여, 단가 마스터 엑셀(`templates/price_master.xlsx`)과 매칭하고 공급사용 발주서 엑셀 및 ERP 업로드용 데이터를 자동 생성/전송하는 웹 애플리케이션입니다.

---

## 🚀 빠른 시작 (원클릭 실행)

1. **`run.bat` 더블 클릭**
   - 시스템에 `uv`가 없으면 **자동으로 초고속 다운로드 및 설치**를 진행합니다.
   - 프로젝트 전용 가상환경(`.venv`)을 생성하고 `pyproject.toml`에 명시된 라이브러리를 동기화(`uv sync`)합니다.
   - 단가 마스터 엑셀 템플릿이 없을 경우 자동 생성합니다.
   - 기본 웹 브라우저(`http://localhost:8000`)가 자동으로 열리며 서버가 시작됩니다.

---

## 🛠️ 수동 명령어 (uv CLI)

터미널에서 직접 실행할 경우:

```bash
# 가상환경 생성 및 의존성 동기화
uv sync

# 패키지 추가
uv add <패키지명>

# 서버 실행
uv run app.py

# 테스트 실행
uv run python test_system.py
```

---

## 📂 주요 파일 구조

* [pyproject.toml](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/pyproject.toml): `uv` 프로젝트 메타데이터 및 의존성 목록
* [run.bat](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/run.bat): uv 자동 감지/설치 및 원클릭 실행 배치 스크립트
* [app.py](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/app.py): FastAPI 백엔드, 파서 로직, 엑셀 생성 및 이메일/다운로드 핸들러
* [init_templates.py](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/init_templates.py): 단가 마스터 및 발주서 양식 템플릿 초기 생성기
* [templates/](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/templates): 단가표 마스터 및 서식 엑셀 파일
* [static/](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/static): 프리미엄 웹 인터페이스 UI (HTML/CSS/JS)
* [output/](file:///c:/Users/JY/Desktop/vibe_coding/Auto%20Purchase%20Order/output): 생성된 발주서 및 ERP 누적 파일 저장소
