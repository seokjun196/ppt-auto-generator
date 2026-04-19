# 📊 PPT 자동 생성기

엑셀 데이터를 PPT 양식 파일에 자동으로 입력하는 Streamlit 웹앱입니다.

## 🚀 실행 방법

### 로컬 실행

```bash
# 1. 레포지토리 클론
git clone https://github.com/<your-username>/ppt-auto-generator.git
cd ppt-auto-generator

# 2. 패키지 설치
pip install -r requirements.txt

# 3. 앱 실행
streamlit run app.py
```

### Streamlit Cloud 배포

1. 이 레포지토리를 GitHub에 업로드합니다.
2. [share.streamlit.io](https://share.streamlit.io) 접속 후 로그인합니다.
3. **New app** → GitHub 레포지토리 선택 → `app.py` 선택 → **Deploy** 클릭합니다.

---

## 📁 파일 구성

```
├── app.py              # Streamlit 메인 앱
├── requirements.txt    # Python 패키지 목록
└── README.md           # 사용 설명서
```

---

## 📋 엑셀 파일 형식

| 직책 | 직책 폰트 | 직책 글씨크기 | 이름 | 이름 폰트 | 이름 글씨크기 |
|------|----------|------------|------|----------|------------|
| 대표이사 | 맑은 고딕 | 24 | 홍길동 | 맑은 고딕 | 36 |
| 부장 | Arial | 20 | 김철수 | Arial | 28 |

- **1행**: 헤더 (컬럼명) — 위 표와 동일하게 입력
- **2행~**: 데이터 (각 행 = PPT 슬라이드 1장)

---

## 🖼️ PPT 양식 파일 규칙

- 첫 번째 슬라이드가 **템플릿**으로 사용됩니다.
- 글상자 안에 `{직책}`, `{이름}` 플레이스홀더를 입력합니다.
- 이미지(배경, 로고 등)는 모든 슬라이드에 동일하게 반복됩니다.
- 엑셀의 행 수만큼 슬라이드가 자동 생성됩니다.

---

## ⚠️ 주의사항

- PPT 양식의 **첫 번째 슬라이드**만 템플릿으로 사용됩니다.
- 글상자 내 플레이스홀더 `{직책}`, `{이름}` 철자를 정확히 입력해야 합니다.
- 폰트명은 시스템에 설치된 폰트를 사용해야 정상 표시됩니다.
