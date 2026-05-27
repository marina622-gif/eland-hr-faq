# 이랜드건설 HR FAQ 챗봇

이랜드건설 임직원을 위한 HR(인사) FAQ 챗봇입니다.  
Claude AI 기반 RAG(Retrieval-Augmented Generation)로 문서를 검색해 답변을 생성하며, 관리자 페이지에서 FAQ를 직접 관리할 수 있습니다.

---

## 주요 기능

- **AI 챗봇**: 연차/휴가, 급여, 경조사, 복리후생 등 HR 관련 질문에 한국어로 답변
- **RAG 기반**: `docs/` 폴더 문서를 벡터 DB에 저장해 관련 내용만 참조
- **문의 접수**: 답변 불가 시 담당자에게 자동 이메일 알림 + 접수 확인 메일 발송
- **관리자 페이지**: FAQ CRUD, 미답변 질문 관리, 답변 등록

---

## 기술 스택

| 항목 | 내용 |
|------|------|
| Backend | Python 3.9 + Flask |
| AI | Anthropic Claude (`claude-haiku-4-5`) |
| Vector DB | ChromaDB + DefaultEmbeddingFunction |
| 문서 저장 | Excel (openpyxl) |
| 이메일 | Gmail SMTP |

---

## 설치 및 실행

### 1. 저장소 클론

```bash
git clone https://github.com/marina622-gif/eland-hr-faq.git
cd eland-hr-faq
```

### 2. 패키지 설치

```bash
pip install flask anthropic chromadb openpyxl python-dotenv
```

### 3. 환경변수 설정

```bash
cp .env.example .env
```

`.env` 파일을 열어 아래 항목을 채웁니다:

```
ANTHROPIC_API_KEY=sk-ant-...       # Anthropic 콘솔에서 발급
ADMIN_EMAIL=your@email.com         # 미답변 질문 알림 수신 이메일
EMAIL_USER=your@gmail.com          # Gmail 계정
EMAIL_PASSWORD=xxxx xxxx xxxx xxxx # Gmail 앱 비밀번호
```

> Gmail 앱 비밀번호 발급: Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호

### 4. 서버 실행

```bash
python app.py
```

---

## 접속

| 페이지 | 주소 |
|--------|------|
| 챗봇 | http://localhost:5000 |
| 관리자 | http://localhost:5000/admin |

관리자 기본 비밀번호: `admin1234`

---

## 문서 추가 방법

`docs/` 폴더에 파일을 추가하면 서버 시작 시 자동으로 인덱싱됩니다.

| 형식 | 지원 여부 |
|------|----------|
| `.xlsx` | ✅ (FAQ 시트) |
| `.txt` | ✅ |
| `.pdf` | ✅ (`pip install pdfplumber`) |
| `.docx` | ✅ (`pip install python-docx`) |

---

## 외부 공개 (ngrok)

```bash
ngrok http --domain=your-domain.ngrok-free.app 5000
```
