# Purchase Management System

Google Apps Script 기반의 구매 관리 시스템입니다. Google Sheets를 데이터베이스로 사용하며, 웹 앱 인터페이스를 통해 구매 내역, 발주서(PO), 거래명세서, 세금계산서 등을 효율적으로 관리할 수 있습니다.

## 🚀 주요 기능

### 📊 대시보드
- 월간 구매 요약 (총 구매액, 신규 등록 건수)
- 미처리 PO 현황
- 카테고리별 구매 현황 시각화

### 📝 구매 관리
- **구매 내역 조회**: 전체 구매 기록 검색 및 필터링
- **구매 등록**: 신규 구매(입고) 등록 및 PO 연동
- **PO 관리**: 발주서 생성, 조회, 상태 관리 (발주 승인 → 입고 완료)

### 📄 문서 관리
- **거래명세서**: 파일 업로드 및 메타데이터 관리
- **세금계산서**: 월별 세금계산서 관리 및 파일 저장

### 📈 분석
- 카테고리별 구매 금액 분석 (기간별)
- 업체별 구매 금액 분석 (기간별)

### 👥 마스터 데이터
- **공급사 관리**: 공급사 정보 CRUD
- **아이템 DB**: 아이템 정보 및 공급사별 단가 관리

## 🛠️ 기술 스택

- **Backend**: Google Apps Script (JavaScript)
- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Database**: Google Sheets
- **Storage**: Google Drive
- **Deployment**: Google Apps Script Web App

## 📁 프로젝트 구조

```
purchase/
├── Code.gs              # 백엔드 로직 (Apps Script)
├── Index.html           # 메인 HTML 템플릿
├── JavaScript.html      # 클라이언트 사이드 JavaScript
├── Stylesheet.html      # CSS 스타일시트
├── GEMINI.md            # 프로젝트 규칙 및 가이드라인
├── README.md            # 이 파일
├── .agent/              # AI 에이전트 설정
│   └── workflows/       # 워크플로우 문서
│       ├── deploy.md    # 배포 가이드
│       ├── development.md # 개발 가이드
│       └── testing.md   # 테스트 가이드
└── .claude/             # Claude AI 설정
```

## 📦 설치 및 배포

### 1. Google Sheets 설정
1. 새 Google Sheets 생성
2. 다음 시트들을 생성:
   - `PurchaseHistory` - 구매 내역
   - `PO_List` - 발주서 목록
   - `Statements` - 거래명세서
   - `Suppliers` - 공급사 정보
   - `TaxInvoices` - 세금계산서
   - `ItemDatabase` - 아이템 DB

### 2. Google Drive 폴더 생성
1. 파일 업로드용 폴더 생성
2. 폴더 ID 복사 (URL에서 확인)

### 3. Google Apps Script 프로젝트 생성
1. [Google Apps Script](https://script.google.com) 접속
2. 새 프로젝트 생성
3. 파일 업로드:
   - `Code.gs`
   - `Index.html`
   - `JavaScript.html`
   - `Stylesheet.html`

### 4. 환경 설정
`Code.gs` 파일 상단의 상수 업데이트:
```javascript
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
const DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";
```

### 5. 웹 앱 배포
1. "배포" → "새 배포" 클릭
2. "웹 앱" 선택
3. 액세스 권한 설정
4. 배포 후 URL 복사

자세한 배포 가이드는 [.agent/workflows/deploy.md](.agent/workflows/deploy.md)를 참조하세요.

## 🎯 사용법

### 기본 워크플로우

1. **공급사 등록**
   - "공급사 관리" 페이지에서 신규 공급사 추가

2. **발주서(PO) 생성**
   - "PO 관리" 페이지에서 새 PO 생성
   - 아이템 정보 입력
   - 상태: "발주 승인"

3. **구매 등록**
   - "구매 등록" 페이지에서 입고 정보 입력
   - PO 선택 시 자동으로 정보 입력
   - 등록 시 PO 상태가 "입고 완료"로 자동 변경

4. **문서 관리**
   - 거래명세서/세금계산서 파일 업로드
   - 메타데이터 입력 및 저장

5. **분석**
   - "분석" 페이지에서 기간별 구매 현황 확인

## 🔧 개발

### 로컬 개발 환경
- Google Apps Script 에디터 사용 (권장)
- 또는 clasp CLI 도구 사용

자세한 개발 가이드는 [.agent/workflows/development.md](.agent/workflows/development.md)를 참조하세요.

### 코딩 규칙
- 네이밍: camelCase (함수/변수), UPPER_SNAKE_CASE (상수)
- ID 접두사: PH-, PO-, ST-, TX-, SUP-, ITEM-
- 에러 처리: try-catch 블록 필수
- 보안: HTML 출력 시 `escapeHtml()` 사용

자세한 규칙은 [GEMINI.md](GEMINI.md)를 참조하세요.

## 🧪 테스트

테스트 체크리스트 및 가이드는 [.agent/workflows/testing.md](.agent/workflows/testing.md)를 참조하세요.

## 📚 문서

- [GEMINI.md](GEMINI.md) - 프로젝트 규칙 및 가이드라인
- [.agent/workflows/deploy.md](.agent/workflows/deploy.md) - 배포 가이드
- [.agent/workflows/development.md](.agent/workflows/development.md) - 개발 가이드
- [.agent/workflows/testing.md](.agent/workflows/testing.md) - 테스트 가이드

## 🔒 보안

- 스프레드시트 및 드라이브 접근 권한 관리
- 사용자 입력 검증 및 이스케이프 처리
- 웹 앱 배포 시 액세스 권한 설정 주의

## 🐛 문제 해결

### 일반적인 오류
- **"시트를 찾을 수 없습니다"**: `SHEET_NAMES` 상수와 실제 시트 이름 확인
- **"권한이 없습니다"**: 스프레드시트/드라이브 접근 권한 확인
- **파일 업로드 실패**: `DRIVE_FOLDER_ID` 및 폴더 권한 확인

## 📝 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.

## 👨‍💻 개발자

Purchase Management System Development Team

## 📞 지원

문제가 발생하거나 질문이 있으시면 개발팀에 문의하세요.

---

**버전**: 1.0.0  
**최종 업데이트**: 2025-11-20
