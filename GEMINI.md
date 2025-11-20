# Purchase Management System - Project Rules

## 프로젝트 개요
Google Apps Script 기반의 구매 관리 시스템입니다. Google Sheets를 데이터베이스로 사용하며, 웹 앱 인터페이스를 통해 구매 내역, PO(Purchase Order), 거래명세서, 세금계산서 등을 관리합니다.

## 기술 스택
- **Backend**: Google Apps Script (JavaScript)
- **Frontend**: HTML, CSS, JavaScript
- **Database**: Google Sheets
- **Storage**: Google Drive (파일 업로드)

## 프로젝트 구조
```
purchase/
├── Code.gs              # 백엔드 로직 (Apps Script)
├── Index.html           # 메인 HTML 템플릿
├── JavaScript.html      # 클라이언트 사이드 JavaScript
├── Stylesheet.html      # CSS 스타일시트
├── README.md            # 프로젝트 설명
├── CLAUDE.md            # Claude AI 관련 문서
├── CODE_REVIEW_REPORT.md
├── IMPROVEMENTS_COMPLETED.md
├── PO_NUMBER_CHANGE.md
└── .claude/             # Claude 설정 디렉토리
```

## 주요 기능
1. **대시보드**: 월간 구매 요약, 카테고리별 현황
2. **구매 내역**: 구매 기록 조회 및 검색
3. **구매 등록**: 신규 구매(입고) 등록
4. **PO 관리**: 발주서 생성, 조회, 상태 관리
5. **거래명세서 관리**: 파일 업로드 및 관리
6. **세금계산서 관리**: 파일 업로드 및 관리
7. **분석**: 카테고리별/업체별 구매 금액 분석
8. **공급사 관리**: 공급사 정보 CRUD
9. **데이터베이스**: 아이템 DB 관리

## 데이터 시트 구조
- **PurchaseHistory**: 구매 내역
- **PO_List**: 발주서 목록
- **Statements**: 거래명세서
- **Suppliers**: 공급사 정보
- **TaxInvoices**: 세금계산서
- **ItemDatabase**: 아이템 데이터베이스

## 코딩 규칙

### 1. Google Apps Script 규칙
- 모든 서버 사이드 함수는 `Code.gs`에 작성
- 상수는 파일 상단에 정의 (SPREADSHEET_ID, DRIVE_FOLDER_ID 등)
- 시트 이름은 `SHEET_NAMES` 객체에서 관리
- 에러 처리는 try-catch 블록 사용 필수
- Logger.log()를 사용하여 디버깅 정보 기록

### 2. HTML/JavaScript 규칙
- 클라이언트 사이드 코드는 `JavaScript.html`에 작성
- 스타일은 `Stylesheet.html`에 작성
- 서버 함수 호출 시 `google.script.run` 사용
- 비동기 처리 시 `.withSuccessHandler()` 및 `.withFailureHandler()` 사용

### 3. 네이밍 컨벤션
- 함수명: camelCase (예: `getPurchaseHistoryPage`)
- 상수: UPPER_SNAKE_CASE (예: `SPREADSHEET_ID`)
- 변수: camelCase
- ID 접두사:
  - PH-: Purchase History
  - PO-: Purchase Order
  - ST-: Statement
  - TX-: Tax Invoice
  - SUP-: Supplier
  - ITEM-: Item Database

### 4. 데이터 검증
- 필수 필드는 반드시 검증
- 숫자 타입은 `Number()` 변환 후 `isNaN()` 체크
- 날짜는 `new Date()` 객체로 변환
- HTML 출력 시 `escapeHtml()` 함수 사용 (XSS 방지)

### 5. 파일 업로드
- 최대 파일 크기: 10MB
- 허용 확장자: pdf, jpg, jpeg, png, gif, doc, docx, xls, xlsx
- Google Drive 폴더 ID: `DRIVE_FOLDER_ID` 상수 사용

### 6. ID 생성 규칙
- 시작 번호: 1000
- 첫 번째 ID 접미사: 1001
- PO 시퀀스 길이: 3자리 (001, 002, ...)
- 형식: `{PREFIX}-{NUMBER}` (예: PH-1001)

## 개발 워크플로우

### 1. 로컬 개발
- Google Apps Script 에디터에서 직접 편집
- 또는 clasp CLI 도구 사용 (선택사항)

### 2. 배포
- Google Apps Script 에디터에서 "배포" → "새 배포" 선택
- 웹 앱으로 배포
- 액세스 권한 설정

### 3. 테스트
- 웹 앱 URL로 접속하여 기능 테스트
- Logger.log() 출력은 Apps Script 에디터의 "실행 로그"에서 확인

## 보안 고려사항
- SPREADSHEET_ID와 DRIVE_FOLDER_ID는 환경에 맞게 설정
- 웹 앱 배포 시 액세스 권한 주의
- 사용자 입력은 항상 검증 및 이스케이프 처리
- 민감한 정보는 Properties Service 사용 권장

## 문제 해결
- 스프레드시트 접근 오류: SPREADSHEET_ID 확인
- 시트를 찾을 수 없음: SHEET_NAMES 확인
- 파일 업로드 실패: DRIVE_FOLDER_ID 및 권한 확인
- 함수 실행 오류: Apps Script 실행 로그 확인

## 참고 문서
- [Google Apps Script 공식 문서](https://developers.google.com/apps-script)
- [Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet)
- [Drive Service](https://developers.google.com/apps-script/reference/drive)
- [HTML Service](https://developers.google.com/apps-script/reference/html)

## 변경 이력
- 최초 작성: 2025-11-20
- 코드베이스 초기화 완료
