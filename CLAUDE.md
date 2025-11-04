# CLAUDE.md

이 파일은 Claude Code (claude.ai/code)가 이 저장소의 코드 작업 시 참고할 가이드입니다.

## 프로젝트 개요

Google Apps Script로 구축된 구매 관리 웹 애플리케이션입니다. 데이터 저장은 Google Sheets를, 파일 업로드는 Google Drive를 사용하는 단일 페이지 애플리케이션(SPA)입니다.

## 아키텍처

### 기술 스택
- **백엔드**: Google Apps Script (JavaScript)
- **프론트엔드**: 서버 사이드 HTML 템플릿을 사용하는 Vanilla JavaScript
- **데이터 저장소**: Google Sheets (여러 시트로 엔티티 분리)
- **파일 저장소**: Google Drive

### 파일 구조
- `Code.gs` - 서버 사이드 로직 (Google Apps Script)
- `Index.html` - 메인 HTML 템플릿
- `JavaScript.html` - 클라이언트 사이드 JavaScript (Index.html에 포함)
- `Stylesheet.html` - CSS 스타일 (Index.html에 포함)

### 주요 설정 (Code.gs)
배포 전 반드시 `Code.gs` 상단의 이 상수들을 업데이트해야 합니다:
```javascript
const SPREADSHEET_ID = "..."; // Google Sheets ID
const DRIVE_FOLDER_ID = "..."; // Google Drive 폴더 ID
```

### 데이터 아키텍처

애플리케이션은 하나의 스프레드시트 내에 **6개의 시트**를 사용합니다:

1. **PurchaseHistory** - 구매 내역 (컬럼):
   - PurchaseID, PurchaseDate, ItemName, Category, Quantity, Unit, UnitPrice, TotalPrice, SupplierID, PO_ID, StatementID, Timestamp

2. **PO_List** - 발주서 (컬럼):
   - PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON (아이템 배열 저장), Notes

3. **Statements** - 거래명세서 (파일 업로드 포함):
   - StatementID, IssueDate, SupplierID, PO_IDs, Amount, FileLink, Timestamp

4. **Suppliers** - 공급사 마스터 데이터:
   - SupplierID, SupplierName, ContactPerson, Email, Phone, Address

5. **TaxInvoices** - 세금계산서 (파일 업로드 포함):
   - TaxInvoiceID, IssueDate, Month, SupplierID, TotalAmount, FileLink, Timestamp

6. **ItemDatabase** - 아이템 마스터 데이터:
   - ItemID, ItemName, Category, Notes, SuppliersJSON (공급사별 가격 배열 저장)

### ID 생성 패턴

#### 일반 엔티티 (순차 증가):
- 구매 내역: `PH-1001`, `PH-1002`, ...
- 공급사: `S-001`, `S-002`, ...
- 아이템: `IT-1001`, `IT-1002`, ...
- 거래명세서: `ST-1001`, `ST-1002`, ...
- 세금계산서: `TAX-1001`, `TAX-1002`, ...

이들은 `getNextId(sheet, prefix)` 함수가 처리합니다.

#### 발주서 (날짜 기반):
- 발주서: `PO-YYYYMMDD001`, `PO-YYYYMMDD002`, ...
- 예시: `PO-20250104001`, `PO-20250104002`, `PO-20250105001` (다음날은 다시 001부터)

**PO 번호 규칙**:
1. 형식: `PO-YYYYMMDD###` (예: PO-20250104001)
2. 매일 001부터 시작
3. 같은 날짜 내에서는 순차적으로 증가 (001, 002, 003, ...)
4. 날짜가 바뀌면 자동으로 새로운 날짜로 리셋되고 001부터 시작
5. 시간대: 베트남 시간(Asia/Ho_Chi_Minh, GMT+7) 기준

`getNextPOId(sheet)` 전용 함수가 처리합니다.

## 애플리케이션 흐름

### 클라이언트-서버 통신
앱은 Google Apps Script의 `google.script.run` API를 사용합니다:

**클라이언트 → 서버:**
```javascript
google.script.run
    .withSuccessHandler(callback)
    .withFailureHandler(errorCallback)
    .serverFunction(data);
```

**서버 → 클라이언트:**
서버 함수는 클라이언트에서 렌더링할 데이터 또는 HTML 문자열을 반환합니다.

### 페이지 로딩 패턴
1. 사용자가 네비게이션 링크 클릭
2. 클라이언트가 `google.script.run.getPage(pageTitle)` 호출
3. 서버의 `getPage()` 함수가 적절한 페이지 함수로 라우팅 (예: `getDashboardPage()`)
4. 서버가 데이터가 포함된 HTML 문자열 생성
5. 클라이언트가 `#content-wrapper`에 HTML 렌더링
6. 클라이언트가 이벤트 위임을 사용하여 이벤트 핸들러 바인딩

### 파일 업로드 흐름
파일 업로드(거래명세서, 세금계산서)는 2단계 프로세스를 사용합니다:

1. **클라이언트**: FileReader를 사용하여 파일을 Base64로 읽음
2. **Drive에 업로드**: `saveFileToDrive(fileData, fileName)`로 파일 저장 후 URL 반환
3. **메타데이터 저장**: `addStatementRecord()` 또는 `addTaxInvoiceRecord()`로 URL과 메타데이터를 Sheets에 저장

## 주요 함수

### 서버 사이드 (Code.gs)

**유틸리티 함수:**
- `getSheet(sheetName)` - 시트 이름으로 시트 가져오기 (에러 처리 포함)
- `getSupplierMap()` - 공급사 ID를 이름으로 매핑하는 객체 반환
- `getNextId(sheet, prefix)` - 다음 순차 ID 생성 (일반 엔티티용)
- `getNextPOId(sheet)` - 날짜 기반 PO ID 생성 (PO 전용)
- `escapeHtml(text)` - XSS 방어를 위한 HTML 이스케이프
- `formatDateKR(date)` - 베트남 시간대로 날짜 포맷 (한국 형식)
- `formatDateInTimeZone(date, timezone)` - 지정된 시간대로 날짜 포맷

**CRUD 작업:**
- `addPurchaseEntry(formData)` - 구매 내역 추가
- `addSupplier(formData)` / `updateSupplier(formData)` - 공급사 CRUD
- `addPO(formData)` - 발주서 생성
- `addItem(formData)` / `updateItem(formData)` - 아이템 DB CRUD
- `deleteRowById(sheetName, id)` - 범용 삭제 함수

**페이지 생성 함수:**
- `getDashboardPage()` - 월간 요약 및 분석
- `getPurchaseHistoryPage()` - 구매 내역 테이블
- `getPurchaseRegisterPage()` - 구매 등록 폼
- `getPOPage()` - 발주서 관리
- `getStatementPage()` / `getTaxInvoicePage()` - 문서 관리 (업로드 포함)
- `getSupplierPage()` - 공급사 관리
- `getDatabasePage()` - 아이템 마스터 데이터베이스
- `getAnalysisPage()` - 기간별 분석

**분석 함수:**
- `getDashboardData()` - 월간 요약 데이터
- `getAnalysisData(options)` - 카테고리/공급사별 구매액 분석 (기간 필터)

### 클라이언트 사이드 (JavaScript.html)

**핵심 함수:**
- `loadPage(pageTitle, linkElement)` - 페이지 네비게이션
- `showModal(title, contentHtml)` / `closeModal()` - 모달 관리
- `showSpinner()` / `hideSpinner()` - 로딩 인디케이터
- `showConfirm(message, onOk)` - 확인 대화상자

**폼 핸들러:**
- `handlePurchaseFormSubmit()` - 구매 등록
- `handleSupplierFormSubmit()` - 공급사 추가/수정
- `handlePOFormSubmit()` - 발주서 생성
- `handleItemDBFormSubmit()` - 아이템 추가/수정
- `handleFileUpload(form, type)` - 거래명세서/세금계산서 파일 업로드

## 개발 참고사항

### 로컬 테스트
Google Apps Script는 로컬에서 실행할 수 없습니다. 개발 워크플로우:
1. 로컬에서 파일 편집
2. 웹 앱으로 배포: **배포 > 새 배포 > 웹 앱**
3. 배포 URL을 사용하여 브라우저에서 테스트
4. 로그 확인: **보기 > 로그** 또는 `Logger.log()`

### 디버깅
- 서버 사이드 디버깅은 `Logger.log()` 사용
- Apps Script 편집기에서 실행 로그 확인: **보기 > 실행**
- 클라이언트 사이드 디버깅은 브라우저 DevTools 콘솔 사용
- 서버 오류는 `showAlert(error.message, true)`로 표시됨

### 일반적인 수정 작업

**폼에 새 필드 추가:**
1. 페이지 생성 함수에서 HTML 업데이트 (예: `getPurchaseRegisterPage()`)
2. JavaScript.html의 폼 핸들러에서 새 필드 수집하도록 업데이트
3. 서버 사이드 함수에서 새 필드를 시트에 저장하도록 업데이트

**새 페이지 추가:**
1. Index.html 사이드바에 네비게이션 링크 추가
2. `getPage()` switch 문에 case 추가
3. Code.gs에 새 `getXxxPage()` 함수 생성
4. 필요한 CRUD 함수 추가

**시트 구조 수정:**
시트에 컬럼을 추가/제거하는 경우 다음을 업데이트:
- 해당 시트를 읽는 페이지 생성 함수
- 해당 시트에 쓰는 CRUD 함수
- `getRange()` 호출의 배열 인덱스

### 알려진 제한사항

1. **인증 없음**: 앱은 Google Apps Script의 내장 인증에 의존
2. **통화**: VND (베트남 동)로 하드코딩됨
3. **데이터 검증 없음**: 시트 수준의 데이터 검증은 수동으로 설정 필요
4. **단일 사용자**: 앱 내에서 동시 편집 처리나 사용자 권한 없음
5. **파일 크기 제한**: Google Drive 업로드 제한 (스크립트는 10MB)

## 배포 체크리스트

1. Code.gs의 `SPREADSHEET_ID` 업데이트
2. Code.gs의 `DRIVE_FOLDER_ID` 업데이트
3. 스프레드시트에 올바른 이름의 6개 시트가 모두 있는지 확인
4. 웹 앱으로 배포: **배포 > 새 배포**
5. 액세스 권한 적절하게 설정
6. 모든 CRUD 작업 테스트
7. 파일 업로드가 정상 작동하는지 확인
