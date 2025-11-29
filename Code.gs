/* eslint-disable no-unused-vars */
// Google Apps Script: 최상위 함수들은 외부에서 호출되므로 "unused" 경고를 무시합니다

// !!! 중요 !!!
// Google 스프레드시트 파일의 URL에서 ID를 복사하여 여기에 붙여넣으세요.
// 예: "https://docs.google.com/spreadsheets/d/1abcdefgHIJKLMN.../edit"
// 이전 대화에서 "1e9dnHTEYIvSwOgjIU_-U3fXWD4JcXZwh2PGwdeA45s0" 이 ID를 사용하신 것을 확인했습니다.
const SPREADSHEET_ID = "1e9dnHTEYIvSwOgjIU_-U3fXWD4JcXZwh2PGwdeA45s0";

// !!! 중요 !!!
// 파일 업로드(거래명세서/세금계산서)를 저장할 Google Drive 폴더 ID
const DRIVE_FOLDER_ID = "1SAs0VJ_3CWyqC_5nVguIeIWEU2nCRu7v";

// 시트 이름 정의 (image_573dff.png 스크린샷 기준)
const SHEET_NAMES = {
  HISTORY: "PurchaseHistory",
  PO: "PO_List",
  STATEMENTS: "Statements",
  SUPPLIERS: "Suppliers",
  TAX: "TaxInvoices",
  DB: "ItemDatabase"
};

// ID 생성 관련 상수
const ID_GENERATION = {
  INITIAL_NUMBER: 1000,      // ID 시작 번호 (일반 ID용)
  FIRST_ID_SUFFIX: "1001",   // 첫 번째 ID의 접미사 (일반 ID용)
  HEADER_ROW: 1,             // 헤더 행 번호
  DATA_START_ROW: 2,         // 데이터 시작 행 번호
  PO_SEQUENCE_LENGTH: 3      // PO 시퀀스 번호 자릿수 (001, 002, ...)
};

// 파일 업로드 관련 상수
const FILE_UPLOAD = {
  MAX_SIZE_MB: 10,                                    // 최대 파일 크기 (MB)
  MAX_SIZE_BYTES: 10 * 1024 * 1024,                  // 최대 파일 크기 (bytes)
  ALLOWED_EXTENSIONS: ["pdf", "jpg", "jpeg", "png", "gif", "doc", "docx", "xls", "xlsx"]
};

// --- 웹 앱 진입점 ---

/**
 * 앱이 로드될 때 'Index.html'을 사용자에게 보여줍니다.
 */
function doGet(_e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("구매 관리 앱")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * HTML 템플릿 내에서 다른 파일(CSS, JS)을 포함시킵니다.
 * 예: <?!= include('Stylesheet'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- 페이지 로드 함수 (클라이언트에서 호출) ---

/**
 * 클라이언트(JavaScript)에서 요청한 페이지 이름에 따라
 * 해당 페이지를 구성하는 HTML을 반환합니다.
 */
function getPage(pageTitle) {
  try {
    switch (pageTitle) {
    case "대시보드":
      return getDashboardPage();
    case "구매 내역":
      return getPurchaseHistoryPage();
    case "구매 등록":
      return getPurchaseRegisterPage();
    case "PO 관리":
      return getPOPage();
    case "거래명세서 관리":
      return getStatementPage();
    case "세금계산서 관리":
      return getTaxInvoicePage();
    case "분석":
      return getAnalysisPage();
    case "공급사 관리":
      return getSupplierPage();
    case "데이터베이스":
      return getDatabasePage();
    default:
      return "<h3>페이지를 찾을 수 없습니다.</h3>";
    }
  } catch (e) {
    Logger.log(e);
    // getPage 자체에서 발생하는 심각한 오류 기록
    return `<p style="color: red;">페이지 로드 중 심각한 오류 발생: ${e.message}</p>`;
  }
}

// --- 유틸리티: 스프레드시트 접근 ---
/**
 * SPREADSHEET_ID가 유효한지 확인하고 시트 객체를 반환합니다.
 * @param {string} sheetName 시트 이름
 * @returns {Sheet} Google 시트 객체
 * @throws {Error} ID가 설정되지 않았거나 시트를 찾을 수 없는 경우
 */
function getSheet(sheetName) {
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || !SPREADSHEET_ID || SPREADSHEET_ID.length < 20) {
    throw new Error("Code.gs 파일 상단의 SPREADSHEET_ID를 실제 ID로 변경해야 합니다.");
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    Logger.log(`스프레드시트 열기 실패. ID: ${SPREADSHEET_ID}, 오류: ${e.message}`);
    throw new Error(`스프레드시트 ID가 잘못되었거나 접근 권한이 없습니다. (ID: ${SPREADSHEET_ID})`);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`시트를 찾을 수 없습니다. 이름: "${sheetName}". Google Sheets 파일을 확인하세요.`);
  }

  return sheet;
}


// --- 개별 페이지 HTML 생성 ---

/**
 * 대시보드 페이지 HTML 생성
 */
function getDashboardPage() {
  try {
    const data = getDashboardData();
    return `
        <h3>월간 요약</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1.5rem;">
            <div class="card">
                <h4 style="color: var(--color-text-secondary); margin-top: 0;">총 구매액 (이번 달)</h4>
                <div style="font-size: 1.8rem; font-weight: 700; color: var(--color-primary);">${data.totalSpentThisMonth.toLocaleString("en-US")} VND</div>
            </div>
            <div class="card">
                <h4 style="color: var(--color-text-secondary); margin-top: 0;">신규 등록 (이번 달)</h4>
                <div style="font-size: 1.8rem; font-weight: 700;">${data.newItemsThisMonth} 건</div>
            </div>
            <div class="card">
                <h4 style="color: var(--color-text-secondary); margin-top: 0;">미처리 PO</h4>
                <div style="font-size: 1.8rem; font-weight: 700; color: #e74c3c;">${data.pendingPOs} 건</div>
            </div>
        </div>
        <h3 style="margin-top: 2rem;">카테고리별 구매 현황 (이번 달)</h3>
        <div class="card" style="min-height: 200px; padding: 1.5rem;">
            ${data.categorySpendHtml}
        </div>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'대시보드' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 구매 내역 페이지 HTML 생성
 */
function getPurchaseHistoryPage() {
  try {
    const sheet = getSheet(SHEET_NAMES.HISTORY);
    const lastRow = sheet.getLastRow();
    let tableRows = "";
    const supplierMap = getSupplierMap(); // 공급사 맵핑

    if (lastRow > 1) {
      // A:N (14개 컬럼)의 모든 데이터를 가져옵니다. 헤더는 제외 (2행부터).
      const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

      data.reverse().forEach(row => {
        let purchaseDate = formatDateKR(row[1]);
        let supplierName = escapeHtml(supplierMap[row[8]] || row[8]);

        // PO 수량 정보 (M열: row[12], N열: row[13])
        const poOrderedQty = row[12];
        const poRemainingQty = row[13];
        const hasPOQty = poOrderedQty !== null && poOrderedQty !== "";

        // PO 수량 표시 (PO가 있고 수량 정보가 있는 경우)
        let poQtyDisplay = "-";
        let remainingDisplay = "-";
        if (hasPOQty) {
          poQtyDisplay = poOrderedQty;
          remainingDisplay = poRemainingQty !== null ? poRemainingQty : "-";
        }

        tableRows += "<tr>" +
            "<td>" + escapeHtml(purchaseDate) + "</td>" +
            "<td>" + escapeHtml(row[2]) + "</td>" +
            "<td>" + escapeHtml(row[3]) + "</td>" +
            "<td>" + escapeHtml(row[4]) + " " + escapeHtml(row[5]) + "</td>" +
            "<td>" + poQtyDisplay + "</td>" +
            "<td>" + remainingDisplay + "</td>" +
            "<td>" + Number(row[6]).toLocaleString("en-US") + " VND</td>" +
            "<td>" + Number(row[7]).toLocaleString("en-US") + " VND</td>" +
            "<td>" + supplierName + "</td>" +
            "<td>" + escapeHtml(row[9] || "-") + "</td>" +
          "</tr>";
      });
    }

    return `
        <div class="card">
            <div class="d-flex justify-between align-center">
                <h3 style="margin: 0;">구매 내역 조회</h3>
                <input type="search" class="form-control" placeholder="아이템명, 공급사, PO 번호 검색..." style="width: 300px;">
            </div>
        </div>
        <table class="styled-table">
            <thead>
                <tr>
                    <th>구매일</th>
                    <th>아이템명</th>
                    <th>카테고리</th>
                    <th>입고수량</th>
                    <th>PO수량</th>
                    <th>미입고</th>
                    <th>단가</th>
                    <th>총액</th>
                    <th>공급사</th>
                    <th>PO 번호</th>
                </tr>
            </thead>
            <tbody>
                ${tableRows || "<tr><td colspan=\"10\" style=\"text-align:center;\">데이터가 없습니다.</td></tr>"}
            </tbody>
        </table>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'구매 내역' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 구매 등록 페이지 HTML 생성
 */
function getPurchaseRegisterPage() {
  try {
    // 공급사 목록을 가져와서 <option> 태그로 만듭니다.
    const supplierMap = getSupplierMap();
    let supplierOptions = "<option value=\"\">공급사를 선택하세요</option>";
    for (const id in supplierMap) {
      supplierOptions += `<option value="${escapeHtml(id)}">${escapeHtml(supplierMap[id])}</option>`;
    }

    // "발주 승인" 상태의 PO 목록 가져오기
    const approvedPOs = getApprovedPOs();
    let poOptions = "<option value=\"\">-- PO 선택 (선택사항) --</option>";
    approvedPOs.forEach(po => {
      const statusLabel = po.status === "부분 입고" ? " [부분입고]" : "";
      poOptions += `<option value="${escapeHtml(po.id)}">${escapeHtml(po.id)} - ${escapeHtml(po.supplierName)} (${po.totalAmount.toLocaleString("en-US")} VND)${statusLabel}</option>`;
    });

    return `
        <h3>신규 구매(입고) 등록</h3>
        <p>입고된 아이템의 정보를 확인하고 등록합니다.</p>
        <form id="purchase-form" style="max-width: 700px;">
            <div class="form-group">
                <label for="reg-po">PO 번호 (선택)</label>
                <select id="reg-po" name="poId" class="form-control" onchange="handlePOSelection(this.value)">
                    ${poOptions}
                </select>
            </div>

            <!-- PO 아이템 목록 (PO 선택 시 표시) -->
            <div id="po-items-section" style="display: none;" class="card mb-2">
                <h5 style="margin-bottom: 10px;">📦 PO 아이템 입고 현황</h5>
                <div id="po-items-list"></div>
                <p style="font-size: 0.85em; color: #666; margin-top: 10px;">* 입고할 아이템을 선택하면 아래 폼에 자동 입력됩니다.</p>
            </div>

            <input type="hidden" id="reg-item-index" name="itemIndex" value="">

            <div class="form-group">
                <label for="reg-item">아이템명</label>
                <input type="text" id="reg-item" name="itemName" class="form-control" placeholder="예: A4 복사용지" required>
            </div>
            <div class="form-group">
                <label for="reg-category">카테고리</label>
                <input type="text" id="reg-category" name="category" class="form-control" placeholder="예: 사무용품">
            </div>
            <div class="form-group">
                <label for="reg-supplier">공급사</label>
                <select id="reg-supplier" name="supplierId" class="form-control" required>
                    ${supplierOptions}
                </select>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem;">
                <div class="form-group">
                    <label for="reg-date">구매(입고)일</label>
                    <input type="date" id="reg-date" name="purchaseDate" class="form-control" required>
                </div>
                <div class="form-group">
                    <label for="reg-qty">수량</label>
                    <input type="number" id="reg-qty" name="quantity" class="form-control" value="1" required>
                    <small id="reg-qty-hint" style="color: #666; display: none;">잔여: <span id="reg-remaining-qty">0</span></small>
                </div>
                <div class="form-group">
                    <label for="reg-unit">단위</label>
                    <input type="text" id="reg-unit" name="unit" class="form-control" placeholder="예: 박스, 개" required>
                </div>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem;">
                <div class="form-group">
                    <label for="reg-price">단가 (VND)</label>
                    <input type="number" id="reg-price" name="unitPrice" class="form-control" placeholder="0" required>
                </div>
                <div class="form-group">
                    <label for="reg-total">총액 (VND)</label>
                    <input type="number" id="reg-total" name="totalPrice" class="form-control" placeholder="0" readonly style="background: #f4f4f4;">
                </div>
            </div>
            <button type="submit" class="btn btn-primary">구매 내역 등록</button>
        </form>
        <div id="form-message" class="mt-1"></div>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'구매 등록' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * PO 관리 페이지 HTML 생성
 */
function getPOPage() {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const lastRow = sheet.getLastRow();
    let tableRows = "";
    const supplierMap = getSupplierMap();

    if (lastRow > 1) {
      // PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

      data.reverse().forEach(row => {
        let issueDate = formatDateKR(row[1]);
        let supplierName = escapeHtml(supplierMap[row[2]] || row[2]);
        let status = escapeHtml(row[4] || "N/A");
        let statusClass = "";

        if (row[4] === "발주 승인") statusClass = "status-pending";
        else if (row[4] === "부분 입고") statusClass = "status-partial";
        else if (row[4] === "입고 완료") statusClass = "status-completed";
        else if (row[4] === "취소") statusClass = "status-cancelled";

        // 취소 버튼은 "취소" 상태가 아닐 때만 표시
        let cancelButton = "";
        if (row[4] !== "취소") {
          cancelButton = `<button class="btn btn-warning btn-sm" onclick="cancelPO('${escapeHtml(row[0])}')">취소</button>`;
        }

        tableRows += `
          <tr>
              <td>${escapeHtml(row[0])}</td> <!-- PO_ID -->
              <td>${escapeHtml(issueDate)}</td> <!-- IssueDate -->
              <td>${supplierName}</td>
              <td>${Number(row[3]).toLocaleString("en-US")} VND</td> <!-- TotalAmount -->
              <td><span class="badge ${statusClass}">${status}</span></td>
              <td>
                  <button class="btn btn-secondary btn-sm" onclick="viewPOItems('${escapeHtml(row[0])}')">아이템 보기</button>
                  ${cancelButton}
              </td>
          </tr>
        `;
      });
    }

    return `
      <h3>발주서(PO) 관리</h3>
      <div class="card mb-2">
          <h4>신규 발주서 생성</h4>
          <button class="btn btn-primary" onclick="showAddPOModal()">+ 신규 PO 생성</button>
      </div>
      <div class="card">
          <h4>PO 목록</h4>
          <table class="styled-table">
              <thead>
                  <tr>
                      <th>PO 번호</th>
                      <th>발행일</th>
                      <th>공급사</th>
                      <th>총액</th>
                      <th>상태</th>
                      <th>작업</th>
                  </tr>
              </thead>
              <tbody>
                  ${tableRows || "<tr><td colspan=\"6\">등록된 PO가 없습니다.</td></tr>"}
              </tbody>
          </table>
      </div>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">PO 페이지 로드 중 오류: ${e.message}</p>`;
  }
}

/**
 * 거래명세서 관리 페이지 HTML 생성
 */
function getStatementPage() {
  try {
    // 공급사 목록 (드롭다운용)
    const supplierMap = getSupplierMap();
    let supplierOptions = "<option value=\"\">-- 공급사 선택 --</option>";
    for (const id in supplierMap) {
      supplierOptions += `<option value="${escapeHtml(id)}">${escapeHtml(supplierMap[id])}</option>`;
    }

    // PO 목록 (체크박스용) - 상태가 "입고 완료"인 PO만
    const poSheet = getSheet(SHEET_NAMES.PO);
    const poLastRow = poSheet.getLastRow();
    let poCheckboxes = "";

    if (poLastRow > 1) {
      const poData = poSheet.getRange(2, 1, poLastRow - 1, 5).getValues();
      poData.forEach(row => {
        const poId = row[0];
        const poStatus = row[4];
        if (poStatus === "입고 완료") {
          poCheckboxes += `
            <label style="display: inline-flex; align-items: center; margin-right: 1rem;">
              <input type="checkbox" name="po_ids" value="${escapeHtml(poId)}" style="margin-right: 0.25rem;">
              ${escapeHtml(poId)}
            </label>
          `;
        }
      });
    }

    const sheet = getSheet(SHEET_NAMES.STATEMENTS);
    const lastRow = sheet.getLastRow();
    let tableRows = "";

    if (lastRow > 1) {
      // StatementID, IssueDate, SupplierID, PO_IDs, Amount, FileLink, Timestamp
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

      data.reverse().forEach(row => {
        let issueDate = formatDateKR(row[1]);
        let supplierName = escapeHtml(supplierMap[row[2]] || row[2]);
        let fileLink = row[5] ? `<a href="${escapeHtml(row[5])}" target="_blank" class="btn btn-secondary btn-sm">파일 보기</a>` : "파일 없음";

        tableRows += `
          <tr>
              <td>${escapeHtml(row[0])}</td> <!-- StatementID -->
              <td>${escapeHtml(issueDate)}</td> <!-- IssueDate -->
              <td>${supplierName}</td>
              <td>${escapeHtml(row[3]) || "-"}</td> <!-- PO_IDs -->
              <td>${Number(row[4]).toLocaleString("en-US")} VND</td> <!-- Amount -->
              <td>${fileLink}</td>
          </tr>
        `;
      });
    }

    return `
      <h3>거래명세서 등록/관리</h3>
      <div class="card mb-2">
          <h4>신규 거래명세서 업로드</h4>
           <form id="upload-form-statement">
              <div class="form-group">
                  <label for="file-input-statement">파일 선택 (PDF, 이미지 등)</label>
                  <input type="file" id="file-input-statement" name="file" class="form-control" required>
              </div>
              <div class="form-group">
                  <label for="statement-supplier">공급사</label>
                  <select id="statement-supplier" name="supplierId" class="form-control" required>
                      ${supplierOptions}
                  </select>
              </div>
              <div class="form-group">
                  <label for="statement-date">발행일</label>
                  <input type="date" id="statement-date" name="issueDate" class="form-control" required>
              </div>
              <div class="form-group">
                  <label>연결된 PO (선택사항)</label>
                  <div style="border: 1px solid #dee2e6; border-radius: 0.25rem; padding: 0.75rem; max-height: 150px; overflow-y: auto;">
                      ${poCheckboxes || "<p style=\"color: #6c757d; margin: 0;\">입고 완료된 PO가 없습니다.</p>"}
                  </div>
              </div>
              <div class="form-group">
                  <label for="statement-amount">총액 (VND)</label>
                  <input type="number" id="statement-amount" name="amount" class="form-control" placeholder="0" required>
              </div>
              <button type="submit" class="btn btn-primary mt-1">업로드 및 저장</button>
           </form>
      </div>
      <table class="styled-table">
          <thead>
              <tr>
                  <th>명세서 ID</th>
                  <th>발행일</th>
                  <th>공급사</th>
                  <th>연결 PO</th>
                  <th>총액</th>
                  <th>파일</th>
              </tr>
          </thead>
          <tbody>
              ${tableRows || "<tr><td colspan=\"6\" style=\"text-align:center;\">데이터가 없습니다.</td></tr>"}
          </tbody>
      </table>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'거래명세서 관리' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 세금계산서 관리 페이지 HTML 생성
 */
function getTaxInvoicePage() {
  try {
    // 공급사 목록 (드롭다운용)
    const supplierMap = getSupplierMap();
    let supplierOptions = "<option value=\"\">-- 공급사 선택 --</option>";
    for (const id in supplierMap) {
      supplierOptions += `<option value="${escapeHtml(id)}">${escapeHtml(supplierMap[id])}</option>`;
    }

    const sheet = getSheet(SHEET_NAMES.TAX);
    const lastRow = sheet.getLastRow();
    let tableRows = "";

    if (lastRow > 1) {
      // TaxInvoiceID, IssueDate, Month, SupplierID, TotalAmount, FileLink
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

      data.reverse().forEach(row => {
        let issueDate = formatDateKR(row[1]);
        let supplierName = escapeHtml(supplierMap[row[3]] || row[3]);
        let fileLink = row[5] ? `<a href="${escapeHtml(row[5])}" target="_blank" class="btn btn-secondary btn-sm">파일 보기</a>` : "파일 없음";

        tableRows += `
          <tr>
              <td>${escapeHtml(row[0])}</td> <!-- TaxInvoiceID -->
              <td>${escapeHtml(issueDate)}</td> <!-- IssueDate -->
              <td>${escapeHtml(row[2])}월</td> <!-- Month -->
              <td>${supplierName}</td>
              <td>${Number(row[4]).toLocaleString("en-US")} VND</td> <!-- TotalAmount -->
              <td>${fileLink}</td>
          </tr>
        `;
      });
    }

    return `
      <h3>세금계산서 등록/관리</h3>
      <div class="card mb-2">
          <h4>신규 세금계산서 업로드</h4>
           <form id="upload-form-tax">
              <div class="form-group">
                  <label for="file-input-tax">파일 선택 (PDF, 이미지 등)</label>
                  <input type="file" id="file-input-tax" name="file" class="form-control" required>
              </div>
              <div class="form-group">
                  <label for="tax-supplier">공급사</label>
                  <select id="tax-supplier" name="supplierId" class="form-control" required>
                      ${supplierOptions}
                  </select>
              </div>
              <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem;">
                  <div class="form-group">
                      <label for="tax-date">발행일</label>
                      <input type="date" id="tax-date" name="issueDate" class="form-control" required>
                  </div>
                  <div class="form-group">
                      <label for="tax-month">귀속월</label>
                      <input type="number" id="tax-month" name="month" class="form-control" placeholder="예: 10" min="1" max="12" required>
                  </div>
                  <div class="form-group">
                      <label for="tax-amount">총액 (VND)</label>
                      <input type="number" id="tax-amount" name="totalAmount" class="form-control" placeholder="0" required>
                  </div>
              </div>
              <button type="submit" class="btn btn-primary mt-1">업로드 및 저장</button>
           </form>
      </div>
      <table class="styled-table">
          <thead>
              <tr>
                  <th>계산서 ID</th>
                  <th>발행일</th>
                  <th>귀속월</th>
                  <th>공급사</th>
                  <th>총액</th>
                  <th>파일</th>
              </tr>
          </thead>
          <tbody>
              ${tableRows || "<tr><td colspan=\"6\" style=\"text-align:center;\">데이터가 없습니다.</td></tr>"}
          </tbody>
      </table>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'세금계산서 관리' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 데이터베이스 페이지 HTML 생성
 */
function getDatabasePage() {
  try {
    const sheet = getSheet(SHEET_NAMES.DB);
    const lastRow = sheet.getLastRow();
    let tableRows = "";

    if (lastRow > 1) {
      // ItemID(A), Code(B), ItemName(C), Category(D), Notes(E), SuppliersJSON(F)
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

      data.forEach(row => {
        let suppliers = "정보 없음";
        try {
          // row[5] = SuppliersJSON
          if (row[5] && row[5] !== "[]") {
            const supplierData = JSON.parse(row[5]);
            suppliers = supplierData.map(s => `${escapeHtml(s.name)} (${Number(s.price).toLocaleString("en-US")} VND)`).join("<br>");
          }
        } catch (_e) {
          // JSON 파싱 실패 시
          suppliers = escapeHtml(row[5]) || "정보 없음 (파싱 오류)";
        }

        tableRows += `
          <tr>
              <td>${escapeHtml(row[0])}</td> <!-- ItemID -->
              <td>${escapeHtml(row[1] || "-")}</td> <!-- Code -->
              <td>${escapeHtml(row[2])}</td> <!-- ItemName -->
              <td>${escapeHtml(row[3])}</td> <!-- Category -->
              <td>${suppliers}</td> <!-- SuppliersJSON 파싱 -->
              <td>
                  <button class="btn btn-secondary btn-sm" onclick="showEditItemModal('${escapeHtml(row[0])}')">수정</button>
              </td>
          </tr>
        `;
      });
    }

    return `
      <div class="d-flex justify-between align-center mb-2">
          <h3 style="margin:0;">아이템 DB 관리</h3>
          <button type="button" class="btn btn-primary" onclick="showAddItemModal()">신규 아이템 추가</button>
      </div>
      <table class="styled-table">
          <thead>
              <tr>
                  <th>아이템 ID</th>
                  <th>Code</th>
                  <th>아이템명</th>
                  <th>카테고리</th>
                  <th>공급사/단가 정보</th>
                  <th>관리</th>
              </tr>
          </thead>
          <tbody>
              ${tableRows || "<tr><td colspan=\"6\" style=\"text-align:center;\">데이터가 없습니다.</td></tr>"}
          </tbody>
      </table>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'데이터베이스' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 공급사 관리 페이지 HTML 생성
 */
function getSupplierPage() {
  try {
    const sheet = getSheet(SHEET_NAMES.SUPPLIERS);
    const lastRow = sheet.getLastRow();
    let tableRows = "";

    if (lastRow > 1) {
      // SupplierID, SupplierName, ContactPerson, Email, Phone
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      data.forEach(row => {
        // row[0] = SupplierID, row[1] = SupplierName, ...
        tableRows += `
          <tr data-id="${escapeHtml(row[0])}">
              <td>${escapeHtml(row[1])}</td> <!-- SupplierName -->
              <td>${escapeHtml(row[2])}</td> <!-- ContactPerson -->
              <td>${escapeHtml(row[4])}</td> <!-- Phone -->
              <td>${escapeHtml(row[3])}</td> <!-- Email -->
              <td>
                  <button class="btn btn-secondary btn-sm" onclick="editSupplier('${escapeHtml(row[0])}')">수정</button>
                  <button class="btn btn-danger btn-sm" onclick="deleteSupplier('${escapeHtml(row[0])}')">삭제</button>
              </td>
          </tr>
        `;
      });
    }

    return `
        <div class="d-flex justify-between align-center mb-2">
            <h3 style="margin:0;">공급사 관리</h3>
            <button type="button" class="btn btn-primary" onclick="showAddSupplierModal()">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
                신규 공급사 추가
            </button>
        </div>
        <table class="styled-table">
            <thead>
                <tr>
                    <th>공급사명</th>
                    <th>담당자</th>
                    <th>연락처</th>
                    <th>이메일</th>
                    <th>관리</th>
                </tr>
            </thead>
            <tbody>
                ${tableRows || "<tr><td colspan=\"5\" style=\"text-align:center;\">데이터가 없습니다.</td></tr>"}
            </tbody>
        </table>
    `;
  } catch (e) {
    Logger.log(e);
    return `<p style="color: red;">'공급사 관리' 로드 실패: ${e.message}</p>`;
  }
}

/**
 * 분석 페이지 HTML 생성
 */
function getAnalysisPage() {
  return `
      <nav class="tab-nav">
          <div class="tab-link active" data-tab="tab-category">카테고리별 정산</div>
          <div class="tab-link" data-tab="tab-supplier">업체별 정산</div>
      </nav>
      <div class="tab-content active" id="tab-category">
          <h4>카테고리별 구매 금액 (기간별)</h4>
          <div class="card d-flex align-center" style="gap: 1rem; flex-wrap: wrap;">
              <label for="cat-start">시작일</label>
              <input type="date" id="cat-start" class="form-control" style="width: 200px;">
              <label for="cat-end">종료일</label>
              <input type="date" id="cat-end" class="form-control" style="width: 200px;">
              <button class="btn btn-primary" onclick="handleAnalysisQuery('category')">조회</button>
          </div>
          <div id="analysis-results-category" class="card" style="min-height: 200px; margin-top: 1rem; color: var(--color-text-secondary);">
              (조회 버튼을 클릭하세요)
          </div>
      </div>
      <div class="tab-content" id="tab-supplier">
          <h4>업체별 구매 금액 (기간별)</h4>
          <div class="card d-flex align-center" style="gap: 1rem; flex-wrap: wrap;">
              <label for="sup-start">시작일</label>
              <input type="date" id="sup-start" class="form-control" style="width: 200px;">
              <label for="sup-end">종료일</label>
              <input type="date" id="sup-end" class="form-control" style="width: 200px;">
              <button class="btn btn-primary" onclick="handleAnalysisQuery('supplier')">조회</button>
          </div>
          <div id="analysis-results-supplier" class="card" style="min-height: 200px; margin-top: 1rem; color: var(--color-text-secondary);">
              (조회 버튼을 클릭하세요)
          </div>
      </div>
  `;
}

// --- CRUD 함수 (데이터 등록/수정/삭제) ---

/**
 * '구매 등록' 폼에서 데이터를 받아 시트에 추가합니다.
 * @param {Object} formData 폼 데이터 객체
 * @returns {String} 성공/실패 메시지
 */
function addPurchaseEntry(formData) {
  try {
    // 필수 필드 검증
    const requiredFields = {
      "itemName": "아이템명",
      "supplierId": "공급사",
      "purchaseDate": "구매일",
      "quantity": "수량",
      "unit": "단위",
      "unitPrice": "단가"
    };

    for (const [field, label] of Object.entries(requiredFields)) {
      if (!formData[field] || formData[field].toString().trim() === "") {
        throw new Error(`필수 항목이 누락되었습니다: ${label}`);
      }
    }

    // 데이터 타입 검증
    const quantity = Number(formData.quantity);
    const unitPrice = Number(formData.unitPrice);

    if (isNaN(quantity) || quantity <= 0) {
      throw new Error("수량은 0보다 큰 숫자여야 합니다.");
    }

    if (isNaN(unitPrice) || unitPrice < 0) {
      throw new Error("단가는 0 이상의 숫자여야 합니다.");
    }

    const sheet = getSheet(SHEET_NAMES.HISTORY);

    // 1. 새 ID 생성 (예: PH-1001)
    const newId = getNextId(sheet, "PH-");

    // 2. 총액 계산
    const totalPrice = Number(formData.quantity) * Number(formData.unitPrice);

    // 3. PO 수량 정보 조회 (PO가 있는 경우)
    let poOrderedQty = null;
    let poRemainingQty = null;

    if (formData.poId && formData.poId.trim() !== "") {
      try {
        const poData = getPOItemsForReceiving(formData.poId);
        const itemIndex = formData.itemIndex !== "" ? Number(formData.itemIndex) : 0;
        if (poData.items && poData.items[itemIndex]) {
          const item = poData.items[itemIndex];
          poOrderedQty = item.orderedQty;
          // 이번 입고 후 남은 수량 계산
          poRemainingQty = item.remainingQty - Number(formData.quantity);
          if (poRemainingQty < 0) poRemainingQty = 0;
        }
      } catch (e) {
        Logger.log(`PO 수량 정보 조회 실패: ${e.message}`);
      }
    }

    // 4. 시트에 추가할 데이터 행
    // 컬럼: A~L(기존) + M(POOrderedQty) + N(PORemainingQty)
    const newRow = [
      newId,
      new Date(formData.purchaseDate),
      formData.itemName,
      formData.category,
      formData.quantity,
      formData.unit,
      formData.unitPrice,
      totalPrice,
      formData.supplierId,
      formData.poId || null, // PO ID는 선택 사항
      null, // StatementID (초기엔 null)
      new Date(), // Timestamp
      poOrderedQty, // M열: PO 주문 수량
      poRemainingQty // N열: 미입고 수량
    ];

    sheet.appendRow(newRow);

    // PO 번호가 있으면 해당 PO의 아이템 입고 수량 업데이트
    if (formData.poId && formData.poId.trim() !== "") {
      try {
        const itemIndex = formData.itemIndex !== "" ? Number(formData.itemIndex) : null;
        updatePOItemReceived(formData.poId, itemIndex, Number(formData.quantity));
        Logger.log(`PO ${formData.poId} 아이템 입고 수량이 업데이트되었습니다.`);
      } catch (e) {
        Logger.log(`PO 입고 수량 업데이트 실패 (계속 진행): ${e.message}`);
        // PO 업데이트 실패해도 구매 등록은 성공으로 처리
      }
    }

    return "구매 내역이 성공적으로 등록되었습니다.";
  } catch (e) {
    Logger.log(e);
    // 클라이언트로 오류 메시지를 전파하기 위해 예외를 다시 던집니다.
    throw new Error(`등록 실패: ${e.message}`);
  }
}

/**
 * '공급사 관리'에서 새 공급사를 시트에 추가합니다.
 * @param {Object} formData 폼 데이터 (name, contactPerson, email, phone, address)
 * @returns {String} 성공 메시지
 */
function addSupplier(formData) {
  try {
    Logger.log("=== addSupplier 호출됨 ===");
    Logger.log("formData: " + JSON.stringify(formData));

    // 필수 필드 검증
    if (!formData.name || formData.name.trim() === "") {
      throw new Error("필수 항목이 누락되었습니다: 공급사명");
    }

    Logger.log("공급사명 검증 통과: " + formData.name);

    const sheet = getSheet(SHEET_NAMES.SUPPLIERS);
    Logger.log("시트 가져오기 완료: " + sheet.getName());

    const newId = getNextId(sheet, "S-");
    Logger.log("새 ID 생성 완료: " + newId);

    const newRow = [
      newId,
      formData.name,
      formData.contactPerson,
      formData.email,
      formData.phone,
      formData.address || null // 주소는 선택 사항
    ];

    Logger.log("새 행 데이터: " + JSON.stringify(newRow));

    sheet.appendRow(newRow);
    Logger.log("✅ 시트에 행 추가 완료");

    return `공급사 '${formData.name}'이(가) 성공적으로 등록되었습니다.`;
  } catch (e) {
    Logger.log("❌ 에러 발생: " + e.message);
    Logger.log(e);
    throw new Error(`공급사 등록 실패: ${e.message}`);
  }
}

/**
 * '공급사 수정'을 위해 특정 ID의 공급사 데이터를 가져옵니다.
 * @param {string} id 공급사 ID
 * @returns {Object} 공급사 데이터
 */
function getSupplierDetails(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.SUPPLIERS);
    const data = sheet.getDataRange().getValues();

    // 헤더(1행) 건너뛰고 ID(A열, index 0) 검색
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        return {
          id: data[i][0],
          name: data[i][1],
          contactPerson: data[i][2],
          email: data[i][3],
          phone: data[i][4],
          address: data[i][5]
        };
      }
    }
    throw new Error("해당 ID의 공급사를 찾을 수 없습니다.");
  } catch (e) {
    Logger.log(e);
    throw new Error(`데이터 조회 실패: ${e.message}`);
  }
}

/**
 * '공급사 수정' 폼에서 데이터를 받아 시트의 기존 행을 업데이트합니다.
 * @param {Object} formData 폼 데이터 (id 포함)
 * @returns {String} 성공 메시지
 */
function updateSupplier(formData) {
  try {
    const sheet = getSheet(SHEET_NAMES.SUPPLIERS);
    const data = sheet.getDataRange().getValues();

    // 헤더(1행) 건너뛰고 ID(A열, index 0) 검색
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == formData.id) {
        // 행을 찾았으면 해당 행(i+1)의 데이터 업데이트
        // (ID, Name, Contact, Email, Phone, Address)
        const range = sheet.getRange(i + 1, 1, 1, 6);
        range.setValues([[
          formData.id,
          formData.name,
          formData.contactPerson,
          formData.email,
          formData.phone,
          formData.address || null
        ]]);
        return `공급사 '${formData.name}' 정보가 업데이트되었습니다.`;
      }
    }
    throw new Error("업데이트할 공급사를 찾지 못했습니다.");
  } catch (e) {
    Logger.log(e);
    throw new Error(`업데이트 실패: ${e.message}`);
  }
}


/**
 * 범용 삭제 함수. 시트 이름과 ID를 받아 해당 행을 삭제합니다.
 * @param {string} sheetName 시트 이름 (SHEET_NAMES 객체의 값)
 * @param {string} id 삭제할 항목의 ID
 * @returns {String} 성공 메시지
 */
function deleteRowById(sheetName, id) {
  if (!sheetName || !id) {
    throw new Error("시트 이름과 ID가 필요합니다.");
  }

  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();

    // 헤더(1행) 건너뛰고 ID(A열, index 0) 검색
    // (참고: 아래에서 위로 반복해야 삭제 시 인덱스 문제가 없습니다)
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1); // (i+1)이 실제 시트 행 번호
        return `ID '${id}' 항목이 ${sheetName} 시트에서 삭제되었습니다.`;
      }
    }

    throw new Error("삭제할 ID를 찾지 못했습니다.");
  } catch (e) {
    Logger.log(e);
    throw new Error(`삭제 실패: ${e.message}`);
  }
}

/**
 * '분석' 페이지용 데이터 집계
 * @param {Object} options { type: 'category' | 'supplier', startDate: 'YYYY-MM-DD', endDate: 'YYYY-MM-DD' }
 * @returns {Object} { title: '...', htmlTable: '...' }
 */
function getAnalysisData(options) {
  try {
    const sheet = getSheet(SHEET_NAMES.HISTORY);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { title: "데이터 없음", htmlTable: "<tr><td>데이터가 없습니다.</td></tr>" };

    // [PurchaseDate(B), Category(D), TotalPrice(H), SupplierID(I)]
    // 인덱스:  1,            3,           7,             8
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    const startDate = new Date(options.startDate);
    const endDate = new Date(options.endDate);
    // 종료일의 시간을 23:59:59로 설정하여 해당 날짜를 포함시킴
    endDate.setHours(23, 59, 59, 999);

    const aggregation = {};
    let total = 0;

    data.forEach(row => {
      const purchaseDate = new Date(row[1]);
      if (purchaseDate >= startDate && purchaseDate <= endDate) {
        const amount = Number(row[7]);
        if (isNaN(amount)) return; // 금액이 숫자가 아니면 스킵

        let key;

        if (options.type === "category") {
          key = row[3] || "미분류"; // Category
        } else { // 'supplier'
          key = row[8] || "미지정"; // SupplierID
        }

        if (!aggregation[key]) {
          aggregation[key] = 0;
        }
        aggregation[key] += amount;
        total += amount;
      }
    });

    // SupplierID를 이름으로 변환 (supplier 타입일 경우)
    let finalAggregation = aggregation;
    if (options.type === "supplier") {
      const supplierMap = getSupplierMap();
      finalAggregation = {};
      for (const key in aggregation) {
        const supplierName = supplierMap[key] || key;
        finalAggregation[supplierName] = aggregation[key];
      }
    }

    // 결과를 금액순으로 정렬
    const sorted = Object.entries(finalAggregation).sort(([, a], [, b]) => b - a);

    // HTML 테이블 생성
    let htmlTable = `
      <thead><tr><th>항목</th><th>금액 (VND)</th><th>비율</th></tr></thead>
      <tbody>
    `;
    sorted.forEach(([key, amount]) => {
      const percentage = total === 0 ? 0 : (amount / total * 100).toFixed(1);
      htmlTable += `
        <tr>
          <td>${key}</td>
          <td>${amount.toLocaleString("en-US")}</td>
          <td>${percentage}%</td>
        </tr>
      `;
    });
    htmlTable += `
      <tr style="font-weight: bold; border-top: 2px solid var(--color-border);">
        <td>총계</td>
        <td>${total.toLocaleString("en-US")}</td>
        <td>100%</td>
      </tr>
      </tbody>
    `;

    return {
      title: `${options.startDate} ~ ${options.endDate} (총 ${total.toLocaleString("en-US")} VND)`,
      htmlTable: htmlTable
    };

  } catch (e) {
    Logger.log(e);
    throw new Error(`분석 데이터 집계 실패: ${e.message}`);
  }
}

/**
 * '대시보드'용 데이터 집계
 * @returns {Object} 대시보드 요약 데이터
 */
function getDashboardData() {
  let totalSpentThisMonth = 0;
  let newItemsThisMonth = 0;
  let categorySpend = {};
  let pendingPOs = 0;

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  try {
    // 1. 구매 내역 분석 (이번 달)
    const historySheet = getSheet(SHEET_NAMES.HISTORY);
    if (historySheet.getLastRow() > 1) {
      // [PurchaseDate(B), Category(D), TotalPrice(H), Timestamp(L)]
      // 인덱스 1, 3, 7, 11
      const historyData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 12).getValues();
      historyData.forEach(row => {
        const timestamp = new Date(row[11]); // Timestamp
        if (timestamp >= startOfMonth) {
          newItemsThisMonth++;
        }

        const purchaseDate = new Date(row[1]);
        if (purchaseDate >= startOfMonth) {
          const amount = Number(row[7]); // TotalPrice
          if(!isNaN(amount)) {
            totalSpentThisMonth += amount;

            const category = row[3] || "미분류";
            if(!categorySpend[category]) categorySpend[category] = 0;
            categorySpend[category] += amount;
          }
        }
      });
    }
  } catch(e) {
    Logger.log(`대시보드(History) 집계 실패: ${e.message}`);
    // 이 시트가 없어도 대시보드의 다른 부분은 작동하도록 계속 진행
  }

  try {
    // 2. PO 내역 분석 (미처리 건)
    const poSheet = getSheet(SHEET_NAMES.PO);
    if (poSheet.getLastRow() > 1) {
      // [Status(E)] 인덱스 4
      const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 5).getValues();
      poData.forEach(row => {
        const status = row[4];
        if (status === "발주 승인") {
          pendingPOs++;
        }
      });
    }
  } catch (e) {
    Logger.log(`대시보드(PO) 집계 실패: ${e.message}`);
  }

  let categorySpendHtml = Object.entries(categorySpend)
    .sort(([,a],[,b]) => b-a)
    .map(([key, val]) => `<div>${key}: <strong>${val.toLocaleString("en-US")} VND</strong></div>`)
    .join("");

  return {
    totalSpentThisMonth,
    newItemsThisMonth,
    pendingPOs,
    categorySpendHtml: categorySpendHtml || "데이터 없음"
  };
}

// --- 유틸리티 함수 ---

/**
 * HTML 특수문자를 이스케이프하여 XSS 공격을 방지합니다.
 * @param {string} text 이스케이프할 텍스트
 * @returns {string} 이스케이프된 텍스트
 */
function escapeHtml(text) {
  if (!text) return "";
  return text.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * 날짜를 베트남 시간대(GMT+7)로 포맷합니다.
 * @param {Date|string} date 포맷할 날짜
 * @param {string} timezone 시간대 (기본값: Asia/Ho_Chi_Minh)
 * @returns {string} 포맷된 날짜 문자열 (YYYY.MM.DD 형식)
 */
function formatDateInTimeZone(date, timezone) {
  if (!date) return "";
  timezone = timezone || "Asia/Ho_Chi_Minh";

  try {
    const dateObj = date instanceof Date ? date : new Date(date);
    return Utilities.formatDate(dateObj, timezone, "yyyy.MM.dd");
  } catch (e) {
    Logger.log(`날짜 포맷 오류: ${e.message}`);
    return "";
  }
}

/**
 * 날짜를 한국 형식으로 포맷합니다 (베트남 시간대 사용).
 * @param {Date|string} date 포맷할 날짜
 * @returns {string} 포맷된 날짜 문자열 (예: 2025. 1. 4.)
 */
function formatDateKR(date) {
  if (!date) return "";

  try {
    const dateObj = date instanceof Date ? date : new Date(date);
    return Utilities.formatDate(dateObj, "Asia/Ho_Chi_Minh", "yyyy. M. d.");
  } catch (e) {
    Logger.log(`날짜 포맷 오류: ${e.message}`);
    return "";
  }
}

/**
 * 공급사 ID와 이름을 매핑하는 객체(Map)를 반환합니다. (드롭다운, 표시에 사용)
 * @returns {Object} { "S-001": "오피스월드", "S-002": "테크링크" }
 */
function getSupplierMap() {
  try {
    const sheet = getSheet(SHEET_NAMES.SUPPLIERS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {}; // 데이터 없음

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A열(ID), B열(Name)

    let map = {};
    data.forEach(row => {
      if (row[0] && row[1]) {
        map[row[0]] = row[1];
      }
    });
    return map;
  } catch (e) {
    Logger.log(e);
    // getSheet에서 발생한 오류(예: ID 틀림) 또는 여기서 발생한 오류를 다시 던짐
    throw new Error(`공급사 목록을 불러오는 데 실패했습니다: ${e.message}`);
  }
}

/**
 * 시트에서 마지막 ID를 찾아 다음 ID를 생성합니다. (예: PH-1000 -> PH-1001)
 * @param {Sheet} sheet 대상 시트
 * @param {String} prefix ID 접두사 (예: "PH-")
 * @returns {String} 신규 ID
 */
function getNextId(sheet, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ID_GENERATION.DATA_START_ROW) {
    return prefix + ID_GENERATION.FIRST_ID_SUFFIX; // 데이터가 없는 경우
  }

  let lastIdNum = ID_GENERATION.INITIAL_NUMBER;
  try {
    // A열 전체에서 마지막 ID를 찾는 더 안전한 방법
    const idColumn = sheet.getRange(ID_GENERATION.DATA_START_ROW, 1, lastRow - 1, 1).getValues();
    idColumn.forEach(row => {
      if (row[0] && typeof row[0] === "string" && row[0].startsWith(prefix)) {
        const parts = row[0].split("-");
        if (parts.length === 2) {
          const num = parseInt(parts[1], 10);
          if (!isNaN(num) && num > lastIdNum) {
            lastIdNum = num;
          }
        }
      }
    });
    return prefix + (lastIdNum + 1);
  } catch(e) {
    Logger.log(`getNextId 오류: ${e.message}. 임시 ID 생성.`);
    // getRange 실패 등
    return prefix + (ID_GENERATION.INITIAL_NUMBER + lastRow + 1); // 임시 ID
  }
}

/**
 * PO 전용 ID 생성 함수: PO-YYYYMMDD001 형식
 * 매일 001부터 시작하며, 같은 날짜 내에서는 순차적으로 증가합니다.
 * @param {Sheet} sheet PO_List 시트
 * @returns {String} 신규 PO ID (예: PO-20250104001)
 */
function getNextPOId(sheet) {
  // 오늘 날짜를 YYYYMMDD 형식으로 가져오기
  const today = new Date();
  const dateStr = Utilities.formatDate(today, "Asia/Ho_Chi_Minh", "yyyyMMdd");
  const prefix = "PO-" + dateStr;

  const lastRow = sheet.getLastRow();
  if (lastRow < ID_GENERATION.DATA_START_ROW) {
    return prefix + "001"; // 첫 PO인 경우
  }

  let maxSequence = 0;

  try {
    // A열에서 오늘 날짜로 시작하는 PO ID를 찾아 최대 시퀀스 번호를 구합니다
    const idColumn = sheet.getRange(ID_GENERATION.DATA_START_ROW, 1, lastRow - 1, 1).getValues();

    idColumn.forEach(row => {
      const poId = row[0];
      if (poId && typeof poId === "string" && poId.startsWith(prefix)) {
        // PO-20250104001 -> 001 추출
        const sequencePart = poId.substring(prefix.length);
        const sequence = parseInt(sequencePart, 10);
        if (!isNaN(sequence) && sequence > maxSequence) {
          maxSequence = sequence;
        }
      }
    });

    // 다음 시퀀스 번호를 지정된 자릿수로 패딩
    const nextSequence = (maxSequence + 1).toString().padStart(ID_GENERATION.PO_SEQUENCE_LENGTH, "0");
    return prefix + nextSequence;

  } catch(e) {
    Logger.log(`getNextPOId 오류: ${e.message}`);
    // 오류 발생 시 기본값 반환
    return prefix + "001";
  }
}

// --- 파일 업로드 및 복잡한 UI 기능 (서버측) ---

/**
 * 'PO 아이템 보기' 버튼 클릭 시 PO의 아이템 목록을 HTML 테이블로 반환
 * @param {string} poId - PO ID
 * @returns {string} HTML 테이블 문자열
 */
function getPOItemsAsHtml(poId) {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const data = sheet.getDataRange().getValues();
    let itemsJSON = null;

    // PO ID (A열)를 찾아 ItemsJSON (F열, 인덱스 5)을 가져옵니다.
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == poId) {
        itemsJSON = data[i][5];
        break;
      }
    }

    if (!itemsJSON) {
      return "<p>해당 PO의 아이템 정보를 찾을 수 없거나 아이템이 없습니다.</p>";
    }

    let items;
    try {
      items = JSON.parse(itemsJSON);
    } catch (e) {
      Logger.log(`JSON 파싱 오류 (PO: ${poId}): ${e.message}`);
      return "<p>아이템 데이터 형식 오류. 관리자에게 문의하세요.</p>";
    }
    let tableHtml = `
      <table class="styled-table" style="width: 100%; box-shadow: none; border: 1px solid var(--color-border);">
        <thead>
          <tr>
            <th>아이템명</th>
            <th>카테고리</th>
            <th>수량</th>
            <th>단위</th>
            <th>단가 (VND)</th>
            <th>합계 (VND)</th>
          </tr>
        </thead>
        <tbody>
    `;
    let total = 0;
    items.forEach(item => {
      const itemTotal = (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0);
      total += itemTotal;
      tableHtml += `
        <tr>
          <td>${escapeHtml(item.name || "")}</td>
          <td>${escapeHtml(item.category || "")}</td>
          <td>${escapeHtml(item.quantity || 0)}</td>
          <td>${escapeHtml(item.unit || "")}</td>
          <td>${(Number(item.unitPrice) || 0).toLocaleString("en-US")}</td>
          <td>${itemTotal.toLocaleString("en-US")}</td>
        </tr>
      `;
    });
    tableHtml += `
        <tr style="font-weight: bold; border-top: 2px solid var(--color-border);">
          <td colspan="5" style="text-align: right;">총계</td>
          <td>${total.toLocaleString("en-US")}</td>
        </tr>
      </tbody>
    </table>
    `;

    return tableHtml;
  } catch (e) {
    Logger.log(`getPOItemsAsHtml 오류: ${e.message}`);
    throw new Error(`아이템 목록 로드 실패: ${e.message}`);
  }
}

/**
 * '새 PO 생성' 모달을 띄우는 데 필요한 데이터를 조회합니다.
 * @returns {Object} { suppliers: [...], items: [...] }
 */
function getPOModalData() {
  try {
    // 1. 공급사 목록 (ID, Name)
    const supplierMap = getSupplierMap();
    const suppliers = Object.entries(supplierMap).map(([id, name]) => ({ id, name }));

    // 2. 아이템 목록 (Code, Name, Category, DefaultUnit, DefaultPrice)
    // ItemDatabase (시트 6)에서 데이터를 가져옵니다.
    const dbSheet = getSheet(SHEET_NAMES.DB);
    const lastRow = dbSheet.getLastRow();
    let items = [];
    if (lastRow > 1) {
      // ItemID(A), Code(B), ItemName(C), Category(D), Notes(E), SuppliersJSON(F)
      const data = dbSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      data.forEach(row => {
        let defaultUnit = "ea";
        let defaultPrice = 0;

        try {
          // SuppliersJSON에서 첫 번째 공급사의 단가/단위를 기본값으로 사용
          if(row[5] && row[5] !== "[]") {
            const supplierInfo = JSON.parse(row[5]);
            if (Array.isArray(supplierInfo) && supplierInfo.length > 0) {
              defaultUnit = supplierInfo[0].unit || "ea";
              defaultPrice = supplierInfo[0].price || 0;
            }
          }
        } catch(e) {
          Logger.log(`SuppliersJSON 파싱 실패 (아이템: ${row[2]}): ${e.message}`);
          // 기본값 사용
        }

        items.push({
          code: row[1] || "", // Code
          name: row[2], // ItemName
          category: row[3] || "", // Category
          defaultUnit: defaultUnit,
          defaultPrice: defaultPrice
        });
      });
    }

    return { suppliers, items };
  } catch (e) {
    Logger.log(`getPOModalData 오류: ${e.message}`);
    throw new Error(`PO 생성 데이터 로드 실패: ${e.message}`);
  }
}

/**
 * '새 PO 생성' 폼 데이터를 받아 시트에 추가합니다.
 * @param {Object} formData { supplierId, notes, items: [...] }
 * @returns {String} 성공 메시지
 */
function addPO(formData) {
  try {
    // 필수 필드 검증
    if (!formData.supplierId || formData.supplierId.trim() === "") {
      throw new Error("필수 항목이 누락되었습니다: 공급사");
    }

    if (!formData.items || !Array.isArray(formData.items) || formData.items.length === 0) {
      throw new Error("최소 1개 이상의 아이템이 필요합니다.");
    }

    // 각 아이템 검증
    formData.items.forEach((item, index) => {
      if (!item.name || item.name.trim() === "") {
        throw new Error(`아이템 ${index + 1}: 아이템명이 필요합니다.`);
      }
      if (!item.category || item.category.trim() === "") {
        throw new Error(`아이템 ${index + 1}: 카테고리가 필요합니다.`);
      }
      if (!item.quantity || Number(item.quantity) <= 0) {
        throw new Error(`아이템 ${index + 1}: 수량은 0보다 커야 합니다.`);
      }
      if (!item.unit || item.unit.trim() === "") {
        throw new Error(`아이템 ${index + 1}: 단위가 필요합니다.`);
      }
      if (isNaN(Number(item.unitPrice)) || Number(item.unitPrice) < 0) {
        throw new Error(`아이템 ${index + 1}: 단가는 0 이상의 숫자여야 합니다.`);
      }
    });

    const sheet = getSheet(SHEET_NAMES.PO);
    const newId = getNextPOId(sheet);  // PO 전용 ID 생성 함수 사용

    // 총액 계산 및 receivedQty 초기화
    let totalAmount = 0;
    const itemsWithReceived = formData.items.map(item => {
      totalAmount += (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0);
      return {
        ...item,
        receivedQty: 0  // 입고 수량 초기화
      };
    });

    const itemsJSON = JSON.stringify(itemsWithReceived);

    const newRow = [
      newId,
      new Date(), // IssueDate (오늘)
      formData.supplierId,
      totalAmount,
      "발주 승인", // 기본 상태
      itemsJSON,
      formData.notes || null
    ];

    sheet.appendRow(newRow);
    return `PO '${newId}'가 성공적으로 생성되었습니다.`;
  } catch (e) {
    Logger.log(`addPO 오류: ${e.message}`);
    throw new Error(`PO 생성 실패: ${e.message}`);
  }
}

/**
 * PO 상태를 업데이트합니다.
 * @param {String} poId PO ID
 * @param {String} newStatus 새 상태 ("발주 승인", "입고 완료", "취소")
 * @returns {String} 성공 메시지
 */
function updatePOStatus(poId, newStatus) {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const data = sheet.getDataRange().getValues();

    // 헤더 제외하고 PO_ID로 검색 (A열, index 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == poId) {
        // E열 (index 4)가 Status
        sheet.getRange(i + 1, 5).setValue(newStatus);
        return `PO '${poId}'의 상태가 '${newStatus}'(으)로 변경되었습니다.`;
      }
    }

    throw new Error(`PO '${poId}'를 찾을 수 없습니다.`);
  } catch (e) {
    Logger.log(`updatePOStatus 오류: ${e.message}`);
    throw new Error(`PO 상태 변경 실패: ${e.message}`);
  }
}

/**
 * PO 아이템의 입고 수량을 업데이트하고 PO 상태를 자동 변경합니다.
 * @param {String} poId PO ID
 * @param {Number|null} itemIndex 아이템 인덱스 (null이면 첫 번째 아이템)
 * @param {Number} receivedQty 이번에 입고된 수량
 * @returns {String} 성공 메시지
 */
function updatePOItemReceived(poId, itemIndex, receivedQty) {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == poId) {
        // PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON, Notes
        const itemsJSON = data[i][5];
        let items = [];

        try {
          items = JSON.parse(itemsJSON);
        } catch(e) {
          Logger.log(`JSON 파싱 오류 (PO: ${poId}): ${e.message}`);
          throw new Error("PO 아이템 데이터 파싱 실패");
        }

        // itemIndex가 null이거나 유효하지 않으면 0 사용
        const targetIndex = (itemIndex !== null && itemIndex >= 0 && itemIndex < items.length)
          ? itemIndex : 0;

        // 해당 아이템의 receivedQty 업데이트
        const currentReceived = Number(items[targetIndex].receivedQty) || 0;
        items[targetIndex].receivedQty = currentReceived + receivedQty;

        // 전체 입고 상태 확인
        let allComplete = true;
        let anyReceived = false;

        items.forEach(item => {
          const ordered = Number(item.quantity) || 0;
          const received = Number(item.receivedQty) || 0;

          if (received > 0) anyReceived = true;
          if (received < ordered) allComplete = false;
        });

        // 상태 결정
        let newStatus;
        if (allComplete) {
          newStatus = "입고 완료";
        } else if (anyReceived) {
          newStatus = "부분 입고";
        } else {
          newStatus = "발주 승인";
        }

        // ItemsJSON과 Status 업데이트 (F열=6, E열=5)
        const updatedItemsJSON = JSON.stringify(items);
        sheet.getRange(i + 1, 6).setValue(updatedItemsJSON);  // ItemsJSON
        sheet.getRange(i + 1, 5).setValue(newStatus);  // Status

        Logger.log(`PO ${poId} 아이템[${targetIndex}] 입고: ${receivedQty}, 새 상태: ${newStatus}`);
        return `입고 수량 업데이트 완료. PO 상태: ${newStatus}`;
      }
    }

    throw new Error(`PO '${poId}'를 찾을 수 없습니다.`);
  } catch (e) {
    Logger.log(`updatePOItemReceived 오류: ${e.message}`);
    throw new Error(`PO 입고 수량 업데이트 실패: ${e.message}`);
  }
}

/**
 * 입고 가능한 PO 목록을 가져옵니다. ("발주 승인" 또는 "부분 입고" 상태)
 * @returns {Array} PO 목록 [{ id, supplierName, totalAmount, status }]
 */
function getApprovedPOs() {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const lastRow = sheet.getLastRow();
    const approvedPOs = [];

    if (lastRow > 1) {
      // PO_ID, IssueDate, SupplierID, TotalAmount, Status
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      const supplierMap = getSupplierMap();

      data.forEach(row => {
        // "발주 승인" 또는 "부분 입고" 상태의 PO만 포함
        if (row[4] === "발주 승인" || row[4] === "부분 입고") {
          approvedPOs.push({
            id: row[0],
            supplierId: row[2],
            supplierName: supplierMap[row[2]] || row[2],
            totalAmount: Number(row[3]),
            status: row[4]
          });
        }
      });
    }

    return approvedPOs;
  } catch (e) {
    Logger.log(`getApprovedPOs 오류: ${e.message}`);
    return [];
  }
}

/**
 * 특정 PO의 상세 정보를 가져옵니다 (아이템 목록 포함).
 * @param {String} poId PO ID
 * @returns {Object} PO 상세 정보 { id, supplierId, items: [...] }
 */
function getPODetails(poId) {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const data = sheet.getDataRange().getValues();

    // 헤더 제외하고 PO_ID로 검색
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == poId) {
        // PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON, Notes
        const itemsJSON = data[i][5];
        let items = [];

        try {
          items = JSON.parse(itemsJSON);
        } catch(e) {
          Logger.log(`JSON 파싱 오류 (PO: ${poId}): ${e.message}`);
          items = [];
        }

        return {
          id: data[i][0],
          supplierId: data[i][2],
          items: items
        };
      }
    }

    throw new Error(`PO '${poId}'를 찾을 수 없습니다.`);
  } catch (e) {
    Logger.log(`getPODetails 오류: ${e.message}`);
    throw new Error(`PO 상세 정보 로드 실패: ${e.message}`);
  }
}

/**
 * 입고 등록용 PO 아이템 목록을 가져옵니다 (잔여 수량 포함).
 * @param {String} poId PO ID
 * @returns {Object} { poId, supplierId, status, items: [...] }
 */
function getPOItemsForReceiving(poId) {
  try {
    const sheet = getSheet(SHEET_NAMES.PO);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == poId) {
        // PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON, Notes
        const itemsJSON = data[i][5];
        let items = [];

        try {
          const parsedItems = JSON.parse(itemsJSON);
          items = parsedItems.map((item, index) => {
            const orderedQty = Number(item.quantity) || 0;
            const receivedQty = Number(item.receivedQty) || 0;
            const remainingQty = orderedQty - receivedQty;

            return {
              index: index,
              name: item.name,
              category: item.category,
              orderedQty: orderedQty,
              receivedQty: receivedQty,
              remainingQty: remainingQty,
              unit: item.unit,
              unitPrice: Number(item.unitPrice) || 0
            };
          });
        } catch(e) {
          Logger.log(`JSON 파싱 오류 (PO: ${poId}): ${e.message}`);
          items = [];
        }

        return {
          poId: data[i][0],
          supplierId: data[i][2],
          status: data[i][4],
          items: items
        };
      }
    }

    throw new Error(`PO '${poId}'를 찾을 수 없습니다.`);
  } catch (e) {
    Logger.log(`getPOItemsForReceiving 오류: ${e.message}`);
    throw new Error(`PO 아이템 목록 로드 실패: ${e.message}`);
  }
}

/**
 * [구현 완료] '신규 아이템 추가' 폼 데이터를 받아 시트에 추가합니다.
 * @param {Object} formData { code, itemName, category, notes, suppliersJSON }
 */
function addItem(formData) {
  try {
    // 필수 필드 검증
    if (!formData.itemName || formData.itemName.trim() === "") {
      throw new Error("필수 항목이 누락되었습니다: 아이템명");
    }

    const sheet = getSheet(SHEET_NAMES.DB);
    const newId = getNextId(sheet, "IT-");

    // ItemID(A), Code(B), ItemName(C), Category(D), Notes(E), SuppliersJSON(F)
    const newRow = [
      newId,
      formData.code || "",
      formData.itemName,
      formData.category,
      formData.notes || null,
      formData.suppliersJSON || "[]"
    ];
    sheet.appendRow(newRow);
    return `아이템 '${formData.itemName}'이(가) 등록되었습니다.`;
  } catch (e) {
    Logger.log(`addItem 오류: ${e.message}`);
    throw new Error(`아이템 등록 실패: ${e.message}`);
  }
}

/**
 * [신규] '아이템 수정'을 위해 특정 ID의 아이템 데이터를 가져옵니다.
 * @param {string} id 아이템 ID
 * @returns {Object} 아이템 데이터
 */
function getItemDetails(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.DB);
    const data = sheet.getDataRange().getValues();

    // 헤더(1행) 건너뛰고 ID(A열, index 0) 검색
    // ItemID(A), Code(B), ItemName(C), Category(D), Notes(E), SuppliersJSON(F)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        return {
          id: data[i][0],
          code: data[i][1] || "",
          name: data[i][2],
          category: data[i][3],
          notes: data[i][4],
          suppliersJSON: data[i][5] // 이 JSON 문자열을 클라이언트에서 파싱
        };
      }
    }
    throw new Error("해당 ID의 아이템을 찾을 수 없습니다.");
  } catch (e) {
    Logger.log(e);
    throw new Error(`아이템 데이터 조회 실패: ${e.message}`);
  }
}

/**
 * [신규] '아이템 수정' 폼에서 데이터를 받아 시트의 기존 행을 업데이트합니다.
 * @param {Object} formData 폼 데이터 (id, code, itemName, category, notes, suppliersJSON 포함)
 * @returns {String} 성공 메시지
 */
function updateItem(formData) {
  try {
    const sheet = getSheet(SHEET_NAMES.DB);
    const data = sheet.getDataRange().getValues();

    // 헤더(1행) 건너뛰고 ID(A열, index 0) 검색
    // ItemID(A), Code(B), ItemName(C), Category(D), Notes(E), SuppliersJSON(F)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == formData.id) {
        // 행을 찾았으면 해당 행(i+1)의 데이터 업데이트
        const range = sheet.getRange(i + 1, 1, 1, 6);
        range.setValues([[
          formData.id,
          formData.code || "",
          formData.itemName,
          formData.category,
          formData.notes || null,
          formData.suppliersJSON || "[]"
        ]]);
        return `아이템 '${formData.itemName}' 정보가 업데이트되었습니다.`;
      }
    }
    throw new Error("업데이트할 아이템을 찾지 못했습니다.");
  } catch (e) {
    Logger.log(e);
    throw new Error(`아이템 업데이트 실패: ${e.message}`);
  }
}

/**
 * [연동됨] 파일 업로드 처리
 * 클라이언트의 fileBlob을 받아 Drive에 저장하고 URL을 반환합니다.
 * (JavaScript.html에서 Base64 인코딩된 데이터가 넘어와야 함)
 */
function saveFileToDrive(fileData, fileName) {
  if (DRIVE_FOLDER_ID === "YOUR_FOLDER_ID_HERE" || !DRIVE_FOLDER_ID || DRIVE_FOLDER_ID.length < 20) {
    throw new Error("Code.gs 파일 상단의 DRIVE_FOLDER_ID를 실제 Google Drive 폴더 ID로 설정해야 합니다.");
  }
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // Base64 데이터를 디코딩하여 Blob 생성
    const contentType = fileData.split(";")[0].replace("data:", "");
    const base64Data = fileData.split(",")[1];
    const decodedData = Utilities.base64Decode(base64Data);
    const fileBlob = Utilities.newBlob(decodedData, contentType, fileName);

    // 동일한 이름의 파일이 있으면 덮어쓰기 대신 새 이름 사용
    const uniqueFileName = `${new Date().toISOString().replace(/:/g, "-")}_${fileName}`;

    const file = folder.createFile(fileBlob);
    file.setName(uniqueFileName);

    // 파일을 '공개' (링크가 있는 모든 사용자가 보기 가능)로 설정
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl(); // 파일 URL 반환
  } catch (e) {
    Logger.log(e);
    throw new Error(`파일 저장 실패: ${e.message}`);
  }
}

/**
 * Drive에 업로드된 파일의 URL과 메타데이터를 'Statements' 시트에 저장합니다.
 * @param {Object} formData { issueDate, supplierId, amount, poIds, fileUrl }
 * @returns {String} 성공 메시지
 */
function addStatementRecord(formData) {
  try {
    const sheet = getSheet(SHEET_NAMES.STATEMENTS);
    const newId = getNextId(sheet, "ST-");

    const newRow = [
      newId,
      new Date(formData.issueDate),
      formData.supplierId,
      formData.poIds || null,
      Number(formData.amount),
      formData.fileUrl,
      new Date() // Timestamp
    ];

    sheet.appendRow(newRow);
    return `거래명세서(ID: ${newId})가 성공적으로 저장되었습니다.`;
  } catch (e) {
    Logger.log(e);
    throw new Error(`거래명세서 정보 저장 실패: ${e.message}`);
  }
}

/**
 * Drive에 업로드된 파일의 URL과 메타데이터를 'TaxInvoices' 시트에 저장합니다.
 * @param {Object} formData { issueDate, month, supplierId, totalAmount, fileUrl }
 * @returns {String} 성공 메시지
 */
function addTaxInvoiceRecord(formData) {
  try {
    const sheet = getSheet(SHEET_NAMES.TAX);
    const newId = getNextId(sheet, "TAX-");

    const newRow = [
      newId,
      new Date(formData.issueDate),
      Number(formData.month),
      formData.supplierId,
      Number(formData.totalAmount),
      formData.fileUrl,
      new Date() // Timestamp
    ];

    sheet.appendRow(newRow);
    return `세금계산서(ID: ${newId})가 성공적으로 저장되었습니다.`;
  } catch (e) {
    Logger.log(e);
    throw new Error(`세금계산서 정보 저장 실패: ${e.message}`);
  }
}

// --- 테스트 함수 (개발 및 디버깅용) ---

/**
 * PO ID 생성 테스트 함수
 * Apps Script 편집기에서 실행: 상단 메뉴 > 실행 > testPOIdGeneration
 *
 * 테스트 시나리오:
 * 1. 오늘 날짜로 PO ID가 생성되는지 확인
 * 2. 여러 번 실행 시 순차적으로 증가하는지 확인
 * 3. 로그에서 생성된 ID 확인
 */
function testPOIdGeneration() {
  try {
    Logger.log("========== PO ID 생성 테스트 시작 ==========");

    const sheet = getSheet(SHEET_NAMES.PO);

    // 테스트 1: 첫 번째 ID 생성
    const firstId = getNextPOId(sheet);
    Logger.log(`✅ 생성된 첫 번째 PO ID: ${firstId}`);

    // 현재 날짜 확인
    const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyyMMdd");
    Logger.log(`📅 오늘 날짜 (YYYYMMDD): ${today}`);

    // 예상 형식 확인
    const expectedPrefix = `PO-${today}`;
    if (firstId.startsWith(expectedPrefix)) {
      Logger.log(`✅ PO ID가 올바른 날짜 형식을 포함하고 있습니다: ${expectedPrefix}`);
    } else {
      Logger.log(`❌ 오류: PO ID가 예상 형식과 다릅니다. 예상: ${expectedPrefix}###`);
    }

    // 시퀀스 번호 추출
    const sequence = firstId.substring(expectedPrefix.length);
    Logger.log(`🔢 추출된 시퀀스 번호: ${sequence}`);

    if (sequence.length === 3 && !isNaN(parseInt(sequence, 10))) {
      Logger.log("✅ 시퀀스 번호가 3자리 숫자 형식입니다");
    } else {
      Logger.log("❌ 오류: 시퀀스 번호 형식이 올바르지 않습니다");
    }

    Logger.log("========== PO ID 생성 테스트 완료 ==========");
    Logger.log("");
    Logger.log("💡 참고:");
    Logger.log("- 실제 PO를 생성하면 시퀀스 번호가 자동으로 증가합니다");
    Logger.log("- 다음 날이 되면 자동으로 새로운 날짜로 001부터 시작합니다");
    Logger.log(`- 오늘 생성되는 PO 번호: ${expectedPrefix}001, ${expectedPrefix}002, ...`);

    return firstId;

  } catch (e) {
    Logger.log(`❌ 테스트 실패: ${e.message}`);
    Logger.log(e.stack);
    throw e;
  }
}

/**
 * PO ID 생성 시뮬레이션 함수 (여러 개 생성 테스트)
 * 실제로 데이터를 추가하지 않고 어떤 ID가 생성될지 미리 확인
 */
function simulatePOIdGeneration() {
  try {
    Logger.log("========== PO ID 생성 시뮬레이션 ==========");

    const sheet = getSheet(SHEET_NAMES.PO);
    const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyyMMdd");

    Logger.log(`📅 기준 날짜: ${today}`);
    Logger.log("");
    Logger.log("오늘 생성 가능한 PO ID 예시:");

    // 다음 5개의 PO ID를 시뮬레이션
    for (let i = 0; i < 5; i++) {
      const nextId = getNextPOId(sheet);
      Logger.log(`  ${i + 1}. ${nextId}`);
    }

    Logger.log("");
    Logger.log("💡 참고: 위 ID들은 시뮬레이션이며 실제로 생성되지 않았습니다");
    Logger.log("실제 PO 생성 시 현재 시트의 마지막 시퀀스 번호를 기준으로 생성됩니다");

  } catch (e) {
    Logger.log(`❌ 시뮬레이션 실패: ${e.message}`);
    throw e;
  }
}

/**
 * 현재 PO_List 시트의 PO ID 현황 조회
 */
function checkCurrentPOIds() {
  try {
    Logger.log("========== 현재 PO ID 현황 ==========");

    const sheet = getSheet(SHEET_NAMES.PO);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log("📝 현재 등록된 PO가 없습니다.");
      Logger.log(`다음 생성될 PO ID: ${getNextPOId(sheet)}`);
      return;
    }

    const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyyMMdd");

    Logger.log(`📅 오늘 날짜: ${today}`);
    Logger.log("");

    // 오늘 생성된 PO 찾기
    const todayPOs = [];
    idColumn.forEach(row => {
      const poId = row[0];
      if (poId && typeof poId === "string" && poId.includes(today)) {
        todayPOs.push(poId);
      }
    });

    if (todayPOs.length > 0) {
      Logger.log(`✅ 오늘 생성된 PO: ${todayPOs.length}개`);
      todayPOs.forEach(id => Logger.log(`  - ${id}`));
    } else {
      Logger.log("📝 오늘 생성된 PO가 없습니다");
    }

    Logger.log("");
    Logger.log(`🔮 다음 생성될 PO ID: ${getNextPOId(sheet)}`);

  } catch (e) {
    Logger.log(`❌ 조회 실패: ${e.message}`);
    throw e;
  }
}

/**
 * PO 정보를 바탕으로 Google Docs 문서를 생성합니다.
 * @param {string} poId PO ID
 * @returns {string} 생성된 문서의 URL
 */
function createPODoc(poId) {
  try {
    // 1. PO 정보 조회
    const poData = getPODetails(poId);
    if (!poData) throw new Error(`PO 정보를 찾을 수 없습니다: ${poId}`);

    // 2. 공급사 정보 조회
    let supplierData = {};
    try {
      supplierData = getSupplierDetails(poData.supplierId);
    } catch (_e) {
      supplierData = { name: poData.supplierId, contactPerson: "", email: "", phone: "", address: "" };
    }

    // 3. Google Doc 생성
    const docName = `PO_${poId}_${supplierData.name}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // 4. 문서 내용 작성
    // 헤더
    body.appendParagraph("PURCHASE ORDER").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("\n");

    // 기본 정보 테이블
    const infoTable = [
      ["PO Number", poId],
      ["Date", Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd")],
      ["Supplier", supplierData.name],
      ["Contact", supplierData.contactPerson || "-"],
      ["Phone", supplierData.phone || "-"],
      ["Address", supplierData.address || "-"]
    ];

    const infoTableElem = body.appendTable(infoTable);
    infoTableElem.setBorderWidth(0);

    body.appendParagraph("\n");
    body.appendParagraph("Items:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // 아이템 테이블
    const itemHeader = ["Item Name", "Category", "Qty", "Unit", "Unit Price", "Total"];
    const itemRows = poData.items.map(item => [
      item.name,
      item.category,
      item.quantity.toString(),
      item.unit,
      Number(item.unitPrice).toLocaleString("en-US") + " VND",
      ((Number(item.quantity) || 0) * (Number(item.unitPrice) || 0)).toLocaleString("en-US") + " VND"
    ]);

    const itemTableData = [itemHeader, ...itemRows];
    const itemTable = body.appendTable(itemTableData);

    // 아이템 테이블 스타일
    itemTable.getRow(0).setBold(true).setBackgroundColor("#EFEFEF");

    // 총계 계산
    const totalAmount = poData.items.reduce((sum, item) => {
      return sum + ((Number(item.quantity) || 0) * (Number(item.unitPrice) || 0));
    }, 0);

    body.appendParagraph("\n");
    const totalPara = body.appendParagraph(`Total Amount: ${totalAmount.toLocaleString("en-US")} VND`);
    totalPara.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    totalPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    body.appendParagraph("\n\n");
    body.appendParagraph("Authorized Signature: __________________________");

    // 5. 저장 및 URL 반환
    doc.saveAndClose();

    // 지정된 폴더로 이동
    const file = DriveApp.getFileById(doc.getId());
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    file.moveTo(folder);

    return file.getUrl();

  } catch (e) {
    Logger.log(`createPODoc 오류: ${e.message}`);
    throw new Error(`문서 생성 실패: ${e.message}`);
  }
}
