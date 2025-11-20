# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Purchase Management web application built with Google Apps Script. It's a Single Page Application (SPA) using Google Sheets for data storage and Google Drive for file uploads.

**Technology Stack:**
- Backend: Google Apps Script (JavaScript)
- Frontend: Vanilla JavaScript with server-side HTML templates
- Database: Google Sheets (entities separated across multiple sheets)
- File Storage: Google Drive

**Core Files:**
- `Code.gs` - Server-side logic (Google Apps Script)
- `Index.html` - Main HTML template
- `JavaScript.html` - Client-side JavaScript (included in Index.html)
- `Stylesheet.html` - CSS styles (included in Index.html)

## Development Workflow

**Note:** Google Apps Script cannot run locally. The development workflow is:

1. Edit files locally
2. Deploy to Google Apps Script:
   - Open [Google Apps Script](https://script.google.com)
   - Copy/paste code into the editor, or use `clasp` CLI
   - Deploy as Web App: **Deploy > New deployment > Web app**
3. Test in browser using the deployment URL
4. Check logs: **View > Logs** or use `Logger.log()`

**Using clasp CLI (optional):**
```bash
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID>
clasp push   # Push local changes to Apps Script
clasp open   # Open in browser
```

**Important Configuration (Code.gs):**
Before deployment, update these constants at the top of `Code.gs`:
```javascript
const SPREADSHEET_ID = "..."; // Your Google Sheets ID
const DRIVE_FOLDER_ID = "..."; // Your Google Drive folder ID for uploads
```

## Data Architecture

The application uses **6 sheets** in a single spreadsheet:

1. **PurchaseHistory** - Purchase records
   - Columns: PurchaseID, PurchaseDate, ItemName, Category, Quantity, Unit, UnitPrice, TotalPrice, SupplierID, PO_ID, StatementID, Timestamp

2. **PO_List** - Purchase Orders
   - Columns: PO_ID, IssueDate, SupplierID, TotalAmount, Status, ItemsJSON (array of items), Notes

3. **Statements** - Transaction statements (with file upload)
   - Columns: StatementID, IssueDate, SupplierID, PO_IDs, Amount, FileLink, Timestamp

4. **Suppliers** - Supplier master data
   - Columns: SupplierID, SupplierName, ContactPerson, Email, Phone, Address

5. **TaxInvoices** - Tax invoices (with file upload)
   - Columns: TaxInvoiceID, IssueDate, Month, SupplierID, TotalAmount, FileLink, Timestamp

6. **ItemDatabase** - Item master data
   - Columns: ItemID, ItemName, Category, Notes, SuppliersJSON (array of supplier prices)

### ID Generation Patterns

**Sequential IDs (most entities):**
- Purchase History: `PH-1001`, `PH-1002`, ...
- Suppliers: `S-001`, `S-002`, ...
- Items: `IT-1001`, `IT-1002`, ...
- Statements: `ST-1001`, `ST-1002`, ...
- Tax Invoices: `TAX-1001`, `TAX-1002`, ...

These are handled by `getNextId(sheet, prefix)` function.

**Date-based IDs (Purchase Orders only):**
- Format: `PO-YYYYMMDD###`
- Examples: `PO-20250104001`, `PO-20250104002`, `PO-20250105001`
- **Rules:**
  1. Resets to 001 each day
  2. Sequential within the same day (001, 002, 003, ...)
  3. Timezone: Vietnam time (Asia/Ho_Chi_Minh, GMT+7)

Handled by dedicated `getNextPOId(sheet)` function.

## Application Architecture & Flow

### Client-Server Communication

The app uses Google Apps Script's `google.script.run` API:

**Client → Server:**
```javascript
google.script.run
    .withSuccessHandler(callback)
    .withFailureHandler(errorCallback)
    .serverFunction(data);
```

**Server → Client:**
Server functions return data or HTML strings that the client renders.

### Page Loading Pattern

1. User clicks navigation link
2. Client calls `google.script.run.getPage(pageTitle)`
3. Server's `getPage()` routes to appropriate page function (e.g., `getDashboardPage()`)
4. Server generates HTML string with data
5. Client renders HTML into `#content-wrapper`
6. Client binds event handlers using event delegation

### File Upload Flow

File uploads (statements, tax invoices) use a 2-step process:

1. **Client**: Read file as Base64 using FileReader
2. **Upload to Drive**: Call `saveFileToDrive(fileData, fileName)` → returns URL
3. **Save metadata**: Call `addStatementRecord()` or `addTaxInvoiceRecord()` with URL and metadata

## Key Functions Reference

### Server-Side (Code.gs)

**Utility Functions:**
- `getSheet(sheetName)` - Get sheet by name with error handling
- `getSupplierMap()` - Returns object mapping supplier IDs to names
- `getNextId(sheet, prefix)` - Generate next sequential ID (for general entities)
- `getNextPOId(sheet)` - Generate date-based PO ID (PO-only)
- `escapeHtml(text)` - HTML escape for XSS protection
- `formatDateKR(date)` - Format date in Vietnam timezone (Korean format)
- `formatDateInTimeZone(date, timezone)` - Format date in specified timezone

**CRUD Operations:**
- `addPurchaseEntry(formData)` - Add purchase record
- `addSupplier(formData)` / `updateSupplier(formData)` - Supplier CRUD
- `addPO(formData)` - Create purchase order
- `addItem(formData)` / `updateItem(formData)` - Item DB CRUD
- `deleteRowById(sheetName, id)` - Generic delete function

**Page Generation Functions:**
- `getDashboardPage()` - Monthly summary and analysis
- `getPurchaseHistoryPage()` - Purchase history table
- `getPurchaseRegisterPage()` - Purchase registration form
- `getPOPage()` - Purchase order management
- `getStatementPage()` / `getTaxInvoicePage()` - Document management (with upload)
- `getSupplierPage()` - Supplier management
- `getDatabasePage()` - Item master database
- `getAnalysisPage()` - Period-based analysis

**Analysis Functions:**
- `getDashboardData()` - Monthly summary data
- `getAnalysisData(options)` - Purchase amount analysis by category/supplier (with date filter)

### Client-Side (JavaScript.html)

**Core Functions:**
- `loadPage(pageTitle, linkElement)` - Page navigation
- `showModal(title, contentHtml)` / `closeModal()` - Modal management
- `showSpinner()` / `hideSpinner()` - Loading indicator
- `showConfirm(message, onOk)` - Confirmation dialog

**Form Handlers:**
- `handlePurchaseFormSubmit()` - Purchase registration
- `handleSupplierFormSubmit()` - Supplier add/edit
- `handlePOFormSubmit()` - PO creation
- `handleItemDBFormSubmit()` - Item add/edit
- `handleFileUpload(form, type)` - Statement/Tax invoice file upload

## Common Development Tasks

### Adding a New Field to a Form

1. Update HTML in page generation function (e.g., `getPurchaseRegisterPage()` in Code.gs)
2. Update form handler in JavaScript.html to collect the new field
3. Update server-side function to save the new field to the sheet

### Adding a New Page

1. Add navigation link in Index.html sidebar
2. Add case to `getPage()` switch statement in Code.gs
3. Create new `getXxxPage()` function in Code.gs
4. Add necessary CRUD functions

### Modifying Sheet Structure

When adding/removing columns from a sheet, update:
- Page generation functions that read from that sheet
- CRUD functions that write to that sheet
- Array indices in `getRange()` calls

## Debugging

**Server-side:**
- Use `Logger.log()` for debugging
- Check logs in Apps Script editor: **View > Executions**

**Client-side:**
- Use browser DevTools console (F12)
- Server errors are displayed via `showAlert(error.message, true)`

## Coding Standards

**Naming Conventions:**
- Functions: camelCase (e.g., `getPurchaseHistoryPage`)
- Constants: UPPER_SNAKE_CASE (e.g., `SPREADSHEET_ID`)
- Variables: camelCase

**ID Prefixes:**
- `PH-` Purchase History
- `PO-` Purchase Order
- `ST-` Statement
- `TAX-` Tax Invoice
- `S-` Supplier
- `IT-` Item

**Security & Validation:**
- Always validate required fields
- Convert numbers using `Number()` and check with `isNaN()`
- Convert dates to `Date` objects
- **Always use `escapeHtml()` when outputting to HTML** (XSS protection)
- Wrap operations in try-catch blocks

**File Upload:**
- Max file size: 10MB
- Allowed extensions: pdf, jpg, jpeg, png, gif, doc, docx, xls, xlsx

## Known Limitations

1. **No authentication**: Relies on Google Apps Script's built-in authentication
2. **Currency**: Hardcoded to VND (Vietnamese Dong)
3. **No data validation**: Sheet-level data validation must be set manually
4. **Single user**: No concurrent editing or user permissions within the app
5. **File size limit**: 10MB for Drive uploads (Apps Script limit)

## Deployment Checklist

1. Update `SPREADSHEET_ID` in Code.gs
2. Update `DRIVE_FOLDER_ID` in Code.gs
3. Ensure spreadsheet has all 6 sheets with correct names
4. Deploy as Web App: **Deploy > New deployment**
5. Set appropriate access permissions
6. Test all CRUD operations
7. Verify file upload works correctly
