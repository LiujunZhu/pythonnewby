// =============================================================================
// 00_Utilities.gs — Shared constants and helper functions
// =============================================================================

// ---------------------------------------------------------------------------
// Sheet name constants
// ---------------------------------------------------------------------------
const SHEETS = {
  DASHBOARD:    'Dashboard',
  COA:          'Chart of Accounts',
  JOURNAL:      'Journal',
  GL:           'General Ledger',
  CUSTOMERS:    'Customers',
  VENDORS:      'Vendors',
  DEPARTMENTS:  'Departments',
  AR_INVOICES:  'AR - Invoices',
  AR_PAYMENTS:  'AR - Payments',
  AP_BILLS:     'AP - Bills',
  AP_PAYMENTS:  'AP - Payments',
  BANK_STMT:    'Bank Statement',
  RECON:        'Reconciliation',
};

// ---------------------------------------------------------------------------
// Colour palette
// ---------------------------------------------------------------------------
const COLORS = {
  HEADER_BG:   '#1565C0',
  HEADER_TEXT: '#FFFFFF',
  ALT_ROW:     '#F5F5F5',
  POSITIVE:    '#E8F5E9',
  NEGATIVE:    '#FFEBEE',
  NEUTRAL:     '#FFF9C4',
  BORDER:      '#BDBDBD',
};

// ---------------------------------------------------------------------------
// Header helper
// ---------------------------------------------------------------------------
function setHeaders(sheet, headers, bgColor) {
  const headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setBackground(bgColor || COLORS.HEADER_BG);
  headerRow.setFontColor(COLORS.HEADER_TEXT);
  headerRow.setFontWeight('bold');
  headerRow.setHorizontalAlignment('center');
  headerRow.setBorder(true, true, true, true, true, true, COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

// ---------------------------------------------------------------------------
// Format columns as currency
// ---------------------------------------------------------------------------
function formatCurrencyColumns(sheet, colNumbers, startRow, numRows) {
  colNumbers.forEach(col => {
    sheet.getRange(startRow, col, numRows, 1).setNumberFormat('$#,##0.00');
  });
}

// ---------------------------------------------------------------------------
// Alternate row shading
// ---------------------------------------------------------------------------
function applyAlternateRowColor(sheet, startRow, endRow, numCols) {
  for (let r = startRow; r <= endRow; r++) {
    const bg = (r % 2 === 0) ? COLORS.ALT_ROW : '#FFFFFF';
    sheet.getRange(r, 1, 1, numCols).setBackground(bg);
  }
}

// ---------------------------------------------------------------------------
// Next sequential ID (e.g. JE-001 → JE-002)
// ---------------------------------------------------------------------------
function getNextId(sheet, col, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '001';

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues()
    .flat()
    .filter(v => String(v).startsWith(prefix));

  if (values.length === 0) return prefix + '001';

  const nums = values.map(v => parseInt(String(v).replace(prefix, ''), 10)).filter(n => !isNaN(n));
  const max  = Math.max(...nums);
  return prefix + String(max + 1).padStart(3, '0');
}

// ---------------------------------------------------------------------------
// Safe get sheet — throws a friendly error if missing
// ---------------------------------------------------------------------------
function getSheet(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" not found. Please run Setup first.`);
  return sh;
}

// ---------------------------------------------------------------------------
// Read all rows from a sheet as objects keyed by header
// ---------------------------------------------------------------------------
function getSheetData(sheetName) {
  const sh = getSheet(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows    = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ---------------------------------------------------------------------------
// Format a JS Date as yyyy-MM-dd
// ---------------------------------------------------------------------------
function fmtDate(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// ---------------------------------------------------------------------------
// Parse a date string or Date object safely
// ---------------------------------------------------------------------------
function parseDate(val) {
  if (val instanceof Date) return val;
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ---------------------------------------------------------------------------
// Number → formatted currency string
// ---------------------------------------------------------------------------
function fmtCurrency(num) {
  return '$' + (Number(num) || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ---------------------------------------------------------------------------
// Show a toast notification
// ---------------------------------------------------------------------------
function toast(msg, title, timeout) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title || 'Accounting', timeout || 5);
}

// ---------------------------------------------------------------------------
// Validate that a row's required cells are filled
// ---------------------------------------------------------------------------
function validateRequired(values, fieldNames) {
  const missing = [];
  fieldNames.forEach((name, i) => {
    if (values[i] === '' || values[i] === null || values[i] === undefined) {
      missing.push(name);
    }
  });
  return missing;
}

// ---------------------------------------------------------------------------
// Update status on AR/AP sheets based on due date and balance
// ---------------------------------------------------------------------------
function refreshInvoiceStatuses(sheetName) {
  const sh = getSheet(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Columns: Due Date = 3, Balance Due = 12, Status = 13
  const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
  const updates = [];

  data.forEach((row, i) => {
    const balance = Number(row[11]) || 0;
    const dueDate = parseDate(row[2]);
    let status    = String(row[12]);

    if (status === 'Draft') {
      updates.push([status]);
      return;
    }

    if (balance <= 0) {
      updates.push(['Paid']);
    } else if (dueDate && dueDate < today) {
      updates.push(['Overdue']);
    } else if (balance > 0) {
      const total = Number(row[9]) || 0;
      updates.push([balance < total ? 'Partial' : 'Sent']);
    } else {
      updates.push([status]);
    }
  });

  sh.getRange(2, 13, updates.length, 1).setValues(updates);
}

// ---------------------------------------------------------------------------
// Colour-code status cells (Status column 13 = col M)
// ---------------------------------------------------------------------------
function colorStatusColumn(sheet, lastRow) {
  if (lastRow < 2) return;
  for (let r = 2; r <= lastRow; r++) {
    const cell = sheet.getRange(r, 13);
    switch (cell.getValue()) {
      case 'Paid':     cell.setBackground('#C8E6C9'); break;
      case 'Overdue':  cell.setBackground('#FFCDD2'); break;
      case 'Partial':  cell.setBackground('#FFF9C4'); break;
      case 'Sent':     cell.setBackground('#BBDEFB'); break;
      case 'Draft':    cell.setBackground('#F5F5F5'); break;
      default:         cell.setBackground('#FFFFFF');
    }
  }
}
