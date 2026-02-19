// =============================================================================
// 01_Setup.gs — Initial spreadsheet setup
// =============================================================================

/**
 * Creates all required sheets with headers and formatting.
 * Run this once when setting up the workbook for the first time.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'Setup Accounting Workbook',
    'This will create all required sheets and sample data. Existing sheets with matching names will be skipped. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  setupChartOfAccountsSheet(ss);
  setupJournalSheet(ss);
  setupGeneralLedgerSheet(ss);
  setupCustomersSheet(ss);
  setupVendorsSheet(ss);
  setupDepartmentsSheet(ss);
  setupARInvoicesSheet(ss);
  setupARPaymentsSheet(ss);
  setupAPBillsSheet(ss);
  setupAPPaymentsSheet(ss);
  setupBankStatementSheet(ss);
  setupReconciliationSheet(ss);
  setupDashboardSheet(ss);

  insertSampleData(ss);

  ui.alert('Setup complete! Your accounting workbook is ready to use.');
}

// ---------------------------------------------------------------------------
// Sheet creators
// ---------------------------------------------------------------------------

function setupChartOfAccountsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.COA);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.COA);

  const headers = ['Account #', 'Account Name', 'Type', 'Sub-Type', 'Normal Balance', 'Department', 'Description', 'Active'];
  setHeaders(sh, headers, '#1565C0');

  // Sample COA
  const coa = [
    // Assets
    ['1000', 'Cash - Checking',         'Asset',     'Current Asset',    'Debit',  '',  'Primary checking account',  'Yes'],
    ['1010', 'Cash - Savings',           'Asset',     'Current Asset',    'Debit',  '',  'Savings account',           'Yes'],
    ['1100', 'Accounts Receivable',      'Asset',     'Current Asset',    'Debit',  '',  'Amounts owed by customers', 'Yes'],
    ['1200', 'Inventory',                'Asset',     'Current Asset',    'Debit',  '',  'Goods held for sale',       'Yes'],
    ['1300', 'Prepaid Expenses',         'Asset',     'Current Asset',    'Debit',  '',  'Prepaid items',             'Yes'],
    ['1500', 'Equipment',                'Asset',     'Fixed Asset',      'Debit',  '',  'Office equipment',          'Yes'],
    ['1510', 'Accum. Depr. - Equipment', 'Asset',     'Fixed Asset',      'Credit', '',  'Accumulated depreciation',  'Yes'],
    // Liabilities
    ['2000', 'Accounts Payable',         'Liability', 'Current Liability','Credit', '',  'Amounts owed to vendors',   'Yes'],
    ['2100', 'Accrued Liabilities',      'Liability', 'Current Liability','Credit', '',  'Accrued expenses',          'Yes'],
    ['2200', 'Sales Tax Payable',        'Liability', 'Current Liability','Credit', '',  'Sales tax collected',       'Yes'],
    ['2500', 'Long-Term Debt',           'Liability', 'Long-Term',        'Credit', '',  'Notes payable',             'Yes'],
    // Equity
    ['3000', 'Owner\'s Capital',         'Equity',    'Equity',           'Credit', '',  'Owner contributions',       'Yes'],
    ['3100', 'Retained Earnings',        'Equity',    'Equity',           'Credit', '',  'Accumulated earnings',      'Yes'],
    ['3200', 'Owner\'s Draw',            'Equity',    'Equity',           'Debit',  '',  'Owner withdrawals',         'Yes'],
    // Revenue
    ['4000', 'Sales Revenue',            'Revenue',   'Operating',        'Credit', '',  'Product/service sales',     'Yes'],
    ['4100', 'Service Revenue',          'Revenue',   'Operating',        'Credit', '',  'Service income',            'Yes'],
    ['4900', 'Other Income',             'Revenue',   'Other',            'Credit', '',  'Miscellaneous income',      'Yes'],
    // Expenses
    ['5000', 'Cost of Goods Sold',       'Expense',   'COGS',             'Debit',  '',  'Direct costs',              'Yes'],
    ['5100', 'Salaries & Wages',         'Expense',   'Operating',        'Debit',  '',  'Employee compensation',     'Yes'],
    ['5200', 'Rent Expense',             'Expense',   'Operating',        'Debit',  '',  'Office/store rent',         'Yes'],
    ['5300', 'Utilities Expense',        'Expense',   'Operating',        'Debit',  '',  'Electric, water, gas',      'Yes'],
    ['5400', 'Office Supplies',          'Expense',   'Operating',        'Debit',  '',  'Office consumables',        'Yes'],
    ['5500', 'Marketing & Advertising',  'Expense',   'Operating',        'Debit',  '',  'Promotion costs',           'Yes'],
    ['5600', 'Insurance Expense',        'Expense',   'Operating',        'Debit',  '',  'Business insurance',        'Yes'],
    ['5700', 'Depreciation Expense',     'Expense',   'Operating',        'Debit',  '',  'Asset depreciation',        'Yes'],
    ['5800', 'Interest Expense',         'Expense',   'Non-Operating',    'Debit',  '',  'Loan interest',             'Yes'],
    ['5900', 'Other Expenses',           'Expense',   'Other',            'Debit',  '',  'Miscellaneous expenses',    'Yes'],
  ];
  sh.getRange(2, 1, coa.length, headers.length).setValues(coa);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function setupJournalSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.JOURNAL);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.JOURNAL);

  const headers = ['Entry ID', 'Date', 'Reference', 'Description', 'Account #', 'Account Name', 'Debit', 'Credit', 'Department', 'Posted', 'Notes'];
  setHeaders(sh, headers, '#2E7D32');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [7, 8], 2, 1000);
}

function setupGeneralLedgerSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.GL);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.GL);

  const headers = ['Account #', 'Account Name', 'Type', 'Opening Balance', 'Total Debits', 'Total Credits', 'Closing Balance'];
  setHeaders(sh, headers, '#4A148C');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [4, 5, 6, 7], 2, 500);
}

function setupCustomersSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.CUSTOMERS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.CUSTOMERS);

  const headers = ['Customer ID', 'Name', 'Email', 'Phone', 'Address', 'City', 'State', 'ZIP', 'Payment Terms', 'Notes', 'Active'];
  setHeaders(sh, headers, '#00695C');

  const sample = [
    ['CUST-001', 'Acme Corp',      'billing@acme.com',   '555-0100', '123 Main St', 'New York',   'NY', '10001', 'Net 30', '', 'Yes'],
    ['CUST-002', 'Globex Inc',     'ap@globex.com',      '555-0200', '456 Oak Ave',  'Los Angeles','CA', '90001', 'Net 15', '', 'Yes'],
    ['CUST-003', 'Initech LLC',    'finance@initech.com','555-0300', '789 Elm Rd',   'Chicago',    'IL', '60601', 'Net 30', '', 'Yes'],
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function setupVendorsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.VENDORS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.VENDORS);

  const headers = ['Vendor ID', 'Name', 'Email', 'Phone', 'Address', 'City', 'State', 'ZIP', 'Payment Terms', 'Expense Account', 'Notes', 'Active'];
  setHeaders(sh, headers, '#BF360C');

  const sample = [
    ['VEND-001', 'Office Depot',      'invoices@officedepot.com', '555-0400', '100 Supply Ln', 'Dallas',    'TX', '75001', 'Net 30', '5400', '', 'Yes'],
    ['VEND-002', 'City Electric Co',  'billing@cityelectric.com', '555-0500', '200 Power Ave', 'Houston',   'TX', '77001', 'Net 15', '5300', '', 'Yes'],
    ['VEND-003', 'First National Bank','loans@fnb.com',           '555-0600', '300 Bank Blvd', 'Austin',    'TX', '78701', 'Net 30', '5800', '', 'Yes'],
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function setupDepartmentsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.DEPARTMENTS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.DEPARTMENTS);

  const headers = ['Dept Code', 'Department Name', 'Manager', 'Cost Center', 'Notes', 'Active'];
  setHeaders(sh, headers, '#37474F');

  const sample = [
    ['ADMIN', 'Administration', 'Jane Smith',  'CC-100', '', 'Yes'],
    ['SALES', 'Sales',          'John Doe',    'CC-200', '', 'Yes'],
    ['OPS',   'Operations',     'Bob Johnson', 'CC-300', '', 'Yes'],
    ['IT',    'Information Technology', 'Alice Wong', 'CC-400', '', 'Yes'],
  ];
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function setupARInvoicesSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.AR_INVOICES);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.AR_INVOICES);

  const headers = [
    'Invoice #', 'Date', 'Due Date', 'Customer ID', 'Customer Name',
    'Description', 'Subtotal', 'Tax Rate %', 'Tax Amount', 'Total',
    'Amount Paid', 'Balance Due', 'Status', 'Revenue Account', 'Department', 'Notes'
  ];
  setHeaders(sh, headers, '#1A237E');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [7, 9, 10, 11, 12], 2, 1000);
}

function setupARPaymentsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.AR_PAYMENTS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.AR_PAYMENTS);

  const headers = ['Payment #', 'Date', 'Invoice #', 'Customer ID', 'Customer Name', 'Amount', 'Method', 'Reference', 'Notes'];
  setHeaders(sh, headers, '#283593');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [6], 2, 1000);
}

function setupAPBillsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.AP_BILLS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.AP_BILLS);

  const headers = [
    'Bill #', 'Date', 'Due Date', 'Vendor ID', 'Vendor Name',
    'Description', 'Subtotal', 'Tax Rate %', 'Tax Amount', 'Total',
    'Amount Paid', 'Balance Due', 'Status', 'Expense Account', 'Department', 'Notes'
  ];
  setHeaders(sh, headers, '#B71C1C');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [7, 9, 10, 11, 12], 2, 1000);
}

function setupAPPaymentsSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.AP_PAYMENTS);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.AP_PAYMENTS);

  const headers = ['Payment #', 'Date', 'Bill #', 'Vendor ID', 'Vendor Name', 'Amount', 'Method', 'Reference', 'Notes'];
  setHeaders(sh, headers, '#C62828');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [6], 2, 1000);
}

function setupBankStatementSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.BANK_STMT);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.BANK_STMT);

  const headers = ['Date', 'Description', 'Reference', 'Debit', 'Credit', 'Balance', 'Matched', 'Journal Entry ID'];
  setHeaders(sh, headers, '#004D40');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  formatCurrencyColumns(sh, [4, 5, 6], 2, 1000);
}

function setupReconciliationSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.RECON);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.RECON);

  sh.getRange('A1').setValue('BANK RECONCILIATION').setFontSize(16).setFontWeight('bold');
  sh.getRange('A3').setValue('Statement Date:');
  sh.getRange('A4').setValue('Bank Account:');
  sh.getRange('A5').setValue('GL Account #:');
  sh.getRange('A7').setValue('Bank Statement Ending Balance:');
  sh.getRange('A8').setValue('Add: Deposits in Transit:');
  sh.getRange('A9').setValue('Less: Outstanding Checks/Payments:');
  sh.getRange('A10').setValue('Adjusted Bank Balance:');
  sh.getRange('A12').setValue('GL Book Balance:');
  sh.getRange('A13').setValue('Add: Unrecorded Deposits:');
  sh.getRange('A14').setValue('Less: Bank Charges / Errors:');
  sh.getRange('A15').setValue('Adjusted Book Balance:');
  sh.getRange('A17').setValue('Difference (should be 0):');

  sh.getRange('B10').setFormula('=B7+B8-B9');
  sh.getRange('B15').setFormula('=B12+B13-B14');
  sh.getRange('B17').setFormula('=B10-B15');

  sh.getRange('B17').setBackground('#FFF9C4');
  sh.getRange('A10:B10').setFontWeight('bold');
  sh.getRange('A15:B15').setFontWeight('bold');
  sh.getRange('A17:B17').setFontWeight('bold').setBackground('#FFF9C4');

  formatCurrencyColumns(sh, [2], 7, 20);
  sh.autoResizeColumns(1, 2);
}

function setupDashboardSheet(ss) {
  let sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (sh) return;
  sh = ss.insertSheet(SHEETS.DASHBOARD, 0); // Insert at front

  sh.getRange('A1').setValue('ACCOUNTING DASHBOARD').setFontSize(20).setFontWeight('bold').setFontColor('#1565C0');
  sh.getRange('A2').setValue('Last Updated:');
  sh.getRange('B2').setFormula('=NOW()').setNumberFormat('yyyy-mm-dd hh:mm');

  const sections = [
    ['A4',  'CURRENT PERIOD SUMMARY',   '#1565C0'],
    ['A10', 'ACCOUNTS RECEIVABLE',       '#1A237E'],
    ['A16', 'ACCOUNTS PAYABLE',          '#B71C1C'],
    ['A22', 'QUICK LINKS',               '#37474F'],
  ];
  sections.forEach(([cell, label, color]) => {
    sh.getRange(cell).setValue(label).setFontSize(12).setFontWeight('bold').setFontColor(color);
  });

  const metrics = [
    ['A5',  'Total Revenue (YTD):',    'B5',  "=IFERROR(SUMIF('"  + SHEETS.JOURNAL + "'!F:F,\"4*\",'" + SHEETS.JOURNAL + "'!H:H)-SUMIF('" + SHEETS.JOURNAL + "'!F:F,\"4*\",'" + SHEETS.JOURNAL + "'!G:G),0)"],
    ['A6',  'Total Expenses (YTD):',   'B6',  "=IFERROR(SUMIF('"  + SHEETS.JOURNAL + "'!F:F,\"5*\",'" + SHEETS.JOURNAL + "'!G:G)-SUMIF('" + SHEETS.JOURNAL + "'!F:F,\"5*\",'" + SHEETS.JOURNAL + "'!H:H),0)"],
    ['A7',  'Net Income (YTD):',       'B7',  '=B5-B6'],
    ['A8',  'Cash Balance:',           'B8',  "=IFERROR(SUMIF('"  + SHEETS.JOURNAL + "'!E:E,\"1000\",'" + SHEETS.JOURNAL + "'!G:G)-SUMIF('" + SHEETS.JOURNAL + "'!E:E,\"1000\",'" + SHEETS.JOURNAL + "'!H:H),0)"],
    ['A11', 'Total AR Outstanding:',   'B11', "=IFERROR(SUMIF('" + SHEETS.AR_INVOICES + "'!M:M,\"<>Paid\",'" + SHEETS.AR_INVOICES + "'!L:L),0)"],
    ['A12', 'Invoices Overdue:',       'B12', "=IFERROR(COUNTIF('" + SHEETS.AR_INVOICES + "'!M:M,\"Overdue\"),0)"],
    ['A17', 'Total AP Outstanding:',   'B17', "=IFERROR(SUMIF('" + SHEETS.AP_BILLS + "'!M:M,\"<>Paid\",'" + SHEETS.AP_BILLS + "'!L:L),0)"],
    ['A18', 'Bills Overdue:',          'B18', "=IFERROR(COUNTIF('" + SHEETS.AP_BILLS + "'!M:M,\"Overdue\"),0)"],
  ];

  metrics.forEach(([labelCell, label, valCell, formula]) => {
    sh.getRange(labelCell).setValue(label).setFontWeight('bold');
    sh.getRange(valCell).setFormula(formula).setNumberFormat('$#,##0.00');
  });

  sh.getRange('B7').setBackground('#E8F5E9');
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 160);
}

// ---------------------------------------------------------------------------
// Sample data
// ---------------------------------------------------------------------------

function insertSampleData(ss) {
  const today = new Date();
  const fmt = Utilities.formatDate;
  const tz  = ss.getSpreadsheetTimeZone();

  const d = (offset) => {
    const dt = new Date(today);
    dt.setDate(dt.getDate() + offset);
    return fmt(dt, tz, 'yyyy-MM-dd');
  };

  // Sample journal entries
  const jSh = ss.getSheetByName(SHEETS.JOURNAL);
  const journalData = [
    ['JE-001', d(-30), 'OP-001', 'Owner capital contribution',        '1000', 'Cash - Checking',    50000,     0,     'ADMIN', 'Yes', ''],
    ['JE-001', d(-30), 'OP-001', 'Owner capital contribution',        '3000', 'Owner\'s Capital',       0, 50000,     'ADMIN', 'Yes', ''],
    ['JE-002', d(-25), 'INV-001','Sales revenue - Acme Corp',         '1100', 'Accounts Receivable',10000,     0,     'SALES', 'Yes', ''],
    ['JE-002', d(-25), 'INV-001','Sales revenue - Acme Corp',         '4000', 'Sales Revenue',          0, 10000,     'SALES', 'Yes', ''],
    ['JE-003', d(-20), 'EXP-001','Office rent - January',             '5200', 'Rent Expense',        2000,     0,     'ADMIN', 'Yes', ''],
    ['JE-003', d(-20), 'EXP-001','Office rent - January',             '2000', 'Accounts Payable',       0,  2000,     'ADMIN', 'Yes', ''],
    ['JE-004', d(-15), 'EXP-002','Utility bill',                      '5300', 'Utilities Expense',    500,     0,     'OPS',   'Yes', ''],
    ['JE-004', d(-15), 'EXP-002','Utility bill',                      '2000', 'Accounts Payable',       0,   500,     'OPS',   'Yes', ''],
    ['JE-005', d(-10), 'PAY-001','Customer payment - Acme Corp',      '1000', 'Cash - Checking',     8000,     0,     'SALES', 'Yes', ''],
    ['JE-005', d(-10), 'PAY-001','Customer payment - Acme Corp',      '1100', 'Accounts Receivable',    0,  8000,     'SALES', 'Yes', ''],
  ];
  jSh.getRange(2, 1, journalData.length, 11).setValues(journalData);

  // Sample AR invoices
  const arSh = ss.getSheetByName(SHEETS.AR_INVOICES);
  const arData = [
    ['INV-001', d(-25), d(5),  'CUST-001','Acme Corp',   'Consulting services Q1', 10000, 0, 0, 10000, 8000, 2000, 'Partial', '4100', 'SALES', ''],
    ['INV-002', d(-15), d(15), 'CUST-002','Globex Inc',  'Software license',        5000, 10, 500, 5500, 0, 5500, 'Sent',    '4000', 'IT',    ''],
    ['INV-003', d(-5),  d(25), 'CUST-003','Initech LLC', 'Training services',       3000, 0,  0,   3000, 0, 3000, 'Draft',   '4100', 'SALES', ''],
  ];
  arSh.getRange(2, 1, arData.length, 16).setValues(arData);

  // Sample AP bills
  const apSh = ss.getSheetByName(SHEETS.AP_BILLS);
  const apData = [
    ['BILL-001', d(-20), d(10),  'VEND-001','Office Depot',     'Office supplies',  500,  0, 0,   500,  0,   500, 'Received','5400', 'ADMIN', ''],
    ['BILL-002', d(-20), d(10),  'VEND-002','City Electric Co', 'Electricity bill', 500,  0, 0,   500,  0,   500, 'Received','5300', 'OPS',   ''],
    ['BILL-003', d(-20), d(10),  'VEND-003','First National Bank','Rent - Main Office', 2000, 0, 0, 2000, 0, 2000,'Received','5200', 'ADMIN', ''],
  ];
  apSh.getRange(2, 1, apData.length, 16).setValues(apData);
}
