// =============================================================================
// Code.gs — Main entry point: menu builder and trigger handlers
// =============================================================================

/**
 * Runs automatically when the spreadsheet is opened.
 * Builds the custom Accounting menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Accounting')

    // ── Setup ────────────────────────────────────────────────────────────────
    .addItem('🔧 Setup Workbook (first time)',     'setupSpreadsheet')
    .addSeparator()

    // ── Chart of Accounts ────────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📋 Chart of Accounts')
        .addItem('Add Account',              'addAccount')
        .addItem('Sort Accounts',            'sortCOA')
        .addItem('Refresh General Ledger',   'refreshGeneralLedger')
    )
    .addSeparator()

    // ── Journal ──────────────────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📝 Journal')
        .addItem('New Journal Entry',        'newJournalEntry')
        .addItem('Void Journal Entry',       'voidJournalEntry')
        .addItem('Validate Journal Balance', 'validateJournal')
    )
    .addSeparator()

    // ── Accounts Receivable ──────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('💰 Accounts Receivable')
        .addItem('Create Invoice',            'createInvoice')
        .addItem('Record Payment Received',   'recordARPayment')
        .addItem('Refresh Invoice Statuses',  'refreshARStatuses')
        .addItem('Generate AR Aging Report',  'generateARAgingReport')
    )
    .addSeparator()

    // ── Accounts Payable ─────────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('🧾 Accounts Payable')
        .addItem('Create Vendor Bill',        'createBill')
        .addItem('Record Vendor Payment',     'recordAPPayment')
        .addItem('Refresh Bill Statuses',     'refreshAPStatuses')
        .addItem('Generate AP Aging Report',  'generateAPAgingReport')
    )
    .addSeparator()

    // ── Bank Reconciliation ──────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('🏦 Bank Reconciliation')
        .addItem('Import Bank Statement',     'importBankStatement')
        .addItem('Auto-Match Transactions',   'autoMatchBankStatement')
        .addItem('Manual Match Selected Row', 'manualMatchBankLine')
        .addItem('Run Reconciliation',        'runBankReconciliation')
        .addItem('Show Unmatched Items',      'showUnmatchedSummary')
    )
    .addSeparator()

    // ── Reports ──────────────────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📈 Reports')
        .addItem('Profit & Loss (Custom Period)',  'generateProfitLoss')
        .addItem('Profit & Loss (Year-to-Date)',   'generatePLYTD')
        .addItem('Balance Sheet',                  'generateBalanceSheet')
        .addItem('Cash Flow Statement',            'generateCashFlow')
        .addItem('Trial Balance',                  'generateTrialBalance')
        .addItem('AR Aging Report',                'generateARAgingReport')
        .addItem('AP Aging Report',                'generateAPAgingReport')
        .addItem('Department Report',              'generateDepartmentReport')
    )
    .addSeparator()

    // ── Utilities ────────────────────────────────────────────────────────────
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('⚙️ Utilities')
        .addItem('Refresh All Statuses',          'refreshAllStatuses')
        .addItem('Refresh Dashboard',             'refreshDashboard')
        .addItem('Refresh General Ledger',        'refreshGeneralLedger')
    )

    .addToUi();
}

// ---------------------------------------------------------------------------
// Utility wrappers called from the menu
// ---------------------------------------------------------------------------

/**
 * Refresh AR and AP statuses and General Ledger in one pass.
 */
function refreshAllStatuses() {
  refreshARStatuses();
  refreshAPStatuses();
  refreshGeneralLedger();
  toast('All statuses and General Ledger refreshed.', 'Refresh');
}

/**
 * Update the Dashboard "Last Updated" timestamp.
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) return;
  sh.getRange('B2').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');
  toast('Dashboard refreshed.');
}

// ---------------------------------------------------------------------------
// Time-driven trigger installer
// ---------------------------------------------------------------------------

/**
 * Install a daily trigger to refresh AR/AP statuses automatically.
 * Run once manually to set up.
 */
function installDailyTrigger() {
  // Remove existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyRefresh')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailyRefresh')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('Daily refresh trigger installed (runs at 8:00 AM every day).');
}

/**
 * Handler called by the daily time trigger.
 */
function dailyRefresh() {
  try {
    refreshInvoiceStatuses(SHEETS.AR_INVOICES);
    refreshInvoiceStatuses(SHEETS.AP_BILLS);
    refreshGeneralLedger();
    refreshDashboard();
    console.log('Daily refresh completed: ' + new Date().toISOString());
  } catch (e) {
    console.error('Daily refresh error: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// onEdit trigger — auto-recalculate invoice balances when edited manually
// ---------------------------------------------------------------------------

function onEdit(e) {
  const sh     = e.range.getSheet();
  const shName = sh.getName();
  const col    = e.range.getColumn();
  const row    = e.range.getRow();

  if (row < 2) return; // ignore header edits

  // When Amount Paid (col 11) is changed on AR or AP sheet, recalc balance
  if (shName === SHEETS.AR_INVOICES && col === 11) {
    const total  = Number(sh.getRange(row, 10).getValue()) || 0;
    const paid   = Number(e.value) || 0;
    const newBal = Math.max(0, total - paid);
    sh.getRange(row, 12).setValue(newBal);

    const dueDate = parseDate(sh.getRange(row, 3).getValue());
    const today   = new Date();
    let status;
    if (newBal <= 0.005) status = 'Paid';
    else if (dueDate && dueDate < today) status = 'Overdue';
    else if (paid > 0) status = 'Partial';
    else status = 'Sent';
    sh.getRange(row, 13).setValue(status);
    colorStatusColumn(sh, row);
  }

  if (shName === SHEETS.AP_BILLS && col === 11) {
    const total  = Number(sh.getRange(row, 10).getValue()) || 0;
    const paid   = Number(e.value) || 0;
    const newBal = Math.max(0, total - paid);
    sh.getRange(row, 12).setValue(newBal);

    const dueDate = parseDate(sh.getRange(row, 3).getValue());
    const today   = new Date();
    let status;
    if (newBal <= 0.005) status = 'Paid';
    else if (dueDate && dueDate < today) status = 'Overdue';
    else if (paid > 0) status = 'Partial';
    else status = 'Received';
    sh.getRange(row, 13).setValue(status);
    colorStatusColumn(sh, row);
  }
}
