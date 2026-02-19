// =============================================================================
// 02_ChartOfAccounts.gs — Chart of Accounts management
// =============================================================================

/**
 * Opens a dialog to add a new account to the Chart of Accounts.
 */
function addAccount() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; }
      label { font-weight: bold; display: block; margin-top: 10px; }
      input, select, textarea { width: 100%; padding: 6px; margin-top: 4px; box-sizing: border-box; }
      .btn { background: #1565C0; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 16px; border-radius: 4px; }
      .btn:hover { background: #0D47A1; }
    </style>
    <h3>Add New Account</h3>
    <label>Account Number *</label>
    <input id="num" type="text" placeholder="e.g. 5100" />
    <label>Account Name *</label>
    <input id="name" type="text" placeholder="e.g. Salaries Expense" />
    <label>Type *</label>
    <select id="type">
      <option>Asset</option><option>Liability</option>
      <option>Equity</option><option>Revenue</option><option>Expense</option>
    </select>
    <label>Sub-Type</label>
    <input id="subtype" type="text" placeholder="e.g. Current Asset" />
    <label>Normal Balance</label>
    <select id="balance">
      <option>Debit</option><option>Credit</option>
    </select>
    <label>Department (optional)</label>
    <input id="dept" type="text" placeholder="e.g. SALES" />
    <label>Description</label>
    <textarea id="desc" rows="2"></textarea>
    <br/>
    <button class="btn" onclick="save()">Save Account</button>
    <script>
      function save() {
        const data = {
          num:     document.getElementById('num').value.trim(),
          name:    document.getElementById('name').value.trim(),
          type:    document.getElementById('type').value,
          subtype: document.getElementById('subtype').value.trim(),
          balance: document.getElementById('balance').value,
          dept:    document.getElementById('dept').value.trim(),
          desc:    document.getElementById('desc').value.trim(),
        };
        if (!data.num || !data.name) { alert('Account Number and Name are required.'); return; }
        google.script.run.withSuccessHandler(() => {
          alert('Account saved!');
          google.script.host.close();
        }).withFailureHandler(e => alert('Error: ' + e.message)).saveNewAccount(data);
      }
    </script>
  `).setWidth(400).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Account');
}

/**
 * Called from the HTML dialog — inserts a new row in the COA sheet.
 */
function saveNewAccount(data) {
  const sh = getSheet(SHEETS.COA);
  const lastRow = sh.getLastRow();

  // Determine normal balance from type if not set
  let normalBal = data.balance;
  if (!normalBal) {
    normalBal = ['Asset', 'Expense'].includes(data.type) ? 'Debit' : 'Credit';
  }

  // Check for duplicate account number
  if (lastRow >= 2) {
    const existing = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    if (existing.includes(data.num)) {
      throw new Error(`Account number ${data.num} already exists.`);
    }
  }

  const newRow = [data.num, data.name, data.type, data.subtype, normalBal, data.dept, data.desc, 'Yes'];
  sh.appendRow(newRow);
  sortCOA();
  toast(`Account ${data.num} - ${data.name} added.`);
}

/**
 * Sort the COA sheet by Account Number.
 */
function sortCOA() {
  const sh = getSheet(SHEETS.COA);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return;
  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).sort({ column: 1, ascending: true });
}

/**
 * Returns an array of { num, name } objects for all active accounts.
 * Used for dropdowns in dialogs.
 */
function getActiveAccounts() {
  const data = getSheetData(SHEETS.COA);
  return data
    .filter(r => r['Active'] === 'Yes')
    .map(r => ({ num: r['Account #'], name: r['Account Name'], type: r['Type'] }));
}

/**
 * Returns a lookup map: accountNum → { name, type, normalBalance }
 */
function buildAccountMap() {
  const data = getSheetData(SHEETS.COA);
  const map  = {};
  data.forEach(r => {
    map[r['Account #']] = {
      name:    r['Account Name'],
      type:    r['Type'],
      subType: r['Sub-Type'],
      normal:  r['Normal Balance'],
    };
  });
  return map;
}

/**
 * Refresh the General Ledger sheet from journal entries.
 */
function refreshGeneralLedger() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const jSh     = getSheet(SHEETS.JOURNAL);
  const glSh    = getSheet(SHEETS.GL);
  const accMap  = buildAccountMap();

  const lastRow = jSh.getLastRow();
  const journalRows = lastRow >= 2
    ? jSh.getRange(2, 1, lastRow - 1, 11).getValues()
    : [];

  // Aggregate debits and credits per account
  const ledger = {};
  journalRows.forEach(row => {
    const acctNum = String(row[4]);   // col E: Account #
    const debit   = Number(row[6]) || 0; // col G
    const credit  = Number(row[7]) || 0; // col H
    const posted  = String(row[9]);   // col J: Posted

    if (posted !== 'Yes') return;
    if (!acctNum) return;

    if (!ledger[acctNum]) ledger[acctNum] = { debits: 0, credits: 0 };
    ledger[acctNum].debits  += debit;
    ledger[acctNum].credits += credit;
  });

  // Build GL rows
  const glRows = Object.keys(accMap).sort().map(num => {
    const acc  = accMap[num];
    const entry = ledger[num] || { debits: 0, credits: 0 };
    let closing;
    if (acc.normal === 'Debit') {
      closing = entry.debits - entry.credits;
    } else {
      closing = entry.credits - entry.debits;
    }
    return [num, acc.name, acc.type, 0, entry.debits, entry.credits, closing];
  });

  // Write to GL
  glSh.clearContents();
  setHeaders(glSh, ['Account #', 'Account Name', 'Type', 'Opening Balance', 'Total Debits', 'Total Credits', 'Closing Balance'], '#4A148C');
  if (glRows.length > 0) {
    glSh.getRange(2, 1, glRows.length, 7).setValues(glRows);
    formatCurrencyColumns(glSh, [4, 5, 6, 7], 2, glRows.length);
    applyAlternateRowColor(glSh, 2, glRows.length + 1, 7);
  }
  glSh.setFrozenRows(1);
  glSh.autoResizeColumns(1, 7);

  toast('General Ledger refreshed.', 'GL Updated');
}

/**
 * Get the current balance of a specific account number.
 */
function getAccountBalance(accountNum) {
  const accMap = buildAccountMap();
  const acc    = accMap[String(accountNum)];
  if (!acc) return 0;

  const jSh    = getSheet(SHEETS.JOURNAL);
  const lastRow = jSh.getLastRow();
  if (lastRow < 2) return 0;

  const rows = jSh.getRange(2, 1, lastRow - 1, 10).getValues();
  let debits = 0, credits = 0;

  rows.forEach(row => {
    if (String(row[4]) === String(accountNum) && row[9] === 'Yes') {
      debits  += Number(row[6]) || 0;
      credits += Number(row[7]) || 0;
    }
  });

  return acc.normal === 'Debit' ? debits - credits : credits - debits;
}
