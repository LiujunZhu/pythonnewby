// =============================================================================
// 06_BankReconciliation.gs — Bank Statement Import & Reconciliation
// =============================================================================

/**
 * Opens a dialog to import bank statement transactions.
 */
function importBankStatement() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { font-weight: bold; display: block; margin-top: 8px; }
      input, select, textarea { padding: 5px; margin-top: 3px; box-sizing: border-box; }
      .btn { background: #004D40; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 14px; border-radius: 4px; }
      .info { background: #E0F2F1; padding: 8px; border-radius: 4px; margin-bottom: 10px; font-size: 12px; }
      .w100 { width: 100%; }
    </style>
    <h3 style="margin:0 0 10px">Import Bank Statement</h3>
    <div class="info">
      Paste CSV data in format: <strong>Date, Description, Reference, Debit, Credit</strong><br/>
      Date format: YYYY-MM-DD | Amounts should be positive numbers.
    </div>
    <label>Opening Balance (Statement)</label>
    <input id="openBal" type="number" step="0.01" placeholder="0.00" class="w100"/>
    <label>CSV Data *</label>
    <textarea id="csvData" rows="12" class="w100" placeholder="2024-01-15,Deposit - Customer Payment,,0,5000.00
2024-01-16,Office Supplies - Vendor,CHK-1001,250.00,0
2024-01-17,Utility Payment,ACH-0045,180.00,0"></textarea>
    <br/>
    <button class="btn" onclick="importData()">Import Transactions</button>
    <script>
      function importData() {
        const csv     = document.getElementById('csvData').value.trim();
        const openBal = parseFloat(document.getElementById('openBal').value) || 0;
        if (!csv) { alert('Please paste CSV data first.'); return; }
        google.script.run
          .withSuccessHandler(count => { alert(count + ' transactions imported!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: ' + e.message))
          .processBankStatementCSV(csv, openBal);
      }
    </script>
  `).setWidth(520).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Bank Statement');
}

/**
 * Parses CSV and appends transactions to the Bank Statement sheet.
 */
function processBankStatementCSV(csvText, openingBalance) {
  const sh = getSheet(SHEETS.BANK_STMT);
  const lines = csvText.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  let runBal   = openingBalance;
  const rows   = [];

  lines.forEach(line => {
    const parts = line.split(',').map(p => p.trim().replace(/^"|"$/g, ''));
    if (parts.length < 5) return;

    const date  = parts[0];
    const desc  = parts[1];
    const ref   = parts[2];
    const debit = parseFloat(parts[3]) || 0;
    const credit= parseFloat(parts[4]) || 0;

    runBal = runBal - debit + credit;
    rows.push([date, desc, ref, debit, credit, runBal, 'No', '']);
  });

  if (rows.length === 0) throw new Error('No valid rows found in CSV data.');

  const lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
  formatCurrencyColumns(sh, [4, 5, 6], lastRow + 1, rows.length);
  applyAlternateRowColor(sh, 2, sh.getLastRow(), 8);

  toast(`${rows.length} bank transactions imported.`);
  return rows.length;
}

// ---------------------------------------------------------------------------
// Auto-match bank statement to journal entries
// ---------------------------------------------------------------------------

/**
 * Attempts to auto-match unmatched bank statement lines to journal entries
 * by amount and date proximity (within 3 days).
 */
function autoMatchBankStatement() {
  const bankSh  = getSheet(SHEETS.BANK_STMT);
  const jSh     = getSheet(SHEETS.JOURNAL);

  const bankLastRow = bankSh.getLastRow();
  const jLastRow    = jSh.getLastRow();

  if (bankLastRow < 2 || jLastRow < 2) {
    SpreadsheetApp.getUi().alert('Not enough data to match.');
    return;
  }

  const bankRows = bankSh.getRange(2, 1, bankLastRow - 1, 8).getValues();
  const jRows    = jSh.getRange(2, 1, jLastRow - 1, 11).getValues();

  // Build a map of journal entries by (amount, date range)
  // Group journal lines by entry ID to get net debit/credit per entry
  const jeMap = {};
  jRows.forEach(r => {
    const id   = String(r[0]);
    const date = parseDate(r[1]);
    if (!jeMap[id]) jeMap[id] = { date, netAmount: 0, id };
    jeMap[id].netAmount += (Number(r[6]) || 0) - (Number(r[7]) || 0);
  });

  let matched = 0;

  bankRows.forEach((bRow, bi) => {
    if (String(bRow[6]) === 'Yes') return; // already matched
    const bDate  = parseDate(bRow[0]);
    const bDebit = Number(bRow[3]) || 0;
    const bCred  = Number(bRow[4]) || 0;
    const bNet   = bCred - bDebit; // positive = credit/inflow

    // Try to find a matching journal entry
    const match = Object.values(jeMap).find(je => {
      if (!je.date || !bDate) return false;
      const dayDiff = Math.abs((je.date - bDate) / 86400000);
      const amtMatch = Math.abs(Math.abs(je.netAmount) - Math.abs(bNet)) < 0.005;
      return dayDiff <= 3 && amtMatch;
    });

    if (match) {
      bankSh.getRange(bi + 2, 7).setValue('Yes');
      bankSh.getRange(bi + 2, 8).setValue(match.id);
      bankSh.getRange(bi + 2, 7, 1, 2).setBackground('#C8E6C9');
      matched++;
    }
  });

  toast(`Auto-match complete: ${matched} transactions matched.`, 'Bank Recon');
}

// ---------------------------------------------------------------------------
// Mark a single bank line as matched (manual)
// ---------------------------------------------------------------------------

function manualMatchBankLine() {
  const ui   = SpreadsheetApp.getUi();
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sh.getName() !== SHEETS.BANK_STMT) {
    ui.alert('Please select a row in the Bank Statement sheet first.');
    return;
  }

  const row = sh.getActiveCell().getRow();
  if (row < 2) { ui.alert('Please select a data row (not the header).'); return; }

  const resp = ui.prompt(
    'Manual Match',
    'Enter the Journal Entry ID to link (e.g. JE-005):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const jeId = resp.getResponseText().trim().toUpperCase();
  sh.getRange(row, 7).setValue('Yes');
  sh.getRange(row, 8).setValue(jeId);
  sh.getRange(row, 7, 1, 2).setBackground('#C8E6C9');
  toast(`Row ${row} matched to ${jeId}.`);
}

// ---------------------------------------------------------------------------
// Generate reconciliation summary
// ---------------------------------------------------------------------------

function runBankReconciliation() {
  const ui      = SpreadsheetApp.getUi();
  const bankSh  = getSheet(SHEETS.BANK_STMT);
  const reconSh = getSheet(SHEETS.RECON);

  // Get bank statement balance range info from user
  const stmtDate = ui.prompt('Reconciliation', 'Enter Statement Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (stmtDate.getSelectedButton() !== ui.Button.OK) return;

  const endBal = ui.prompt('Reconciliation', 'Enter Bank Statement Ending Balance:', ui.ButtonSet.OK_CANCEL);
  if (endBal.getSelectedButton() !== ui.Button.OK) return;

  const bankEndBal = parseFloat(endBal.getResponseText()) || 0;

  // Calculate unmatched transactions
  const lastRow  = bankSh.getLastRow();
  if (lastRow < 2) { ui.alert('No bank statement data found.'); return; }

  const bankRows = bankSh.getRange(2, 1, lastRow - 1, 8).getValues();
  let unmatched_deposits = 0;
  let unmatched_payments = 0;

  bankRows.forEach(row => {
    if (String(row[6]) !== 'Yes') {
      const credit = Number(row[4]) || 0;
      const debit  = Number(row[3]) || 0;
      if (credit > 0) unmatched_deposits += credit;
      if (debit  > 0) unmatched_payments += debit;
    }
  });

  // Get GL balance for cash account (1000)
  const glBal = getAccountBalance('1000');

  // Fill reconciliation sheet
  reconSh.getRange('B3').setValue(stmtDate.getResponseText());
  reconSh.getRange('B4').setValue('Checking Account');
  reconSh.getRange('B5').setValue('1000');
  reconSh.getRange('B7').setValue(bankEndBal);
  reconSh.getRange('B8').setValue(unmatched_deposits);
  reconSh.getRange('B9').setValue(unmatched_payments);
  reconSh.getRange('B12').setValue(glBal);
  reconSh.getRange('B13').setValue(0);
  reconSh.getRange('B14').setValue(0);

  // Colour the diff cell
  const diff = reconSh.getRange('B17').getValue();
  reconSh.getRange('B17').setBackground(Math.abs(diff) < 0.005 ? '#C8E6C9' : '#FFCDD2');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(reconSh);

  toast(
    Math.abs(diff) < 0.005 ? 'Reconciled! Difference is $0.00.' : `Difference: ${fmtCurrency(diff)}`,
    'Bank Reconciliation',
    6
  );
}

// ---------------------------------------------------------------------------
// Summary of unmatched items
// ---------------------------------------------------------------------------

function showUnmatchedSummary() {
  const bankSh  = getSheet(SHEETS.BANK_STMT);
  const lastRow = bankSh.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No bank statement data.');
    return;
  }

  const bankRows = bankSh.getRange(2, 1, lastRow - 1, 8).getValues();
  const unmatched = bankRows.filter(r => String(r[6]) !== 'Yes');

  if (unmatched.length === 0) {
    SpreadsheetApp.getUi().alert('All bank statement items are matched!');
    return;
  }

  let msg = `${unmatched.length} unmatched transactions:\n\n`;
  unmatched.slice(0, 20).forEach(r => {
    msg += `${r[0]}  ${r[1]}  DR:${r[3]}  CR:${r[4]}\n`;
  });
  if (unmatched.length > 20) msg += `\n... and ${unmatched.length - 20} more.`;

  SpreadsheetApp.getUi().alert(msg);
}
