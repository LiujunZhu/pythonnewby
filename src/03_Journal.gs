// =============================================================================
// 03_Journal.gs — Double-entry journal entry management
// =============================================================================

/**
 * Opens a dialog to create a new multi-line journal entry.
 */
function newJournalEntry() {
  const accounts = getActiveAccounts();
  const depts    = getSheetData(SHEETS.DEPARTMENTS).filter(r => r['Active'] === 'Yes');

  const accountOptions = accounts.map(a =>
    `<option value="${a.num}">${a.num} - ${a.name}</option>`
  ).join('');

  const deptOptions = ['<option value=""></option>'].concat(
    depts.map(d => `<option value="${d['Dept Code']}">${d['Dept Code']} - ${d['Department Name']}</option>`)
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { font-weight: bold; display: block; margin-top: 8px; }
      input, select { padding: 5px; margin-top: 3px; }
      .row-section { border: 1px solid #ccc; padding: 10px; margin-top: 10px; border-radius: 4px; }
      .row-header { font-weight: bold; background: #E3F2FD; padding: 6px; margin-bottom: 8px; }
      .grid { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 8px; align-items: end; }
      .totals { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 8px; background: #F5F5F5; padding: 8px; margin-top: 8px; }
      .btn { background: #2E7D32; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 12px; border-radius: 4px; font-size: 13px; }
      .add-btn { background: #1565C0; color: white; padding: 6px 14px; border: none; cursor: pointer; margin-top: 8px; border-radius: 4px; }
      .del-btn { background: #C62828; color: white; padding: 4px 10px; border: none; cursor: pointer; border-radius: 4px; font-size: 11px; }
      .diff-ok { color: #2E7D32; font-weight: bold; }
      .diff-bad { color: #C62828; font-weight: bold; }
      select.acct { width: 100%; }
    </style>
    <h3 style="margin:0 0 12px">New Journal Entry</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
      <div><label>Date *</label><input id="date" type="date" value="${fmtDate(new Date())}" style="width:100%"/></div>
      <div><label>Reference</label><input id="ref" type="text" placeholder="e.g. INV-001" style="width:100%"/></div>
      <div><label>Description *</label><input id="desc" type="text" placeholder="Transaction description" style="width:100%"/></div>
    </div>

    <div id="lines"></div>
    <button class="add-btn" onclick="addLine()">+ Add Line</button>

    <div class="totals" style="margin-top:10px">
      <strong>Totals</strong>
      <div>Debit: <span id="totalDebit" class="diff-ok">0.00</span></div>
      <div>Credit: <span id="totalCredit" class="diff-ok">0.00</span></div>
      <div>Diff: <span id="diff" class="diff-ok">0.00</span></div>
    </div>

    <br/>
    <button class="btn" onclick="saveEntry()">Post Journal Entry</button>

    <script>
      const accountOptions = \`${accountOptions}\`;
      const deptOptions    = \`${deptOptions}\`;
      let lineCount = 0;

      function addLine(acct, dr, cr, dept) {
        lineCount++;
        const id = 'line' + lineCount;
        const div = document.createElement('div');
        div.className = 'row-section';
        div.id = id;
        div.innerHTML = \`
          <div class="row-header">Line \${lineCount} <button class="del-btn" onclick="removeLine('\${id}')">Remove</button></div>
          <div class="grid">
            <div><label>Account *</label><select class="acct" id="acct\${lineCount}">\${accountOptions}</select></div>
            <div><label>Debit</label><input id="dr\${lineCount}" type="number" min="0" step="0.01" value="\${dr||''}" oninput="updateTotals()" style="width:100%"/></div>
            <div><label>Credit</label><input id="cr\${lineCount}" type="number" min="0" step="0.01" value="\${cr||''}" oninput="updateTotals()" style="width:100%"/></div>
            <div><label>Department</label><select id="dept\${lineCount}" style="width:100%">\${deptOptions}</select></div>
          </div>
        \`;
        document.getElementById('lines').appendChild(div);
        if (acct) document.getElementById('acct'+lineCount).value = acct;
        if (dept) document.getElementById('dept'+lineCount).value = dept;
        updateTotals();
      }

      function removeLine(id) {
        document.getElementById(id).remove();
        updateTotals();
      }

      function updateTotals() {
        let dr = 0, cr = 0;
        for (let i = 1; i <= lineCount; i++) {
          const dEl = document.getElementById('dr'+i);
          const cEl = document.getElementById('cr'+i);
          if (dEl) dr += parseFloat(dEl.value) || 0;
          if (cEl) cr += parseFloat(cEl.value) || 0;
        }
        const diff = Math.abs(dr - cr);
        document.getElementById('totalDebit').textContent  = dr.toFixed(2);
        document.getElementById('totalCredit').textContent = cr.toFixed(2);
        const diffEl = document.getElementById('diff');
        diffEl.textContent = diff.toFixed(2);
        diffEl.className = diff < 0.005 ? 'diff-ok' : 'diff-bad';
      }

      function saveEntry() {
        const date = document.getElementById('date').value;
        const ref  = document.getElementById('ref').value.trim();
        const desc = document.getElementById('desc').value.trim();
        if (!date || !desc) { alert('Date and Description are required.'); return; }

        const lines = [];
        for (let i = 1; i <= lineCount; i++) {
          const aEl = document.getElementById('acct'+i);
          if (!aEl) continue;
          const dr = parseFloat(document.getElementById('dr'+i).value) || 0;
          const cr = parseFloat(document.getElementById('cr'+i).value) || 0;
          if (dr === 0 && cr === 0) continue;
          lines.push({ acct: aEl.value, dr, cr, dept: document.getElementById('dept'+i).value });
        }

        if (lines.length < 2) { alert('A journal entry must have at least 2 lines.'); return; }

        const totalDr = lines.reduce((s, l) => s + l.dr, 0);
        const totalCr = lines.reduce((s, l) => s + l.cr, 0);
        if (Math.abs(totalDr - totalCr) > 0.005) {
          alert('Debits must equal Credits. Current difference: ' + Math.abs(totalDr-totalCr).toFixed(2));
          return;
        }

        google.script.run
          .withSuccessHandler(() => { alert('Journal entry posted!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: ' + e.message))
          .saveJournalEntry({ date, ref, desc, lines });
      }

      // Start with 2 blank lines
      addLine(); addLine();
    </script>
  `).setWidth(700).setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, 'New Journal Entry');
}

/**
 * Called from HTML dialog — writes journal lines to the Journal sheet.
 */
function saveJournalEntry(data) {
  const sh      = getSheet(SHEETS.JOURNAL);
  const accMap  = buildAccountMap();
  const entryId = getNextId(sh, 1, 'JE-');

  const rows = data.lines.map(line => {
    const acc = accMap[line.acct] || { name: '' };
    return [
      entryId,
      data.date,
      data.ref,
      data.desc,
      line.acct,
      acc.name,
      line.dr || 0,
      line.cr || 0,
      line.dept || '',
      'Yes',
      '',
    ];
  });

  const lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, rows.length, 11).setValues(rows);
  formatCurrencyColumns(sh, [7, 8], lastRow + 1, rows.length);
  applyAlternateRowColor(sh, 2, sh.getLastRow(), 11);

  // Refresh the GL
  refreshGeneralLedger();

  toast(`Journal Entry ${entryId} posted with ${rows.length} lines.`);
  return entryId;
}

/**
 * Void a journal entry by entry ID — adds reversing lines.
 */
function voidJournalEntry() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt('Void Journal Entry', 'Enter Entry ID to void (e.g. JE-003):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const entryId = resp.getResponseText().trim().toUpperCase();
  if (!entryId) return;

  const sh      = getSheet(SHEETS.JOURNAL);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('No journal entries found.'); return; }

  const rows = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  const entry = rows.filter(r => String(r[0]) === entryId);

  if (entry.length === 0) {
    ui.alert(`Entry ${entryId} not found.`);
    return;
  }

  const voidId   = entryId + '-VOID';
  const today    = fmtDate(new Date());
  const voidRows = entry.map(row => [
    voidId,
    today,
    row[2],
    'VOID: ' + row[3],
    row[4],
    row[5],
    row[7], // swap debit/credit
    row[6],
    row[8],
    'Yes',
    `Void of ${entryId}`,
  ]);

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, voidRows.length, 11).setValues(voidRows);
  formatCurrencyColumns(sh, [7, 8], startRow, voidRows.length);
  applyAlternateRowColor(sh, 2, sh.getLastRow(), 11);

  refreshGeneralLedger();
  toast(`Entry ${entryId} voided as ${voidId}.`);
}

/**
 * Validate that all posted journal entries are balanced.
 */
function validateJournal() {
  const sh      = getSheet(SHEETS.JOURNAL);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No journal entries to validate.');
    return;
  }

  const rows    = sh.getRange(2, 1, lastRow - 1, 10).getValues();
  const entries = {};

  rows.forEach(row => {
    const id = String(row[0]);
    if (!entries[id]) entries[id] = { debits: 0, credits: 0 };
    entries[id].debits  += Number(row[6]) || 0;
    entries[id].credits += Number(row[7]) || 0;
  });

  const unbalanced = Object.entries(entries)
    .filter(([, v]) => Math.abs(v.debits - v.credits) > 0.005)
    .map(([id, v]) => `${id}: DR=${v.debits.toFixed(2)} CR=${v.credits.toFixed(2)}`);

  if (unbalanced.length === 0) {
    SpreadsheetApp.getUi().alert('All journal entries are balanced.');
  } else {
    SpreadsheetApp.getUi().alert(
      `${unbalanced.length} unbalanced entries found:\n\n${unbalanced.join('\n')}`
    );
  }
}

/**
 * Auto-generate a journal entry from an AR invoice (posted when invoice is created).
 */
function postARJournalEntry(invoiceNum, date, customer, amount, revenueAcct, dept) {
  return saveJournalEntry({
    date,
    ref:  invoiceNum,
    desc: `Invoice ${invoiceNum} - ${customer}`,
    lines: [
      { acct: '1100', dr: amount, cr: 0,      dept },
      { acct: revenueAcct || '4000', dr: 0, cr: amount, dept },
    ],
  });
}

/**
 * Auto-generate a journal entry when an AR payment is received.
 */
function postARPaymentJournalEntry(paymentNum, date, customer, amount, dept) {
  return saveJournalEntry({
    date,
    ref:  paymentNum,
    desc: `Payment received - ${customer}`,
    lines: [
      { acct: '1000', dr: amount, cr: 0,      dept },
      { acct: '1100', dr: 0,      cr: amount, dept },
    ],
  });
}

/**
 * Auto-generate a journal entry from an AP bill.
 */
function postAPJournalEntry(billNum, date, vendor, amount, expenseAcct, dept) {
  return saveJournalEntry({
    date,
    ref:  billNum,
    desc: `Bill ${billNum} - ${vendor}`,
    lines: [
      { acct: expenseAcct || '5900', dr: amount, cr: 0,      dept },
      { acct: '2000',                dr: 0,      cr: amount, dept },
    ],
  });
}

/**
 * Auto-generate a journal entry when an AP payment is made.
 */
function postAPPaymentJournalEntry(paymentNum, date, vendor, amount, dept) {
  return saveJournalEntry({
    date,
    ref:  paymentNum,
    desc: `Payment to vendor - ${vendor}`,
    lines: [
      { acct: '2000', dr: amount, cr: 0,      dept },
      { acct: '1000', dr: 0,      cr: amount, dept },
    ],
  });
}
