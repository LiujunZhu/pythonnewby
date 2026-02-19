// =============================================================================
// 05_AP.gs — Accounts Payable: Bills & Payments
// =============================================================================

// ---------------------------------------------------------------------------
// Create Bill
// ---------------------------------------------------------------------------

function createBill() {
  const vendors  = getSheetData(SHEETS.VENDORS).filter(r => r['Active'] === 'Yes');
  const depts    = getSheetData(SHEETS.DEPARTMENTS).filter(r => r['Active'] === 'Yes');
  const accounts = getActiveAccounts().filter(a => a.type === 'Expense');

  const vendorOpts = vendors.map(v =>
    `<option value="${v['Vendor ID']}|${v['Name']}|${v['Payment Terms']}|${v['Expense Account']}">${v['Vendor ID']} - ${v['Name']}</option>`
  ).join('');
  const deptOpts = ['<option value=""></option>'].concat(
    depts.map(d => `<option value="${d['Dept Code']}">${d['Dept Code']} - ${d['Department Name']}</option>`)
  ).join('');
  const acctOpts = accounts.map(a =>
    `<option value="${a.num}">${a.num} - ${a.name}</option>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { font-weight: bold; display: block; margin-top: 8px; }
      input, select, textarea { padding: 5px; margin-top: 3px; box-sizing: border-box; }
      .grid3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }
      .line-grid { display: grid; grid-template-columns: 3fr 1fr 1fr 30px; gap: 6px; align-items: end; margin-top: 6px; }
      .totals { background: #FFF3E0; padding: 10px; margin-top: 10px; text-align: right; }
      .btn { background: #B71C1C; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 12px; border-radius: 4px; }
      .add-btn { background: #37474F; color: white; padding: 5px 12px; border: none; cursor: pointer; border-radius: 3px; font-size: 12px; }
      .del-btn { background: #C62828; color: white; border: none; cursor: pointer; border-radius: 3px; padding: 4px 8px; }
      .w100 { width: 100%; }
    </style>
    <h3 style="margin:0 0 12px">Create Vendor Bill</h3>
    <div class="grid3">
      <div><label>Vendor *</label><select id="vendor" class="w100" onchange="prefillAcct()">${vendorOpts}</select></div>
      <div><label>Bill Date *</label><input id="billDate" type="date" value="${fmtDate(new Date())}" class="w100"/></div>
      <div><label>Due Date *</label><input id="dueDate" type="date" value="${fmtDate(addDays(new Date(), 30))}" class="w100"/></div>
    </div>
    <div class="grid3" style="margin-top:8px">
      <div><label>Expense Account</label><select id="expAcct" class="w100">${acctOpts}</select></div>
      <div><label>Department</label><select id="dept" class="w100">${deptOpts}</select></div>
      <div><label>Tax Rate %</label><input id="taxRate" type="number" min="0" max="100" step="0.1" value="0" class="w100" oninput="calcTotals()"/></div>
    </div>
    <div style="margin-top:8px"><label>Notes</label><textarea id="notes" class="w100" rows="2"></textarea></div>

    <h4 style="margin:12px 0 4px">Line Items</h4>
    <div style="display:grid;grid-template-columns:3fr 1fr 1fr 30px;gap:6px;font-weight:bold;font-size:12px;padding:0 2px">
      <span>Description</span><span>Qty</span><span>Unit Price</span><span></span>
    </div>
    <div id="lineItems"></div>
    <button class="add-btn" onclick="addLineItem()" style="margin-top:6px">+ Add Line</button>

    <div class="totals">
      <div>Subtotal: <strong id="subtotal">$0.00</strong></div>
      <div>Tax: <strong id="taxAmt">$0.00</strong></div>
      <div style="font-size:15px;margin-top:4px">Total: <strong id="total">$0.00</strong></div>
    </div>

    <button class="btn" onclick="save()">Create Bill</button>

    <script>
      const vendorData = {};
      document.querySelectorAll('#vendor option').forEach(o => {
        const p = o.value.split('|');
        vendorData[p[0]] = { expAcct: p[3] };
      });

      function prefillAcct() {
        const p = document.getElementById('vendor').value.split('|');
        const acctEl = document.getElementById('expAcct');
        if (p[3]) {
          for (let i = 0; i < acctEl.options.length; i++) {
            if (acctEl.options[i].value === p[3]) { acctEl.selectedIndex = i; break; }
          }
        }
      }

      let lineCount = 0;
      function addLineItem() {
        lineCount++;
        const n = lineCount;
        const div = document.createElement('div');
        div.className = 'line-grid';
        div.id = 'li'+n;
        div.innerHTML = \`
          <input id="ldesc\${n}" type="text" placeholder="Item description" style="width:100%" oninput="calcTotals()"/>
          <input id="lqty\${n}"  type="number" value="1" min="0" step="0.01" style="width:100%" oninput="calcTotals()"/>
          <input id="lprice\${n}" type="number" min="0" step="0.01" placeholder="0.00" style="width:100%" oninput="calcTotals()"/>
          <button class="del-btn" onclick="document.getElementById('li\${n}').remove();calcTotals()">x</button>
        \`;
        document.getElementById('lineItems').appendChild(div);
      }

      function calcTotals() {
        let sub = 0;
        for (let i = 1; i <= lineCount; i++) {
          const qty   = parseFloat(document.getElementById('lqty'+i)?.value)   || 0;
          const price = parseFloat(document.getElementById('lprice'+i)?.value) || 0;
          sub += qty * price;
        }
        const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
        const tax   = sub * taxRate / 100;
        document.getElementById('subtotal').textContent = '$'+sub.toFixed(2);
        document.getElementById('taxAmt').textContent   = '$'+tax.toFixed(2);
        document.getElementById('total').textContent    = '$'+(sub+tax).toFixed(2);
      }

      function save() {
        const vp = document.getElementById('vendor').value.split('|');
        const lines = [];
        for (let i = 1; i <= lineCount; i++) {
          const d = document.getElementById('ldesc'+i);
          if (!d) continue;
          lines.push({ desc: d.value.trim(), qty: parseFloat(document.getElementById('lqty'+i).value)||0, price: parseFloat(document.getElementById('lprice'+i).value)||0 });
        }
        const sub     = lines.reduce((s,l) => s + l.qty*l.price, 0);
        const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
        const tax     = sub * taxRate / 100;
        const data = {
          vendorId:    vp[0],
          vendorName:  vp[1],
          billDate:    document.getElementById('billDate').value,
          dueDate:     document.getElementById('dueDate').value,
          expAcct:     document.getElementById('expAcct').value,
          dept:        document.getElementById('dept').value,
          taxRate, tax, subtotal: sub, total: sub + tax,
          notes:       document.getElementById('notes').value.trim(),
          description: lines.map(l => l.qty+'x '+l.desc+' @'+l.price).join('; '),
        };
        if (!data.vendorId || data.total <= 0) { alert('Vendor and at least one line item are required.'); return; }
        google.script.run
          .withSuccessHandler(id => { alert('Bill '+id+' created!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: '+e.message))
          .saveBill(data);
      }

      prefillAcct();
      addLineItem();
    </script>
  `).setWidth(680).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Vendor Bill');
}

function saveBill(data) {
  const sh     = getSheet(SHEETS.AP_BILLS);
  const billId = getNextId(sh, 1, 'BILL-');

  const row = [
    billId,
    data.billDate,
    data.dueDate,
    data.vendorId,
    data.vendorName,
    data.description,
    data.subtotal,
    data.taxRate,
    data.tax,
    data.total,
    0,
    data.total,
    'Received',
    data.expAcct,
    data.dept,
    data.notes,
  ];

  sh.appendRow(row);
  formatCurrencyColumns(sh, [7, 9, 10, 11, 12], sh.getLastRow(), 1);
  colorStatusColumn(sh, sh.getLastRow());

  postAPJournalEntry(billId, data.billDate, data.vendorName, data.total, data.expAcct, data.dept);

  toast(`Bill ${billId} created for ${fmtCurrency(data.total)}.`);
  return billId;
}

// ---------------------------------------------------------------------------
// Record AP Payment
// ---------------------------------------------------------------------------

function recordAPPayment() {
  const bills = getSheetData(SHEETS.AP_BILLS)
    .filter(r => ['Received', 'Partial', 'Overdue'].includes(String(r['Status'])));

  if (bills.length === 0) {
    SpreadsheetApp.getUi().alert('No open bills found.');
    return;
  }

  const billOpts = bills.map(b =>
    `<option value="${b['Bill #']}|${b['Vendor Name']}|${b['Balance Due']}">${b['Bill #']} - ${b['Vendor Name']} (${fmtCurrency(b['Balance Due'])} due)</option>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { font-weight: bold; display: block; margin-top: 8px; }
      input, select { padding: 5px; margin-top: 3px; width: 100%; box-sizing: border-box; }
      .btn { background: #B71C1C; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 14px; border-radius: 4px; }
    </style>
    <h3 style="margin:0 0 12px">Record AP Payment</h3>
    <label>Bill *</label>
    <select id="bill" onchange="prefillAmount()">${billOpts}</select>
    <label>Payment Date *</label>
    <input id="date" type="date" value="${fmtDate(new Date())}"/>
    <label>Amount *</label>
    <input id="amount" type="number" min="0.01" step="0.01" placeholder="0.00"/>
    <label>Payment Method</label>
    <select id="method">
      <option>Check</option><option>Bank Transfer</option><option>ACH</option>
      <option>Credit Card</option><option>Cash</option>
    </select>
    <label>Reference / Check #</label>
    <input id="ref" type="text" placeholder="Optional"/>
    <label>Notes</label>
    <input id="notes" type="text" placeholder="Optional"/>
    <br/>
    <button class="btn" onclick="save()">Record Payment</button>
    <script>
      function prefillAmount() {
        const parts = document.getElementById('bill').value.split('|');
        document.getElementById('amount').value = parseFloat(parts[2]) || '';
      }
      function save() {
        const parts = document.getElementById('bill').value.split('|');
        const data = {
          billId:     parts[0],
          vendorName: parts[1],
          date:       document.getElementById('date').value,
          amount:     parseFloat(document.getElementById('amount').value),
          method:     document.getElementById('method').value,
          ref:        document.getElementById('ref').value.trim(),
          notes:      document.getElementById('notes').value.trim(),
        };
        if (!data.date || !data.amount) { alert('Date and Amount are required.'); return; }
        google.script.run
          .withSuccessHandler(id => { alert('Payment '+id+' recorded!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: '+e.message))
          .saveAPPayment(data);
      }
      prefillAmount();
    </script>
  `).setWidth(440).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record AP Payment');
}

function saveAPPayment(data) {
  const pmtSh  = getSheet(SHEETS.AP_PAYMENTS);
  const billSh = getSheet(SHEETS.AP_BILLS);

  const pmtId = getNextId(pmtSh, 1, 'APMT-');

  const lastBillRow = billSh.getLastRow();
  const billRows    = billSh.getRange(2, 1, lastBillRow - 1, 16).getValues();
  let   billRowIdx  = -1;

  for (let i = 0; i < billRows.length; i++) {
    if (String(billRows[i][0]) === data.billId) { billRowIdx = i; break; }
  }
  if (billRowIdx < 0) throw new Error(`Bill ${data.billId} not found.`);

  const billRow  = billRows[billRowIdx];
  const total    = Number(billRow[9])  || 0;
  const prevPaid = Number(billRow[10]) || 0;
  const newPaid  = prevPaid + data.amount;
  const newBal   = Math.max(0, total - newPaid);
  const newStatus = newBal <= 0.005 ? 'Paid' : 'Partial';

  const sheetRow = billRowIdx + 2;
  billSh.getRange(sheetRow, 11).setValue(newPaid);
  billSh.getRange(sheetRow, 12).setValue(newBal);
  billSh.getRange(sheetRow, 13).setValue(newStatus);
  colorStatusColumn(billSh, sheetRow);

  pmtSh.appendRow([pmtId, data.date, data.billId, billRow[3], data.vendorName, data.amount, data.method, data.ref, data.notes]);
  formatCurrencyColumns(pmtSh, [6], pmtSh.getLastRow(), 1);

  postAPPaymentJournalEntry(pmtId, data.date, data.vendorName, data.amount, billRow[14]);

  toast(`Payment ${pmtId} of ${fmtCurrency(data.amount)} applied to ${data.billId}.`);
  return pmtId;
}

// ---------------------------------------------------------------------------
// AP Aging Report
// ---------------------------------------------------------------------------

function generateAPAgingReport() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const billData = getSheetData(SHEETS.AP_BILLS);
  const today   = new Date();
  today.setHours(0, 0, 0, 0);

  refreshInvoiceStatuses(SHEETS.AP_BILLS);

  let reportSh = ss.getSheetByName('AP Aging Report');
  if (reportSh) ss.deleteSheet(reportSh);
  reportSh = ss.insertSheet('AP Aging Report');

  setHeaders(reportSh,
    ['Vendor', 'Bill #', 'Bill Date', 'Due Date', 'Total', 'Balance Due', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days'],
    '#B71C1C'
  );

  const rows   = [];
  const totals = { current: 0, d30: 0, d60: 0, d90: 0, over90: 0 };

  billData.forEach(bill => {
    const bal     = Number(bill['Balance Due']) || 0;
    if (bal <= 0) return;
    const dueDate = parseDate(bill['Due Date']);
    if (!dueDate) return;
    const days    = Math.floor((today - dueDate) / 86400000);

    let current = 0, d30 = 0, d60 = 0, d90 = 0, over90 = 0;
    if (days <= 0)         { current = bal; totals.current += bal; }
    else if (days <= 30)   { d30  = bal; totals.d30  += bal; }
    else if (days <= 60)   { d60  = bal; totals.d60  += bal; }
    else if (days <= 90)   { d90  = bal; totals.d90  += bal; }
    else                   { over90 = bal; totals.over90 += bal; }

    rows.push([
      bill['Vendor Name'], bill['Bill #'], bill['Date'], bill['Due Date'],
      Number(bill['Total']) || 0, bal, current, d30, d60, d90, over90
    ]);
  });

  if (rows.length > 0) {
    reportSh.getRange(2, 1, rows.length, 11).setValues(rows);
    formatCurrencyColumns(reportSh, [5, 6, 7, 8, 9, 10, 11], 2, rows.length);
    applyAlternateRowColor(reportSh, 2, rows.length + 1, 11);
  }

  const totRow = rows.length + 2;
  reportSh.getRange(totRow, 1).setValue('TOTAL').setFontWeight('bold');
  reportSh.getRange(totRow, 6, 1, 6).setValues([[
    Object.values(totals).reduce((s, v) => s + v, 0),
    totals.current, totals.d30, totals.d60, totals.d90, totals.over90
  ]]).setFontWeight('bold').setBackground('#FFEBEE');
  formatCurrencyColumns(reportSh, [6, 7, 8, 9, 10, 11], totRow, 1);

  reportSh.setFrozenRows(1);
  reportSh.autoResizeColumns(1, 11);
  ss.setActiveSheet(reportSh);
  toast('AP Aging Report generated.', 'AP Report');
}

// ---------------------------------------------------------------------------
// Refresh AP statuses
// ---------------------------------------------------------------------------
function refreshAPStatuses() {
  refreshInvoiceStatuses(SHEETS.AP_BILLS);
  const sh = getSheet(SHEETS.AP_BILLS);
  colorStatusColumn(sh, sh.getLastRow());
  toast('AP bill statuses refreshed.');
}
