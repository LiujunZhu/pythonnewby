// =============================================================================
// 04_AR.gs — Accounts Receivable: Invoicing & Payments
// =============================================================================

// ---------------------------------------------------------------------------
// Create Invoice
// ---------------------------------------------------------------------------

function createInvoice() {
  const customers = getSheetData(SHEETS.CUSTOMERS).filter(r => r['Active'] === 'Yes');
  const depts     = getSheetData(SHEETS.DEPARTMENTS).filter(r => r['Active'] === 'Yes');
  const accounts  = getActiveAccounts().filter(a => a.type === 'Revenue');

  const customerOpts = customers.map(c =>
    `<option value="${c['Customer ID']}|${c['Customer Name']}|${c['Payment Terms']}">${c['Customer ID']} - ${c['Customer Name']}</option>`
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
      .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
      .grid3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }
      .line-grid { display: grid; grid-template-columns: 3fr 1fr 1fr 30px; gap: 6px; align-items: end; margin-top: 6px; }
      .totals { background: #F5F5F5; padding: 10px; margin-top: 10px; text-align: right; }
      .btn { background: #1A237E; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 12px; border-radius: 4px; }
      .add-btn { background: #2E7D32; color: white; padding: 5px 12px; border: none; cursor: pointer; border-radius: 3px; font-size: 12px; }
      .del-btn { background: #C62828; color: white; border: none; cursor: pointer; border-radius: 3px; padding: 4px 8px; }
      .w100 { width: 100%; }
    </style>
    <h3 style="margin:0 0 12px">Create Invoice</h3>
    <div class="grid3">
      <div><label>Customer *</label><select id="customer" class="w100">${customerOpts}</select></div>
      <div><label>Invoice Date *</label><input id="invDate" type="date" value="${fmtDate(new Date())}" class="w100"/></div>
      <div><label>Due Date *</label><input id="dueDate" type="date" value="${fmtDate(addDays(new Date(), 30))}" class="w100"/></div>
    </div>
    <div class="grid3" style="margin-top:8px">
      <div><label>Revenue Account</label><select id="revAcct" class="w100">${acctOpts}</select></div>
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

    <button class="btn" onclick="save()">Create Invoice</button>

    <script>
      let lineCount = 0;
      function addLineItem(desc, qty, price) {
        lineCount++;
        const n = lineCount;
        const div = document.createElement('div');
        div.className = 'line-grid';
        div.id = 'li'+n;
        div.innerHTML = \`
          <input id="ldesc\${n}" type="text" placeholder="Service/product description" value="\${desc||''}" style="width:100%" oninput="calcTotals()"/>
          <input id="lqty\${n}"  type="number" value="\${qty||1}" min="0" step="0.01" style="width:100%" oninput="calcTotals()"/>
          <input id="lprice\${n}" type="number" value="\${price||''}" min="0" step="0.01" placeholder="0.00" style="width:100%" oninput="calcTotals()"/>
          <button class="del-btn" onclick="document.getElementById('li\${n}').remove();calcTotals()">x</button>
        \`;
        document.getElementById('lineItems').appendChild(div);
        calcTotals();
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
        const total = sub + tax;
        document.getElementById('subtotal').textContent = '$'+sub.toFixed(2);
        document.getElementById('taxAmt').textContent   = '$'+tax.toFixed(2);
        document.getElementById('total').textContent    = '$'+total.toFixed(2);
      }
      function save() {
        const custRaw = document.getElementById('customer').value.split('|');
        const lines = [];
        for (let i = 1; i <= lineCount; i++) {
          const d = document.getElementById('ldesc'+i);
          if (!d) continue;
          lines.push({
            desc:  d.value.trim(),
            qty:   parseFloat(document.getElementById('lqty'+i).value)   || 0,
            price: parseFloat(document.getElementById('lprice'+i).value) || 0,
          });
        }
        const sub     = lines.reduce((s,l) => s + l.qty*l.price, 0);
        const taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
        const tax     = sub * taxRate / 100;
        const data = {
          customerId:   custRaw[0],
          customerName: custRaw[1],
          invDate:      document.getElementById('invDate').value,
          dueDate:      document.getElementById('dueDate').value,
          revAcct:      document.getElementById('revAcct').value,
          dept:         document.getElementById('dept').value,
          taxRate,
          tax,
          subtotal:     sub,
          total:        sub + tax,
          notes:        document.getElementById('notes').value.trim(),
          description:  lines.map(l => l.qty+'x '+l.desc+' @'+l.price).join('; '),
          lines,
        };
        if (!data.customerId || data.total <= 0) { alert('Customer and at least one line item are required.'); return; }
        google.script.run
          .withSuccessHandler(id => { alert('Invoice '+id+' created!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: '+e.message))
          .saveInvoice(data);
      }
      addLineItem();
    </script>
  `).setWidth(680).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Invoice');
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function saveInvoice(data) {
  const sh       = getSheet(SHEETS.AR_INVOICES);
  const invoiceId = getNextId(sh, 1, 'INV-');

  const row = [
    invoiceId,
    data.invDate,
    data.dueDate,
    data.customerId,
    data.customerName,
    data.description,
    data.subtotal,
    data.taxRate,
    data.tax,
    data.total,
    0,            // Amount Paid
    data.total,   // Balance Due
    'Draft',
    data.revAcct,
    data.dept,
    data.notes,
  ];

  sh.appendRow(row);
  formatCurrencyColumns(sh, [7, 9, 10, 11, 12], sh.getLastRow(), 1);
  colorStatusColumn(sh, sh.getLastRow());

  // Post journal entry
  postARJournalEntry(invoiceId, data.invDate, data.customerName, data.total, data.revAcct, data.dept);

  toast(`Invoice ${invoiceId} created for ${fmtCurrency(data.total)}.`);
  return invoiceId;
}

// ---------------------------------------------------------------------------
// Record AR Payment
// ---------------------------------------------------------------------------

function recordARPayment() {
  const invoices = getSheetData(SHEETS.AR_INVOICES)
    .filter(r => ['Sent', 'Partial', 'Overdue'].includes(String(r['Status'])));

  if (invoices.length === 0) {
    SpreadsheetApp.getUi().alert('No open invoices found.');
    return;
  }

  const invOpts = invoices.map(i =>
    `<option value="${i['Invoice #']}|${i['Customer Name']}|${i['Balance Due']}">${i['Invoice #']} - ${i['Customer Name']} (${fmtCurrency(i['Balance Due'])} due)</option>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { font-weight: bold; display: block; margin-top: 8px; }
      input, select { padding: 5px; margin-top: 3px; width: 100%; box-sizing: border-box; }
      .btn { background: #1A237E; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 14px; border-radius: 4px; }
    </style>
    <h3 style="margin:0 0 12px">Record AR Payment</h3>
    <label>Invoice *</label>
    <select id="invoice" onchange="prefillAmount()">${invOpts}</select>
    <label>Payment Date *</label>
    <input id="date" type="date" value="${fmtDate(new Date())}"/>
    <label>Amount *</label>
    <input id="amount" type="number" min="0.01" step="0.01" placeholder="0.00"/>
    <label>Payment Method</label>
    <select id="method">
      <option>Check</option><option>Bank Transfer</option><option>Credit Card</option>
      <option>Cash</option><option>Online Payment</option>
    </select>
    <label>Reference / Check #</label>
    <input id="ref" type="text" placeholder="Optional"/>
    <label>Notes</label>
    <input id="notes" type="text" placeholder="Optional"/>
    <br/>
    <button class="btn" onclick="save()">Record Payment</button>
    <script>
      function prefillAmount() {
        const parts = document.getElementById('invoice').value.split('|');
        document.getElementById('amount').value = parseFloat(parts[2]) || '';
      }
      function save() {
        const parts = document.getElementById('invoice').value.split('|');
        const data = {
          invoiceId:    parts[0],
          customerName: parts[1],
          date:         document.getElementById('date').value,
          amount:       parseFloat(document.getElementById('amount').value),
          method:       document.getElementById('method').value,
          ref:          document.getElementById('ref').value.trim(),
          notes:        document.getElementById('notes').value.trim(),
        };
        if (!data.date || !data.amount) { alert('Date and Amount are required.'); return; }
        google.script.run
          .withSuccessHandler(id => { alert('Payment '+id+' recorded!'); google.script.host.close(); })
          .withFailureHandler(e => alert('Error: '+e.message))
          .saveARPayment(data);
      }
      prefillAmount();
    </script>
  `).setWidth(440).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record AR Payment');
}

function saveARPayment(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const pmtSh = getSheet(SHEETS.AR_PAYMENTS);
  const invSh = getSheet(SHEETS.AR_INVOICES);

  const pmtId = getNextId(pmtSh, 1, 'RCPT-');

  // Find invoice row and update
  const lastInvRow = invSh.getLastRow();
  const invRows    = invSh.getRange(2, 1, lastInvRow - 1, 16).getValues();
  let   invRowIdx  = -1;

  for (let i = 0; i < invRows.length; i++) {
    if (String(invRows[i][0]) === data.invoiceId) { invRowIdx = i; break; }
  }
  if (invRowIdx < 0) throw new Error(`Invoice ${data.invoiceId} not found.`);

  const invRow   = invRows[invRowIdx];
  const total    = Number(invRow[9])  || 0;
  const prevPaid = Number(invRow[10]) || 0;
  const newPaid  = prevPaid + data.amount;
  const newBal   = Math.max(0, total - newPaid);
  const newStatus = newBal <= 0.005 ? 'Paid' : 'Partial';

  const sheetRow = invRowIdx + 2;
  invSh.getRange(sheetRow, 11).setValue(newPaid);
  invSh.getRange(sheetRow, 12).setValue(newBal);
  invSh.getRange(sheetRow, 13).setValue(newStatus);
  colorStatusColumn(invSh, sheetRow);

  // Record payment
  pmtSh.appendRow([pmtId, data.date, data.invoiceId, invRow[3], data.customerName, data.amount, data.method, data.ref, data.notes]);
  formatCurrencyColumns(pmtSh, [6], pmtSh.getLastRow(), 1);

  // Post journal entry
  postARPaymentJournalEntry(pmtId, data.date, data.customerName, data.amount, invRow[14]);

  toast(`Payment ${pmtId} of ${fmtCurrency(data.amount)} applied to ${data.invoiceId}.`);
  return pmtId;
}

// ---------------------------------------------------------------------------
// AR Aging Report (30/60/90 days)
// ---------------------------------------------------------------------------

function generateARAgingReport() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const invData = getSheetData(SHEETS.AR_INVOICES);
  const today   = new Date();
  today.setHours(0, 0, 0, 0);

  // Refresh statuses first
  refreshInvoiceStatuses(SHEETS.AR_INVOICES);

  let reportSh = ss.getSheetByName('AR Aging Report');
  if (reportSh) ss.deleteSheet(reportSh);
  reportSh = ss.insertSheet('AR Aging Report');

  setHeaders(reportSh,
    ['Customer', 'Invoice #', 'Invoice Date', 'Due Date', 'Total', 'Balance Due', 'Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days'],
    '#1A237E'
  );

  const rows = [];
  const totals = { current: 0, d30: 0, d60: 0, d90: 0, over90: 0 };

  invData.forEach(inv => {
    const bal     = Number(inv['Balance Due']) || 0;
    if (bal <= 0) return;
    const dueDate = parseDate(inv['Due Date']);
    if (!dueDate) return;
    const days    = Math.floor((today - dueDate) / 86400000);

    let current = 0, d30 = 0, d60 = 0, d90 = 0, over90 = 0;
    if (days <= 0)         { current = bal; totals.current += bal; }
    else if (days <= 30)   { d30  = bal; totals.d30  += bal; }
    else if (days <= 60)   { d60  = bal; totals.d60  += bal; }
    else if (days <= 90)   { d90  = bal; totals.d90  += bal; }
    else                   { over90 = bal; totals.over90 += bal; }

    rows.push([
      inv['Customer Name'], inv['Invoice #'], inv['Date'], inv['Due Date'],
      Number(inv['Total']) || 0, bal, current, d30, d60, d90, over90
    ]);
  });

  if (rows.length > 0) {
    reportSh.getRange(2, 1, rows.length, 11).setValues(rows);
    formatCurrencyColumns(reportSh, [5, 6, 7, 8, 9, 10, 11], 2, rows.length);
    applyAlternateRowColor(reportSh, 2, rows.length + 1, 11);
  }

  // Totals row
  const totRow = rows.length + 2;
  reportSh.getRange(totRow, 1).setValue('TOTAL').setFontWeight('bold');
  reportSh.getRange(totRow, 6, 1, 6).setValues([[
    Object.values(totals).reduce((s, v) => s + v, 0),
    totals.current, totals.d30, totals.d60, totals.d90, totals.over90
  ]]).setFontWeight('bold').setBackground('#E3F2FD');
  formatCurrencyColumns(reportSh, [6, 7, 8, 9, 10, 11], totRow, 1);

  reportSh.setFrozenRows(1);
  reportSh.autoResizeColumns(1, 11);
  ss.setActiveSheet(reportSh);
  toast('AR Aging Report generated.', 'AR Report');
}

// ---------------------------------------------------------------------------
// Refresh AR statuses
// ---------------------------------------------------------------------------
function refreshARStatuses() {
  refreshInvoiceStatuses(SHEETS.AR_INVOICES);
  const sh = getSheet(SHEETS.AR_INVOICES);
  colorStatusColumn(sh, sh.getLastRow());
  toast('AR invoice statuses refreshed.');
}
