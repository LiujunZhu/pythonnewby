// =============================================================================
// 07_Reports.gs — Financial Reports: P&L, Balance Sheet, Cash Flow, Departments
// =============================================================================

// ---------------------------------------------------------------------------
// Profit & Loss (Income Statement)
// ---------------------------------------------------------------------------

function generateProfitLoss() {
  const ui = SpreadsheetApp.getUi();

  const fromResp = ui.prompt('P&L Report', 'Start Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  const toResp = ui.prompt('P&L Report', 'End Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (toResp.getSelectedButton() !== ui.Button.OK) return;

  const fromDate = new Date(fromResp.getResponseText());
  const toDate   = new Date(toResp.getResponseText());
  toDate.setHours(23, 59, 59);

  _generatePL(fromDate, toDate);
}

function generatePLYTD() {
  const now   = new Date();
  const start = new Date(now.getFullYear(), 0, 1); // Jan 1
  const end   = new Date(now.getFullYear(), 11, 31, 23, 59, 59); // Dec 31
  _generatePL(start, end, 'YTD');
}

function _generatePL(fromDate, toDate, suffix) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const accMap  = buildAccountMap();
  const jData   = getSheetData(SHEETS.JOURNAL);

  // Aggregate by account for the period
  const accountTotals = {};
  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const txDate = parseDate(row['Date']);
    if (!txDate || txDate < fromDate || txDate > toDate) return;

    const acctNum = String(row['Account #']);
    const acc     = accMap[acctNum];
    if (!acc) return;
    if (!['Revenue', 'Expense'].includes(acc.type)) return;

    if (!accountTotals[acctNum]) {
      accountTotals[acctNum] = { name: acc.name, type: acc.type, subType: acc.subType, net: 0 };
    }
    const dr = Number(row['Debit'])  || 0;
    const cr = Number(row['Credit']) || 0;
    if (acc.type === 'Revenue') {
      accountTotals[acctNum].net += cr - dr;
    } else {
      accountTotals[acctNum].net += dr - cr;
    }
  });

  const shName = suffix ? `P&L Report (${suffix})` : 'P&L Report';
  let sh = ss.getSheetByName(shName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(shName);

  let r = 1;
  const write = (label, value, bold, bg, indent) => {
    const cell = sh.getRange(r, indent ? 2 : 1);
    cell.setValue(label);
    if (bold) cell.setFontWeight('bold');
    if (bg)   sh.getRange(r, 1, 1, 3).setBackground(bg);
    if (value !== undefined && value !== null) {
      sh.getRange(r, 3).setValue(value).setNumberFormat('$#,##0.00');
      if (bold && bg) sh.getRange(r, 3).setFontWeight('bold');
    }
    r++;
  };

  // Title
  sh.getRange(r, 1).setValue('PROFIT & LOSS STATEMENT').setFontSize(16).setFontWeight('bold').setFontColor('#2E7D32');
  r++;
  sh.getRange(r, 1).setValue(
    `Period: ${fmtDate(fromDate)} to ${fmtDate(toDate)}`
  ).setFontStyle('italic');
  r += 2;

  // Revenue section
  write('REVENUE', null, true, '#E8F5E9');
  const revenues = Object.entries(accountTotals).filter(([, v]) => v.type === 'Revenue').sort();
  let totalRevenue = 0;
  revenues.forEach(([num, v]) => {
    write(`  ${num} - ${v.name}`, v.net, false, null, false);
    totalRevenue += v.net;
  });
  write('Total Revenue', totalRevenue, true, '#C8E6C9');
  r++;

  // COGS section
  const cogsAccounts = Object.entries(accountTotals)
    .filter(([, v]) => v.type === 'Expense' && v.subType === 'COGS').sort();
  let totalCOGS = 0;
  if (cogsAccounts.length > 0) {
    write('COST OF GOODS SOLD', null, true, '#FFF3E0');
    cogsAccounts.forEach(([num, v]) => {
      write(`  ${num} - ${v.name}`, v.net, false, null, false);
      totalCOGS += v.net;
    });
    write('Total COGS', totalCOGS, true, '#FFE0B2');
    r++;
  }

  const grossProfit = totalRevenue - totalCOGS;
  write('GROSS PROFIT', grossProfit, true, grossProfit >= 0 ? '#A5D6A7' : '#FFCDD2');
  r++;

  // Operating Expenses
  write('OPERATING EXPENSES', null, true, '#FFEBEE');
  const opExpAccounts = Object.entries(accountTotals)
    .filter(([, v]) => v.type === 'Expense' && v.subType !== 'COGS').sort();
  let totalOpExp = 0;
  opExpAccounts.forEach(([num, v]) => {
    write(`  ${num} - ${v.name}`, v.net, false, null, false);
    totalOpExp += v.net;
  });
  write('Total Operating Expenses', totalOpExp, true, '#FFCDD2');
  r++;

  const netIncome = grossProfit - totalOpExp;
  sh.getRange(r, 1, 1, 3)
    .setValues([['NET INCOME', '', netIncome]])
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground(netIncome >= 0 ? '#1B5E20' : '#B71C1C')
    .setFontColor('#FFFFFF');
  sh.getRange(r, 3).setNumberFormat('$#,##0.00');
  r += 2;

  sh.getRange(r, 1).setValue(`Generated: ${fmtDate(new Date())}`).setFontStyle('italic').setFontColor('#757575');

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 20);
  sh.setColumnWidth(3, 140);
  sh.setFrozenRows(0);
  ss.setActiveSheet(sh);
  toast('P&L Report generated.', 'Reports');
}

// ---------------------------------------------------------------------------
// Balance Sheet
// ---------------------------------------------------------------------------

function generateBalanceSheet() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const accMap = buildAccountMap();
  const jData  = getSheetData(SHEETS.JOURNAL);

  // Compute balance for ALL accounts (cumulative, no date filter for B/S)
  const accountBalances = {};
  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const acctNum = String(row['Account #']);
    const acc     = accMap[acctNum];
    if (!acc) return;

    if (!accountBalances[acctNum]) {
      accountBalances[acctNum] = { name: acc.name, type: acc.type, subType: acc.subType, normal: acc.normal, net: 0 };
    }
    const dr = Number(row['Debit'])  || 0;
    const cr = Number(row['Credit']) || 0;
    if (acc.normal === 'Debit') {
      accountBalances[acctNum].net += dr - cr;
    } else {
      accountBalances[acctNum].net += cr - dr;
    }
  });

  let sh = ss.getSheetByName('Balance Sheet');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Balance Sheet');

  let r = 1;
  const write = (label, value, bold, bg) => {
    sh.getRange(r, 1).setValue(label);
    if (bold) sh.getRange(r, 1).setFontWeight('bold');
    if (bg)   sh.getRange(r, 1, 1, 3).setBackground(bg);
    if (value !== undefined && value !== null) {
      sh.getRange(r, 3).setValue(value).setNumberFormat('$#,##0.00');
      if (bold) sh.getRange(r, 3).setFontWeight('bold');
    }
    r++;
  };

  sh.getRange(r, 1).setValue('BALANCE SHEET').setFontSize(16).setFontWeight('bold').setFontColor('#1565C0');
  r++;
  sh.getRange(r, 1).setValue(`As of: ${fmtDate(new Date())}`).setFontStyle('italic');
  r += 2;

  const sectionTotals = {};
  const writeSection = (label, types, subTypes, color, totLabel) => {
    write(label, null, true, color);
    let total = 0;
    Object.entries(accountBalances)
      .filter(([, v]) => types.includes(v.type) && (!subTypes || subTypes.includes(v.subType)))
      .sort(([a], [b]) => a.localeCompare(b))
      .forEach(([num, v]) => {
        write(`  ${num} - ${v.name}`, v.net);
        total += v.net;
      });
    write(totLabel || `Total ${label}`, total, true, color);
    sectionTotals[label] = total;
    r++;
    return total;
  };

  // Assets
  const currentAssets = writeSection('CURRENT ASSETS',  ['Asset'], ['Current Asset'], '#E3F2FD', 'Total Current Assets');
  const fixedAssets   = writeSection('FIXED ASSETS',    ['Asset'], ['Fixed Asset'],   '#BBDEFB', 'Total Fixed Assets');
  const otherAssets   = writeSection('OTHER ASSETS',    ['Asset'], ['Other Asset'],   '#90CAF9', 'Total Other Assets');
  const totalAssets   = currentAssets + fixedAssets + otherAssets;
  write('TOTAL ASSETS', totalAssets, true, '#1565C0');
  sh.getRange(r - 1, 1, 1, 3).setFontColor('#FFFFFF');
  r += 2;

  // Liabilities
  const curLiab   = writeSection('CURRENT LIABILITIES', ['Liability'], ['Current Liability'], '#FFF3E0', 'Total Current Liabilities');
  const ltLiab    = writeSection('LONG-TERM LIABILITIES', ['Liability'], ['Long-Term'],       '#FFE0B2', 'Total Long-Term Liabilities');
  const totalLiab = curLiab + ltLiab;
  write('TOTAL LIABILITIES', totalLiab, true, '#E65100');
  sh.getRange(r - 1, 1, 1, 3).setFontColor('#FFFFFF');
  r += 2;

  // Equity
  const totalEquity = writeSection('EQUITY', ['Equity'], null, '#F3E5F5', 'Total Equity');
  r++;

  // Net income (current year P&L adds to retained earnings)
  const now     = new Date();
  const ytdFrom = new Date(now.getFullYear(), 0, 1);
  const ytdTo   = new Date(now.getFullYear(), 11, 31, 23, 59, 59);
  let netIncome = 0;
  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const txDate = parseDate(row['Date']);
    if (!txDate || txDate < ytdFrom || txDate > ytdTo) return;
    const acc = accMap[String(row['Account #'])];
    if (!acc) return;
    const dr = Number(row['Debit']) || 0;
    const cr = Number(row['Credit']) || 0;
    if (acc.type === 'Revenue') netIncome += cr - dr;
    if (acc.type === 'Expense') netIncome -= dr - cr;
  });
  write('  Current Year Net Income', netIncome, false, '#F3E5F5');
  const totalEquityAndNI = totalEquity + netIncome;
  write('TOTAL EQUITY', totalEquityAndNI, true, '#7B1FA2');
  sh.getRange(r - 1, 1, 1, 3).setFontColor('#FFFFFF');
  r += 2;

  const totalLiabEquity = totalLiab + totalEquityAndNI;
  write('TOTAL LIABILITIES & EQUITY', totalLiabEquity, true, '#212121');
  sh.getRange(r - 1, 1, 1, 3).setFontColor('#FFFFFF');
  r++;

  const diff = totalAssets - totalLiabEquity;
  write(`Check (Assets - L&E): ${diff.toFixed(2)}`, null, false, Math.abs(diff) < 0.005 ? '#C8E6C9' : '#FFCDD2');
  r++;
  sh.getRange(r, 1).setValue(`Generated: ${fmtDate(new Date())}`).setFontStyle('italic').setFontColor('#757575');

  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 20);
  sh.setColumnWidth(3, 150);
  ss.setActiveSheet(sh);
  toast('Balance Sheet generated.', 'Reports');
}

// ---------------------------------------------------------------------------
// Cash Flow Statement (indirect method)
// ---------------------------------------------------------------------------

function generateCashFlow() {
  const ui = SpreadsheetApp.getUi();

  const fromResp = ui.prompt('Cash Flow', 'Start Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  const toResp = ui.prompt('Cash Flow', 'End Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (toResp.getSelectedButton() !== ui.Button.OK) return;

  const fromDate = new Date(fromResp.getResponseText());
  const toDate   = new Date(toResp.getResponseText());
  toDate.setHours(23, 59, 59);

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const accMap = buildAccountMap();
  const jData  = getSheetData(SHEETS.JOURNAL);

  // Cash flows = changes in cash account (1000)
  let cashInflows  = 0;
  let cashOutflows = 0;
  const inFlowDetails  = [];
  const outFlowDetails = [];

  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const txDate = parseDate(row['Date']);
    if (!txDate || txDate < fromDate || txDate > toDate) return;
    if (String(row['Account #']) !== '1000') return;

    const dr = Number(row['Debit'])  || 0;
    const cr = Number(row['Credit']) || 0;
    if (dr > 0) { cashInflows  += dr; inFlowDetails.push([row['Date'], row['Description'], dr]); }
    if (cr > 0) { cashOutflows += cr; outFlowDetails.push([row['Date'], row['Description'], cr]); }
  });

  let sh = ss.getSheetByName('Cash Flow Statement');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Cash Flow Statement');

  let r = 1;
  const write = (label, value, bold, bg) => {
    sh.getRange(r, 1).setValue(label);
    if (bold) sh.getRange(r, 1).setFontWeight('bold');
    if (bg)   sh.getRange(r, 1, 1, 3).setBackground(bg);
    if (value !== undefined && value !== null) {
      sh.getRange(r, 3).setValue(value).setNumberFormat('$#,##0.00');
      if (bold) sh.getRange(r, 3).setFontWeight('bold');
    }
    r++;
  };

  sh.getRange(r, 1).setValue('CASH FLOW STATEMENT').setFontSize(16).setFontWeight('bold').setFontColor('#004D40');
  r++;
  sh.getRange(r, 1).setValue(`Period: ${fmtDate(fromDate)} to ${fmtDate(toDate)}`).setFontStyle('italic');
  r += 2;

  write('CASH INFLOWS', null, true, '#E0F2F1');
  inFlowDetails.forEach(d => write(`  ${d[0]}  ${d[1]}`, d[2]));
  write('Total Cash Inflows', cashInflows, true, '#B2DFDB');
  r++;

  write('CASH OUTFLOWS', null, true, '#FFF3E0');
  outFlowDetails.forEach(d => write(`  ${d[0]}  ${d[1]}`, d[2]));
  write('Total Cash Outflows', cashOutflows, true, '#FFE0B2');
  r++;

  const netCashFlow = cashInflows - cashOutflows;
  write('NET CASH FLOW', netCashFlow, true, netCashFlow >= 0 ? '#004D40' : '#B71C1C');
  sh.getRange(r - 1, 1, 1, 3).setFontColor('#FFFFFF');
  r++;

  const openingCash = 0; // Would need to be calculated from prior periods
  write('Opening Cash Balance', openingCash);
  write('Closing Cash Balance', openingCash + netCashFlow, true, '#E0F2F1');
  r++;
  sh.getRange(r, 1).setValue(`Generated: ${fmtDate(new Date())}`).setFontStyle('italic').setFontColor('#757575');

  sh.setColumnWidth(1, 350);
  sh.setColumnWidth(2, 20);
  sh.setColumnWidth(3, 150);
  ss.setActiveSheet(sh);
  toast('Cash Flow Statement generated.', 'Reports');
}

// ---------------------------------------------------------------------------
// Department Report
// ---------------------------------------------------------------------------

function generateDepartmentReport() {
  const ui = SpreadsheetApp.getUi();

  const fromResp = ui.prompt('Department Report', 'Start Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  const toResp = ui.prompt('Department Report', 'End Date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (toResp.getSelectedButton() !== ui.Button.OK) return;

  const fromDate = new Date(fromResp.getResponseText());
  const toDate   = new Date(toResp.getResponseText());
  toDate.setHours(23, 59, 59);

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const accMap  = buildAccountMap();
  const jData   = getSheetData(SHEETS.JOURNAL);
  const depts   = getSheetData(SHEETS.DEPARTMENTS);

  // Build dept map: code → name
  const deptNames = {};
  depts.forEach(d => { deptNames[d['Dept Code']] = d['Department Name']; });

  // Aggregate by dept → account
  const deptData = {};
  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const txDate = parseDate(row['Date']);
    if (!txDate || txDate < fromDate || txDate > toDate) return;

    const dept    = String(row['Department'] || 'UNALLOCATED');
    const acctNum = String(row['Account #']);
    const acc     = accMap[acctNum];
    if (!acc) return;
    if (!['Revenue', 'Expense'].includes(acc.type)) return;

    if (!deptData[dept]) deptData[dept] = {};
    if (!deptData[dept][acctNum]) {
      deptData[dept][acctNum] = { name: acc.name, type: acc.type, net: 0 };
    }
    const dr = Number(row['Debit'])  || 0;
    const cr = Number(row['Credit']) || 0;
    if (acc.type === 'Revenue') deptData[dept][acctNum].net += cr - dr;
    else                        deptData[dept][acctNum].net += dr - cr;
  });

  let sh = ss.getSheetByName('Department Report');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Department Report');

  let r = 1;
  sh.getRange(r, 1).setValue('DEPARTMENT REPORT').setFontSize(16).setFontWeight('bold').setFontColor('#37474F');
  r++;
  sh.getRange(r, 1).setValue(`Period: ${fmtDate(fromDate)} to ${fmtDate(toDate)}`).setFontStyle('italic');
  r += 2;

  const deptColors = ['#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC'];
  let colorIdx = 0;

  Object.entries(deptData).sort().forEach(([deptCode, accounts]) => {
    const deptName = deptNames[deptCode] || deptCode;
    const bg = deptColors[colorIdx % deptColors.length];
    colorIdx++;

    sh.getRange(r, 1).setValue(`DEPARTMENT: ${deptCode} - ${deptName}`)
      .setFontSize(12).setFontWeight('bold').setBackground(bg);
    sh.getRange(r, 2, 1, 2).setBackground(bg);
    r++;

    let deptRevenue = 0, deptExpense = 0;
    let hasRevenue = false, hasExpense = false;

    const revAccts = Object.entries(accounts).filter(([, v]) => v.type === 'Revenue').sort();
    const expAccts = Object.entries(accounts).filter(([, v]) => v.type === 'Expense').sort();

    if (revAccts.length > 0) {
      sh.getRange(r, 1).setValue('  Revenue').setFontWeight('bold').setFontColor('#2E7D32');
      r++;
      revAccts.forEach(([num, v]) => {
        sh.getRange(r, 1).setValue(`    ${num} - ${v.name}`);
        sh.getRange(r, 3).setValue(v.net).setNumberFormat('$#,##0.00');
        deptRevenue += v.net;
        r++;
      });
      sh.getRange(r, 1).setValue('  Total Revenue').setFontWeight('bold');
      sh.getRange(r, 3).setValue(deptRevenue).setNumberFormat('$#,##0.00').setFontWeight('bold');
      r++;
    }

    if (expAccts.length > 0) {
      sh.getRange(r, 1).setValue('  Expenses').setFontWeight('bold').setFontColor('#C62828');
      r++;
      expAccts.forEach(([num, v]) => {
        sh.getRange(r, 1).setValue(`    ${num} - ${v.name}`);
        sh.getRange(r, 3).setValue(v.net).setNumberFormat('$#,##0.00');
        deptExpense += v.net;
        r++;
      });
      sh.getRange(r, 1).setValue('  Total Expenses').setFontWeight('bold');
      sh.getRange(r, 3).setValue(deptExpense).setNumberFormat('$#,##0.00').setFontWeight('bold');
      r++;
    }

    const deptNI = deptRevenue - deptExpense;
    sh.getRange(r, 1).setValue(`  Net Income / (Loss)`).setFontWeight('bold');
    sh.getRange(r, 1, 1, 3).setBackground(deptNI >= 0 ? '#C8E6C9' : '#FFCDD2');
    sh.getRange(r, 3).setValue(deptNI).setNumberFormat('$#,##0.00').setFontWeight('bold');
    r += 2;
  });

  sh.getRange(r, 1).setValue(`Generated: ${fmtDate(new Date())}`).setFontStyle('italic').setFontColor('#757575');

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 20);
  sh.setColumnWidth(3, 150);
  ss.setActiveSheet(sh);
  toast('Department Report generated.', 'Reports');
}

// ---------------------------------------------------------------------------
// Trial Balance
// ---------------------------------------------------------------------------

function generateTrialBalance() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const accMap = buildAccountMap();
  const jData  = getSheetData(SHEETS.JOURNAL);

  // Sum debits and credits per account
  const totals = {};
  jData.forEach(row => {
    if (row['Posted'] !== 'Yes') return;
    const acctNum = String(row['Account #']);
    if (!totals[acctNum]) totals[acctNum] = { debits: 0, credits: 0 };
    totals[acctNum].debits  += Number(row['Debit'])  || 0;
    totals[acctNum].credits += Number(row['Credit']) || 0;
  });

  let sh = ss.getSheetByName('Trial Balance');
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet('Trial Balance');

  setHeaders(sh, ['Account #', 'Account Name', 'Type', 'Debit', 'Credit', 'Net Balance'], '#37474F');

  const rows = [];
  let totDr = 0, totCr = 0;

  Object.keys(accMap).sort().forEach(num => {
    const acc = accMap[num];
    const t   = totals[num] || { debits: 0, credits: 0 };
    if (t.debits === 0 && t.credits === 0) return; // skip zero accounts
    const net = acc.normal === 'Debit' ? t.debits - t.credits : t.credits - t.debits;
    rows.push([num, acc.name, acc.type, t.debits, t.credits, net]);
    totDr += t.debits;
    totCr += t.credits;
  });

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
    formatCurrencyColumns(sh, [4, 5, 6], 2, rows.length);
    applyAlternateRowColor(sh, 2, rows.length + 1, 6);
  }

  // Totals row
  const totRow = rows.length + 2;
  sh.getRange(totRow, 1).setValue('TOTALS').setFontWeight('bold');
  sh.getRange(totRow, 4, 1, 2).setValues([[totDr, totCr]]).setFontWeight('bold').setNumberFormat('$#,##0.00');
  sh.getRange(totRow, 1, 1, 6).setBackground(Math.abs(totDr - totCr) < 0.005 ? '#C8E6C9' : '#FFCDD2');

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 6);
  ss.setActiveSheet(sh);
  toast('Trial Balance generated.', 'Reports');
}
