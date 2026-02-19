# Google Sheets Accounting App

A full double-entry bookkeeping system built with Google Sheets + Google Apps Script.

## Features

| Module | Capabilities |
|---|---|
| **Chart of Accounts** | Add/manage accounts (Assets, Liabilities, Equity, Revenue, Expenses) |
| **Journal** | Double-entry journal entries with auto-balancing validation |
| **General Ledger** | Auto-generated from posted journal entries |
| **AR — Invoicing** | Create invoices, record payments, track outstanding balances |
| **AP — Bills** | Create vendor bills, record payments, track payables |
| **Bank Reconciliation** | Import CSV bank statements, auto-match to journal entries |
| **Reports** | P&L, Balance Sheet, Cash Flow, Trial Balance, AR/AP Aging, Department |
| **Departments** | Allocate transactions by department with department-level P&L |

---

## Getting Started

### Option A — Using clasp (recommended for developers)

1. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
2. Log in: `clasp login`
3. Create a new Google Apps Script project linked to a Spreadsheet:
   ```
   clasp create --type sheets --title "Accounting App"
   ```
4. Update `.clasp.json` with the generated `scriptId`
5. Push the code: `clasp push`
6. Open the spreadsheet → **Extensions → Apps Script** to verify
7. Reload the spreadsheet — the **📊 Accounting** menu will appear

### Option B — Manual copy-paste

1. Create a new Google Spreadsheet
2. Go to **Extensions → Apps Script**
3. Create one `.gs` file per source file in `src/`
4. Paste the contents of each file
5. Save and reload the spreadsheet

### First Run

1. Open the spreadsheet
2. Click **📊 Accounting → 🔧 Setup Workbook (first time)**
3. Approve the required permissions
4. All sheets will be created with sample data

---

## Sheet Structure

```
├── Dashboard              — Key metrics overview
├── Chart of Accounts      — Account master list
├── Journal                — All posted journal entries
├── General Ledger         — Account balances (auto-refreshed)
├── Customers              — Customer master
├── Vendors                — Vendor master
├── Departments            — Department list
├── AR - Invoices          — Customer invoices
├── AR - Payments          — Customer payments received
├── AP - Bills             — Vendor bills
├── AP - Payments          — Vendor payments made
├── Bank Statement         — Imported bank transactions
└── Reconciliation         — Bank reconciliation worksheet
```

Generated report sheets (created on demand):
- `P&L Report`, `Balance Sheet`, `Cash Flow Statement`
- `Trial Balance`, `AR Aging Report`, `AP Aging Report`, `Department Report`

---

## Account Numbering Convention

| Range | Type |
|---|---|
| 1000–1999 | Assets |
| 2000–2999 | Liabilities |
| 3000–3999 | Equity |
| 4000–4999 | Revenue |
| 5000–5999 | Expenses |

---

## Workflow

### Daily Transactions
1. **Invoices**: Accounting → AR → Create Invoice
   → Auto-posts a Debit AR / Credit Revenue journal entry
2. **Customer Payments**: Accounting → AR → Record Payment Received
   → Auto-posts Debit Cash / Credit AR
3. **Vendor Bills**: Accounting → AP → Create Vendor Bill
   → Auto-posts Debit Expense / Credit AP
4. **Vendor Payments**: Accounting → AP → Record Vendor Payment
   → Auto-posts Debit AP / Credit Cash
5. **Manual Entries**: Accounting → Journal → New Journal Entry
   → Multi-line double-entry form with debit = credit validation

### Month-End Close
1. **Refresh statuses**: Accounting → Utilities → Refresh All Statuses
2. **Bank Recon**: Import bank CSV → Auto-match → Run Reconciliation
3. **Reports**: Generate P&L, Balance Sheet, Trial Balance

### Optional: Daily Auto-Refresh
Run `installDailyTrigger()` once from the Apps Script editor to automatically refresh AR/AP statuses and the GL every morning at 8:00 AM.

---

## Bank Statement CSV Format

```
Date,Description,Reference,Debit,Credit
2024-01-15,Customer Deposit,,0,5000.00
2024-01-16,Office Supplies - Check #1001,CHK-1001,250.00,0
2024-01-17,Utility Payment - ACH,ACH-045,180.00,0
```

- **Date**: `YYYY-MM-DD`
- **Debit**: money leaving the account (positive number)
- **Credit**: money entering the account (positive number)

---

## File Structure

```
src/
├── Code.gs                — Menu, onOpen, onEdit, trigger installer
├── 00_Utilities.gs        — Shared constants & helper functions
├── 01_Setup.gs            — Sheet initialization & sample data
├── 02_ChartOfAccounts.gs  — COA management & General Ledger refresh
├── 03_Journal.gs          — Double-entry journal entries
├── 04_AR.gs               — Invoicing & AR payments
├── 05_AP.gs               — Vendor bills & AP payments
├── 06_BankReconciliation.gs — Bank statement import & matching
└── 07_Reports.gs          — P&L, Balance Sheet, Cash Flow, Dept reports
appsscript.json            — Apps Script manifest
.clasp.json                — clasp configuration (add your scriptId)
```

---

## License

MIT — free to use, modify, and distribute.
