# Financial Statement Auditing Through VBA

> **Audit Smarter, Not Harder** - Complete VBA toolkit to audit financial statements in accordance with **GAAS** and **GAAP**

---

## Overview

This section provides ready-to-use VBA macros that automate audit procedures for every major financial statement line item. Each module is designed to:

- Accept **standardized inputs** (GL Detail + Sub-Ledger)
- Perform **substantive testing** per GAAS requirements
- Test **management assertions** (Existence, Completeness, Valuation, Rights, Presentation)
- Generate **exception reports** for follow-up
- Document **audit evidence** automatically

---

## Quick Navigation

### Balance Sheet Accounts

| Account | Key Audit Procedures | Assertions Tested |
|---------|---------------------|-------------------|
| [**Cash & Cash Equivalents**](./balance-sheet/cash.md) | Bank reconciliation, cutoff, confirmation | Existence, Completeness, Valuation |
| [**Accounts Receivable**](./balance-sheet/accounts-receivable.md) | Aging analysis, confirmations, subsequent receipts, allowance | Existence, Valuation, Completeness |
| [**Inventory**](./balance-sheet/inventory.md) | Test counts, cost testing, NRV, obsolescence | Existence, Valuation, Completeness |
| [**Prepaid Expenses**](./balance-sheet/prepaids.md) | Recalculation, amortization, support | Existence, Valuation, Accuracy |
| [**Property, Plant & Equipment**](./balance-sheet/ppe.md) | Additions, disposals, depreciation, impairment | Existence, Valuation, Completeness |
| [**Intangible Assets**](./balance-sheet/intangibles.md) | Amortization, impairment testing | Valuation, Existence |
| [**Accounts Payable**](./balance-sheet/accounts-payable.md) | Search for unrecorded liabilities, cutoff, confirmations | Completeness, Existence, Accuracy |
| [**Accrued Expenses**](./balance-sheet/accrued-expenses.md) | Recalculation, subsequent payments, reasonableness | Completeness, Valuation, Accuracy |
| [**Debt**](./balance-sheet/debt.md) | Confirmation, recalculation, covenant testing | Existence, Completeness, Valuation |
| [**Equity**](./balance-sheet/equity.md) | Roll-forward, authorization, dividends | Existence, Completeness, Presentation |

### Income Statement Accounts

| Account | Key Audit Procedures | Assertions Tested |
|---------|---------------------|-------------------|
| [**Revenue**](./income-statement/revenue.md) | Cutoff testing, analytical procedures, detail testing | Occurrence, Completeness, Accuracy |
| [**Cost of Goods Sold**](./income-statement/cogs.md) | Gross margin analysis, inventory tie-out, cost testing | Accuracy, Completeness, Cutoff |
| [**Operating Expenses**](./income-statement/operating-expenses.md) | Analytical procedures, vouching, search for unusual | Occurrence, Accuracy, Classification |
| [**Payroll Expense**](./income-statement/payroll.md) | Analytical review, recalculation, authorization | Occurrence, Accuracy, Completeness |
| [**Depreciation & Amortization**](./income-statement/depreciation.md) | Recalculation, policy consistency, useful lives | Accuracy, Consistency |
| [**Interest Expense**](./income-statement/interest.md) | Recalculation, debt agreement tie-out | Accuracy, Completeness |
| [**Income Tax Expense**](./income-statement/income-tax.md) | Provision recalculation, rate analysis, deferred taxes | Accuracy, Valuation |

---

## Standardized Input Formats

### GL Detail Format (Required for all modules)

Your General Ledger detail export should have these columns:

| Column | Header Name | Description | Format |
|--------|-------------|-------------|--------|
| A | `Date` | Transaction date | Date |
| B | `JE_Number` | Journal entry reference | Text |
| C | `Account` | GL account number | Text |
| D | `Account_Name` | GL account description | Text |
| E | `Description` | Transaction description | Text |
| F | `Debit` | Debit amount | Number |
| G | `Credit` | Credit amount | Number |
| H | `Source` | Source module (AP, AR, etc.) | Text |

### Sub-Ledger Formats

Each account section specifies its required sub-ledger format. Common formats:

**Accounts Receivable Sub-Ledger:**
| Column | Header | Description |
|--------|--------|-------------|
| A | `Customer_ID` | Unique customer identifier |
| B | `Customer_Name` | Customer name |
| C | `Invoice_Number` | Invoice reference |
| D | `Invoice_Date` | Date of invoice |
| E | `Due_Date` | Payment due date |
| F | `Amount` | Invoice amount |
| G | `Paid_Amount` | Amount collected |
| H | `Balance` | Outstanding balance |

**Accounts Payable Sub-Ledger:**
| Column | Header | Description |
|--------|--------|-------------|
| A | `Vendor_ID` | Unique vendor identifier |
| B | `Vendor_Name` | Vendor name |
| C | `Invoice_Number` | Vendor invoice reference |
| D | `Invoice_Date` | Date of invoice |
| E | `Due_Date` | Payment due date |
| F | `Amount` | Invoice amount |
| G | `Paid_Amount` | Amount paid |
| H | `Balance` | Outstanding balance |

---

## Management Assertions (GAAS)

Every audit procedure tests one or more assertions:

### Balance Sheet Assertions
| Assertion | What It Means | How We Test |
|-----------|---------------|-------------|
| **Existence** | Assets/liabilities actually exist | Confirmations, physical inspection, subsequent receipts |
| **Completeness** | All items are recorded | Search for unrecorded items, cutoff testing |
| **Valuation** | Recorded at correct amounts | Recalculation, market comparison, impairment testing |
| **Rights & Obligations** | Company owns assets / owes liabilities | Title review, debt agreements, legal confirmations |
| **Presentation** | Properly classified and disclosed | Review classifications, note disclosures |

### Income Statement Assertions
| Assertion | What It Means | How We Test |
|-----------|---------------|-------------|
| **Occurrence** | Transactions actually happened | Vouching to support, confirmations |
| **Completeness** | All transactions recorded | Analytical procedures, cutoff testing |
| **Accuracy** | Amounts are mathematically correct | Recalculation, footing, extensions |
| **Cutoff** | In correct period | Test transactions around period end |
| **Classification** | In correct accounts | Review coding, unusual items |

---

## How to Use These Audit Modules

### Step 1: Prepare Your Data
1. Export GL detail to Excel in the standard format above
2. Export relevant sub-ledger (AR aging, AP aging, inventory listing, etc.)
3. Place each on its own worksheet with headers in Row 1

### Step 2: Name Your Worksheets
The VBA expects these worksheet names:
- `GL_Detail` - General ledger detail
- `AR_Aging` - Accounts receivable sub-ledger
- `AP_Aging` - Accounts payable sub-ledger
- `Inventory_Listing` - Inventory detail
- `Fixed_Assets` - PP&E sub-ledger
- `Bank_Statements` - Bank reconciliation support

### Step 3: Run the Audit Module
1. Press **Alt+F11** to open VBA Editor
2. Insert a new Module
3. Copy/paste the relevant audit code
4. Press **F5** to run
5. Review the generated audit report

### Step 4: Review Exceptions
Each module creates an exception report highlighting:
- Items exceeding materiality thresholds
- Reconciling differences
- Potential misstatements
- Items requiring additional procedures

---

## Audit Workflow

```
┌─────────────────────────────────────────────────────────────┐
│                    AUDIT WORKFLOW                           │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. PLANNING                                                │
│     ├── Set materiality thresholds                          │
│     ├── Identify significant accounts                       │
│     └── Determine sample sizes                              │
│                                                             │
│  2. DATA PREPARATION                                        │
│     ├── Export GL detail (standard format)                  │
│     ├── Export sub-ledgers                                  │
│     └── Obtain supporting documents                         │
│                                                             │
│  3. RUN AUDIT MODULES                                       │
│     ├── Balance Sheet accounts                              │
│     │   ├── Cash                                            │
│     │   ├── Receivables                                     │
│     │   ├── Inventory                                       │
│     │   ├── Fixed Assets                                    │
│     │   ├── Payables                                        │
│     │   ├── Accruals                                        │
│     │   └── Debt & Equity                                   │
│     │                                                       │
│     └── Income Statement accounts                           │
│         ├── Revenue                                         │
│         ├── COGS                                            │
│         ├── Operating Expenses                              │
│         └── Other Income/Expense                            │
│                                                             │
│  4. REVIEW EXCEPTIONS                                       │
│     ├── Investigate flagged items                           │
│     ├── Request client explanations                         │
│     └── Propose adjustments if needed                       │
│                                                             │
│  5. CONCLUDE                                                │
│     ├── Summarize findings                                  │
│     ├── Document conclusions                                │
│     └── Generate audit report                               │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Setting Materiality

Each module uses configurable materiality thresholds:

```vba
' Set these values at the top of each module
Const MATERIALITY As Double = 50000        ' Overall materiality
Const TRIVIAL_THRESHOLD As Double = 2500   ' Below this = trivial
Const PERFORMANCE_MAT As Double = 37500    ' Performance materiality (75% of overall)
```

**Common Materiality Benchmarks:**
| Benchmark | Typical Range | Example ($10M Revenue) |
|-----------|---------------|------------------------|
| Revenue | 0.5% - 1% | $50,000 - $100,000 |
| Total Assets | 0.5% - 1% | $50,000 - $100,000 |
| Net Income | 5% - 10% | Variable |
| Equity | 1% - 2% | Variable |

---

## Sample Size Determination

The modules include sample size calculators based on:

```vba
' Factors affecting sample size
Const CONFIDENCE_LEVEL As Double = 0.95    ' 95% confidence
Const TOLERABLE_ERROR As Double = 0.05     ' 5% tolerable misstatement
Const EXPECTED_ERROR As Double = 0.01      ' 1% expected misstatement

' Sample size formula (simplified)
' n = (Confidence Factor * Population Value) / Tolerable Misstatement
```

---

## Quick Start: Run Full Balance Sheet Audit

```vba
Sub RunFullBalanceSheetAudit()
    '================================================
    ' Master Audit Controller - Balance Sheet
    ' Runs all balance sheet audit modules
    '================================================

    Dim startTime As Double
    startTime = Timer

    MsgBox "Starting Full Balance Sheet Audit..." & vbCrLf & _
           "Ensure all required worksheets are prepared.", vbInformation

    ' Run each module
    Call AuditCash
    Call AuditAccountsReceivable
    Call AuditInventory
    Call AuditPrepaids
    Call AuditFixedAssets
    Call AuditAccountsPayable
    Call AuditAccruedExpenses
    Call AuditDebt
    Call AuditEquity

    MsgBox "Balance Sheet Audit Complete!" & vbCrLf & _
           "Time: " & Round(Timer - startTime, 1) & " seconds" & vbCrLf & _
           "Review exception reports on each audit worksheet.", vbInformation

End Sub
```

---

## File Organization

```
financial-statement-auditing/
├── README.md (this file)
├── balance-sheet/
│   ├── cash.md
│   ├── accounts-receivable.md
│   ├── inventory.md
│   ├── prepaids.md
│   ├── ppe.md
│   ├── intangibles.md
│   ├── accounts-payable.md
│   ├── accrued-expenses.md
│   ├── debt.md
│   └── equity.md
└── income-statement/
    ├── revenue.md
    ├── cogs.md
    ├── operating-expenses.md
    ├── payroll.md
    ├── depreciation.md
    ├── interest.md
    └── income-tax.md
```

---

## Important Notes

### GAAS Compliance
These procedures are designed to comply with **Generally Accepted Auditing Standards** (GAAS), including:
- AU-C Section 330: Performing Audit Procedures
- AU-C Section 500: Audit Evidence
- AU-C Section 505: External Confirmations
- AU-C Section 520: Analytical Procedures
- AU-C Section 530: Audit Sampling

### GAAP Compliance
Procedures test for compliance with **Generally Accepted Accounting Principles** (GAAP), including:
- ASC 606: Revenue Recognition
- ASC 842: Leases
- ASC 330: Inventory
- ASC 360: Property, Plant & Equipment
- ASC 310: Receivables
- ASC 450: Contingencies

### Professional Judgment Required
These VBA modules are **tools to assist** the auditor, not replace professional judgment. Always:
- Apply professional skepticism
- Consider entity-specific risks
- Modify procedures as needed
- Document your conclusions

---

## Next Steps

1. **[Start with Cash](./balance-sheet/cash.md)** - The most straightforward account to audit
2. **[Accounts Receivable](./balance-sheet/accounts-receivable.md)** - High risk, detailed procedures
3. **[Revenue Testing](./income-statement/revenue.md)** - Critical income statement account

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
