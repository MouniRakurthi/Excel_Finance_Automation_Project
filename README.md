# Finance Operations Automation Project

Excel-based automation system for a mid-size company's finance operations — reconciliation, AP/AR management, compliance, risk scoring, procurement analysis, forecasting, and one-click recurring reporting, built entirely in native Excel formulas + VBA (no Power Query, no XLOOKUP, compatible with older non-365 Excel builds).

## The Problem

Finance teams often rely on manual, disconnected processes to reconcile transactions, track payables/receivables, monitor compliance, and assess risk — leading to delayed exception detection and no standardized way to catch fraud-adjacent patterns like duplicate payments, segregation-of-duty violations, or unmatched receipts.

**This project answers:** *How can a finance operations team use Excel to detect discrepancies, monitor payables and receivables, assess compliance and risk, and generate accurate recurring reports — with minimal manual effort?*

## What's Inside

| Module | What it does |
|---|---|
| **3-Way Match** | Matches PO ↔ Goods Receipt ↔ Invoice, flags mismatches and goods invoiced without receipt |
| **Bank Reconciliation** | Matches AP payments to bank statement, flags unmatched items and unexplained bank activity |
| **AP / AR Aging** | Buckets outstanding invoices 0-30 / 31-60 / 61-90 / 90+ days |
| **Duplicate Payment Detection** | Flags invoices paid more than once |
| **DPO / DSO** | Days Payable/Sales Outstanding, headline metric + monthly trend |
| **Segregation of Duties** | Flags POs where requester = approver |
| **Risk & Control Matrix (RCM)** | Internal-audit style pass/fail panel, live-linked to every module |
| **AML-Style Risk Scoring** | Weighted score per transaction (large amount, round-dollar, new vendor, geography) |
| **Vendor Spend / Pareto Analysis** | Which vendors drive 80% of spend |
| **New Vendor Risk Flagging** | Recently onboarded vendors with unusually large orders |
| **Cash Outflow Forecast** | 3-month forward AP forecast using linear regression |
| **Recurring Reports** | Weekly / Monthly / Quarterly / Yearly summaries, formula-driven |
| **KPI Dashboard** | One-page snapshot with headline KPIs, aging charts, DPO/DSO trend, top vendors, live audit status |
| **VBA Automation** | One-click macro: refreshes every formula and exports the reports + dashboard as a dated PDF |

## Tech Stack

- Excel (formulas only — SUMIFS, INDEX/MATCH, COUNTIFS, IFERROR — no Power Query, no XLOOKUP, no dynamic arrays, for compatibility with older Excel builds)
- VBA for one-click refresh + PDF export automation
- Windows Task Scheduler pattern documented for unattended scheduling

## Repo Contents

- `Finance_Automation_Workbook.xlsx` — the full workbook (32 tabs: raw data, every module, recurring reports, KPI dashboard)
- `Finance_Automation_Macros.bas` — VBA module (import into a `.xlsm` copy to enable the one-click button)
- `Project_Documentation.pdf` — full write-up: business problem, methodology per module, assumptions, limitations
- `How_To_Run_Guide.pdf` — step-by-step setup: importing the macro, adding a button, Task Scheduler

## How It Works

Everything lives in one workbook so every module recalculates together — no external file links. Tabs are color-coded by function (Reconciliation, AP, AR, Compliance, Risk, Procurement, Forecasting, Reporting). All calculations are live formulas, not pasted values, so the whole workbook re-derives itself the moment the raw data or the editable Assumptions tab (as-of date, risk thresholds, tolerances) changes.

The VBA macro (`RunAllReports`) recalculates every formula, refreshes any PivotTables, and exports the four recurring reports + KPI dashboard as one dated PDF — ready to save, share, or attach manually.

## Known Limitations

- Risk-scoring weights and the geography watch-list are illustrative starting points, not a validated AML model
- Reporting tabs are formula-built summaries rather than native Excel PivotTables, by design, for older-Excel compatibility
- Built on a static one-time data extract; production use would connect the raw-data tabs to a live source

---

*Built as a portfolio project demonstrating Excel-based financial operations automation, internal controls design, and VBA scripting.*
