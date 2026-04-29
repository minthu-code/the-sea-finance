# ExhibitLedger THB Bot — Perfection Pass Validation Summary

Author: **Manus AI**  
Validation date: **2026-04-28**

## Scope validated

This validation covers the refined Telegram bot after the full workflow improvement pass. The bot remains **Telegram-based** and **THB-only**, but now includes a guided button menu, conversational data-entry states, stronger exhibition finance controls, sale collection tracking, inventory KPIs, account-head budgets, readiness checks, and a richer Excel workbook export.

| Area | Validation result |
|---|---|
| Python syntax | Passed for `main.py`, `exhibitledger.py`, `local_workflow_test.py`, and `local_report_test.py`. |
| New workflow regression | Passed through exhibition creation, split setup, artwork registration, partial sale collection, receipt approval, budget setup, P&L totals, readiness report, and Excel export. |
| Existing Shwe Dagon sample | Passed existing report generation and export using the previous seed/sample structure. |
| Expanded workbook export | Verified sheet coverage: Executive Summary, P&L, Inventory, Artist Payables, Confirmed Expenses, Budget vs Actual, Pending Receipts, Sales Allocations, Data Quality, and Audit Log. |
| Data-quality refinement | Corrected a false-positive split warning so collaborator-inclusive 45/10/45 sales reconcile properly. |

## Key workflow test figures

The expanded workflow regression test used a temporary SQLite database at `/tmp/exhibitledger_workflow_test.db`. It created a test exhibition with two registered artworks, one partial sale, two approved receipts, and two expense budgets. The sale used a **45% gallery / 10% collaborator / 45% artist** split.

| Metric | Expected and validated value |
|---|---:|
| Gross artwork sales | ฿100,000.00 THB |
| Gallery revenue | ฿45,000.00 THB |
| Artist payable | ฿45,000.00 THB |
| Collaborator share | ฿10,000.00 THB |
| Cash collected | ฿60,000.00 THB |
| Sale receivable | ฿40,000.00 THB |
| Direct costs | ฿57,500.00 THB |
| Operating expenses | ฿3,500.00 THB |
| Net profit / loss | ฿-16,000.00 THB |
| Inventory sell-through | 50.0% |
| Unsold asking value | ฿80,000.00 THB |

## Validation commands run

```bash
cd /home/ubuntu/exhibitledger_thb_bot
python3.11 -m py_compile main.py exhibitledger.py local_workflow_test.py local_report_test.py
python3.11 local_workflow_test.py > /tmp/exhibitledger_workflow_test_output.txt 2>&1
python3.11 local_report_test.py --code SHWEDAGON2024 --export > /tmp/exhibitledger_local_report_output.txt 2>&1
```

## Generated validation artifacts

| Artifact | Purpose |
|---|---|
| `/tmp/exhibitledger_workflow_test_output.txt` | Full new workflow regression output. |
| `/tmp/exhibitledger_local_report_output.txt` | Backward-compatibility output for the existing Shwe Dagon sample. |
| `/tmp/exhibitledger_workbook_sheets.txt` | Expanded Excel workbook sheet-name verification. |
| `/tmp/exhibitledger_workflow_exports/TEST2026_pnl_report.xlsx` | Generated workbook from the new workflow regression test. |

## Notes

The existing Shwe Dagon sample still reports warnings that are appropriate for legacy seed data: it has imported non-THB source values, does not yet have an active split rule, and does not yet have registered artwork inventory rows. These are not validation failures; they are now visible as management review prompts so future exhibitions can be controlled more tightly from the start.
