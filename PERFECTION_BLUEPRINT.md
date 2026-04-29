# ExhibitLedger THB Bot — Refinement Blueprint

Author: **Manus AI**
Date: **2026-04-28**

## Refinement Objective

The current bot is stable and already supports exhibition-level THB P&L reporting, commission splits, artwork registration, sale allocation, receipt capture, approval, expense reports, and Excel export. This refinement pass upgrades it from a command-driven accounting helper into a **guided gallery finance assistant**. The priority is to reduce typing, prevent accounting mistakes, and make the P&L file more useful for daily operation and final management review.

## Practical Gallery Workflow to Add

The improved bot should open with a **button-based main menu**. Each major area should redirect into a submenu and, where useful, ask follow-up questions in normal chat. Commands must remain available for power users and backward compatibility, but the preferred workflow should be: tap a menu, choose an action, then answer one question at a time.

| Area | Button workflow | Practical outcome |
|---|---|---|
| Exhibition | Current exhibition, list previous exhibitions, add new exhibition, switch exhibition, readiness check | User always knows which ledger they are posting into. |
| Splits | View split, common presets, custom split entry | Sale allocation is confirmed before artwork sales are recorded. |
| Artworks & Sales | Register artwork, list available works, record sale, view inventory dashboard | Inventory, sell-through, revenue, artist payable, and receivables are controlled. |
| Receipts & Expenses | Add text receipt, send photo receipt, pending approvals, account heads, expense report | Expenses are staged first and only hit P&L after confirmation. |
| Reports & Export | Executive dashboard, P&L, artist payouts, budget report, data check, Excel export | The user can review financial health without memorizing commands. |
| Help & Settings | Command list, current exhibition, Google Sheets preview status | The bot is self-explanatory during live use. |

## Accounting and P&L Enhancements

The existing P&L is clear but can be strengthened with management controls that galleries usually need during an exhibition. The enhancements should not change the THB-only principle, and they should not post pending receipts into final P&L until approval.

| Enhancement | Why it matters | Implementation approach |
|---|---|---|
| Inventory KPIs | P&L alone does not show unsold stock or sell-through. | Add dashboard functions for total artworks, sold artworks, unsold asking value, sell-through rate, average sale price, gross sales, cash collected, receivables, and pending expense count. |
| Sale payment status | A sale is not always fully collected immediately. | Add migration-safe columns to `artwork_sales` for buyer name, payment status, collected amount, balance due, payment method, and notes. Default old sales to fully collected. |
| Budget vs actual | The user may need to compare venue, freight, marketing, and hospitality costs against expectation. | Add `expense_budgets` table by exhibition and account head, plus formatter and Excel sheet. |
| Richer account heads | The previous heads were good, but travel, accommodation, insurance, tax/duty, and repairs are common exhibition costs. | Expand `ACCOUNT_HEADS` while preserving existing names. |
| Better P&L summary | The existing P&L is accountant-facing, but needs a top management snapshot. | Add dashboard/report functions and add Excel sheets for executive summary, KPI dashboard, inventory, budget variance, pending receipts, data quality, and audit trail. |
| Readiness check | Before final export, the bot should detect missing splits, unsold inventory, pending receipts, uncollected sale balances, negative costs, and weak data. | Extend `data_quality_checks()` with inventory, receipt, and receivable checks. |

## Conversational Flow Strategy

The bot can be improved without introducing a complex Telegram conversation framework. A lightweight `context.user_data["flow"]` state can store the current guided action and the step being awaited. This keeps the implementation predictable and easy to test locally.

| Guided flow | User path | Input format accepted |
|---|---|---|
| New exhibition | Main Menu → Exhibition → Add New Exhibition | `CODE | Name | Location | Start date | End date | Notes`; only `CODE | Name` required. |
| Custom split | Main Menu → Splits → Custom Split | `gallery 45 collaborator Curator 10 artist 45` or `gallery 50 artist 50`. |
| Add artwork | Main Menu → Artworks & Sales → Register Artwork | `Title | Artist | Asking Price THB`. |
| Record sale | Main Menu → Artworks & Sales → Record Sale | `Artwork ID | Sale Price | Buyer | Collected Amount | Payment Method | Notes`; first two fields required. |
| Add text receipt | Main Menu → Receipts & Expenses → Add Text Receipt | `Amount Description`; creates pending approval card. |
| Set budget | Reports or Expenses → Set Budget | `Account Head | Amount THB` or `Number | Amount THB`. |

## Backward Compatibility Rules

All existing slash commands must still work. Existing validation scripts must continue passing. Legacy Shwe Dagon seed data must still generate reports, even though its early prototype assumptions differ from the newer sale-allocation workflow. The new functionality should add safe tables/columns through `CREATE TABLE IF NOT EXISTS` and `ALTER TABLE` checks instead of destructive migrations.

## Deliverables for This Pass

This pass should update `exhibitledger.py`, `main.py`, `README.md`, and the local regression tests. The final package should include a clean ZIP, a validation summary, and documentation explaining both the new guided menu and the raw command fallback.
