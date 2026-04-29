# ExhibitLedger THB Telegram Bot Prototype

This is a practical Telegram bot prototype for **exhibition-by-exhibition Profit & Loss reporting** for The Sea Art Gallery. The first test case is the **Shwe Dagon Platform Exhibition**. The bot reports in **Thai Baht only** and is designed to become the foundation for a Google Sheets-connected finance assistant.

## What This Prototype Does

The prototype stores exhibition finance data in a local SQLite database, generates a THB-only P&L report, shows artist payable control, performs accountant-style data-quality checks, exports the report to Excel, and now includes read-only Google Sheets setup/preview commands. It remains intentionally read-only against Google Sheets, because live write-back should only be enabled after the report logic is approved.

| Feature | Current Prototype Status |
|---|---|
| Telegram commands | Included |
| Shwe Dagon seed data | Included |
| THB-only reporting | Included |
| Artist payable summary | Included |
| Excel export | Included |
| Google Sheets read-only preview | Added as safe scaffolding |
| Google Sheets live import/write-back | Not enabled yet; requires mapping approval |
| Confirmation-based sheet write-back | Next phase |

## Important Currency Note

Your Shwe Dagon PDFs are denominated in **MMK**, while the final bot must report in **THB**. For the current Shwe Dagon test run, the confirmed conversion is **150 MMK = 1 THB**, so the seed script should use `SEED_MMK_TO_THB_RATE=0.006666666666666667`. The regenerated prototype database and Excel export included with this package already use that rate.

If you later decide to use a different historical rate or verified THB amounts, reseed the database and regenerate the export before relying on the report for real financial decisions.

## Local Setup

Create and activate a Python environment if desired, then install dependencies.

```bash
cd /home/ubuntu/exhibitledger_thb_bot
pip install -r requirements.txt
```

Copy the environment template and set your values.

```bash
cp .env.example .env
```

For a first local test without Telegram, seed the Shwe Dagon data using the confirmed Shwe Dagon rate.

```bash
export SEED_MMK_TO_THB_RATE=0.006666666666666667
python seed_shwedagon.py
python local_report_test.py --code SHWEDAGON2024 --export
```

## Telegram Setup

Create a bot with Telegram **BotFather**, copy the token, and set it as an environment variable.

```bash
export TELEGRAM_BOT_TOKEN="your_botfather_token"
python main.py
```

Then open the bot in Telegram and run:

```text
/start
/exhibitions
/summary SHWEDAGON2024
/pl SHWEDAGON2024
/artist_payouts SHWEDAGON2024
/data_check SHWEDAGON2024
/export SHWEDAGON2024
/sheets_status
/sync_preview
```

## Deployment With Docker

Build and run the bot container.

```bash
docker build -t exhibitledger-thb-bot .
docker run --env TELEGRAM_BOT_TOKEN="your_botfather_token" \
  --env SEED_MMK_TO_THB_RATE="0.006666666666666667" \
  -v $(pwd)/data:/data \
  exhibitledger-thb-bot
```

Seed the database before running the worker in production. On Render, use a **Background Worker** service and add a persistent disk mounted at `/data`.

## Google Sheets Read-Only Preview

This version includes a safe read-only Google Sheets scaffold. The `/sheets_status` command checks whether the spreadsheet ID, service-account credential path, and mapping file are present. The `/sync_preview` command attempts to read a few rows from the mapped tabs and shows a dry-run preview only; it does not import, update, delete, or overwrite any sheet data.

To use it, create a Google service account, download its JSON credential file outside the project repository, share the finance workbook with the service-account email as a viewer, set `GOOGLE_SHEETS_SPREADSHEET_ID`, set `GOOGLE_APPLICATION_CREDENTIALS`, and adjust `sheets_mapping.example.json` into your approved live mapping file.

## Next Phase: Controlled Google Sheets Import

After the read-only preview looks correct, the next phase should convert preview rows into a controlled staging import. The recommended approach is service-account read access first, then a manual approval step before database import, and only much later confirmation-based sheet write-back.

| Needed From You | Purpose |
|---|---|
| Telegram BotFather token | Run the live bot. |
| Google Spreadsheet ID | Connect the bot to the finance workbook. |
| Google service-account JSON | Allow the bot to read sheets securely. |
| Confirmed conversion method | Current Shwe Dagon test uses 150 MMK = 1 THB; confirm any future exhibition rates separately. |
| Final column mapping approval | Avoid misclassifying cash book, commission, and stock rows. |

## Prototype Command List

| Command | Description |
|---|---|
| `/start` | Shows welcome and common commands. |
| `/exhibitions` | Lists available exhibitions. |
| `/summary SHWEDAGON2024` | Shows the accountant-facing executive summary and control points. |
| `/pl SHWEDAGON2024` | Generates the Shwe Dagon P&L. |
| `/artist_payouts SHWEDAGON2024` | Shows artist settlement summary. |
| `/data_check SHWEDAGON2024` | Shows warnings and validation notes. |
| `/export SHWEDAGON2024` | Sends an Excel report file. |
| `/sheets_status` | Checks the read-only Google Sheets setup. |
| `/sync_preview` | Previews mapped Google Sheet tabs without importing or writing. |


## Workflow Update: Receipt Approval, Commission Splits, Artworks, and Sales

This version keeps the original Telegram bot and P&L engine, but adds the real gallery workflow requested for exhibition-by-exhibition accounting. The system remains **THB-only**. Google Sheets remains optional and read-only; the priority workflow is now Telegram-first receipt capture, approval, and sale allocation.

### Recommended operating sequence

The practical finance workflow is to create the exhibition first, define the commission split before any sale is recorded, register the consigned artworks, capture expenses as pending receipts, approve those receipts under the correct account head, record sales, and then generate the exhibition P&L.

| Step | Telegram command or action | Accounting result |
|---|---|---|
| Create exhibition | `/new_exhibition BKK2026 Bangkok Art Fair 2026` | Creates a THB-only exhibition ledger. |
| Set working exhibition | `/use BKK2026` | Makes future receipt text/photo capture default to this exhibition. |
| Define split rule | `/set_split BKK2026 gallery 45 collaborator Curator 10 artist 45` | Stores the exhibition commission allocation. Percentages must total 100%. |
| Register artwork | `/add_artwork BKK2026 Quiet River \| Artist A \| 120000` | Adds an available consigned artwork. |
| Capture receipt | `/receipt BKK2026 3500 coffee and snacks opening night` | Creates a pending expense, not yet posted to P&L. |
| Approve receipt | Tap `Confirm`, `Change Account`, `Change Amount`, or `Ignore` | Only confirmed receipts are inserted into `confirmed_expenses` and `pnl_lines`. |
| Record sale | `/sold 1 100000` | Marks artwork as sold, posts gross sale, gallery revenue, artist payable, and collaborator share. |
| Review expenses | `/expense_report BKK2026` | Groups confirmed expenses by account head. |
| Review pending | `/pending BKK2026` | Shows receipts still awaiting approval. |
| Generate P&L | `/pl BKK2026` | Produces the exhibition-level P&L in THB. |
| Export Excel | `/export BKK2026` | Exports P&L, artist payables, confirmed expenses, and sale allocations. |

### Receipt capture behavior

The bot intentionally does **not** save OCR or AI guesses directly into accounting records. Text receipts and photo receipts create rows in `pending_expenses`. The approval card shows the exhibition, amount, suggested account head, P&L section, description, and status. A confirmed receipt is then posted to `confirmed_expenses` and to the P&L as a `pnl_lines` row. A changed or ignored receipt remains auditable through the pending receipt status.

For photo receipts, Telegram provides the image file reference, and the bot uses the caption as the controlled source text. If a photo has no amount in the caption, the bot creates a pending receipt with a zero amount so the user can press **Change Amount** before confirming.

### Expense account heads

| Account head | Default P&L section |
|---|---|
| Transport & Local Logistics | `direct_cost` |
| Air Cargo & Freight | `operating_expense` |
| Venue Rental | `operating_expense` |
| Installation & Production | `direct_cost` |
| Framing & Artwork Preparation | `direct_cost` |
| Food & Beverage / Hospitality | `operating_expense` |
| Marketing & PR | `operating_expense` |
| Office & Admin Supplies | `operating_expense` |
| Staff & Helpers / Labor | `operating_expense` |
| Banking & Payment Fees | `operating_expense` |
| Miscellaneous (Needs Review) | `operating_expense` |

### New local validation

Run the workflow regression test without Telegram credentials:

```bash
cd /home/ubuntu/exhibitledger_thb_bot
python3.11 local_workflow_test.py
```

The test uses a temporary SQLite database at `/tmp/exhibitledger_workflow_test.db`, creates a sample exhibition, applies a 45/10/45 gallery-collaborator-artist split, records an artwork sale, approves two receipts, verifies the P&L totals, and exports a workbook to `/tmp/exhibitledger_workflow_exports`.


## Perfection Pass: Guided Menus, Controls, and Stronger P&L Review

This refinement turns the bot from a command-only prototype into a **guided Telegram finance assistant**. The recommended entry point is now `/menu`, which opens button sections for exhibitions, split rules, artworks and sales, receipts and expenses, reports, export, and help. Slash commands still work for power users, but day-to-day operators can now tap a section and answer the bot's follow-up prompt instead of memorizing command syntax.

| Menu section | Main guided actions | Practical purpose |
|---|---|---|
| Exhibitions | Current exhibition, list previous exhibitions, add new exhibition, switch exhibition, readiness check | Keeps every transaction attached to the correct exhibition ledger. |
| Splits | View split, set 50/50 preset, set 45/10/45 preset, enter custom split | Prevents sale allocation mistakes before artworks are sold. |
| Artworks & Sales | Register artwork, list artwork inventory, record sale, inventory dashboard | Connects consigned artworks, sales, collections, and artist payables. |
| Receipts & Expenses | Add text receipt, pending receipts, expense report, account heads, budget setup | Forces approval before expenses enter the P&L and supports budget control. |
| Reports & Export | Executive dashboard, P&L, artist payouts, budget vs actual, data check, Excel export | Provides accountant-style review and final reporting. |
| Help & Settings | Command guide, current exhibition, read-only Sheets status and preview | Keeps Google Sheets safe and read-only until import/write-back rules are approved. |

### New management controls added

The P&L format now includes practical gallery controls that were missing from the earlier command-only version. Sale records can now store buyer name, collected amount, payment method, and notes, which allows the bot to show **sale receivables** instead of treating every sale as fully collected. Inventory metrics now show registered artworks, sold artworks, available artworks, sell-through rate, unsold asking value, cash collected, and outstanding sale receivables. Expense budgets can be set per account head, and the bot can compare confirmed actual expenses against those budgets.

| Control | Why it matters | Where to use it |
|---|---|---|
| Receipt approval gate | Stops accidental or guessed expenses from posting directly into finance records. | Pending receipt cards and `/pending`. |
| Account-head classification | Makes the P&L readable by grouping costs into gallery-relevant heads. | `/accounts`, `/expense_report`, receipt approval. |
| Budget vs actual | Shows whether hospitality, installation, logistics, marketing, or other costs are over plan. | `/budget` or Reports menu. |
| Inventory dashboard | Tracks sell-through and unsold asking value, not only accounting profit. | `/inventory` or Artworks & Sales menu. |
| Sale receivables | Separates sold value from cash collected. | `/summary`, `/inventory`, Excel export. |
| Final readiness check | Creates a pre-export checklist for pending receipts, split rules, sale collection gaps, and other review issues. | `/readiness` or Exhibitions menu. |

### Guided examples

A normal operator flow can now be completed by tapping buttons. Run `/menu`, tap **Exhibitions**, tap **Add New Exhibition**, and send `BKK2026 | Bangkok Art Fair 2026 | Bangkok | 2026-03-01 | 2026-03-15 | Main fair booth`. Then tap **Splits** and choose a preset or enter a custom rule such as `gallery 45 collaborator Curator 10 artist 45`. Next, tap **Artworks & Sales**, register artworks using `Quiet River | Artist A | 120000`, and record sales using `1 | 100000 | Buyer Name | 60000 | Bank Transfer | Balance due next week`. Expenses are captured from **Receipts & Expenses** using text such as `3500 coffee and snacks opening night`, then confirmed, reclassified, corrected, or ignored from the approval card.

### Added slash commands

| Command | Description |
|---|---|
| `/menu` | Opens the guided button menu. |
| `/cancel` | Cancels the current guided data-entry action. |
| `/inventory <CODE>` | Shows collection and sales KPIs for the exhibition. |
| `/budget <CODE>` | Shows budget vs actual by account head. |
| `/budget <CODE> <ACCOUNT_HEAD_OR_NUMBER> <AMOUNT_THB>` | Sets an account-head budget. |
| `/readiness <CODE>` | Shows final readiness checks before export or management review. |

### Expanded Excel workbook

The Excel export now goes beyond a simple P&L. The workbook includes additional management-control sheets for executive summary, P&L lines, artist payables, confirmed expenses, sale allocations, artworks, pending receipts, inventory metrics, budget vs actual, data-quality checks, and audit log. This makes the export more useful for gallery management, accountant review, and later reconciliation with Google Sheets.

### Operator safeguard

The bot is intentionally conservative. If a message contains no recognizable THB amount and no guided action is active, it does not invent a transaction. If a receipt is uncertain, it stays pending. If a split rule does not total 100%, the bot rejects it. The correct operating habit is to keep `/menu` open, clear pending receipts regularly, record collected amounts separately from sale price, run `/readiness`, then export the final workbook.
