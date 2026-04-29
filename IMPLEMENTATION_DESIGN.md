# ExhibitLedger THB Bot — Prototype Implementation Design

## Objective

This prototype is a deployment-ready Telegram bot for testing exhibition-based P&L reporting using the Shwe Dagon exhibition as the first seed case. The bot reports only in **Thai Baht (THB)** and treats every financial row as belonging to an exhibition.

## Prototype Scope

The first version is intentionally practical and testable. It uses SQLite as the local prototype database, supports Telegram commands, generates a Shwe Dagon P&L report, exports the report to Excel, and includes deployment files for Render.com using Docker. Google Sheets integration is prepared as a later module because the user must provide credentials and sheet IDs before live read/write can be enabled.

## Data Model

| Table | Purpose |
|---|---|
| `exhibitions` | Stores exhibition code, name, location, period, status, and notes. |
| `pnl_lines` | Stores revenue, direct cost, operating expense, and optional overhead lines in THB. |
| `artist_payables` | Stores artist settlement lines, gross sale amount, gallery commission, artist payable, paid amount, and outstanding amount in THB. |
| `audit_log` | Stores generated reports and future write actions for traceability. |

## P&L Logic

The report is structured as follows:

| Section | Calculation |
|---|---|
| Gross Artwork Sales | Sum of sales bridge lines, used as activity metric. |
| Gallery Revenue | For consignment, only gallery commission is revenue. For owned works, revenue can be full sale price. |
| Direct Costs | Artist fees/payables, partner commission, artwork-specific preparation, canvas, catalog printing, and local transportation. |
| Gross Profit | Gallery Revenue minus Direct Costs. |
| Operating Expenses | Air cargo, tickets, venue rental, event supplies, snacks, photographer, and other exhibition running costs. |
| Net Profit / Loss | Gross Profit minus Operating Expenses and optional overhead. |

## THB Conversion Handling for Shwe Dagon Test Data

The attached Shwe Dagon source PDFs are denominated in MMK. Because the final bot must report only in THB, the seed script uses a configurable conversion rate read from `SEED_MMK_TO_THB_RATE`. The current validated Shwe Dagon prototype uses the user-confirmed rate **150 MMK = 1 THB**, implemented as `SEED_MMK_TO_THB_RATE=0.006666666666666667`. For any future exhibition or revised source file, the rate should be confirmed before reseeding.

## Commands

| Command | Purpose |
|---|---|
| `/start` | Show welcome and command menu. |
| `/exhibitions` | List available exhibitions. |
| `/pl SHWEDAGON2024` | Generate the exhibition P&L report. |
| `/artist_payouts SHWEDAGON2024` | Show artist payable summary. |
| `/export SHWEDAGON2024` | Export P&L and artist payables to Excel. |
| `/data_check SHWEDAGON2024` | Show data-quality warnings. |
| `/help` | Show command help. |

## Safety Notes

This version does not write to Google Sheets yet. That is intentional for the first prototype, because the bot should be tested against Shwe Dagon data before live editing is enabled. The next build phase can add Google Sheets read/write using `gspread`, service-account credentials, sheet mapping, and confirmation steps before every edit.
