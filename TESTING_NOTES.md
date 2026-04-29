# ExhibitLedger THB Prototype — Testing Notes

## Build Completed

A deployment-ready Telegram bot prototype has been created for exhibition-based P&L reporting. The first test exhibition is `SHWEDAGON2024`, based on the Shwe Dagon estimated P&L and artist commission PDFs previously attached by the user. The next-version hardening pass added an executive summary command, stronger accountant-style reconciliation checks, and read-only Google Sheets setup/preview scaffolding.

## Validation Performed

The prototype was seeded locally, the Shwe Dagon P&L was generated, artist payables were calculated, data-quality checks were run, and an Excel export was produced.

| Validation Step | Result |
|---|---|
| Python dependency installation | Passed |
| Database initialization | Passed |
| Shwe Dagon seed script | Passed |
| Local P&L generation | Passed |
| Artist payable generation | Passed |
| Data-quality check | Passed with expected currency warning |
| Excel export | Passed |
| Executive summary command | Added; pending final live Telegram user review |
| Google Sheets read-only setup check | Added; awaits real spreadsheet ID and service-account credential |
| Google Sheets dry-run preview | Added; no write-back enabled |
| Python syntax compilation | Passed |

## Current Shwe Dagon Test Output

The local test generated the following headline figures using the confirmed Shwe Dagon conversion rate: **150 MMK = 1 THB**, implemented as `SEED_MMK_TO_THB_RATE=0.006666666666666667`. The source PDFs remain MMK-origin files, but the prototype output below is recalculated into THB.

| Metric | Prototype Output |
|---|---:|
| Gross artwork sales / activity | ฿320,640.00 THB |
| Gallery revenue | ฿320,640.00 THB |
| Direct costs | ฿193,554.34 THB |
| Gross profit | ฿127,085.66 THB |
| Operating expenses | ฿89,875.33 THB |
| Net profit / loss | ฿37,210.33 THB |
| Artist payable total from sample commission statement | ฿25,000.00 THB |
| Artist payable outstanding | ฿25,000.00 THB |

## Important Currency Warning

The bot itself is THB-only, but the seed data was extracted from MMK source files. The current package has now been reseeded with the user-confirmed Shwe Dagon rate:

```bash
export SEED_MMK_TO_THB_RATE=0.006666666666666667
python seed_shwedagon.py
python local_report_test.py --code SHWEDAGON2024 --export
```

If a later exhibition uses a different source currency or rate, reseed that exhibition separately before relying on the output.

## Telegram Testing Requirements

A live polling bot process was started in the sandbox using the provided BotFather token without writing the token into project source files. The token should still be treated as sensitive; after testing, consider regenerating it in BotFather if this chat or any logs may be shared. For future local testing, set `TELEGRAM_BOT_TOKEN` in the terminal before running `python main.py`.

## Recommended Next Step

The next practical step is to test this hardened version in Telegram with `/summary SHWEDAGON2024`, `/pl SHWEDAGON2024`, `/data_check SHWEDAGON2024`, and `/export SHWEDAGON2024`. After the report layout is approved, provide the Google Spreadsheet ID and a service-account credential so `/sheets_status` and `/sync_preview` can be tested against the real finance workbook in read-only mode. Confirmation-based write-back should remain disabled until the live import mapping is proven correct.
