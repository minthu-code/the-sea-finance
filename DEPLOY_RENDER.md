# Render Deployment Guide for TheSeaFinance Bot

This repository is ready to deploy as a **Telegram worker bot** on Render. The bot uses **Thai Baht (THB)** as the operating currency and stores its finance data in SQLite.

## 1. Upload to GitHub

Unzip the GitHub upload package locally and upload the extracted files into your `the-sea-finance` GitHub repository. Do **not** upload the ZIP as the only file, because GitHub will store it as a ZIP instead of creating the repository files.

The repository should contain files such as `main.py`, `exhibitledger.py`, `Dockerfile`, `render.yaml`, `requirements.txt`, `.env.example`, `.gitignore`, `.dockerignore`, and this deployment guide.

## 2. Required Render settings

The included `render.yaml` creates a Docker-based worker service named `the-sea-finance-bot`. It also defines a persistent disk mounted at `/data`.

| Setting | Value |
|---|---|
| Service type | Worker |
| Runtime | Docker |
| Start command | Managed by `Dockerfile` with `python main.py` |
| Database path | `/data/exhibitledger.db` |
| Export path | `/data/exports` |
| Persistent disk | Required |

## 3. Environment variables

Set these in Render. Never commit real tokens to GitHub.

| Variable | Required | Recommended value |
|---|---:|---|
| `TELEGRAM_BOT_TOKEN` | Yes | Your BotFather token |
| `DB_PATH` | Yes | `/data/exhibitledger.db` |
| `DEFAULT_EXHIBITION` | Yes | `SHWEDAGON2024` or your preferred default exhibition code |
| `EXPORT_DIR` | Yes | `/data/exports` |
| `SEED_MMK_TO_THB_RATE` | Optional | Only needed when reseeding old MMK source data |
| `LOG_LEVEL` | Optional | `INFO` |

## 4. Important safety rules

The `.env`, local SQLite database, logs, generated exports, process IDs, and temporary validation outputs are excluded from the GitHub upload package. Keep it that way. Your production finance database should live on Render's persistent disk, not in GitHub.

If you later change the bot code, the recommended workflow is to edit locally, run validation, push to GitHub, then let Render redeploy from the repository.

## 5. First bot test after deployment

After Render shows the worker as live, open Telegram and message your bot:

```text
/menu
```

Use `/menu` for the guided button workflow. Use `/readiness`, `/summary`, `/pl`, and `/export` after creating or selecting an exhibition.
