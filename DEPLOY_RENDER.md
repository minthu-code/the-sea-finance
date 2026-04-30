# Render Deployment Guide for TheSeaFinance Bot

This repository is ready to deploy as a **Docker Web Service** on Render. The Telegram bot still runs in **polling mode**, but it also opens a tiny HTTP health endpoint so Render can detect a live port. The bot uses **Thai Baht (THB)** as the operating currency and stores its finance data in SQLite.

## 1. Upload to GitHub

Unzip the GitHub upload package locally and upload the extracted files into your `the-sea-finance` GitHub repository. Do **not** upload the ZIP as the only file, because GitHub will store it as a ZIP instead of creating the repository files.

The repository should contain files such as `main.py`, `exhibitledger.py`, `Dockerfile`, `render.yaml`, `requirements.txt`, `.env.example`, `.gitignore`, `.dockerignore`, and this deployment guide.

## 2. Required Render settings

The included `render.yaml` creates a Docker-based Web Service named `the-sea-finance-bot`. It defines a persistent disk mounted at `/data` and a health check path at `/health`.

| Setting | Value |
| --- | --- |
| Service type | **Web Service** |
| Runtime | Docker |
| Start command | Managed by `Dockerfile` with `python main.py` |
| Health check path | `/health` |
| Render port | Render provides `PORT`; the bot opens a small health server on that port |
| Database path | `/data/exhibitledger.db` |
| Export path | `/data/exports` |
| Persistent disk | Required |

> If a previous Render deployment failed with **"No open ports detected"**, it was created as a Web Service while the earlier bot only ran Telegram polling. This version fixes that by serving `/health` in the background while the Telegram bot continues polling.

## 3. Environment variables

Set these in Render. Never commit real tokens to GitHub.

| Variable | Required | Recommended value |
| --- | --- | --- |
| `TELEGRAM_BOT_TOKEN` | Yes | Your BotFather token |
| `DB_PATH` | Yes | `/data/exhibitledger.db` |
| `DEFAULT_EXHIBITION` | Yes | `SHWEDAGON2024` or your preferred default exhibition code |
| `EXPORT_DIR` | Yes | `/data/exports` |
| `SEED_MMK_TO_THB_RATE` | Optional | Only needed when reseeding old MMK source data |
| `LOG_LEVEL` | Optional | `INFO` |

Render automatically provides `PORT` for Web Services. You do not need to set `PORT` manually.

## 4. Important safety rules

The `.env`, local SQLite database, logs, generated exports, process IDs, and temporary validation outputs are excluded from the GitHub upload package. Keep it that way. Your production finance database should live on Render's persistent disk, not in GitHub.

If you later change the bot code, the recommended workflow is to edit locally, run validation, push to GitHub, then let Render redeploy from the repository.

## 5. First bot test after deployment

After Render shows the Web Service as live, open Telegram and message your bot:

```text
/menu
```

Use `/menu` for the guided button workflow. Use `/readiness`, `/summary`, `/pl`, and `/export` after creating or selecting an exhibition.

## 6. If you created the wrong Render service type

If your existing service is a **Background Worker**, it may run the bot but will not use Render's Web Service health checks. If your existing service is a **Web Service** and previously failed, upload this corrected package and redeploy it. The safest path is to create or update a **Web Service**, set `TELEGRAM_BOT_TOKEN`, and keep the health check path as `/health`.
