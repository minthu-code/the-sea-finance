# GitHub Web Upload Steps

You are currently on GitHub's **Drag files here to add them to your repository** screen. Use the package I provided as follows.

## Correct upload method

Download the attached ZIP file, unzip it on your computer, then drag **all extracted files and folders** into the GitHub upload box.

Do **not** drag the ZIP file itself into GitHub as the only upload, because GitHub will store the ZIP instead of creating the bot repository files.

## Files you should see after unzipping

The extracted folder should include these important files at the top level.

| File | Purpose |
|---|---|
| `main.py` | Telegram bot entrypoint with guided menus |
| `exhibitledger.py` | THB finance engine and P&L logic |
| `Dockerfile` | Render container startup file |
| `render.yaml` | Render worker and persistent-disk blueprint |
| `requirements.txt` | Python dependencies |
| `.env.example` | Safe example environment file, without real token |
| `.gitignore` | Prevents secrets and databases from being committed |
| `.dockerignore` | Prevents secrets and databases from entering Docker image |
| `DEPLOY_RENDER.md` | Render deployment instructions |
| `README.md` | Bot usage guide |

## Commit message

For the first upload, use this commit message:

```text
Initial upload of TheSeaFinance Telegram bot
```

After the GitHub upload is complete, connect the repository to Render and set the required environment variables from `DEPLOY_RENDER.md`.
