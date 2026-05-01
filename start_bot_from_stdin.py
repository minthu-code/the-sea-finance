import getpass
import os
import subprocess
import sys
import time
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent


def main() -> None:
    print("TOKEN_INPUT_READY", flush=True)
    token = getpass.getpass("Telegram token: ").strip()
    if not token:
        raise SystemExit("No token received")

    env = os.environ.copy()
    env["TELEGRAM_BOT_TOKEN"] = token
    env["DB_PATH"] = str(PROJECT_DIR / "exhibitledger.db")
    env["DEFAULT_EXHIBITION"] = "SHWEDAGON2024"
    env["EXPORT_DIR"] = str(PROJECT_DIR / "exports")
    env["LOG_LEVEL"] = env.get("LOG_LEVEL", "INFO")

    log_path = PROJECT_DIR / "bot.log"
    pid_path = PROJECT_DIR / "bot.pid"
    log = open(log_path, "ab", buffering=0)
    process = subprocess.Popen(
        [sys.executable, "main.py"],
        cwd=str(PROJECT_DIR),
        stdout=log,
        stderr=subprocess.STDOUT,
        env=env,
        start_new_session=True,
    )
    pid_path.write_text(str(process.pid), encoding="utf-8")
    time.sleep(4)
    print(f"BOT_PID {process.pid}", flush=True)
    print(f"BOT_RUNNING {process.poll() is None}", flush=True)


if __name__ == "__main__":
    main()
