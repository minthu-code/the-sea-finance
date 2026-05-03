import base64
import io
import logging
import os
import re
import sqlite3
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread
from typing import List, Tuple

import pytesseract
from PIL import Image
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.error import BadRequest, Conflict, Forbidden, NetworkError, TelegramError, TimedOut
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

from exhibitledger import (
    account_head_names,
    add_artwork,
    confirm_pending_expense,
    connect,
    create_exhibition,
    create_pending_expense,
    data_quality_checks,
    export_report_xlsx,
    format_account_heads_markdown,
    format_artist_payables_markdown,
    format_artworks_markdown,
    format_budget_report_markdown,
    format_executive_summary_markdown,
    format_expense_report_markdown,
    format_inventory_dashboard_markdown,
    format_pending_expense_card,
    format_pending_expenses_markdown,
    format_readiness_markdown,
    format_report_markdown,
    format_sale_markdown,
    format_split_rules_markdown,
    get_exhibition,
    get_pending_expense,
    ignore_pending_expense,
    init_db,
    list_exhibitions,
    log_action,
    money,
    normalize_code,
    parse_amount_thb,
    record_sale,
    set_commission_splits,
    set_expense_budget,
    update_pending_account,
    update_pending_amount,
)
from sheets_integration import format_sheets_status_markdown, format_sync_preview_markdown

# Optional OpenAI client for enhanced receipt processing
ai_client = None
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
if OPENAI_API_KEY and OPENAI_API_KEY.strip():
    try:
        from openai import OpenAI

        ai_client = OpenAI()
        logging.info("OpenAI client initialized for enhanced receipt scanning.")
    except Exception as e:
        logging.warning(f"Could not initialize OpenAI client: {e}. Using OCR fallback.")


logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=os.environ.get("LOG_LEVEL", "INFO"),
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("telegram.ext").setLevel(logging.INFO)


DEFAULT_EXHIBITION = re.sub(r"[^A-Z0-9_]", "", (os.environ.get("DEFAULT_EXHIBITION", "SHWEDAGON2024") or "SHWEDAGON2024").split()[0].upper()) or "SHWEDAGON2024"
ALLOWED_SPLIT_TYPES = {"gallery", "artist", "collaborator", "collector"}


# ---------------------------------------------------------------------------
# Render Web Service health port
# ---------------------------------------------------------------------------


class _HealthHandler(BaseHTTPRequestHandler):
    """Tiny HTTP endpoint so Render Web Services can detect an open port.

    The Telegram bot still runs in polling mode. This server only answers Render's
    health/port check and does not expose any finance data.
    """

    def do_GET(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler API
        if self.path in {"/", "/health", "/healthz"}:
            body = b"TheSeaFinance bot is running.\n"
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_response(404)
        self.end_headers()

    def log_message(self, fmt: str, *args: object) -> None:
        logger.debug("Health check: " + fmt, *args)


def start_render_health_server() -> None:
    """Start a background HTTP health server when Render provides PORT.

    Render Web Services must bind to a port. Background workers normally do not
    set PORT, so this function quietly does nothing for worker deployments.
    """

    raw_port = os.environ.get("PORT")
    if not raw_port:
        return
    try:
        port = int(raw_port)
    except ValueError:
        logger.warning("Ignoring invalid PORT value: %s", raw_port)
        return

    server = ThreadingHTTPServer(("0.0.0.0", port), _HealthHandler)
    thread = Thread(target=server.serve_forever, name="render-health-server", daemon=True)
    thread.start()
    logger.info("Render health server listening on port %s", port)


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------


def _current_exhibition(context: ContextTypes.DEFAULT_TYPE) -> str:
    raw = context.user_data.get("current_exhibition") or DEFAULT_EXHIBITION
    # Strip anything that isn't a valid exhibition code character
    clean = re.sub(r"[^A-Z0-9_]", "", raw.split()[0].upper()) if raw else ""
    return clean or DEFAULT_EXHIBITION


def _get_code(context: ContextTypes.DEFAULT_TYPE) -> str:
    return context.args[0].strip().upper() if context.args else _current_exhibition(context)


def _set_current_exhibition(context: ContextTypes.DEFAULT_TYPE, code: str) -> None:
    context.user_data["current_exhibition"] = normalize_code(code)


def _clear_flow(context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("flow", None)


def _set_flow(context: ContextTypes.DEFAULT_TYPE, name: str) -> None:
    context.user_data["flow"] = {"name": name}


def _parse_float(raw: str, label: str) -> float:
    try:
        return float(str(raw).replace(",", ""))
    except ValueError:
        raise ValueError(f"{label} must be a number.")


def _looks_number(token: str) -> bool:
    try:
        float(str(token).replace(",", ""))
        return True
    except ValueError:
        return False


def _pipe_parts(text: str) -> List[str]:
    return [part.strip() for part in (text or "").split("|")]


def _account_name_from_text(raw: str) -> str:
    names = account_head_names()
    cleaned = (raw or "").strip()
    if cleaned.isdigit():
        index = int(cleaned) - 1
        if 0 <= index < len(names):
            return names[index]
    lowered = cleaned.lower()
    for name in names:
        if lowered == name.lower():
            return name
    partial = [name for name in names if lowered and lowered in name.lower()]
    if len(partial) == 1:
        return partial[0]
    raise ValueError("Unknown account head. Use /accounts or choose from the guided menu.")


async def _send_long_text(message, text: str, parse_mode=None, reply_markup=None) -> None:
    max_len = 3900
    if len(text) <= max_len:
        await message.reply_text(text, parse_mode=parse_mode, reply_markup=reply_markup)
        return
    chunks = [text[i : i + max_len] for i in range(0, len(text), max_len)]
    for idx, chunk in enumerate(chunks):
        await message.reply_text(chunk, parse_mode=parse_mode, reply_markup=reply_markup if idx == len(chunks) - 1 else None)


# ---------------------------------------------------------------------------
# Menus and keyboards
# ---------------------------------------------------------------------------


def _main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Exhibitions", callback_data="menu:exhibitions"), InlineKeyboardButton("Splits", callback_data="menu:splits")],
            [InlineKeyboardButton("Artworks & Sales", callback_data="menu:artworks"), InlineKeyboardButton("Receipts & Expenses", callback_data="menu:expenses")],
            [InlineKeyboardButton("Reports & Export", callback_data="menu:reports"), InlineKeyboardButton("Help & Settings", callback_data="menu:help")],
        ]
    )


def _back_home_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("Back to Main Menu", callback_data="menu:home")]])


def _exhibition_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Current Exhibition", callback_data="quick:current"), InlineKeyboardButton("List Previous", callback_data="quick:list_exhibitions")],
            [InlineKeyboardButton("Add New Exhibition", callback_data="flow:new_exhibition"), InlineKeyboardButton("Switch Exhibition", callback_data="flow:use_exhibition")],
            [InlineKeyboardButton("Final Readiness Check", callback_data="report:readiness")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _split_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("View Current Split", callback_data="report:split")],
            [InlineKeyboardButton("Preset 50/50 Gallery-Artist", callback_data="preset_split:5050")],
            [InlineKeyboardButton("Preset 45/10/45 Gallery-Collaborator-Artist", callback_data="preset_split:451045")],
            [InlineKeyboardButton("Custom Split", callback_data="flow:custom_split")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _artwork_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Register Artwork", callback_data="artwork:register_options"), InlineKeyboardButton("List Artworks", callback_data="report:artworks")],
            [InlineKeyboardButton("Edit Artwork", callback_data="artwork:edit_hint"), InlineKeyboardButton("Inventory Dashboard", callback_data="report:inventory")],
            [InlineKeyboardButton("Record Sale", callback_data="flow:record_sale")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _artwork_register_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Manual Input", callback_data="flow:add_artwork")],
            [InlineKeyboardButton("Bulk Import (Excel)", callback_data="artwork:bulk_import")],
            [InlineKeyboardButton("Back", callback_data="menu:artworks")],
        ]
    )


def _expense_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Add Text Receipt", callback_data="flow:text_receipt"), InlineKeyboardButton("Pending Receipts", callback_data="report:pending")],
            [InlineKeyboardButton("Expense Report", callback_data="report:expenses"), InlineKeyboardButton("Account Heads", callback_data="report:accounts")],
            [InlineKeyboardButton("Set Expense Budget", callback_data="flow:set_budget")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _reports_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Executive Dashboard", callback_data="report:summary"), InlineKeyboardButton("P&L", callback_data="report:pl")],
            [InlineKeyboardButton("Artist Payouts", callback_data="report:artists"), InlineKeyboardButton("Budget vs Actual", callback_data="report:budget")],
            [InlineKeyboardButton("Data Check", callback_data="report:data_check"), InlineKeyboardButton("Export Excel", callback_data="report:export")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _help_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Command Guide", callback_data="quick:command_guide"), InlineKeyboardButton("Current Exhibition", callback_data="quick:current")],
            [InlineKeyboardButton("Sheets Status", callback_data="report:sheets_status"), InlineKeyboardButton("Sheets Preview", callback_data="report:sync_preview")],
            [InlineKeyboardButton("Back", callback_data="menu:home")],
        ]
    )


def _pending_keyboard(pending_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Confirm", callback_data=f"expense:confirm:{pending_id}")],
            [
                InlineKeyboardButton("Change Account", callback_data=f"expense:account:{pending_id}"),
                InlineKeyboardButton("Change Amount", callback_data=f"expense:amount:{pending_id}"),
            ],
            [InlineKeyboardButton("Ignore", callback_data=f"expense:ignore:{pending_id}")],
        ]
    )


def _account_keyboard(pending_id: int) -> InlineKeyboardMarkup:
    names = account_head_names()
    buttons = []
    # Pair account heads two per row so the list isn't endless
    for i in range(0, len(names), 2):
        row = [InlineKeyboardButton(names[i][:40], callback_data=f"expense:setacct:{pending_id}:{i}")]
        if i + 1 < len(names):
            row.append(InlineKeyboardButton(names[i + 1][:40], callback_data=f"expense:setacct:{pending_id}:{i + 1}"))
        buttons.append(row)
    buttons.append([InlineKeyboardButton("← Back", callback_data=f"expense:back:{pending_id}")])
    return InlineKeyboardMarkup(buttons)


def _exhibition_picker_keyboard() -> InlineKeyboardMarkup:
    rows = []
    for row in list_exhibitions()[:20]:
        rows.append([InlineKeyboardButton(f"{row['code']} — {row['name']}"[:60], callback_data=f"useexh:{row['code']}")])
    rows.append([InlineKeyboardButton("Back", callback_data="menu:exhibitions")])
    return InlineKeyboardMarkup(rows)


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------


def _parse_split_tokens(code: str, tokens: List[str]) -> Tuple[str, List[dict]]:
    if len(tokens) < 2:
        raise ValueError("Use: gallery 50 artist 50, or gallery 45 collaborator Curator 10 artist 45")
    entries = []
    idx = 0
    while idx < len(tokens):
        party_type = tokens[idx].lower().strip()
        if party_type not in ALLOWED_SPLIT_TYPES:
            raise ValueError(f"Unknown party type: {tokens[idx]}. Use gallery, artist, collaborator, or collector.")
        idx += 1
        if idx >= len(tokens):
            raise ValueError(f"Missing percentage after {party_type}.")
        party_name = party_type.title()
        if _looks_number(tokens[idx]):
            percent = _parse_float(tokens[idx], "Split percent")
            idx += 1
        else:
            party_name = tokens[idx]
            idx += 1
            if idx >= len(tokens):
                raise ValueError(f"Missing percentage after {party_type} {party_name}.")
            percent = _parse_float(tokens[idx], "Split percent")
            idx += 1
        entries.append({"party_type": party_type, "party_name": party_name, "percent": percent})
    return normalize_code(code), entries


def _parse_split_args(args: List[str]) -> Tuple[str, List[dict]]:
    if len(args) < 3:
        raise ValueError(
            "Usage: /set_split <EXHIBITION_CODE> gallery 50 artist 50\n"
            "Optional named collaborator example: /set_split EXH gallery 40 collaborator Rotary 10 artist 50"
        )
    return _parse_split_tokens(normalize_code(args[0]), args[1:])


def _extract_receipt_code_and_text(context: ContextTypes.DEFAULT_TYPE, text: str) -> Tuple[str, str]:
    parts = (text or "").strip().split(maxsplit=1)
    if parts and get_exhibition(parts[0]):
        code = normalize_code(parts[0])
        rest = parts[1] if len(parts) > 1 else ""
        _set_current_exhibition(context, code)
        return code, rest
    return _current_exhibition(context), text


def generate_artwork_template_xlsx() -> str:
    """Generate and return path to the artwork bulk import template Excel file."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Artworks"

    headers = ["Title", "Artist", "Asking Price THB", "Notes (optional)"]
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Sample rows so the user understands the format
    samples = [
        ["Red Eye Balls", "Aung Myint", 400000, ""],
        ["In & Out", "Aung Myint", 304000, "Special framing required"],
        ["Morning Light", "Artist Name", 120000, ""],
    ]
    for row in samples:
        ws.append(row)

    # Instructions sheet
    ws2 = wb.create_sheet("Instructions")
    ws2.append(["Column", "Required", "Notes"])
    ws2.append(["Title", "Yes", "Artwork title — special characters like & are fine"])
    ws2.append(["Artist", "Yes", "Artist full name"])
    ws2.append(["Asking Price THB", "Yes", "Number only, no commas or currency symbols"])
    ws2.append(["Notes (optional)", "No", "Any extra notes about the artwork"])
    ws2.append(["", "", ""])
    ws2.append(["", "", "Delete the 3 sample rows before uploading."])
    ws2.append(["", "", "Upload this file back to the bot after filling it in."])

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 35

    export_dir = os.environ.get("EXPORT_DIR", "./exports")
    os.makedirs(export_dir, exist_ok=True)
    path = os.path.join(export_dir, "artwork_import_template.xlsx")
    wb.save(path)
    return path


async def handle_document_upload(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle Excel file uploads for bulk artwork import.

    Accepts the file if:
    - The bulk_artwork flow is active (user tapped Bulk Import then uploaded), OR
    - The file is named like our template (user filled it in and sent it back), OR
    - The file is any .xlsx and no other flow is active (user uploaded directly without going through menu).

    Never ignores silently — always replies so the user knows what happened.
    """
    doc = update.message.document
    if not doc:
        return

    fname = doc.file_name or ""
    if not fname.lower().endswith((".xlsx", ".xls")):
        # Only respond if there's a bulk flow active — otherwise silently ignore non-Excel docs
        if context.user_data.get("flow", {}).get("name") == "bulk_artwork":
            await update.message.reply_text(
                "Please upload an Excel (.xlsx) file, not this file type.\n"
                "Use the template sent above, fill it in, and upload it back."
            )
        return

    in_bulk_flow = context.user_data.get("flow", {}).get("name") == "bulk_artwork"
    is_template_name = "artwork_import_template" in fname.lower()
    any_other_flow = context.user_data.get("flow", {}).get("name") not in {None, "bulk_artwork"}

    # Ignore if user is mid-way through a completely different flow (e.g. adding an exhibition)
    if any_other_flow:
        return

    # If none of our signals match — still accept it and try, because the user clearly
    # intended to send an artwork file. Don't leave them hanging silently.
    # Worst case: it fails gracefully and tells them what went wrong.

    code = _current_exhibition(context)
    ex = get_exhibition(code)

    # Always acknowledge immediately so the user knows the bot received the file
    await update.message.reply_text(
        f"Received {fname}. Processing artwork import for {code}"
        + (f" — {ex['name']}" if ex else "")
        + "..."
    )
    _clear_flow(context)

    try:
        from openpyxl import load_workbook

        file = await doc.get_file()
        raw = await file.download_as_bytearray()

        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header row

        registered = []
        skipped = []
        sample_titles = {"Red Eye Balls", "In & Out", "Morning Light"}  # template sample rows

        for i, row in enumerate(rows, start=2):
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue  # blank row

            title = str(row[0]).strip() if row[0] is not None else ""
            artist = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            price_raw = row[2] if len(row) > 2 else None

            # Skip the sample rows from the template so user doesn't have to delete them
            if title in sample_titles and artist in {"Aung Myint", "Artist Name"}:
                skipped.append(f"Row {i} ({title}): skipped sample row — please replace with your real artworks")
                continue

            if not title or not artist:
                skipped.append(f"Row {i}: missing title or artist name")
                continue

            try:
                price = float(str(price_raw).replace(",", "").strip())
                if price <= 0:
                    raise ValueError("price must be greater than zero")
            except Exception:
                skipped.append(f"Row {i} ({title}): invalid price '{price_raw}' — must be a number like 120000")
                continue

            try:
                artwork_row = add_artwork(code, title, artist, price)
                registered.append(f"#{artwork_row['id']}  {title}  |  {artist}  |  {money(price)}")
            except Exception as exc:
                skipped.append(f"Row {i} ({title}): {exc}")

        if not registered and not skipped:
            await update.message.reply_text(
                "The file appears to be empty or only contains the header row.\n\n"
                "Make sure you filled in your artworks below the header row and saved the file before uploading.",
                reply_markup=_artwork_menu_keyboard(),
            )
            return

        lines = [f"Bulk import complete — {code}", ""]
        if registered:
            lines.append(f"Registered {len(registered)} artwork(s):")
            lines.extend(f"  ✅ {r}" for r in registered)
        if skipped:
            if registered:
                lines.append("")
            lines.append(f"Skipped {len(skipped)} row(s):")
            lines.extend(f"  ⚠️ {s}" for s in skipped)
        if registered:
            lines += ["", "Use List Artworks to confirm or /export to generate the full report."]

        await update.message.reply_text("\n".join(lines), reply_markup=_artwork_menu_keyboard())

    except Exception as exc:
        logger.exception("Bulk artwork import failed")
        await update.message.reply_text(
            f"Could not read the file: {exc}\n\n"
            "Make sure the file is a valid .xlsx and uses the artwork import template format.\n"
            "Tap Bulk Import again to get a fresh template."
        )


async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Quick one-line status: current exhibition + key numbers."""
    code = _current_exhibition(context)
    try:
        ex = get_exhibition(code)
        if not ex:
            rows = list_exhibitions()
            if rows:
                names = "\n".join(f"  • {r['code']} — {r['name']}" for r in rows[:10])
                await update.message.reply_text(
                    f"No active exhibition set.\n\nAvailable exhibitions:\n{names}\n\n"
                    "Use /use <CODE> or tap Menu → Exhibitions → Switch Exhibition.",
                    reply_markup=_exhibition_menu_keyboard(),
                )
            else:
                await update.message.reply_text("No exhibitions found. Create one first via /menu.")
            return

        with connect() as conn:
            artworks_total = conn.execute("SELECT COUNT(*) FROM artworks WHERE exhibition_code=?", (code,)).fetchone()[0]
            artworks_sold = conn.execute("SELECT COUNT(*) FROM artworks WHERE exhibition_code=? AND status='sold'", (code,)).fetchone()[0]
            pending_count = conn.execute("SELECT COUNT(*) FROM pending_expenses WHERE exhibition_code=? AND status='pending'", (code,)).fetchone()[0]
            net_row = conn.execute(
                "SELECT SUM(CASE WHEN section IN ('sales_bridge','gallery_revenue') THEN amount_thb ELSE -amount_thb END) "
                "FROM pnl_lines WHERE exhibition_code=?", (code,)
            ).fetchone()[0] or 0

        lines = [
            f"Current exhibition: {code} — {ex['name']}",
            f"Artworks: {artworks_sold} sold / {artworks_total} total",
            f"Pending receipts: {pending_count}",
            f"Net P&L so far: {money(net_row)}",
        ]
        if pending_count > 5:
            lines.append(f"⚠️ {pending_count} receipts awaiting approval — review with /pending {code}")
        await update.message.reply_text("\n".join(lines), reply_markup=_main_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not get status: {exc}")


# ---------------------------------------------------------------------------
# Core command handlers
# ---------------------------------------------------------------------------


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _clear_flow(context)
    code = _current_exhibition(context)
    ex = get_exhibition(code)
    if ex:
        exhibition_line = f"Current working exhibition: {code} — {ex['name']}"
    else:
        # Context was wiped (redeploy) — try to restore most recently created exhibition
        rows = list_exhibitions()
        if rows:
            # Auto-set to the most recently created exhibition (last in list)
            latest = rows[-1]
            _set_current_exhibition(context, latest["code"])
            exhibition_line = (
                f"⚠️ Session reset after update. Auto-switched to your most recent exhibition:\n"
                f"{latest['code']} — {latest['name']}\n\n"
                "All your exhibitions:\n"
                + "\n".join(f"  • {r['code']} — {r['name']}" for r in rows[:10])
                + "\n\nUse /use <CODE> to switch to a different one."
            )
        else:
            exhibition_line = "No exhibitions yet. Tap Exhibitions → Add New Exhibition to get started."

    text = (
        "THE SEA ART GALLERY — ExhibitLedger THB\n\n"
        "Tap /menu for the guided workflow. All reports in THB.\n\n"
        + exhibition_line
    )
    await update.message.reply_text(text, reply_markup=_main_menu_keyboard())


async def menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _clear_flow(context)
    code = _current_exhibition(context)
    ex = get_exhibition(code)
    label = f"{code} — {ex['name']}" if ex else code
    await update.message.reply_text(f"Main menu. Current exhibition: {label}", reply_markup=_main_menu_keyboard())


async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    _clear_flow(context)
    context.user_data.pop("awaiting_amount_for_pending_id", None)
    await update.message.reply_text("Cancelled the current guided action.", reply_markup=_main_menu_keyboard())


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = _command_guide_text()
    await _send_long_text(update.message, text, reply_markup=_help_menu_keyboard())


def _command_guide_text() -> str:
    return (
        "ExhibitLedger THB Help\n\n"
        "Preferred use: tap /menu and choose an action. Power-user slash commands remain available.\n\n"
        "Setup and master data:\n"
        "/new_exhibition <CODE> <NAME>\n"
        "/use <EXHIBITION_CODE>\n"
        "/set_split <CODE> gallery <PCT> artist <PCT>\n"
        "/set_split <CODE> gallery 40 collaborator Curator 10 artist 50\n"
        "/split <CODE>\n"
        "/accounts\n\n"
        "Artwork and sales:\n"
        "/add_artwork <CODE> <TITLE> | <ARTIST> | <PRICE_THB>\n"
        "/artworks <CODE>\n"
        "/inventory <CODE>\n"
        "/sold <ARTWORK_ID> <ACTUAL_PRICE_THB> [BUYER] [COLLECTED_THB]\n\n"
        "Receipts, expenses, and budgets:\n"
        "/receipt <CODE> <AMOUNT_THB> <DESCRIPTION>\n"
        "/pending <CODE>\n"
        "/expense_report <CODE>\n"
        "/budget <CODE> — show budget report\n"
        "/budget <CODE> <ACCOUNT_HEAD_OR_NUMBER> <AMOUNT_THB> — set budget\n\n"
        "Reports:\n"
        "/summary <CODE>\n"
        "/pl <CODE>\n"
        "/artist_payouts <CODE>\n"
        "/readiness <CODE>\n"
        "/data_check <CODE>\n"
        "/export <CODE>\n\n"
        "Edit artworks (fix title, artist, or price):\n"
        "/edit_artwork <ID> title <NEW TITLE>\n"
        "/edit_artwork <ID> artist <NEW ARTIST NAME>\n"
        "/edit_artwork <ID> price <NEW PRICE THB>\n\n"
        "Artist payments:\n"
        "/pay_artist <CODE> <ARTIST NAME> <AMOUNT THB>\n\n"
        "Google Sheets commands remain read-only and optional: /sheets_status and /sync_preview. Use /cancel at any time to stop a guided action."
    )


async def exhibitions(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    rows = list_exhibitions()
    if not rows:
        await update.message.reply_text("No exhibitions found. Create one from /menu or with /new_exhibition CODE Name")
        return
    lines = ["Available exhibitions:\n"]
    for row in rows:
        lines.append(f"• {row['code']} — {row['name']} ({row['status']}; {row['currency']}; {row['start_date'] or '-'} to {row['end_date'] or '-'})")
    await update.message.reply_text("\n".join(lines), reply_markup=_exhibition_picker_keyboard())


async def new_exhibition(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if len(context.args) < 2:
            raise ValueError("Usage: /new_exhibition <CODE> <NAME>")
        code = normalize_code(context.args[0])
        name = " ".join(context.args[1:]).strip()
        row = create_exhibition(code, name)
        _set_current_exhibition(context, row["code"])
        await update.message.reply_text(f"Created exhibition {row['code']} — {row['name']}\nCurrent working exhibition set to {row['code']}.", reply_markup=_main_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to create exhibition")
        await update.message.reply_text(f"Could not create exhibition: {exc}")


async def use_exhibition(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if not context.args:
            await update.message.reply_text(f"Current working exhibition: {_current_exhibition(context)}", reply_markup=_exhibition_menu_keyboard())
            return
        code = normalize_code(context.args[0])
        ex = get_exhibition(code)
        if not ex:
            raise ValueError(f"Exhibition not found: {code}")
        _set_current_exhibition(context, code)
        await update.message.reply_text(f"Current working exhibition set to {code} — {ex['name']}.", reply_markup=_main_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not set current exhibition: {exc}")


async def set_split(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        code, entries = _parse_split_args(context.args)
        set_commission_splits(code, entries)
        _set_current_exhibition(context, code)
        text = format_split_rules_markdown(code)
        await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN, reply_markup=_split_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to set commission split")
        await update.message.reply_text(f"Could not set split: {exc}")


async def split(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_split_rules_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_split_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not show split for {code}: {exc}")


async def add_artwork_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if len(context.args) < 2:
            raise ValueError("Usage: /add_artwork <CODE> <TITLE> | <ARTIST> | <PRICE_THB>")
        code = normalize_code(context.args[0])
        rest = " ".join(context.args[1:])
        if "|" in rest:
            parts = _pipe_parts(rest)
            if len(parts) != 3:
                raise ValueError("Use exactly three pipe-separated fields: Title | Artist | PriceTHB")
            title, artist, price_raw = parts
        else:
            if len(context.args) < 4:
                raise ValueError("Usage: /add_artwork <CODE> <TITLE> | <ARTIST> | <PRICE_THB>")
            title, artist, price_raw = context.args[1], context.args[2], context.args[3]
        price = _parse_float(price_raw, "Price")
        row = add_artwork(code, title, artist, price)
        _set_current_exhibition(context, code)
        await update.message.reply_text(
            f"Artwork registered.\nID: #{row['id']}\nExhibition: {row['exhibition_code']}\nTitle: {row['title']}\nArtist: {row['artist']}\nAsking Price: {money(row['asking_price_thb'])}",
            reply_markup=_artwork_menu_keyboard(),
        )
    except Exception as exc:
        logger.exception("Failed to add artwork")
        await update.message.reply_text(f"Could not add artwork: {exc}")


async def edit_artwork_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if len(context.args) < 3:
            raise ValueError(
                "Usage:\n"
                "/edit_artwork <ID> title New Title Here\n"
                "/edit_artwork <ID> artist New Artist Name\n"
                "/edit_artwork <ID> price 350000\n\n"
                "Example: /edit_artwork 3 title In & Out"
            )
        artwork_id = int(context.args[0])
        field = context.args[1].lower().strip()
        new_value = " ".join(context.args[2:]).strip()

        if field not in {"title", "artist", "price"}:
            raise ValueError("Field must be one of: title, artist, price")
        if not new_value:
            raise ValueError("New value cannot be empty.")

        with connect() as conn:
            row = conn.execute("SELECT * FROM artworks WHERE id = ?", (artwork_id,)).fetchone()
            if not row:
                raise ValueError(f"Artwork #{artwork_id} not found.")
            row = dict(row)
            if row.get("status") == "sold":
                raise ValueError(f"Artwork #{artwork_id} is already sold and cannot be edited.")

            if field == "price":
                price = _parse_float(new_value, "Price")
                conn.execute("UPDATE artworks SET asking_price_thb = ? WHERE id = ?", (price, artwork_id))
                log_action("edit_artwork", row["exhibition_code"], f"Updated artwork #{artwork_id} price to {money(price)}")
                await update.message.reply_text(
                    f"Artwork #{artwork_id} updated.\nTitle: {row['title']}\nArtist: {row['artist']}\nNew Asking Price: {money(price)}",
                    reply_markup=_artwork_menu_keyboard(),
                )
            elif field == "title":
                conn.execute("UPDATE artworks SET title = ? WHERE id = ?", (new_value, artwork_id))
                log_action("edit_artwork", row["exhibition_code"], f"Updated artwork #{artwork_id} title from '{row['title']}' to '{new_value}'")
                await update.message.reply_text(
                    f"Artwork #{artwork_id} updated.\nNew Title: {new_value}\nArtist: {row['artist']}\nAsking Price: {money(row['asking_price_thb'])}",
                    reply_markup=_artwork_menu_keyboard(),
                )
            elif field == "artist":
                conn.execute("UPDATE artworks SET artist = ? WHERE id = ?", (new_value, artwork_id))
                log_action("edit_artwork", row["exhibition_code"], f"Updated artwork #{artwork_id} artist from '{row['artist']}' to '{new_value}'")
                await update.message.reply_text(
                    f"Artwork #{artwork_id} updated.\nTitle: {row['title']}\nNew Artist: {new_value}\nAsking Price: {money(row['asking_price_thb'])}",
                    reply_markup=_artwork_menu_keyboard(),
                )
    except Exception as exc:
        logger.exception("Failed to edit artwork")
        await update.message.reply_text(f"Could not edit artwork: {exc}")


async def artworks(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_artworks_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_artwork_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not list artworks for {code}: {exc}")


async def inventory(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_inventory_dashboard_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_artwork_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not show inventory dashboard for {code}: {exc}")


async def sold(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if len(context.args) < 2:
            raise ValueError("Usage: /sold <ARTWORK_ID> <ACTUAL_PRICE_THB> [BUYER] [COLLECTED_THB]")
        artwork_id = int(context.args[0])
        actual_price = _parse_float(context.args[1], "Actual sale price")
        buyer_name = None
        collected = None
        if len(context.args) >= 3:
            if _looks_number(context.args[-1]) and len(context.args) >= 4:
                collected = _parse_float(context.args[-1], "Collected amount")
                buyer_name = " ".join(context.args[2:-1]).strip() or None
            else:
                buyer_name = " ".join(context.args[2:]).strip() or None
        result = record_sale(artwork_id, actual_price, buyer_name=buyer_name, amount_collected_thb=collected)
        _set_current_exhibition(context, result["sale"]["exhibition_code"])
        await update.message.reply_text(format_sale_markdown(result), parse_mode=ParseMode.MARKDOWN, reply_markup=_artwork_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to record sale")
        await update.message.reply_text(f"Could not record sale: {exc}")


async def receipt(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if not context.args:
            raise ValueError("Usage: /receipt <EXHIBITION_CODE> <AMOUNT_THB> <DESCRIPTION>")
        code, raw_text = _extract_receipt_code_and_text(context, " ".join(context.args))
        pending = create_pending_expense(code, raw_text)
        await update.message.reply_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending["id"]))
    except Exception as exc:
        logger.exception("Failed to create receipt")
        await update.message.reply_text(f"Could not create receipt: {exc}")


async def pending(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_pending_expenses_markdown(code), reply_markup=_expense_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not show pending receipts for {code}: {exc}")


async def expense_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_expense_report_markdown(code), reply_markup=_expense_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to generate expense report")
        await update.message.reply_text(f"Could not generate expense report for {code}: {exc}")


async def accounts(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(format_account_heads_markdown(), reply_markup=_expense_menu_keyboard())


async def budget(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        if len(context.args) >= 3:
            code = normalize_code(context.args[0])
            amount = _parse_float(context.args[-1], "Budget amount")
            account = _account_name_from_text(" ".join(context.args[1:-1]))
            set_expense_budget(code, account, amount)
            _set_current_exhibition(context, code)
            await update.message.reply_text(format_budget_report_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
            return
        code = _get_code(context)
        await update.message.reply_text(format_budget_report_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not update or show budget report: {exc}")


async def summary(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        text = format_executive_summary_markdown(code)
        log_action("telegram_summary", code, "Generated executive summary in Telegram")
        await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to generate executive summary")
        await update.message.reply_text(f"Could not generate executive summary for {code}: {exc}")


async def pl(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        text = format_report_markdown(code)
        log_action("telegram_pl", code, "Generated P&L in Telegram")
        await _send_long_text(update.message, text, parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to generate P&L")
        await update.message.reply_text(f"Could not generate P&L for {code}: {exc}")


async def artist_payouts(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        text = format_artist_payables_markdown(code)
        log_action("telegram_artist_payouts", code, "Generated artist payout summary in Telegram")
        await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to generate artist payouts")
        await update.message.reply_text(f"Could not generate artist payouts for {code}: {exc}")


async def data_check(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        warnings = data_quality_checks(code)
        text = "Data quality check — " + code + "\n\n" + "\n".join(f"• {w}" for w in warnings)
        log_action("telegram_data_check", code, "Generated data quality check")
        await update.message.reply_text(text, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to run data check")
        await update.message.reply_text(f"Could not run data check for {code}: {exc}")


async def readiness(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        await update.message.reply_text(format_readiness_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        await update.message.reply_text(f"Could not run readiness check for {code}: {exc}")


async def export(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = _get_code(context)
    try:
        export_dir = os.environ.get("EXPORT_DIR", "./exports")
        file_path = export_report_xlsx(code, export_dir)
        with open(file_path, "rb") as f:
            await update.message.reply_document(document=f, filename=Path(file_path).name)
    except Exception as exc:
        logger.exception("Failed to export report")
        await update.message.reply_text(f"Could not export report for {code}: {exc}")


async def pay_artist(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Mark an artist payable as paid (partially or in full).

    Usage:
        /pay_artist <EXHIBITION_CODE> <ARTIST_NAME> <AMOUNT_THB>
    Example:
        /pay_artist SHWEDAGON2024 "Aye Nyein Myint" 25000
        /pay_artist CLASH090526 "Aung Myint" 50000
    """
    try:
        if len(context.args) < 3:
            raise ValueError(
                "Usage: /pay_artist <CODE> <ARTIST NAME> <AMOUNT THB>\n"
                "Example: /pay_artist SHWEDAGON2024 \"U Lu Min\" 126000\n\n"
                "Use /artist_payouts <CODE> to see outstanding amounts."
            )
        code = normalize_code(context.args[0])
        # Last arg is amount, everything between code and amount is artist name
        amount = _parse_float(context.args[-1], "Payment amount")
        artist_name = " ".join(context.args[1:-1]).strip().strip('"').strip("'")
        if not artist_name:
            raise ValueError("Artist name cannot be empty.")
        if amount <= 0:
            raise ValueError("Payment amount must be greater than zero.")

        with connect() as conn:
            row = conn.execute(
                "SELECT * FROM artist_payables WHERE exhibition_code=? AND LOWER(artist)=LOWER(?)",
                (code, artist_name),
            ).fetchone()
            if not row:
                # Try partial match
                row = conn.execute(
                    "SELECT * FROM artist_payables WHERE exhibition_code=? AND LOWER(artist) LIKE LOWER(?)",
                    (code, f"%{artist_name}%"),
                ).fetchone()
            if not row:
                raise ValueError(
                    f"Artist '{artist_name}' not found in {code}.\n"
                    f"Use /artist_payouts {code} to see exact artist names."
                )
            row = dict(row)
            new_paid = round(row["paid_thb"] + amount, 2)
            new_outstanding = round(row["artist_payable_thb"] - new_paid, 2)
            if new_paid > row["artist_payable_thb"]:
                raise ValueError(
                    f"Payment of {money(amount)} exceeds outstanding amount of {money(row['outstanding_thb'])}.\n"
                    "If this is intentional, adjust the artist payable first."
                )
            new_status = "Paid" if new_outstanding <= 0 else "Partial"
            conn.execute(
                "UPDATE artist_payables SET paid_thb=?, outstanding_thb=?, status=? WHERE id=?",
                (new_paid, max(new_outstanding, 0), new_status, row["id"]),
            )
            log_action("pay_artist", code, f"Paid {money(amount)} to {row['artist']}. Paid: {money(new_paid)}, Outstanding: {money(max(new_outstanding, 0))}")

        lines = [
            f"Payment recorded for {row['artist']}.",
            f"Exhibition: {code}",
            f"This payment: {money(amount)}",
            f"Total paid: {money(new_paid)} / {money(row['artist_payable_thb'])}",
            f"Outstanding: {money(max(new_outstanding, 0))}",
            f"Status: {new_status}",
        ]
        await update.message.reply_text("\n".join(lines), reply_markup=_reports_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to record artist payment")
        await update.message.reply_text(f"Could not record payment: {exc}")


async def sheets_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        text = format_sheets_status_markdown()
        log_action("telegram_sheets_status", None, "Checked Google Sheets read-only setup")
        await update.message.reply_text(text, reply_markup=_help_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to check Google Sheets setup")
        await update.message.reply_text(f"Could not check Google Sheets setup: {exc}")


async def sync_preview(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        text = format_sync_preview_markdown(limit_rows=3)
        log_action("telegram_sync_preview", None, "Previewed Google Sheets workbook in read-only mode")
        await update.message.reply_text(text, reply_markup=_help_menu_keyboard())
    except Exception as exc:
        logger.exception("Failed to preview Google Sheets workbook")
        await update.message.reply_text(
            "Google Sheets preview is not ready yet. Please set GOOGLE_SHEETS_SPREADSHEET_ID, GOOGLE_APPLICATION_CREDENTIALS, "
            f"and share the sheet with the service-account email. Detail: {exc}"
        )


# ---------------------------------------------------------------------------
# Guided callback and conversation flows
# ---------------------------------------------------------------------------


def _flow_prompt(flow_name: str, context: ContextTypes.DEFAULT_TYPE) -> str:
    code = _current_exhibition(context)
    prompts = {
        "new_exhibition": (
            "Add new exhibition.\n"
            "Send: CODE | Name | Location | Start date | End date | Notes\n"
            "Only CODE | Name is required.\n"
            "Example: BKK2027 | Bangkok Art Fair 2027 | Bangkok | 2027-03-01 | 2027-03-15"
        ),
        "use_exhibition": "Switch exhibition. Send the exhibition code, or tap one below.",
        "custom_split": (
            f"Set custom split for {code}.\n"
            "Send: gallery <pct> collaborator <name> <pct> artist <pct>\n"
            "Example: gallery 45 collaborator Supples 25 artist 30\n"
            "Percentages must total 100%."
        ),
        "add_artwork": (
            f"Register artwork for {code}.\n"
            "Send: Title | Artist | Asking Price THB\n"
            "Example: In & Out | Aung Myint | 304000"
        ),
        "record_sale": (
            "Record sale.\n"
            "Send: Artwork ID | Sale Price | Buyer | Collected Amount | Payment Method | Notes\n"
            "Only Artwork ID and Sale Price are required.\n"
            "Example: 3 | 304000 | John Smith | 150000 | Bank Transfer | Balance due June"
        ),
        "text_receipt": (
            f"Add text receipt for {code}.\n"
            "Send: Amount Description\n"
            "Example: 3500 coffee and snacks opening night"
        ),
        "set_budget": (
            f"Set budget for {code}.\n"
            "Send: Account Head | Amount THB\n"
            "You may use the account number from /accounts.\n"
            "Example: Venue Rental | 50000"
        ),
        "bulk_artwork": (
            "Upload the filled Excel template to register all artworks at once.\n"
            "Use /cancel if you want to stop."
        ),
    }
    return prompts.get(flow_name, f"Guided action: {flow_name}. Send your input or use /cancel to stop.")


async def menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    try:
        await query.answer()
    except BadRequest as e:
        if "Query is too old" in str(e) or "query id is invalid" in str(e).lower():
            logger.warning("Callback query too old — ignoring stale button tap")
            return
        else:
            raise
    data = query.data or ""
    try:
        if data == "menu:home":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Main menu. Current exhibition: {label}", reply_markup=_main_menu_keyboard())
            return
        if data == "menu:exhibitions":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Exhibitions. Current: {label}", reply_markup=_exhibition_menu_keyboard())
            return
        if data == "menu:splits":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Commission splits. Current exhibition: {label}", reply_markup=_split_menu_keyboard())
            return
        if data == "menu:artworks":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Artworks and sales. Current exhibition: {label}", reply_markup=_artwork_menu_keyboard())
            return
        if data == "menu:expenses":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Receipts and expenses. Current exhibition: {label}", reply_markup=_expense_menu_keyboard())
            return
        if data == "menu:reports":
            _clear_flow(context)
            code = _current_exhibition(context)
            ex = get_exhibition(code)
            label = f"{code} — {ex['name']}" if ex else code
            await query.edit_message_text(f"Reports and export. Current exhibition: {label}", reply_markup=_reports_menu_keyboard())
            return
        if data == "menu:help":
            _clear_flow(context)
            await query.edit_message_text("Help and settings.", reply_markup=_help_menu_keyboard())
            return

        if data.startswith("flow:"):
            flow_name = data.split(":", 1)[1]
            _set_flow(context, flow_name)
            reply_markup = _exhibition_picker_keyboard() if flow_name == "use_exhibition" else _back_home_keyboard()
            await query.edit_message_text(_flow_prompt(flow_name, context), reply_markup=reply_markup)
            return

        if data.startswith("artwork:"):
            action = data.split(":", 1)[1]
            if action == "register_options":
                await query.edit_message_text(
                    "Register artwork — choose input method:",
                    reply_markup=_artwork_register_keyboard(),
                )
                return
            if action == "bulk_import":
                _set_flow(context, "bulk_artwork")
                try:
                    template_path = generate_artwork_template_xlsx()
                    await query.edit_message_text(
                        "Bulk artwork import:\n\n"
                        "1. Download the template file below\n"
                        "2. Fill in your artworks (delete the 3 sample rows first)\n"
                        "3. Upload the filled file back to this chat\n\n"
                        "Artwork titles with special characters like & work fine."
                    )
                    with open(template_path, "rb") as f:
                        await query.message.reply_document(
                            document=f,
                            filename="artwork_import_template.xlsx",
                            caption="Fill this in and upload it back here to register all artworks at once.",
                        )
                except Exception as exc:
                    await query.message.reply_text(f"Could not generate template: {exc}")
                return
            if action == "edit_hint":
                await query.edit_message_text(
                    "To edit an artwork, use the command:\n\n"
                    "/edit_artwork <ID> title New Title\n"
                    "/edit_artwork <ID> artist New Artist Name\n"
                    "/edit_artwork <ID> price 350000\n\n"
                    "Check artwork IDs with List Artworks.",
                    reply_markup=_artwork_menu_keyboard(),
                )
                return

        if data.startswith("useexh:"):
            code = normalize_code(data.split(":", 1)[1])
            ex = get_exhibition(code)
            if not ex:
                raise ValueError(f"Exhibition not found: {code}")
            _set_current_exhibition(context, code)
            _clear_flow(context)
            await query.edit_message_text(f"Current working exhibition set to {code} — {ex['name']}.", reply_markup=_main_menu_keyboard())
            return

        if data.startswith("preset_split:"):
            code = _current_exhibition(context)
            preset = data.split(":", 1)[1]
            if preset == "5050":
                entries = [{"party_type": "gallery", "party_name": "Gallery", "percent": 50}, {"party_type": "artist", "party_name": "Artist", "percent": 50}]
            elif preset == "451045":
                entries = [
                    {"party_type": "gallery", "party_name": "Gallery", "percent": 45},
                    {"party_type": "collaborator", "party_name": "Collaborator", "percent": 10},
                    {"party_type": "artist", "party_name": "Artist", "percent": 45},
                ]
            else:
                raise ValueError("Unknown split preset.")
            set_commission_splits(code, entries)
            await query.edit_message_text(format_split_rules_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_split_menu_keyboard())
            return

        if data.startswith("quick:"):
            quick = data.split(":", 1)[1]
            if quick == "current":
                code = _current_exhibition(context)
                ex = get_exhibition(code)
                text = f"Current working exhibition: {code}" + (f" — {ex['name']}" if ex else "")
                await query.edit_message_text(text, reply_markup=_main_menu_keyboard())
                return
            if quick == "list_exhibitions":
                rows = list_exhibitions()
                text = "Choose an exhibition:\n\n" + "\n".join(f"• {row['code']} — {row['name']}" for row in rows[:20]) if rows else "No exhibitions found."
                await query.edit_message_text(text, reply_markup=_exhibition_picker_keyboard())
                return
            if quick == "command_guide":
                await query.edit_message_text("Command guide sent below.", reply_markup=_help_menu_keyboard())
                await query.message.reply_text(_command_guide_text())
                return

        if data.startswith("report:"):
            await _handle_report_callback(query, context, data.split(":", 1)[1])
            return

        raise ValueError(f"Unknown menu action: {data}")
    except Exception as exc:
        logger.exception("Menu callback failed for data=%s", data)
        try:
            await query.edit_message_text(f"Could not complete action: {exc}", reply_markup=_main_menu_keyboard())
        except Exception:
            try:
                await query.message.reply_text(f"Could not complete action: {exc}", reply_markup=_main_menu_keyboard())
            except Exception:
                pass


async def _handle_report_callback(query, context: ContextTypes.DEFAULT_TYPE, report_name: str) -> None:
    code = _current_exhibition(context)
    parse_mode = None
    if report_name == "split":
        text = format_split_rules_markdown(code)
        keyboard = _split_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "artworks":
        text = format_artworks_markdown(code)
        keyboard = _artwork_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "inventory":
        text = format_inventory_dashboard_markdown(code)
        keyboard = _artwork_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "pending":
        text = format_pending_expenses_markdown(code)
        keyboard = _expense_menu_keyboard()
    elif report_name == "expenses":
        text = format_expense_report_markdown(code)
        keyboard = _expense_menu_keyboard()
    elif report_name == "accounts":
        text = format_account_heads_markdown()
        keyboard = _expense_menu_keyboard()
    elif report_name == "summary":
        text = format_executive_summary_markdown(code)
        keyboard = _reports_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "pl":
        text = format_report_markdown(code)
        keyboard = _reports_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "artists":
        text = format_artist_payables_markdown(code)
        keyboard = _reports_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "budget":
        text = format_budget_report_markdown(code)
        keyboard = _reports_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "readiness":
        text = format_readiness_markdown(code)
        keyboard = _reports_menu_keyboard()
        parse_mode = ParseMode.MARKDOWN
    elif report_name == "data_check":
        warnings = data_quality_checks(code)
        text = "Data quality check — " + code + "\n\n" + "\n".join(f"• {w}" for w in warnings)
        keyboard = _reports_menu_keyboard()
    elif report_name == "export":
        await query.edit_message_text(f"Preparing Excel export for {code}...")
        export_dir = os.environ.get("EXPORT_DIR", "./exports")
        file_path = export_report_xlsx(code, export_dir)
        with open(file_path, "rb") as f:
            await query.message.reply_document(document=f, filename=Path(file_path).name)
        await query.message.reply_text("Export complete.", reply_markup=_reports_menu_keyboard())
        return
    elif report_name == "sheets_status":
        text = format_sheets_status_markdown()
        keyboard = _help_menu_keyboard()
    elif report_name == "sync_preview":
        text = format_sync_preview_markdown(limit_rows=3)
        keyboard = _help_menu_keyboard()
    else:
        raise ValueError("Unknown report action.")

    if len(text) <= 3900:
        await query.edit_message_text(text, parse_mode=parse_mode, reply_markup=keyboard)
    else:
        await query.edit_message_text("Report is long, so I sent it below.", reply_markup=keyboard)
        await _send_long_text(query.message, text, parse_mode=parse_mode)


async def _handle_guided_flow(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> bool:
    flow = context.user_data.get("flow")
    if not flow:
        return False
    flow_name = flow["name"]
    try:
        if flow_name == "new_exhibition":
            parts = _pipe_parts(text)
            if len(parts) < 2:
                raise ValueError("Send at least CODE | Name")
            row = create_exhibition(parts[0], parts[1], parts[2] if len(parts) > 2 else None, parts[3] if len(parts) > 3 else None, parts[4] if len(parts) > 4 else None, parts[5] if len(parts) > 5 else None)
            _set_current_exhibition(context, row["code"])
            _clear_flow(context)
            await update.message.reply_text(f"Created exhibition {row['code']} — {row['name']}. Current exhibition updated.", reply_markup=_main_menu_keyboard())
            return True

        if flow_name == "use_exhibition":
            code = normalize_code(text.strip())
            ex = get_exhibition(code)
            if not ex:
                raise ValueError(f"Exhibition not found: {code}")
            _set_current_exhibition(context, code)
            _clear_flow(context)
            await update.message.reply_text(f"Current working exhibition set to {code} — {ex['name']}.", reply_markup=_main_menu_keyboard())
            return True

        if flow_name == "custom_split":
            code, entries = _parse_split_tokens(_current_exhibition(context), text.split())
            set_commission_splits(code, entries)
            _clear_flow(context)
            await update.message.reply_text(format_split_rules_markdown(code), parse_mode=ParseMode.MARKDOWN, reply_markup=_split_menu_keyboard())
            return True

        if flow_name == "add_artwork":
            parts = _pipe_parts(text)
            if len(parts) != 3:
                raise ValueError("Send exactly: Title | Artist | Asking Price THB")
            row = add_artwork(_current_exhibition(context), parts[0], parts[1], _parse_float(parts[2], "Asking price"))
            _clear_flow(context)
            await update.message.reply_text(f"Artwork registered.\nID: #{row['id']}\nTitle: {row['title']}\nArtist: {row['artist']}\nAsking Price: {money(row['asking_price_thb'])}", reply_markup=_artwork_menu_keyboard())
            return True

        if flow_name == "record_sale":
            parts = _pipe_parts(text) if "|" in text else text.split(maxsplit=2)
            if len(parts) < 2:
                raise ValueError("Send at least: Artwork ID | Sale Price")
            artwork_id = int(parts[0].strip())
            sale_price = _parse_float(parts[1], "Sale price")
            buyer = parts[2].strip() if len(parts) > 2 else None
            collected = _parse_float(parts[3], "Collected amount") if len(parts) > 3 and parts[3].strip() else None
            method = parts[4].strip() if len(parts) > 4 else None
            notes = parts[5].strip() if len(parts) > 5 else None
            result = record_sale(artwork_id, sale_price, buyer_name=buyer, amount_collected_thb=collected, payment_method=method, notes=notes)
            _set_current_exhibition(context, result["sale"]["exhibition_code"])
            _clear_flow(context)
            await update.message.reply_text(format_sale_markdown(result), parse_mode=ParseMode.MARKDOWN, reply_markup=_artwork_menu_keyboard())
            return True

        if flow_name == "text_receipt":
            pending = create_pending_expense(_current_exhibition(context), text)
            _clear_flow(context)
            await update.message.reply_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending["id"]))
            return True

        if flow_name == "set_budget":
            parts = _pipe_parts(text)
            if len(parts) != 2:
                raise ValueError("Send exactly: Account Head | Amount THB")
            account = _account_name_from_text(parts[0])
            amount = _parse_float(parts[1], "Budget amount")
            set_expense_budget(_current_exhibition(context), account, amount)
            _clear_flow(context)
            await update.message.reply_text(format_budget_report_markdown(_current_exhibition(context)), parse_mode=ParseMode.MARKDOWN, reply_markup=_reports_menu_keyboard())
            return True

        if flow_name == "bulk_artwork":
            # User typed text while in bulk mode — remind them to upload the Excel file
            await update.message.reply_text(
                "I'm waiting for you to upload the filled Excel template.\n\n"
                "Please attach the .xlsx file. Use /cancel to stop."
            )
            return True
    except Exception as exc:
        await update.message.reply_text(f"Could not complete guided action: {exc}\n\n{_flow_prompt(flow_name, context)}")
        return True
    return False


# ---------------------------------------------------------------------------
# Receipt approval callbacks and passive receipt capture
# ---------------------------------------------------------------------------


async def expense_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    try:
        await query.answer()
    except BadRequest as e:
        if "Query is too old" in str(e) or "query id is invalid" in str(e).lower():
            logger.warning("Expense callback query too old — ignoring stale button tap")
            return
        else:
            raise
    try:
        parts = query.data.split(":")
        action = parts[1]
        pending_id = int(parts[2])

        if action == "confirm":
            confirmed = confirm_pending_expense(pending_id)
            await query.edit_message_text(
                f"Expense confirmed and posted to P&L.\n"
                f"Expense ID: #{confirmed['id']}\n"
                f"Exhibition: {confirmed['exhibition_code']}\n"
                f"Account Head: {confirmed['account_head']}\n"
                f"Amount: {money(confirmed['amount_thb'])}",
                reply_markup=_expense_menu_keyboard(),
            )
            return

        if action == "ignore":
            row = ignore_pending_expense(pending_id)
            await query.edit_message_text(f"Pending receipt #{row['id']} was ignored. It was not posted to the P&L.", reply_markup=_expense_menu_keyboard())
            return

        if action == "account":
            await query.edit_message_text(f"Choose account head for pending receipt #{pending_id}:", reply_markup=_account_keyboard(pending_id))
            return

        if action == "setacct":
            account_idx = int(parts[3])
            names = account_head_names()
            if account_idx < 0 or account_idx >= len(names):
                raise ValueError("Account index out of range.")
            pending = update_pending_account(pending_id, names[account_idx])
            await query.edit_message_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending_id))
            return

        if action == "amount":
            context.user_data["awaiting_amount_for_pending_id"] = pending_id
            await query.edit_message_text(f"Send the corrected THB amount for pending receipt #{pending_id}. Example: 3500")
            return

        if action == "back":
            pending = get_pending_expense(pending_id)
            if not pending:
                raise ValueError(f"Pending expense not found: {pending_id}")
            await query.edit_message_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending_id))
            return

        raise ValueError(f"Unknown callback action: {action}")
    except Exception as exc:
        logger.exception("Expense callback failed")
        try:
            await query.edit_message_text(f"Could not update receipt: {exc}", reply_markup=_expense_menu_keyboard())
        except Exception:
            try:
                await query.message.reply_text(f"Could not update receipt: {exc}", reply_markup=_expense_menu_keyboard())
            except Exception:
                pass


async def handle_text_expense(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = update.message.text or ""
    if text.strip().lower() in {"cancel", "stop", "back"}:
        _clear_flow(context)
        context.user_data.pop("awaiting_amount_for_pending_id", None)
        await update.message.reply_text("Cancelled current action.", reply_markup=_main_menu_keyboard())
        return

    if context.user_data.get("awaiting_amount_for_pending_id"):
        pending_id = int(context.user_data.pop("awaiting_amount_for_pending_id"))
        try:
            amount = parse_amount_thb(text)
            pending = update_pending_amount(pending_id, amount)
            await update.message.reply_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending_id))
        except Exception as exc:
            await update.message.reply_text(f"Could not update amount: {exc}")
        return

    if await _handle_guided_flow(update, context, text):
        return

    # Smart detection: "Title | Artist | Price" typed outside a flow should never
    # be treated as a receipt (this is what caused "In & Out" to be classified as income).
    if text.count("|") == 2:
        parts = _pipe_parts(text)
        try:
            price = _parse_float(parts[2], "price")
            if price > 0 and parts[0].strip() and parts[1].strip():
                # Looks like artwork data — route to artwork registration
                _set_flow(context, "add_artwork")
                row = add_artwork(_current_exhibition(context), parts[0], parts[1], price)
                _clear_flow(context)
                await update.message.reply_text(
                    f"Artwork registered from your message.\n"
                    f"ID: #{row['id']}\nTitle: {row['title']}\nArtist: {row['artist']}\nAsking Price: {money(row['asking_price_thb'])}\n\n"
                    "If this was meant as an expense instead, use /pending to review receipts.",
                    reply_markup=_artwork_menu_keyboard(),
                )
                return
        except Exception:
            pass  # Not a valid artwork pattern — fall through to receipt handling

    if parse_amount_thb(text) <= 0:
        await update.message.reply_text(
            "I did not find a THB amount. Tap /menu to choose an action, or record an expense by sending for example: "
            f"{_current_exhibition(context)} 3500 coffee and snacks.",
            reply_markup=_main_menu_keyboard(),
        )
        return

    try:
        code, raw_text = _extract_receipt_code_and_text(context, text)
        pending = create_pending_expense(code, raw_text)
        await update.message.reply_text(format_pending_expense_card(pending), reply_markup=_pending_keyboard(pending["id"]))
    except Exception as exc:
        logger.exception("Failed to capture text expense")
        await update.message.reply_text(f"Could not capture this expense: {exc}")


async def _process_receipt_image(photo_file) -> str:
    """Extract text from a photo using OCR or AI."""
    # Read image into memory
    image_bytes = await photo_file.download_as_bytearray()
    image = Image.open(io.BytesIO(image_bytes))

    # Try AI first if available
    if ai_client:
        try:
            # Convert to base64 for GPT-4o-mini
            buffered = io.BytesIO()
            image.save(buffered, format="JPEG")
            base64_image = base64.b64encode(buffered.getvalue()).decode("utf-8")

            response = ai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "Extract all text from this receipt. Focus on the total amount and items."},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}},
                        ],
                    }
                ],
                max_tokens=300,
            )
            ai_text = response.choices[0].message.content
            if ai_text:
                return ai_text
        except Exception as e:
            logger.warning(f"AI receipt extraction failed: {e}. Falling back to Tesseract.")

    # Fallback to Tesseract OCR
    try:
        return pytesseract.image_to_string(image)
    except Exception as e:
        logger.error(f"Tesseract OCR failed: {e}")
        return ""


async def handle_photo_expense(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        # Send a "processing" message
        status_msg = await update.message.reply_text("Scanning receipt... 🔍")

        caption = update.message.caption or ""
        photo = update.message.photo[-1]
        photo_file = await photo.get_file()

        # Extract text from image
        ocr_text = await _process_receipt_image(photo_file)

        # Combine caption and OCR text
        combined_text = f"{caption}\n\n{ocr_text}".strip()

        code, raw_text = _extract_receipt_code_and_text(context, combined_text)
        photo_file_id = photo.file_id

        pending = create_pending_expense(code, raw_text, photo_file_id=photo_file_id)

        # Delete status message and send result
        try:
            await status_msg.delete()
        except Exception:
            pass  # Delete is cosmetic — don't crash if it fails
        await update.message.reply_text(
            format_pending_expense_card(pending),
            reply_markup=_pending_keyboard(pending["id"])
        )
    except Exception as exc:
        logger.exception("Failed to capture photo receipt")
        await update.message.reply_text(f"Could not capture this receipt photo: {exc}")


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Log errors and notify the user where possible."""
    logger.error("Exception while handling an update:", exc_info=context.error)

    # These are network/infrastructure errors — no user message needed
    if isinstance(context.error, (Conflict, NetworkError, TimedOut)):
        if isinstance(context.error, Conflict):
            logger.error("Conflict: another bot instance may be running.")
        return

    # Bot was blocked — nothing to do
    if isinstance(context.error, Forbidden):
        logger.warning("Bot was blocked by the user.")
        return

    # For all other errors, try to notify the user
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "Something went wrong. Please try again or tap /menu to restart."
            )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Application registration
# ---------------------------------------------------------------------------


def _seed_shwedagon_if_missing() -> None:
    """Seed the Shwe Dagon exhibition using raw sqlite3 — works even if exhibitledger path differs."""
    CODE = "SHWEDAGON2024"
    conversion_rate = float(os.environ.get("SEED_MMK_TO_THB_RATE", "0.006666666666666667"))
    db_path = os.environ.get("DB_PATH", "./exhibitledger.db")

    def thb(mmk: float) -> float:
        return round(mmk * conversion_rate, 2)

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    try:
        conn.execute(
            """INSERT INTO exhibitions
               (code, name, location, start_date, end_date, status, currency, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (CODE, "Shwe Dagon Platform Exhibition", "Bangkok / Yangon logistics",
             "2024-09-01", "2024-09-28", "completed", "THB",
             f"Auto-seeded. MMK lines at rate={conversion_rate}. "
             "Artist THB prices from Sheet2 artist list."),
        )

        pnl_lines = [
            ("sales_bridge",      "Gross artwork sales",                  "Sales of paintings — Sheet1",                             48_096_000, 10),
            ("gallery_revenue",   "Gallery artwork revenue",              "Gross revenue — Sheet1. 50% commission basis.",           48_096_000, 20),
            ("direct_cost",       "Artists' fees / artist share",         "50% artist share — Sheet1",                               25_708_000, 30),
            ("direct_cost",       "Rotary commission",                    "Partner commission — Sheet1",                                432_000, 31),
            ("direct_cost",       "Blank canvas",                         "Artwork preparation — Sheet1",                             1_357_000, 32),
            ("direct_cost",       "Catalog printing",                     "Catalog printing — Sheet1",                                1_200_000, 33),
            ("direct_cost",       "Local transportation (BKK paintings)", "Local artwork transport — Sheet1",                           336_150, 34),
            ("operating_expense", "Air cargo YGN→BKK",                   "Inbound exhibition logistics — Sheet1",                    2_355_000, 40),
            ("operating_expense", "Air cargo BKK→YGN",                   "Return exhibition logistics — Sheet1",                       495_000, 41),
            ("operating_expense", "Air tickets for CS",                   "Travel cost — Sheet1",                                     2_889_050, 42),
            ("operating_expense", "Rental of exhibition space",           "Venue rental — Sheet1",                                    3_932_250, 43),
            ("operating_expense", "Utensil renting",                      "Event supply rental — Sheet1",                               465_000, 44),
            ("operating_expense", "Coffee & snacks",                      "Opening/event hospitality — Sheet1",                       2_850_000, 45),
            ("operating_expense", "Photographer",                         "Photography cost — Sheet1",                                  495_000, 46),
        ]
        for section, category, description, amount_mmk, sort_order in pnl_lines:
            conn.execute(
                """INSERT INTO pnl_lines
                   (exhibition_code, section, category, description,
                    amount_thb, source_amount, source_currency, source_ref, sort_order)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (CODE, section, category, description,
                 thb(amount_mmk), amount_mmk, "MMK",
                 "22ShweDagonPlatformEstimatedP&L-Sheet1", sort_order),
            )

        artists = [
            ("U Lu Min",            2,  126_000),
            ("Zaw Win Phay",        2,  133_000),
            ("Min Zayar Oo",        2,   29_750),
            ("Kyi Hlaing Aung",     3,   29_750),
            ("Kaung Paing",         1,   22_750),
            ("Kyaw Lin",            2,   29_750),
            ("Aye Nyein Myint",     2,   29_750),
            ("Nu Nu",               3,   29_750),
            ("Ye Aung Myat",        2,   29_750),
            ("Orient Thant Zin",    1,   35_000),
            ("Maung Maung Yin Min", 1,   35_000),
            ("Myoe Kyaw",           2,   52_500),
            ("Aung Ko",             2,   28_000),
            ("Hla Phone Aung",      1,   28_000),
            ("Win Myint Moe",       4,   42_000),
            ("Aye Min",             2,   29_750),
            ("Nann Nann",           4,   61_250),
            ("CNK",                 3,   52_500),
            ("Ba Sai Wunna",        3,   35_000),
            ("Mann Zar Hein",       2,   35_000),
            ("Saw Lin Aung",        2,   70_000),
            ("Mor Mor",             2,  126_000),
            ("U Thu Won",           1,   28_000),
            ("U Hla Htun Aung",     2,   42_000),
            ("Thee Zar",            2,   70_000),
            ("Nyi Htut",            2,   70_000),
        ]
        for artist_name, num_paintings, unit_thb in artists:
            gross = round(num_paintings * unit_thb, 2)
            gallery_commission = round(gross * 0.50, 2)
            artist_payable = round(gross - gallery_commission, 2)
            conn.execute(
                """INSERT INTO artist_payables
                   (exhibition_code, artist, invoice_ref,
                    gross_sale_thb, gallery_commission_thb, artist_payable_thb,
                    paid_thb, outstanding_thb, status,
                    source_amount, source_currency, notes)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (CODE, artist_name, f"ShweDagon-{artist_name.replace(' ', '')}",
                 gross, gallery_commission, artist_payable,
                 0.0, artist_payable, "Pending",
                 gross, "THB",
                 f"{num_paintings} painting(s) at ฿{unit_thb:,.0f} each. "
                 "Sheet2 THB price. 50% gallery commission."),
            )

        conn.execute(
            "INSERT INTO audit_log (timestamp, action, exhibition_code, details) VALUES (?, ?, ?, ?)",
            (datetime.utcnow().isoformat(timespec="seconds") + "Z",
             "seed_shwedagon", CODE,
             f"{len(pnl_lines)} P&L lines + {len(artists)} artists seeded via raw sqlite3."),
        )
        conn.commit()
        logger.info("SHWEDAGON2024 seeded successfully into %s", db_path)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


async def reseed_shwedagon(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Force a full reseed of SHWEDAGON2024 — wipes old data and reloads all 26 artists."""
    await update.message.reply_text("Reseeding SHWEDAGON2024... please wait.")
    try:
        db_path = os.environ.get("DB_PATH", "./exhibitledger.db")
        raw_conn = sqlite3.connect(db_path)
        try:
            raw_conn.execute("DELETE FROM artist_payables WHERE exhibition_code = 'SHWEDAGON2024'")
            raw_conn.execute("DELETE FROM pnl_lines WHERE exhibition_code = 'SHWEDAGON2024'")
            raw_conn.execute("DELETE FROM exhibitions WHERE code = 'SHWEDAGON2024'")
            raw_conn.commit()
        finally:
            raw_conn.close()
        _seed_shwedagon_if_missing()
        await update.message.reply_text(
            "SHWEDAGON2024 reseeded successfully.\n"
            "26 artists and all P&L lines are now loaded.\n\n"
            "Use /export SHWEDAGON2024 to get the full Excel report."
        )
    except Exception as exc:
        logger.exception("Reseed failed")
        await update.message.reply_text(f"Reseed failed: {exc}")


def build_application() -> Application:
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not set. Create a bot with BotFather and set the token first.")
    init_db()

    # Auto-seed SHWEDAGON2024 on startup if missing or incomplete (< 5 artists = bad old seed).
    try:
        db_path = os.environ.get("DB_PATH", "./exhibitledger.db")
        _raw = sqlite3.connect(db_path)
        try:
            _ex = _raw.execute("SELECT code FROM exhibitions WHERE code='SHWEDAGON2024'").fetchone()
            _artist_count = _raw.execute(
                "SELECT COUNT(*) FROM artist_payables WHERE exhibition_code='SHWEDAGON2024'"
            ).fetchone()[0]
            if not _ex or _artist_count < 5:
                # Delete stale/partial data in the same connection before reseeding
                _raw.execute("DELETE FROM artist_payables WHERE exhibition_code='SHWEDAGON2024'")
                _raw.execute("DELETE FROM pnl_lines WHERE exhibition_code='SHWEDAGON2024'")
                _raw.execute("DELETE FROM exhibitions WHERE code='SHWEDAGON2024'")
                _raw.commit()
                _needs_seed = True
            else:
                _needs_seed = False
        finally:
            _raw.close()

        if _needs_seed:
            _seed_shwedagon_if_missing()
            logger.info("SHWEDAGON2024 seeded/reseeded on startup (had %d artists).", _artist_count if _ex else 0)
        else:
            logger.info("SHWEDAGON2024 already present with %d artists — skipping seed.", _artist_count)
    except Exception as _seed_err:
        logger.warning("Could not check/seed SHWEDAGON2024 on startup: %s", _seed_err)

    application = Application.builder().token(token).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("menu", menu_command))
    application.add_handler(CommandHandler("status", status_command))
    application.add_handler(CommandHandler("cancel", cancel_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("exhibitions", exhibitions))
    application.add_handler(CommandHandler("new_exhibition", new_exhibition))
    application.add_handler(CommandHandler("use", use_exhibition))
    application.add_handler(CommandHandler("set_split", set_split))
    application.add_handler(CommandHandler("split", split))
    application.add_handler(CommandHandler("add_artwork", add_artwork_command))
    application.add_handler(CommandHandler("edit_artwork", edit_artwork_command))
    application.add_handler(CommandHandler("artworks", artworks))
    application.add_handler(CommandHandler("inventory", inventory))
    application.add_handler(CommandHandler("sold", sold))
    application.add_handler(CommandHandler("receipt", receipt))
    application.add_handler(CommandHandler("pending", pending))
    application.add_handler(CommandHandler("expense_report", expense_report))
    application.add_handler(CommandHandler("accounts", accounts))
    application.add_handler(CommandHandler("budget", budget))
    application.add_handler(CommandHandler("summary", summary))
    application.add_handler(CommandHandler("pl", pl))
    application.add_handler(CommandHandler("artist_payouts", artist_payouts))
    application.add_handler(CommandHandler("pay_artist", pay_artist))
    application.add_handler(CommandHandler("readiness", readiness))
    application.add_handler(CommandHandler("data_check", data_check))
    application.add_handler(CommandHandler("export", export))
    application.add_handler(CommandHandler("sheets_status", sheets_status))
    application.add_handler(CommandHandler("sync_preview", sync_preview))
    application.add_handler(CommandHandler("reseed_shwedagon", reseed_shwedagon))
    application.add_handler(CallbackQueryHandler(expense_callback, pattern=r"^expense:"))
    application.add_handler(CallbackQueryHandler(menu_callback, pattern=r"^(menu:|flow:|artwork:|preset_split:|useexh:|quick:|report:)"))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document_upload))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo_expense))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_expense))

    # Register error handler
    application.add_error_handler(error_handler)

    return application


def main() -> None:
    start_render_health_server()
    app = build_application()
    logger.info("Starting ExhibitLedger THB bot in polling mode")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
