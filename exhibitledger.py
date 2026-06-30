import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


DEFAULT_DB_PATH = os.environ.get("DB_PATH") or os.environ.get("DATABASE_PATH") or "./exhibitledger.db"

ACCOUNT_HEADS = [
    {
        "name": "Transport & Local Logistics",
        "section": "direct_cost",
        "keywords": ["transport", "taxi", "truck", "van", "grab", "delivery", "logistics", "local", "moving"],
    },
    {
        "name": "Air Cargo & Freight",
        "section": "operating_expense",
        "keywords": ["air cargo", "freight", "cargo", "shipping", "courier", "dhl", "fedex", "ems", "export", "import"],
    },
    {
        "name": "Customs Duties & Taxes",
        "section": "operating_expense",
        "keywords": ["customs", "duty", "duties", "tax", "vat", "import tax", "export tax", "clearance"],
    },
    {
        "name": "Venue Rental",
        "section": "operating_expense",
        "keywords": ["venue", "rental", "rent", "space", "hall", "booth", "gallery hire"],
    },
    {
        "name": "Installation & Production",
        "section": "direct_cost",
        "keywords": ["installation", "install", "production", "lighting", "plinth", "wall", "label", "catalog", "printing", "setup"],
    },
    {
        "name": "Framing & Artwork Preparation",
        "section": "direct_cost",
        "keywords": ["frame", "framing", "canvas", "stretch", "varnish", "mount", "artwork preparation", "preparation"],
    },
    {
        "name": "Repairs & Conservation",
        "section": "direct_cost",
        "keywords": ["repair", "restore", "restoration", "conservation", "damage", "touch up", "fix"],
    },
    {
        "name": "Travel & Accommodation",
        "section": "operating_expense",
        "keywords": ["flight", "air ticket", "ticket", "hotel", "accommodation", "travel", "visa", "per diem", "taxi airport"],
    },
    {
        "name": "Insurance",
        "section": "operating_expense",
        "keywords": ["insurance", "insured", "policy", "premium", "coverage"],
    },
    {
        "name": "Security",
        "section": "operating_expense",
        "keywords": ["security", "guard", "cctv", "safety", "supervision"],
    },
    {
        "name": "Food & Beverage / Hospitality",
        "section": "operating_expense",
        "keywords": ["food", "drink", "coffee", "snack", "catering", "wine", "water", "hospitality", "beverage", "meal"],
    },
    {
        "name": "Opening Event / VIP Relations",
        "section": "operating_expense",
        "keywords": ["opening", "vip", "guest", "invitation", "ceremony", "reception", "flowers", "gift"],
    },
    {
        "name": "Marketing & PR",
        "section": "operating_expense",
        "keywords": ["marketing", "pr", "advert", "ads", "facebook", "instagram", "boost", "poster", "media", "press", "photographer", "photo"],
    },
    {
        "name": "Office & Admin Supplies",
        "section": "operating_expense",
        "keywords": ["office", "admin", "paper", "ink", "stationery", "supplies", "receipt book", "folder"],
    },
    {
        "name": "Staff & Helpers / Labor",
        "section": "operating_expense",
        "keywords": ["staff", "helper", "labor", "labour", "assistant", "wage", "salary", "overtime", "crew"],
    },
    {
        "name": "Banking & Payment Fees",
        "section": "operating_expense",
        "keywords": ["bank", "fee", "fees", "transfer", "payment", "credit card", "card", "promptpay", "charge"],
    },
    {
        "name": "Miscellaneous (Needs Review)",
        "section": "operating_expense",
        "keywords": [],
    },
]
ACCOUNT_HEAD_BY_NAME = {row["name"].lower(): row for row in ACCOUNT_HEADS}
ALLOWED_PARTY_TYPES = {"gallery", "artist", "collaborator", "collector"}


def db_path() -> str:
    return os.environ.get("DB_PATH") or os.environ.get("DATABASE_PATH") or DEFAULT_DB_PATH


@contextmanager
def connect(path: str | None = None):
    conn = sqlite3.connect(path or db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def _table_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}


def _utc_now() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def _insert_audit(conn: sqlite3.Connection, action: str, exhibition_code: str | None, details: str) -> None:
    conn.execute(
        "INSERT INTO audit_log (timestamp, action, exhibition_code, details) VALUES (?, ?, ?, ?)",
        (_utc_now(), action, exhibition_code, details),
    )


def init_db(path: str | None = None) -> None:
    """Create all legacy and workflow tables without destroying existing data."""
    with connect(path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS exhibitions (
                code TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                location TEXT,
                start_date TEXT,
                end_date TEXT,
                status TEXT DEFAULT 'active',
                currency TEXT DEFAULT 'THB',
                notes TEXT
            );

            CREATE TABLE IF NOT EXISTS pnl_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                section TEXT NOT NULL,
                category TEXT NOT NULL,
                description TEXT,
                amount_thb REAL NOT NULL DEFAULT 0,
                source_amount REAL,
                source_currency TEXT,
                source_ref TEXT,
                sort_order INTEGER DEFAULT 0,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS artist_payables (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                artist TEXT NOT NULL,
                invoice_ref TEXT,
                gross_sale_thb REAL NOT NULL DEFAULT 0,
                gallery_commission_thb REAL NOT NULL DEFAULT 0,
                artist_payable_thb REAL NOT NULL DEFAULT 0,
                paid_thb REAL NOT NULL DEFAULT 0,
                outstanding_thb REAL NOT NULL DEFAULT 0,
                status TEXT DEFAULT 'Pending',
                source_amount REAL,
                source_currency TEXT,
                notes TEXT,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                action TEXT NOT NULL,
                exhibition_code TEXT,
                details TEXT
            );

            CREATE TABLE IF NOT EXISTS commission_split_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                party_type TEXT NOT NULL,
                party_name TEXT NOT NULL,
                percent REAL NOT NULL,
                sort_order INTEGER DEFAULT 0,
                active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS artworks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                title TEXT NOT NULL,
                artist TEXT NOT NULL,
                asking_price_thb REAL NOT NULL DEFAULT 0,
                status TEXT DEFAULT 'available',
                created_at TEXT NOT NULL,
                sold_at TEXT,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS artwork_sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                artwork_id INTEGER NOT NULL,
                exhibition_code TEXT NOT NULL,
                actual_price_thb REAL NOT NULL,
                sale_date TEXT NOT NULL,
                gallery_share_thb REAL NOT NULL DEFAULT 0,
                collaborator_share_thb REAL NOT NULL DEFAULT 0,
                artist_payable_thb REAL NOT NULL DEFAULT 0,
                split_summary TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY (artwork_id) REFERENCES artworks(id),
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS sale_allocations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id INTEGER NOT NULL,
                exhibition_code TEXT NOT NULL,
                artwork_id INTEGER NOT NULL,
                party_type TEXT NOT NULL,
                party_name TEXT NOT NULL,
                percent REAL NOT NULL,
                amount_thb REAL NOT NULL,
                FOREIGN KEY (sale_id) REFERENCES artwork_sales(id),
                FOREIGN KEY (artwork_id) REFERENCES artworks(id),
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS pending_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                raw_text TEXT,
                description TEXT,
                suggested_amount_thb REAL NOT NULL DEFAULT 0,
                suggested_account_head TEXT NOT NULL,
                suggested_section TEXT NOT NULL,
                status TEXT DEFAULT 'pending',
                photo_file_id TEXT,
                created_at TEXT NOT NULL,
                confirmed_at TEXT,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS confirmed_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                account_head TEXT NOT NULL,
                section TEXT NOT NULL,
                description TEXT,
                amount_thb REAL NOT NULL,
                receipt_ref TEXT,
                raw_text TEXT,
                pending_expense_id INTEGER,
                created_at TEXT NOT NULL,
                pnl_line_id INTEGER,
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code),
                FOREIGN KEY (pending_expense_id) REFERENCES pending_expenses(id),
                FOREIGN KEY (pnl_line_id) REFERENCES pnl_lines(id)
            );


            CREATE TABLE IF NOT EXISTS expense_budgets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                exhibition_code TEXT NOT NULL,
                account_head TEXT NOT NULL,
                section TEXT NOT NULL,
                budget_thb REAL NOT NULL DEFAULT 0,
                notes TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(exhibition_code, account_head),
                FOREIGN KEY (exhibition_code) REFERENCES exhibitions(code)
            );

            CREATE TABLE IF NOT EXISTS user_states (
                chat_id INTEGER PRIMARY KEY,
                current_exhibition TEXT,
                active_flow TEXT,
                flow_step INTEGER DEFAULT 0,
                flow_data TEXT DEFAULT '{}',
                updated_at TEXT
            );
            """
        )

        sale_column_additions = {
            "buyer_name": "TEXT",
            "payment_status": "TEXT DEFAULT 'collected'",
            "amount_collected_thb": "REAL DEFAULT 0",
            "balance_due_thb": "REAL DEFAULT 0",
            "payment_method": "TEXT",
            "notes": "TEXT",
        }
        sale_columns = _table_columns(conn, "artwork_sales")
        for column, definition in sale_column_additions.items():
            if column not in sale_columns:
                conn.execute(f"ALTER TABLE artwork_sales ADD COLUMN {column} {definition}")

        artwork_column_additions = {
            "inventory_code": "TEXT",
            "medium": "TEXT",
            "dimensions": "TEXT",
            "year_created": "TEXT",
        }
        artwork_columns = _table_columns(conn, "artworks")
        for column, definition in artwork_column_additions.items():
            if column not in artwork_columns:
                conn.execute(f"ALTER TABLE artworks ADD COLUMN {column} {definition}")



def money(amount: float) -> str:
    return f"฿{float(amount or 0):,.2f} THB"


def compact_money(amount: float) -> str:
    return f"฿{float(amount or 0):,.0f}"


def normalize_code(code: str) -> str:
    cleaned = (code or "").strip().upper()
    if not cleaned:
        raise ValueError("Exhibition code is required.")
    return cleaned


def list_exhibitions() -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            "SELECT code, name, location, start_date, end_date, status, currency FROM exhibitions ORDER BY start_date DESC, code"
        ).fetchall()


def get_exhibition(code: str) -> sqlite3.Row | None:
    with connect() as conn:
        return conn.execute("SELECT * FROM exhibitions WHERE UPPER(code) = UPPER(?)", (code,)).fetchone()


def create_exhibition(
    code: str,
    name: str,
    location: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    notes: str | None = None,
) -> sqlite3.Row:
    code = normalize_code(code)
    name = (name or "").strip()
    if not name:
        raise ValueError("Exhibition name is required.")
    with connect() as conn:
        existing = conn.execute("SELECT code FROM exhibitions WHERE UPPER(code) = UPPER(?)", (code,)).fetchone()
        if existing:
            raise ValueError(f"Exhibition already exists: {code}")
        conn.execute(
            """
            INSERT INTO exhibitions (code, name, location, start_date, end_date, status, currency, notes)
            VALUES (?, ?, ?, ?, ?, 'active', 'THB', ?)
            """,
            (code, name, location, start_date, end_date, notes),
        )
        _insert_audit(conn, "create_exhibition", code, f"Created exhibition {name}")
        return conn.execute("SELECT * FROM exhibitions WHERE code = ?", (code,)).fetchone()


def get_lines(code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM pnl_lines
            WHERE UPPER(exhibition_code) = UPPER(?)
            ORDER BY sort_order, id
            """,
            (code,),
        ).fetchall()


def get_artist_payables(code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM artist_payables
            WHERE UPPER(exhibition_code) = UPPER(?)
            ORDER BY outstanding_thb DESC, artist
            """,
            (code,),
        ).fetchall()


def section_total(lines: List[sqlite3.Row], section: str) -> float:
    return sum(float(row["amount_thb"] or 0) for row in lines if row["section"] == section)


def calculate_report(code: str) -> Dict:
    exhibition = get_exhibition(code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {code}")

    lines = get_lines(code)
    payables = get_artist_payables(code)

    gross_sales = section_total(lines, "sales_bridge")
    gallery_revenue = section_total(lines, "gallery_revenue")
    direct_costs = section_total(lines, "direct_cost")
    operating_expenses = section_total(lines, "operating_expense")
    overhead = section_total(lines, "allocated_overhead")
    gross_profit = gallery_revenue - direct_costs
    contribution_profit = gross_profit - operating_expenses
    net_profit = contribution_profit - overhead

    artist_payable_total = sum(float(row["artist_payable_thb"] or 0) for row in payables)
    artist_paid_total = sum(float(row["paid_thb"] or 0) for row in payables)
    artist_outstanding_total = sum(float(row["outstanding_thb"] or 0) for row in payables)

    gross_margin_pct = (gross_profit / gallery_revenue * 100) if gallery_revenue else 0
    contribution_margin_pct = (contribution_profit / gallery_revenue * 100) if gallery_revenue else 0
    net_margin_pct = (net_profit / gallery_revenue * 100) if gallery_revenue else 0
    expense_ratio_pct = ((direct_costs + operating_expenses + overhead) / gallery_revenue * 100) if gallery_revenue else 0

    return {
        "exhibition": dict(exhibition),
        "lines": [dict(row) for row in lines],
        "payables": [dict(row) for row in payables],
        "totals": {
            "gross_sales": gross_sales,
            "gallery_revenue": gallery_revenue,
            "direct_costs": direct_costs,
            "operating_expenses": operating_expenses,
            "allocated_overhead": overhead,
            "gross_profit": gross_profit,
            "contribution_profit": contribution_profit,
            "net_profit": net_profit,
            "artist_payable_total": artist_payable_total,
            "artist_paid_total": artist_paid_total,
            "artist_outstanding_total": artist_outstanding_total,
            "gross_margin_pct": gross_margin_pct,
            "contribution_margin_pct": contribution_margin_pct,
            "net_margin_pct": net_margin_pct,
            "expense_ratio_pct": expense_ratio_pct,
        },
    }


def grouped_lines(report: Dict, section: str) -> List[Dict]:
    return [row for row in report["lines"] if row["section"] == section]


def format_report_markdown(code: str) -> str:
    report = calculate_report(code)
    ex = report["exhibition"]
    totals = report["totals"]

    def line_items(section: str) -> str:
        rows = grouped_lines(report, section)
        if not rows:
            return "_No lines recorded._\n"
        return "\n".join(f"• {row['category']}: {money(float(row['amount_thb']))}" for row in rows) + "\n"

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    text = [
        "*THE SEA ART GALLERY*",
        "*Profit & Loss Statement*",
        f"Exhibition: *{ex['name']}* (`{ex['code']}`)",
        f"Location: {ex['location'] or '-'}",
        f"Period: {ex['start_date'] or '-'} to {ex['end_date'] or '-'}",
        "Currency: *THB only*",
        f"Generated: {generated}",
        "",
        "*A. Sales Bridge*",
        line_items("sales_bridge"),
        f"Gross Artwork Sales / Activity: *{money(totals['gross_sales'])}*",
        "",
        "*B. Gallery Revenue*",
        line_items("gallery_revenue"),
        f"Total Gallery Revenue: *{money(totals['gallery_revenue'])}*",
        "",
        "*C. Direct Costs*",
        line_items("direct_cost"),
        f"Total Direct Costs: *{money(totals['direct_costs'])}*",
        "",
        f"*Gross Profit:* *{money(totals['gross_profit'])}* ({totals['gross_margin_pct']:.1f}%)",
        "",
        "*D. Exhibition Operating Expenses*",
        line_items("operating_expense"),
        f"Total Operating Expenses: *{money(totals['operating_expenses'])}*",
        "",
        f"*Contribution Profit:* *{money(totals['contribution_profit'])}* ({totals['contribution_margin_pct']:.1f}%)",
        "",
        "*E. Allocated Overhead*",
        line_items("allocated_overhead"),
        f"Allocated Overhead: *{money(totals['allocated_overhead'])}*",
        "",
        f"*NET PROFIT / (LOSS):* *{money(totals['net_profit'])}* ({totals['net_margin_pct']:.1f}%)",
        "",
        "*Artist Payable Control*",
        f"Artist Payable Total: *{money(totals['artist_payable_total'])}*",
        f"Paid: *{money(totals['artist_paid_total'])}*",
        f"Outstanding: *{money(totals['artist_outstanding_total'])}*",
        "",
        "_Note: This report is exhibition-by-exhibition and THB-only. Confirm receipt classifications and commission splits before relying on final numbers._",
    ]
    return "\n".join(text)


def format_artist_payables_markdown(code: str) -> str:
    exhibition = get_exhibition(code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {code}")
    payables = get_artist_payables(code)
    if not payables:
        return f"No artist payable rows found for `{code}`."

    lines = [f"*Artist Payables — {exhibition['name']}*", ""]
    for row in payables:
        lines.append(
            f"• *{row['artist']}* — Gross {compact_money(float(row['gross_sale_thb']))}; "
            f"Gallery commission {compact_money(float(row['gallery_commission_thb']))}; "
            f"Artist payable {compact_money(float(row['artist_payable_thb']))}; "
            f"Outstanding {compact_money(float(row['outstanding_thb']))}; Status: {row['status']}"
        )
    return "\n".join(lines)


def data_quality_checks(code: str) -> List[str]:
    report = calculate_report(code)
    warnings: List[str] = []
    ex = report["exhibition"]
    lines = report["lines"]
    payables = report["payables"]
    totals = report["totals"]

    if ex.get("currency") != "THB":
        warnings.append("Exhibition currency is not THB.")
    if not lines:
        warnings.append("No P&L lines recorded.")
    if not any(row["section"] == "gallery_revenue" for row in lines):
        warnings.append("No gallery revenue lines recorded.")
    if not any(row["section"] == "sales_bridge" for row in lines):
        warnings.append("No gross sales bridge line recorded.")
    if any((row.get("source_currency") or "THB") != "THB" for row in lines):
        warnings.append("Some rows were imported from non-THB sources. Confirm conversion rate before final use.")

    if payables:
        payable_outstanding_sum = sum(float(r["outstanding_thb"] or 0) for r in payables)
        payable_gross_sum = sum(float(r["gross_sale_thb"] or 0) for r in payables)
        payable_commission_sum = sum(float(r["gallery_commission_thb"] or 0) for r in payables)
        payable_artist_sum = sum(float(r["artist_payable_thb"] or 0) for r in payables)
        if abs(totals["artist_outstanding_total"] - payable_outstanding_sum) > 0.01:
            warnings.append("Artist payable outstanding total does not reconcile.")
        if totals["gross_sales"] and abs(totals["gross_sales"] - payable_gross_sum) > max(1.0, totals["gross_sales"] * 0.02):
            warnings.append("Artist commission records appear partial or unmapped; gross sales bridge does not yet reconcile to all artist sale records.")
        split_rows_for_reconcile = get_split_rules(code)
        collaborator_pct = sum(float(r["percent"] or 0) for r in split_rows_for_reconcile if str(r["party_type"]).lower() not in {"gallery", "artist"})
        expected_gallery_artist_share = payable_gross_sum * max(0.0, (100.0 - collaborator_pct)) / 100.0
        if abs((payable_commission_sum + payable_artist_sum) - expected_gallery_artist_share) > max(1.0, payable_gross_sum * 0.02):
            warnings.append("Gallery and artist portions do not reconcile to the active split rule.")

    split_rows = get_split_rules(code)
    if not split_rows:
        warnings.append("No active commission split rule has been set for this exhibition.")
    else:
        split_total = sum(float(r["percent"] or 0) for r in split_rows)
        if abs(split_total - 100) > 0.01:
            warnings.append(f"Commission split totals {split_total:.2f}%, not 100%.")

    pending_count = count_pending_expenses(code)
    if pending_count:
        warnings.append(f"There are {pending_count} pending expense receipt(s) awaiting approval.")

    metrics = calculate_inventory_metrics(code)
    if metrics["total_artworks"] == 0:
        warnings.append("No artwork inventory has been registered for this exhibition.")
    if metrics["receivables_thb"] > 0:
        warnings.append(f"There are uncollected or partially collected sale balances totaling {money(metrics['receivables_thb'])}.")
    if metrics["sold_artworks"] > 0 and metrics["total_artworks"] > 0 and metrics["sell_through_rate_pct"] < 30:
        warnings.append(f"Sell-through is currently {metrics['sell_through_rate_pct']:.1f}%; review whether final reporting should mention unsold inventory.")

    budget_rows = calculate_budget_report(code)
    over_budget = [row for row in budget_rows if row["budget_thb"] > 0 and row["variance_thb"] < 0]
    if over_budget:
        worst = sorted(over_budget, key=lambda row: row["variance_thb"])[0]
        warnings.append(f"Budget overrun: {worst['account_head']} is over budget by {money(abs(worst['variance_thb']))}.")

    if totals["gallery_revenue"] <= 0:
        warnings.append("Gallery revenue is zero or negative.")
    if totals["direct_costs"] < 0 or totals["operating_expenses"] < 0 or totals["allocated_overhead"] < 0:
        warnings.append("One or more cost sections are negative; confirm sign convention.")
    if totals["net_profit"] < 0:
        warnings.append("Exhibition currently reports a net loss under current assumptions.")
    if totals["expense_ratio_pct"] > 100:
        warnings.append("Total expense ratio is above 100% of gallery revenue; review cost classification.")
    if not warnings:
        warnings.append("No blocking issues found for current exhibition data.")
    return warnings


def format_executive_summary_markdown(code: str) -> str:
    report = calculate_report(code)
    ex = report["exhibition"]
    totals = report["totals"]
    metrics = calculate_inventory_metrics(code)
    checks = data_quality_checks(code)
    status = "PROFIT" if totals["net_profit"] >= 0 else "LOSS"
    text = [
        f"*Executive Dashboard — {ex['name']}*",
        f"Code: `{ex['code']}`",
        "Currency: *THB only*",
        "",
        f"P&L status: *{status}*",
        f"Gross Artwork Sales / Activity: *{money(totals['gross_sales'])}*",
        f"Gallery Revenue: *{money(totals['gallery_revenue'])}*",
        f"Direct Costs: *{money(totals['direct_costs'])}*",
        f"Operating Expenses: *{money(totals['operating_expenses'])}*",
        f"Net Profit / (Loss): *{money(totals['net_profit'])}*",
        f"Net Margin on Gallery Revenue: *{totals['net_margin_pct']:.1f}%*",
        "",
        "Inventory and collection controls:",
        f"• Artworks registered: {metrics['total_artworks']} | Sold: {metrics['sold_artworks']} | Available: {metrics['available_artworks']}",
        f"• Sell-through rate: {metrics['sell_through_rate_pct']:.1f}%",
        f"• Unsold asking value: {money(metrics['unsold_asking_value_thb'])}",
        f"• Cash collected from sales: {money(metrics['cash_collected_thb'])}",
        f"• Sale receivables outstanding: {money(metrics['receivables_thb'])}",
        f"• Pending receipts awaiting approval: {metrics['pending_receipts']}",
        "",
        "Control points:",
    ]
    text.extend(f"• {check}" for check in checks[:8])
    text.append("")
    text.append("Recommended next action: clear pending receipts, collect outstanding sale balances, verify account heads, and run /export before final review.")
    return "\n".join(text)


def log_action(action: str, exhibition_code: str | None, details: str) -> None:
    with connect() as conn:
        _insert_audit(conn, action, exhibition_code, details)


# ---------------------------------------------------------------------------
# Commission split, artwork, and sale workflow
# ---------------------------------------------------------------------------


def set_commission_splits(exhibition_code: str, entries: Sequence[Dict]) -> List[sqlite3.Row]:
    exhibition_code = normalize_code(exhibition_code)
    if not get_exhibition(exhibition_code):
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    if not entries:
        raise ValueError("At least one split entry is required.")

    cleaned = []
    for idx, entry in enumerate(entries, start=1):
        party_type = (entry.get("party_type") or "").strip().lower()
        party_name = (entry.get("party_name") or party_type.title()).strip()
        try:
            percent = float(entry.get("percent"))
        except (TypeError, ValueError):
            raise ValueError(f"Invalid percent for split entry {entry!r}")
        if party_type not in ALLOWED_PARTY_TYPES:
            raise ValueError(f"Unsupported party type: {party_type}. Use gallery, artist, collaborator, or collector.")
        if percent <= 0:
            raise ValueError("Split percentages must be greater than zero.")
        cleaned.append({"party_type": party_type, "party_name": party_name, "percent": percent, "sort_order": idx})

    total = sum(row["percent"] for row in cleaned)
    if abs(total - 100.0) > 0.01:
        raise ValueError(f"Split percentages must total 100%. Current total is {total:.2f}%.")

    with connect() as conn:
        conn.execute("DELETE FROM commission_split_rules WHERE UPPER(exhibition_code) = UPPER(?)", (exhibition_code,))
        for row in cleaned:
            conn.execute(
                """
                INSERT INTO commission_split_rules
                (exhibition_code, party_type, party_name, percent, sort_order, active, created_at)
                VALUES (?, ?, ?, ?, ?, 1, ?)
                """,
                (exhibition_code, row["party_type"], row["party_name"], row["percent"], row["sort_order"], _utc_now()),
            )
        _insert_audit(conn, "set_commission_splits", exhibition_code, format_split_summary(cleaned))
        return conn.execute(
            "SELECT * FROM commission_split_rules WHERE UPPER(exhibition_code) = UPPER(?) ORDER BY sort_order, id",
            (exhibition_code,),
        ).fetchall()


def get_split_rules(exhibition_code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM commission_split_rules
            WHERE UPPER(exhibition_code) = UPPER(?) AND active = 1
            ORDER BY sort_order, id
            """,
            (exhibition_code,),
        ).fetchall()


def format_split_summary(entries: Sequence[Dict | sqlite3.Row]) -> str:
    parts = []
    for row in entries:
        party_type = row["party_type"]
        party_name = row["party_name"]
        percent = float(row["percent"])
        parts.append(f"{party_type}:{party_name} {percent:g}%")
    return "; ".join(parts)


def format_split_rules_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    rows = get_split_rules(exhibition_code)
    if not rows:
        return f"No commission split rule has been set for `{normalize_code(exhibition_code)}`."
    total = sum(float(row["percent"] or 0) for row in rows)
    lines = [f"*Commission Split — {exhibition['name']}*", ""]
    for row in rows:
        lines.append(f"• {row['party_type'].title()} — {row['party_name']}: {float(row['percent']):g}%")
    lines.append("")
    lines.append(f"Total: *{total:g}%*")
    return "\n".join(lines)


def add_artwork(exhibition_code: str, title: str, artist: str, asking_price_thb: float) -> sqlite3.Row:
    exhibition_code = normalize_code(exhibition_code)
    if not get_exhibition(exhibition_code):
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    title = (title or "").strip()
    artist = (artist or "").strip()
    price = float(asking_price_thb)
    if not title or not artist:
        raise ValueError("Artwork title and artist are required.")
    if price < 0:
        raise ValueError("Artwork price cannot be negative.")
    with connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO artworks (exhibition_code, title, artist, asking_price_thb, status, created_at)
            VALUES (?, ?, ?, ?, 'available', ?)
            """,
            (exhibition_code, title, artist, price, _utc_now()),
        )
        artwork_id = cur.lastrowid
        _insert_audit(conn, "add_artwork", exhibition_code, f"Added artwork #{artwork_id}: {title} by {artist} at {money(price)}")
        return conn.execute("SELECT * FROM artworks WHERE id = ?", (artwork_id,)).fetchone()


def get_artwork(artwork_id: int) -> sqlite3.Row | None:
    with connect() as conn:
        return conn.execute("SELECT * FROM artworks WHERE id = ?", (int(artwork_id),)).fetchone()


def list_artworks(exhibition_code: str, include_sold: bool = True) -> List[sqlite3.Row]:
    with connect() as conn:
        if include_sold:
            return conn.execute(
                """
                SELECT * FROM artworks WHERE UPPER(exhibition_code) = UPPER(?)
                ORDER BY id
                """,
                (exhibition_code,),
            ).fetchall()
        return conn.execute(
            """
            SELECT * FROM artworks WHERE UPPER(exhibition_code) = UPPER(?) AND status <> 'sold'
            ORDER BY id
            """,
            (exhibition_code,),
        ).fetchall()


def format_artworks_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    rows = list_artworks(exhibition_code)
    if not rows:
        return f"No artworks have been registered for `{normalize_code(exhibition_code)}`."
    lines = [f"*Artworks — {exhibition['name']}*", ""]
    for row in rows:
        lines.append(
            f"• #{row['id']} — {row['title']} / {row['artist']} / Asking {compact_money(row['asking_price_thb'])} / {row['status']}"
        )
    return "\n".join(lines)


def _allocation_amounts(actual_price_thb: float, split_rows: Sequence[sqlite3.Row]) -> List[Dict]:
    allocations = []
    running = 0.0
    for idx, row in enumerate(split_rows):
        if idx == len(split_rows) - 1:
            amount = round(actual_price_thb - running, 2)
        else:
            amount = round(actual_price_thb * float(row["percent"]) / 100.0, 2)
            running += amount
        allocations.append(
            {
                "party_type": row["party_type"],
                "party_name": row["party_name"],
                "percent": float(row["percent"]),
                "amount_thb": amount,
            }
        )
    return allocations


def record_sale(
    artwork_id: int,
    actual_price_thb: float,
    sale_date: str | None = None,
    buyer_name: str | None = None,
    amount_collected_thb: float | None = None,
    payment_method: str | None = None,
    notes: str | None = None,
) -> Dict:
    artwork_id = int(artwork_id)
    actual_price = float(actual_price_thb)
    if actual_price <= 0:
        raise ValueError("Actual sale price must be greater than zero.")
    artwork = get_artwork(artwork_id)
    if not artwork:
        raise ValueError(f"Artwork not found: {artwork_id}")
    if artwork["status"] == "sold":
        raise ValueError(f"Artwork #{artwork_id} is already marked as sold.")
    exhibition_code = artwork["exhibition_code"]
    splits = get_split_rules(exhibition_code)
    if not splits:
        raise ValueError(f"Set a commission split first with /set_split {exhibition_code} ...")

    sale_date = sale_date or datetime.now().strftime("%Y-%m-%d")
    collected = actual_price if amount_collected_thb is None else float(amount_collected_thb)
    if collected < 0:
        raise ValueError("Collected amount cannot be negative.")
    if collected > actual_price:
        raise ValueError("Collected amount cannot exceed actual sale price.")
    balance_due = round(actual_price - collected, 2)
    if balance_due <= 0:
        payment_status = "collected"
    elif collected > 0:
        payment_status = "partial"
    else:
        payment_status = "uncollected"

    allocations = _allocation_amounts(actual_price, splits)
    gallery_share = sum(row["amount_thb"] for row in allocations if row["party_type"] == "gallery")
    collaborator_share = sum(row["amount_thb"] for row in allocations if row["party_type"] in {"collaborator", "collector"})
    artist_payable = sum(row["amount_thb"] for row in allocations if row["party_type"] == "artist")
    split_summary = format_split_summary(allocations)

    with connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO artwork_sales
            (artwork_id, exhibition_code, actual_price_thb, sale_date, gallery_share_thb, collaborator_share_thb,
             artist_payable_thb, split_summary, created_at, buyer_name, payment_status, amount_collected_thb,
             balance_due_thb, payment_method, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                artwork_id,
                exhibition_code,
                actual_price,
                sale_date,
                gallery_share,
                collaborator_share,
                artist_payable,
                split_summary,
                _utc_now(),
                (buyer_name or "").strip() or None,
                payment_status,
                collected,
                balance_due,
                (payment_method or "").strip() or None,
                (notes or "").strip() or None,
            ),
        )
        sale_id = cur.lastrowid
        for row in allocations:
            party_name = row["party_name"]
            if row["party_type"] == "artist" and party_name.lower() in {"artist", "default artist"}:
                party_name = artwork["artist"]
            conn.execute(
                """
                INSERT INTO sale_allocations
                (sale_id, exhibition_code, artwork_id, party_type, party_name, percent, amount_thb)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (sale_id, exhibition_code, artwork_id, row["party_type"], party_name, row["percent"], row["amount_thb"]),
            )

        conn.execute("UPDATE artworks SET status = 'sold', sold_at = ? WHERE id = ?", (_utc_now(), artwork_id))
        source_ref = f"sale:{sale_id}; artwork:{artwork_id}"
        conn.execute(
            """
            INSERT INTO pnl_lines (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
            VALUES (?, 'sales_bridge', 'Gross artwork sales', ?, ?, ?, 'THB', ?, 10)
            """,
            (exhibition_code, f"Sold artwork #{artwork_id}: {artwork['title']}", actual_price, actual_price, source_ref),
        )
        if gallery_share:
            conn.execute(
                """
                INSERT INTO pnl_lines (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
                VALUES (?, 'gallery_revenue', 'Gallery commission from sold artwork', ?, ?, ?, 'THB', ?, 20)
                """,
                (exhibition_code, f"Gallery share from artwork #{artwork_id}: {artwork['title']}", gallery_share, gallery_share, source_ref),
            )
        if artist_payable:
            conn.execute(
                """
                INSERT INTO pnl_lines (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
                VALUES (?, 'direct_cost', 'Artist payable from sold artwork', ?, ?, ?, 'THB', ?, 30)
                """,
                (exhibition_code, f"Artist share for {artwork['artist']} on artwork #{artwork_id}", artist_payable, artist_payable, source_ref),
            )
            conn.execute(
                """
                INSERT INTO artist_payables
                (exhibition_code, artist, invoice_ref, gross_sale_thb, gallery_commission_thb, artist_payable_thb,
                 paid_thb, outstanding_thb, status, source_amount, source_currency, notes)
                VALUES (?, ?, ?, ?, ?, ?, 0, ?, 'Pending', ?, 'THB', ?)
                """,
                (
                    exhibition_code,
                    artwork["artist"],
                    f"Sale #{sale_id} / Artwork #{artwork_id}",
                    actual_price,
                    gallery_share,
                    artist_payable,
                    artist_payable,
                    actual_price,
                    f"Auto-created from sale using split: {split_summary}; buyer={buyer_name or '-'}; payment={payment_status}",
                ),
            )
        if collaborator_share:
            conn.execute(
                """
                INSERT INTO pnl_lines (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
                VALUES (?, 'direct_cost', 'Collaborator / collector share from sold artwork', ?, ?, ?, 'THB', ?, 31)
                """,
                (exhibition_code, f"Collaborator or collector share from artwork #{artwork_id}", collaborator_share, collaborator_share, source_ref),
            )
        _insert_audit(conn, "record_sale", exhibition_code, f"Recorded sale #{sale_id} for artwork #{artwork_id} at {money(actual_price)}; collected {money(collected)}")
        sale = conn.execute("SELECT * FROM artwork_sales WHERE id = ?", (sale_id,)).fetchone()
        sale_allocations = conn.execute("SELECT * FROM sale_allocations WHERE sale_id = ? ORDER BY id", (sale_id,)).fetchall()
        return {"sale": dict(sale), "artwork": dict(artwork), "allocations": [dict(row) for row in sale_allocations]}


def format_sale_markdown(sale_result: Dict) -> str:
    sale = sale_result["sale"]
    artwork = sale_result["artwork"]
    lines = [
        f"*Sale Recorded — Artwork #{artwork['id']}*",
        f"Title: {artwork['title']}",
        f"Artist: {artwork['artist']}",
        f"Actual sale price: *{money(sale['actual_price_thb'])}*",
        f"Buyer: {sale.get('buyer_name') or '-'}",
        f"Payment status: *{sale.get('payment_status') or 'collected'}*",
        f"Collected: *{money(sale.get('amount_collected_thb') or sale['actual_price_thb'])}*",
        f"Balance due: *{money(sale.get('balance_due_thb') or 0)}*",
        "",
        "Split allocation:",
    ]
    for row in sale_result["allocations"]:
        lines.append(f"• {row['party_type'].title()} — {row['party_name']}: {float(row['percent']):g}% = {money(row['amount_thb'])}")
    lines.extend(
        [
            "",
            f"Gallery revenue posted: *{money(sale['gallery_share_thb'])}*",
            f"Artist payable posted: *{money(sale['artist_payable_thb'])}*",
            f"Collaborator / collector share posted: *{money(sale['collaborator_share_thb'])}*",
            "",
            "If payment status is partial or uncollected, the P&L still records the sale, while the dashboard shows the receivable balance.",
        ]
    )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Receipt capture# ---------------------------------------------------------------------------
# Receipt capture, approval, and expense reporting workflow
# ---------------------------------------------------------------------------


def parse_amount_thb(text: str | None) -> float:
    text = text or ""
    preferred = re.findall(r"(?:฿|THB|thb)\s*([0-9][0-9,]*(?:\.\d{1,2})?)", text)
    matches = preferred or re.findall(r"\b([0-9][0-9,]*(?:\.\d{1,2})?)\b", text)
    values = []
    for raw in matches:
        try:
            value = float(raw.replace(",", ""))
        except ValueError:
            continue
        if value > 0:
            values.append(value)
    if not values:
        return 0.0
    return round(max(values), 2)


def suggest_account_head(text: str | None) -> Dict:
    lowered = (text or "").lower()
    for row in ACCOUNT_HEADS:
        for keyword in row["keywords"]:
            if keyword in lowered:
                return row
    return ACCOUNT_HEADS[-1]


def account_head_names() -> List[str]:
    return [row["name"] for row in ACCOUNT_HEADS]


def get_account_head(name: str) -> Dict:
    row = ACCOUNT_HEAD_BY_NAME.get((name or "").strip().lower())
    if not row:
        raise ValueError(f"Unknown account head: {name}")
    return row


def clean_expense_description(raw_text: str | None) -> str:
    text = (raw_text or "Receipt photo / expense").strip()
    text = re.sub(r"(?:฿|THB|thb)?\s*[0-9][0-9,]*(?:\.\d{1,2})?", "", text).strip(" -:;,")
    return text[:180] or "Receipt photo / expense"


def create_pending_expense(exhibition_code: str, raw_text: str | None, photo_file_id: str | None = None) -> sqlite3.Row:
    exhibition_code = normalize_code(exhibition_code)
    if not get_exhibition(exhibition_code):
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    amount = parse_amount_thb(raw_text)
    account = suggest_account_head(raw_text)
    description = clean_expense_description(raw_text)
    with connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO pending_expenses
            (exhibition_code, raw_text, description, suggested_amount_thb, suggested_account_head, suggested_section,
             status, photo_file_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?, ?)
            """,
            (exhibition_code, raw_text, description, amount, account["name"], account["section"], photo_file_id, _utc_now()),
        )
        pending_id = cur.lastrowid
        _insert_audit(conn, "create_pending_expense", exhibition_code, f"Created pending expense #{pending_id}")
        return conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (pending_id,)).fetchone()


def get_pending_expense(pending_id: int) -> sqlite3.Row | None:
    with connect() as conn:
        return conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()


def list_pending_expenses(exhibition_code: str | None = None) -> List[sqlite3.Row]:
    with connect() as conn:
        if exhibition_code:
            return conn.execute(
                """
                SELECT * FROM pending_expenses
                WHERE UPPER(exhibition_code) = UPPER(?) AND status = 'pending'
                ORDER BY id
                """,
                (exhibition_code,),
            ).fetchall()
        return conn.execute("SELECT * FROM pending_expenses WHERE status = 'pending' ORDER BY exhibition_code, id").fetchall()


def count_pending_expenses(exhibition_code: str | None = None) -> int:
    return len(list_pending_expenses(exhibition_code))


def update_pending_account(pending_id: int, account_head: str) -> sqlite3.Row:
    account = get_account_head(account_head)
    with connect() as conn:
        row = conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()
        if not row:
            raise ValueError(f"Pending expense not found: {pending_id}")
        if row["status"] != "pending":
            raise ValueError(f"Pending expense #{pending_id} is already {row['status']}.")
        conn.execute(
            "UPDATE pending_expenses SET suggested_account_head = ?, suggested_section = ? WHERE id = ?",
            (account["name"], account["section"], int(pending_id)),
        )
        _insert_audit(conn, "update_pending_account", row["exhibition_code"], f"Pending expense #{pending_id} account changed to {account['name']}")
        return conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()


def update_pending_amount(pending_id: int, amount_thb: float) -> sqlite3.Row:
    amount = float(amount_thb)
    if amount <= 0:
        raise ValueError("Expense amount must be greater than zero.")
    with connect() as conn:
        row = conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()
        if not row:
            raise ValueError(f"Pending expense not found: {pending_id}")
        if row["status"] != "pending":
            raise ValueError(f"Pending expense #{pending_id} is already {row['status']}.")
        conn.execute("UPDATE pending_expenses SET suggested_amount_thb = ? WHERE id = ?", (amount, int(pending_id)))
        _insert_audit(conn, "update_pending_amount", row["exhibition_code"], f"Pending expense #{pending_id} amount changed to {money(amount)}")
        return conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()


def ignore_pending_expense(pending_id: int) -> sqlite3.Row:
    with connect() as conn:
        row = conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()
        if not row:
            raise ValueError(f"Pending expense not found: {pending_id}")
        conn.execute("UPDATE pending_expenses SET status = 'ignored' WHERE id = ?", (int(pending_id),))
        _insert_audit(conn, "ignore_pending_expense", row["exhibition_code"], f"Ignored pending expense #{pending_id}")
        return conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()


def confirm_pending_expense(pending_id: int) -> sqlite3.Row:
    with connect() as conn:
        pending = conn.execute("SELECT * FROM pending_expenses WHERE id = ?", (int(pending_id),)).fetchone()
        if not pending:
            raise ValueError(f"Pending expense not found: {pending_id}")
        if pending["status"] != "pending":
            raise ValueError(f"Pending expense #{pending_id} is already {pending['status']}.")
        amount = float(pending["suggested_amount_thb"] or 0)
        if amount <= 0:
            raise ValueError("Please change the amount before confirming. The current amount is zero.")
        source_ref = f"expense:{pending_id}"
        cur = conn.execute(
            """
            INSERT INTO pnl_lines
            (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
            VALUES (?, ?, ?, ?, ?, ?, 'THB', ?, 50)
            """,
            (
                pending["exhibition_code"],
                pending["suggested_section"],
                pending["suggested_account_head"],
                pending["description"],
                amount,
                amount,
                source_ref,
            ),
        )
        pnl_line_id = cur.lastrowid
        cur = conn.execute(
            """
            INSERT INTO confirmed_expenses
            (exhibition_code, account_head, section, description, amount_thb, receipt_ref, raw_text,
             pending_expense_id, created_at, pnl_line_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                pending["exhibition_code"],
                pending["suggested_account_head"],
                pending["suggested_section"],
                pending["description"],
                amount,
                source_ref,
                pending["raw_text"],
                int(pending_id),
                _utc_now(),
                pnl_line_id,
            ),
        )
        confirmed_id = cur.lastrowid
        conn.execute("UPDATE pending_expenses SET status = 'confirmed', confirmed_at = ? WHERE id = ?", (_utc_now(), int(pending_id)))
        _insert_audit(conn, "confirm_pending_expense", pending["exhibition_code"], f"Confirmed expense #{confirmed_id} from pending #{pending_id}")
        return conn.execute("SELECT * FROM confirmed_expenses WHERE id = ?", (confirmed_id,)).fetchone()


def list_confirmed_expenses(exhibition_code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM confirmed_expenses
            WHERE UPPER(exhibition_code) = UPPER(?)
            ORDER BY account_head, id
            """,
            (exhibition_code,),
        ).fetchall()


def format_pending_expense_card(pending: sqlite3.Row) -> str:
    return "\n".join(
        [
            f"Receipt Pending Approval #{pending['id']}",
            f"Exhibition: {pending['exhibition_code']}",
            f"Amount: {money(pending['suggested_amount_thb'])}",
            f"Suggested Account Head: {pending['suggested_account_head']}",
            f"P&L Section: {pending['suggested_section']}",
            f"Description: {pending['description'] or '-'}",
            f"Status: {pending['status']}",
            "",
            "Please confirm, change the account head, change the amount, or ignore this receipt.",
        ]
    )


def format_pending_expenses_markdown(exhibition_code: str | None = None) -> str:
    rows = list_pending_expenses(exhibition_code)
    title = f"Pending Receipts — {normalize_code(exhibition_code)}" if exhibition_code else "Pending Receipts"
    if not rows:
        return f"{title}\n\nNo pending receipts awaiting approval."
    lines = [title, ""]
    for row in rows:
        lines.append(
            f"• #{row['id']} / {row['exhibition_code']} / {money(row['suggested_amount_thb'])} / "
            f"{row['suggested_account_head']} / {row['description']}"
        )
    return "\n".join(lines)


def format_expense_report_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    rows = list_confirmed_expenses(exhibition_code)
    if not rows:
        return f"Expense Report — {exhibition['name']}\n\nNo confirmed expenses have been recorded."
    grouped: Dict[str, List[sqlite3.Row]] = {}
    for row in rows:
        grouped.setdefault(row["account_head"], []).append(row)
    grand_total = sum(float(row["amount_thb"] or 0) for row in rows)
    lines = [f"Expense Report — {exhibition['name']} ({exhibition['code']})", "Currency: THB only", ""]
    for account_head, account_rows in grouped.items():
        subtotal = sum(float(row["amount_thb"] or 0) for row in account_rows)
        lines.append(f"{account_head}: {money(subtotal)}")
        for row in account_rows:
            lines.append(f"  • #{row['id']} {row['description']} — {money(row['amount_thb'])}")
        lines.append("")
    lines.append(f"Total Confirmed Expenses: {money(grand_total)}")
    return "\n".join(lines)


def format_account_heads_markdown() -> str:
    lines = ["Expense Account Heads", "", "Use these classifications when approving receipts:", ""]
    for idx, row in enumerate(ACCOUNT_HEADS, start=1):
        lines.append(f"{idx}. {row['name']} — P&L section: {row['section']}")
    return "\n".join(lines)



# ---------------------------------------------------------------------------
# Management controls, budgets, inventory, and readiness dashboards
# ---------------------------------------------------------------------------


def list_sales(exhibition_code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT s.*, a.title, a.artist, a.asking_price_thb
            FROM artwork_sales s
            LEFT JOIN artworks a ON a.id = s.artwork_id
            WHERE UPPER(s.exhibition_code) = UPPER(?)
            ORDER BY s.sale_date, s.id
            """,
            (exhibition_code,),
        ).fetchall()


def calculate_inventory_metrics(exhibition_code: str) -> Dict:
    artworks = list_artworks(exhibition_code)
    sales = list_sales(exhibition_code)
    total_artworks = len(artworks)
    sold_artworks = len([row for row in artworks if row["status"] == "sold"])
    available_artworks = total_artworks - sold_artworks
    total_asking_value = sum(float(row["asking_price_thb"] or 0) for row in artworks)
    unsold_asking_value = sum(float(row["asking_price_thb"] or 0) for row in artworks if row["status"] != "sold")
    gross_sales = sum(float(row["actual_price_thb"] or 0) for row in sales)
    cash_collected = sum(float((row["amount_collected_thb"] if "amount_collected_thb" in row.keys() else row["actual_price_thb"]) or 0) for row in sales)
    receivables = sum(float((row["balance_due_thb"] if "balance_due_thb" in row.keys() else 0) or 0) for row in sales)
    average_sale_price = gross_sales / len(sales) if sales else 0.0
    sell_through = sold_artworks / total_artworks * 100 if total_artworks else 0.0
    return {
        "total_artworks": total_artworks,
        "sold_artworks": sold_artworks,
        "available_artworks": available_artworks,
        "total_asking_value_thb": total_asking_value,
        "unsold_asking_value_thb": unsold_asking_value,
        "gross_sales_thb": gross_sales,
        "cash_collected_thb": cash_collected,
        "receivables_thb": receivables,
        "average_sale_price_thb": average_sale_price,
        "sell_through_rate_pct": sell_through,
        "pending_receipts": count_pending_expenses(exhibition_code),
    }


def set_expense_budget(exhibition_code: str, account_head: str, budget_thb: float, notes: str | None = None) -> sqlite3.Row:
    exhibition_code = normalize_code(exhibition_code)
    if not get_exhibition(exhibition_code):
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    account = get_account_head(account_head)
    amount = float(budget_thb)
    if amount < 0:
        raise ValueError("Budget amount cannot be negative.")
    now = _utc_now()
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO expense_budgets (exhibition_code, account_head, section, budget_thb, notes, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(exhibition_code, account_head)
            DO UPDATE SET budget_thb = excluded.budget_thb, notes = excluded.notes, section = excluded.section, updated_at = excluded.updated_at
            """,
            (exhibition_code, account["name"], account["section"], amount, notes, now, now),
        )
        _insert_audit(conn, "set_expense_budget", exhibition_code, f"Set budget {account['name']} to {money(amount)}")
        return conn.execute(
            "SELECT * FROM expense_budgets WHERE UPPER(exhibition_code) = UPPER(?) AND account_head = ?",
            (exhibition_code, account["name"]),
        ).fetchone()


def list_expense_budgets(exhibition_code: str) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM expense_budgets
            WHERE UPPER(exhibition_code) = UPPER(?)
            ORDER BY account_head
            """,
            (exhibition_code,),
        ).fetchall()


def calculate_budget_report(exhibition_code: str) -> List[Dict]:
    confirmed = list_confirmed_expenses(exhibition_code) if get_exhibition(exhibition_code) else []
    budgets = {row["account_head"]: row for row in list_expense_budgets(exhibition_code)} if get_exhibition(exhibition_code) else {}
    names = sorted(set(account_head_names()) | set(row["account_head"] for row in confirmed) | set(budgets.keys()))
    rows: List[Dict] = []
    for name in names:
        account = ACCOUNT_HEAD_BY_NAME.get(name.lower()) or {"section": budgets.get(name, {}).get("section", "operating_expense") if budgets.get(name) else "operating_expense"}
        actual = sum(float(row["amount_thb"] or 0) for row in confirmed if row["account_head"] == name)
        budget = float(budgets[name]["budget_thb"] or 0) if name in budgets else 0.0
        variance = budget - actual
        utilization = actual / budget * 100 if budget else 0.0
        rows.append(
            {
                "account_head": name,
                "section": account["section"],
                "budget_thb": budget,
                "actual_thb": actual,
                "variance_thb": variance,
                "utilization_pct": utilization,
            }
        )
    return rows


def format_budget_report_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    rows = calculate_budget_report(exhibition_code)
    budgeted = [row for row in rows if row["budget_thb"] or row["actual_thb"]]
    if not budgeted:
        return f"No budgets or confirmed expenses found for `{normalize_code(exhibition_code)}`. Set one with /budget {normalize_code(exhibition_code)} <account> <amount>."
    lines = [f"*Budget vs Actual — {exhibition['name']}*", ""]
    for row in budgeted:
        marker = "OVER" if row["budget_thb"] > 0 and row["variance_thb"] < 0 else "OK"
        lines.append(
            f"• {row['account_head']}: Actual {compact_money(row['actual_thb'])} / Budget {compact_money(row['budget_thb'])} / Variance {compact_money(row['variance_thb'])} ({marker})"
        )
    return "\n".join(lines)


def format_inventory_dashboard_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    metrics = calculate_inventory_metrics(exhibition_code)
    lines = [
        f"*Inventory & Cash Dashboard — {exhibition['name']}*",
        "",
        f"Registered artworks: *{metrics['total_artworks']}*",
        f"Sold artworks: *{metrics['sold_artworks']}*",
        f"Available artworks: *{metrics['available_artworks']}*",
        f"Sell-through rate: *{metrics['sell_through_rate_pct']:.1f}%*",
        f"Total asking value: *{money(metrics['total_asking_value_thb'])}*",
        f"Unsold asking value: *{money(metrics['unsold_asking_value_thb'])}*",
        f"Gross sale value: *{money(metrics['gross_sales_thb'])}*",
        f"Average sale price: *{money(metrics['average_sale_price_thb'])}*",
        f"Cash collected: *{money(metrics['cash_collected_thb'])}*",
        f"Receivables outstanding: *{money(metrics['receivables_thb'])}*",
        f"Pending receipt approvals: *{metrics['pending_receipts']}*",
    ]
    return "\n".join(lines)


def format_readiness_markdown(exhibition_code: str) -> str:
    exhibition = get_exhibition(exhibition_code)
    if not exhibition:
        raise ValueError(f"Exhibition not found: {exhibition_code}")
    checks = data_quality_checks(exhibition_code)
    blocking = [c for c in checks if "No blocking issues" not in c]
    status = "READY FOR REVIEW" if not blocking else "NEEDS ATTENTION"
    lines = [
        f"*Final Review Readiness — {exhibition['name']}*",
        f"Status: *{status}*",
        "",
    ]
    lines.extend(f"• {check}" for check in checks)
    lines.append("")
    lines.append("Use the guided menu to clear pending receipts, set missing splits, correct budgets, or export the final workbook.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------


def export_report_xlsx(code: str, output_dir: str = "./exports") -> str:
    report = calculate_report(code)
    ex = report["exhibition"]
    totals = report["totals"]
    metrics = calculate_inventory_metrics(code)
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    file_path = Path(output_dir) / f"{ex['code']}_pnl_report.xlsx"

    wb = Workbook()
    summary = wb.active
    summary.title = "Executive Summary"
    summary.append(["THE SEA ART GALLERY"])
    summary.append(["Exhibition Finance Dashboard"])
    summary.append(["Exhibition", ex["name"], "Code", ex["code"]])
    summary.append(["Location", ex["location"], "Period", f"{ex['start_date']} to {ex['end_date']}"])
    summary.append(["Currency", "THB only"])
    summary.append([])
    summary.append(["Metric", "Value"])
    summary_rows = [
        ("Gross Artwork Sales / Activity", totals["gross_sales"]),
        ("Gallery Revenue", totals["gallery_revenue"]),
        ("Direct Costs", totals["direct_costs"]),
        ("Operating Expenses", totals["operating_expenses"]),
        ("Allocated Overhead", totals["allocated_overhead"]),
        ("Net Profit / (Loss)", totals["net_profit"]),
        ("Net Margin %", totals["net_margin_pct"]),
        ("Registered Artworks", metrics["total_artworks"]),
        ("Sold Artworks", metrics["sold_artworks"]),
        ("Sell-through %", metrics["sell_through_rate_pct"]),
        ("Unsold Asking Value", metrics["unsold_asking_value_thb"]),
        ("Cash Collected", metrics["cash_collected_thb"]),
        ("Sale Receivables", metrics["receivables_thb"]),
        ("Pending Receipts", metrics["pending_receipts"]),
    ]
    for label, value in summary_rows:
        summary.append([label, value])

    ws = wb.create_sheet("P&L")
    ws.append(["THE SEA ART GALLERY"])
    ws.append(["Profit & Loss Statement"])
    ws.append(["Exhibition", ex["name"], "Code", ex["code"]])
    ws.append(["Location", ex["location"], "Period", f"{ex['start_date']} to {ex['end_date']}"])
    ws.append(["Currency", "THB only"])
    ws.append([])

    def add_section(title: str, section: str, total_label: str, total_value: float):
        ws.append([title])
        ws.append(["Category", "Description", "Amount (THB)", "Source Ref"])
        for row in grouped_lines(report, section):
            ws.append([row["category"], row.get("description") or "", float(row["amount_thb"]), row.get("source_ref") or ""])
        ws.append([total_label, "", float(total_value), ""])
        ws.append([])

    add_section("A. Sales Bridge", "sales_bridge", "Gross Sales / Activity", totals["gross_sales"])
    add_section("B. Gallery Revenue", "gallery_revenue", "Total Gallery Revenue", totals["gallery_revenue"])
    add_section("C. Direct Costs", "direct_cost", "Total Direct Costs", totals["direct_costs"])
    ws.append(["Gross Profit", "", float(totals["gross_profit"]), f"{totals['gross_margin_pct']:.1f}%"])
    ws.append([])
    add_section("D. Operating Expenses", "operating_expense", "Total Operating Expenses", totals["operating_expenses"])
    ws.append(["Contribution Profit", "", float(totals["contribution_profit"]), f"{totals['contribution_margin_pct']:.1f}%"])
    ws.append([])
    add_section("E. Allocated Overhead", "allocated_overhead", "Total Allocated Overhead", totals["allocated_overhead"])
    ws.append(["NET PROFIT / (LOSS)", "", float(totals["net_profit"]), f"{totals['net_margin_pct']:.1f}%"])

    inventory = wb.create_sheet("Inventory")
    inventory.append(["Artwork ID", "Inventory Code", "Title", "Artist", "Medium", "Dimensions", "Year", "Asking Price THB", "Status", "Sold At"])
    for row in list_artworks(code):
        inventory.append([
            row["id"],
            row["inventory_code"] if "inventory_code" in row.keys() else "",
            row["title"],
            row["artist"],
            row["medium"] if "medium" in row.keys() else "",
            row["dimensions"] if "dimensions" in row.keys() else "",
            row["year_created"] if "year_created" in row.keys() else "",
            float(row["asking_price_thb"] or 0),
            row["status"],
            row["sold_at"],
        ])

    ap = wb.create_sheet("Artist Payables")
    ap.append(["Artist", "Invoice Ref", "Gross Sale THB", "Gallery Commission THB", "Artist Payable THB", "Paid THB", "Outstanding THB", "Status"])
    for row in report["payables"]:
        ap.append([
            row["artist"], row.get("invoice_ref") or "", float(row["gross_sale_thb"]), float(row["gallery_commission_thb"]),
            float(row["artist_payable_thb"]), float(row["paid_thb"]), float(row["outstanding_thb"]), row["status"],
        ])

    expenses = wb.create_sheet("Confirmed Expenses")
    expenses.append(["ID", "Account Head", "P&L Section", "Description", "Amount THB", "Receipt Ref", "Created At"])
    for row in list_confirmed_expenses(code):
        expenses.append([row["id"], row["account_head"], row["section"], row["description"], float(row["amount_thb"]), row["receipt_ref"], row["created_at"]])

    budget = wb.create_sheet("Budget vs Actual")
    budget.append(["Account Head", "P&L Section", "Budget THB", "Actual THB", "Variance THB", "Utilization %"])
    for row in calculate_budget_report(code):
        if row["budget_thb"] or row["actual_thb"]:
            budget.append([row["account_head"], row["section"], float(row["budget_thb"]), float(row["actual_thb"]), float(row["variance_thb"]), float(row["utilization_pct"])])

    pending = wb.create_sheet("Pending Receipts")
    pending.append(["ID", "Exhibition", "Description", "Amount THB", "Suggested Account", "Suggested Section", "Status", "Created At", "Raw Text"])
    for row in list_pending_expenses(code):
        pending.append([row["id"], row["exhibition_code"], row["description"], float(row["suggested_amount_thb"]), row["suggested_account_head"], row["suggested_section"], row["status"], row["created_at"], row["raw_text"]])

    sales = wb.create_sheet("Sales Allocations")
    sales.append(["Sale ID", "Artwork ID", "Title", "Artist", "Buyer", "Sale Price THB", "Collected THB", "Balance Due THB", "Payment Status", "Party Type", "Party Name", "Percent", "Allocation Amount THB"])
    with connect() as conn:
        sale_rows = conn.execute(
            """
            SELECT sa.sale_id, sa.artwork_id, a.title, a.artist, s.buyer_name, s.actual_price_thb, s.amount_collected_thb,
                   s.balance_due_thb, s.payment_status, sa.party_type, sa.party_name, sa.percent, sa.amount_thb
            FROM sale_allocations sa
            LEFT JOIN artwork_sales s ON s.id = sa.sale_id
            LEFT JOIN artworks a ON a.id = sa.artwork_id
            WHERE UPPER(sa.exhibition_code) = UPPER(?)
            ORDER BY sa.sale_id, sa.id
            """,
            (code,),
        ).fetchall()
    for row in sale_rows:
        sales.append([
            row["sale_id"], row["artwork_id"], row["title"], row["artist"], row["buyer_name"], float(row["actual_price_thb"] or 0),
            float(row["amount_collected_thb"] or 0), float(row["balance_due_thb"] or 0), row["payment_status"], row["party_type"],
            row["party_name"], float(row["percent"]), float(row["amount_thb"]),
        ])

    quality = wb.create_sheet("Data Quality")
    quality.append(["Check"])
    for warning in data_quality_checks(code):
        quality.append([warning])

    audit = wb.create_sheet("Audit Log")
    audit.append(["Timestamp", "Action", "Exhibition", "Details"])
    with connect() as conn:
        audit_rows = conn.execute(
            """
            SELECT timestamp, action, exhibition_code, details
            FROM audit_log
            WHERE exhibition_code IS NULL OR UPPER(exhibition_code) = UPPER(?)
            ORDER BY id DESC LIMIT 250
            """,
            (code,),
        ).fetchall()
    for row in audit_rows:
        audit.append([row["timestamp"], row["action"], row["exhibition_code"], row["details"]])

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col_idx in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col_idx)].width = 24
        for cell in sheet[1]:
            cell.font = Font(bold=True, size=14)
        thin = Side(style="thin", color="DDDDDD")
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    wb.save(file_path)
    log_action("export_xlsx", code, str(file_path))
    return str(file_path)


# ===========================================================================
# User State Management (DB Persisted)
# ===========================================================================

import json

def get_user_state(chat_id: int) -> dict:
    with connect() as conn:
        row = conn.execute("SELECT * FROM user_states WHERE chat_id = ?", (chat_id,)).fetchone()
        if row:
            return dict(row)
        # Default state
        default_exh = resolve_default_exhibition()
        return {
            "chat_id": chat_id,
            "current_exhibition": default_exh,
            "active_flow": None,
            "flow_step": 0,
            "flow_data": "{}",
        }

def set_user_exhibition(chat_id: int, exhibition_code: str) -> None:
    exhibition_code = normalize_code(exhibition_code)
    now = _utc_now()
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO user_states (chat_id, current_exhibition, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(chat_id) DO UPDATE SET current_exhibition = excluded.current_exhibition, updated_at = excluded.updated_at
            """,
            (chat_id, exhibition_code, now),
        )

def set_user_flow(chat_id: int, active_flow: str | None, flow_step: int = 0, flow_data: dict | None = None) -> None:
    now = _utc_now()
    data_str = json.dumps(flow_data or {})
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO user_states (chat_id, active_flow, flow_step, flow_data, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(chat_id) DO UPDATE SET active_flow = excluded.active_flow, flow_step = excluded.flow_step,
                                               flow_data = excluded.flow_data, updated_at = excluded.updated_at
            """,
            (chat_id, active_flow, flow_step, data_str, now),
        )

def clear_user_flow(chat_id: int) -> None:
    now = _utc_now()
    with connect() as conn:
        conn.execute(
            """
            UPDATE user_states 
            SET active_flow = NULL, flow_step = 0, flow_data = '{}', updated_at = ?
            WHERE chat_id = ?
            """,
            (now, chat_id),
        )

def resolve_default_exhibition() -> str:
    try:
        with connect() as conn:
            row = conn.execute(
                "SELECT code FROM exhibitions WHERE status NOT IN ('prototype', 'completed') "
                "ORDER BY COALESCE(end_date, start_date, '9999') DESC, code LIMIT 1"
            ).fetchone()
            if row:
                return row[0]
            row = conn.execute("SELECT code FROM exhibitions ORDER BY rowid DESC LIMIT 1").fetchone()
            if row:
                return row[0]
    except Exception:
        pass
    raw_default = os.environ.get("DEFAULT_EXHIBITION", "SHWEDAGON2024") or "SHWEDAGON2024"
    return re.sub(r"[^A-Z0-9_]", "", raw_default.split()[0].upper()) or "SHWEDAGON2024"


# ===========================================================================
# Advanced P&L, Cash Flow, Portfolio, and closeout features
# ===========================================================================

def get_cash_flow_timeline(code: str) -> List[Dict]:
    timeline = []
    with connect() as conn:
        # Sales cash inflow
        sales = conn.execute(
            """
            SELECT s.sale_date as date, a.title, s.actual_price_thb, s.amount_collected_thb
            FROM artwork_sales s
            LEFT JOIN artworks a ON a.id = s.artwork_id
            WHERE UPPER(s.exhibition_code) = UPPER(?)
            """, (code,)
        ).fetchall()
        for r in sales:
            timeline.append({
                "type": "sale",
                "date": r["date"],
                "description": f"Sale: {r['title']}",
                "amount": float(r["amount_collected_thb"] or 0),
                "total_value": float(r["actual_price_thb"] or 0),
            })
        
        # Expenses cash outflow
        expenses = conn.execute(
            """
            SELECT SUBSTR(created_at, 1, 10) as date, description, amount_thb, account_head
            FROM confirmed_expenses
            WHERE UPPER(exhibition_code) = UPPER(?)
            """, (code,)
        ).fetchall()
        for r in expenses:
            timeline.append({
                "type": "expense",
                "date": r["date"],
                "description": f"{r['account_head']} - {r['description']}",
                "amount": -float(r["amount_thb"] or 0),
            })
            
    # Sort by date
    timeline.sort(key=lambda x: x["date"])
    return timeline

def format_cash_flow_timeline_markdown(code: str) -> str:
    timeline = get_cash_flow_timeline(code)
    if not timeline:
        return f"No cash flow transactions recorded for `{code}`."
    
    lines = [f"*Cash Flow Timeline — {code}*", "Only actual cash collected / paid is shown here.", ""]
    running_balance = 0.0
    for item in timeline:
        amount = item["amount"]
        running_balance += amount
        sign = "⬆️ +" if amount >= 0 else "⬇️ -"
        lines.append(f"• `{item['date']}` {sign}{compact_money(abs(amount))} : {item['description']}")
    
    lines.append("")
    lines.append(f"Net Cash Position: *{money(running_balance)}*")
    return "\n".join(lines)

def format_multi_exhibition_dashboard() -> str:
    with connect() as conn:
        exhibitions = conn.execute("SELECT code, name, status FROM exhibitions ORDER BY start_date DESC").fetchall()
        if not exhibitions:
            return "No exhibitions found."
        
        lines = ["*THE SEA ART GALLERY — Portfolio Dashboard*", ""]
        portfolio_revenue = 0.0
        portfolio_expenses = 0.0
        portfolio_net = 0.0
        
        for ex in exhibitions:
            code = ex["code"]
            # Calculate P&L lines
            gallery_rev = conn.execute("SELECT SUM(amount_thb) FROM pnl_lines WHERE exhibition_code=? AND section='gallery_revenue'", (code,)).fetchone()[0] or 0.0
            direct_cost = conn.execute("SELECT SUM(amount_thb) FROM pnl_lines WHERE exhibition_code=? AND section='direct_cost'", (code,)).fetchone()[0] or 0.0
            op_exp = conn.execute("SELECT SUM(amount_thb) FROM pnl_lines WHERE exhibition_code=? AND section='operating_expense'", (code,)).fetchone()[0] or 0.0
            overhead = conn.execute("SELECT SUM(amount_thb) FROM pnl_lines WHERE exhibition_code=? AND section='allocated_overhead'", (code,)).fetchone()[0] or 0.0
            
            net_pnl = gallery_rev - direct_cost - op_exp - overhead
            
            portfolio_revenue += gallery_rev
            portfolio_expenses += (direct_cost + op_exp + overhead)
            portfolio_net += net_pnl
            
            status_emoji = "🟢" if ex["status"] == "active" else "⚪"
            lines.append(
                f"{status_emoji} *{code}* — {ex['name']} ({ex['status'].upper()})\n"
                f"  Rev: {compact_money(gallery_rev)} | Exp: {compact_money(direct_cost + op_exp + overhead)}\n"
                f"  Net P&L: *{compact_money(net_pnl)}*"
            )
            lines.append("")
            
        lines.append("─────────────────────────")
        lines.append("🏆 *PORTFOLIO TOTALS*")
        lines.append(f"• Total Gallery Revenue: *{money(portfolio_revenue)}*")
        lines.append(f"• Total Expenses: *{money(portfolio_expenses)}*")
        lines.append(f"• Net Profit/Loss: *{money(portfolio_net)}*")
        return "\n".join(lines)

def check_budget_alert(exhibition_code: str, account_head: str) -> str | None:
    with connect() as conn:
        budget_row = conn.execute(
            "SELECT budget_thb FROM expense_budgets WHERE exhibition_code=? AND account_head=?",
            (exhibition_code, account_head)
        ).fetchone()
        if not budget_row or not budget_row[0]:
            return None
        
        budget = float(budget_row[0])
        actual = conn.execute(
            "SELECT SUM(amount_thb) FROM confirmed_expenses WHERE exhibition_code=? AND account_head=?",
            (exhibition_code, account_head)
        ).fetchone()[0] or 0.0
        
        pct = (actual / budget) * 100
        if pct >= 100:
            return f"🚨 *Budget Overrun Alert!* Category *{account_head}* has spent {money(actual)} out of its {money(budget)} budget ({pct:.1f}%)."
        elif pct >= 80:
            return f"⚠️ *Budget Warning!* Category *{account_head}* has spent {money(actual)} out of its {money(budget)} budget ({pct:.1f}%)."
        return None

def get_exhibition_closeout_status(code: str) -> Dict:
    with connect() as conn:
        ex = get_exhibition(code)
        if not ex:
            raise ValueError(f"Exhibition not found: {code}")
        
        artworks_total = conn.execute("SELECT COUNT(*) FROM artworks WHERE exhibition_code=?", (code,)).fetchone()[0]
        artworks_sold = conn.execute("SELECT COUNT(*) FROM artworks WHERE exhibition_code=? AND status='sold'", (code,)).fetchone()[0]
        artworks_avail = artworks_total - artworks_sold
        
        pending_count = conn.execute("SELECT COUNT(*) FROM pending_expenses WHERE exhibition_code=? AND status='pending'", (code,)).fetchone()[0]
        
        unpaid_artists = conn.execute("SELECT COUNT(*) FROM artist_payables WHERE exhibition_code=? AND outstanding_thb > 0.01", (code,)).fetchone()[0]
        unpaid_artists_list = conn.execute("SELECT artist, outstanding_thb FROM artist_payables WHERE exhibition_code=? AND outstanding_thb > 0.01", (code,)).fetchall()
        
        outstanding_receivables = conn.execute("SELECT SUM(balance_due_thb) FROM artwork_sales WHERE exhibition_code=?", (code,)).fetchone()[0] or 0.0
        
        return {
            "exhibition": dict(ex),
            "artworks_total": artworks_total,
            "artworks_sold": artworks_sold,
            "artworks_avail": artworks_avail,
            "pending_count": pending_count,
            "unpaid_artists_count": unpaid_artists,
            "unpaid_artists": [dict(r) for r in unpaid_artists_list],
            "outstanding_receivables": outstanding_receivables,
        }

def format_closeout_status_markdown(code: str) -> str:
    status = get_exhibition_closeout_status(code)
    ex = status["exhibition"]
    lines = [
        f"🏁 *Close-Out Status — {ex['name']}*",
        f"Code: `{ex['code']}`",
        "",
        "📋 *Close-Out Checklist:*",
    ]
    
    if status["artworks_avail"] == 0:
        lines.append(f"✅ Artworks: All registered artworks are sold ({status['artworks_sold']}/{status['artworks_total']})")
    else:
        lines.append(f"⚠️ Artworks: {status['artworks_avail']} unsold artworks remain in inventory")
        
    if status["pending_count"] == 0:
        lines.append("✅ Expenses: No pending receipt approvals")
    else:
        lines.append(f"❌ Expenses: {status['pending_count']} pending receipts must be confirmed or ignored")
        
    if status["outstanding_receivables"] < 0.01:
        lines.append("✅ Sales: All sale payments fully collected")
    else:
        lines.append(f"⚠️ Sales: {money(status['outstanding_receivables'])} in receivables outstanding")
        
    if status["unpaid_artists_count"] == 0:
        lines.append("✅ Artists: All artist payables settled in full")
    else:
        lines.append(f"❌ Artists: {status['unpaid_artists_count']} artists outstanding ({money(sum(r['outstanding_thb'] for r in status['unpaid_artists']))})")
        for artist in status["unpaid_artists"]:
            lines.append(f"  • {artist['artist']}: {money(artist['outstanding_thb'])}")
            
    lines.append("")
    
    ready = (status["pending_count"] == 0) and (status["unpaid_artists_count"] == 0)
    if ready:
        lines.append("🎉 *Exhibition is ready to close!* Use `/close_exhibition` to change status to Completed.")
    else:
        lines.append("⚠️ *Exhibition is not ready to close.* Please resolve red items before closing.")
        
    return "\n".join(lines)

def close_exhibition_in_db(code: str) -> None:
    with connect() as conn:
        conn.execute("UPDATE exhibitions SET status = 'completed', end_date = ? WHERE code = ?", (datetime.now().strftime("%Y-%m-%d"), code))
        _insert_audit(conn, "close_exhibition", code, "Exhibition closed out and status set to completed.")

def seed_shwedagon_if_missing() -> None:
    """Seed the Shwe Dagon exhibition using raw sqlite3."""
    CODE = "SHWEDAGON2024"
    conversion_rate = float(os.environ.get("SEED_MMK_TO_THB_RATE", "0.006666666666666667"))
    db_path_val = db_path()

    def thb(mmk: float) -> float:
        return round(mmk * conversion_rate, 2)

    conn = sqlite3.connect(db_path_val)
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
            ("Win Myint Moe",       4,   42_000), # Wait, duplicate wins? Handover notes say 26 artists, let's verify Win Myint Moe is once
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
        # Filter duplicates (e.g. Win Myint Moe is listed twice in main.py, let's keep only unique)
        seen_artists = set()
        unique_artists = []
        for name, num, price in artists:
            if name not in seen_artists:
                seen_artists.add(name)
                unique_artists.append((name, num, price))
                
        for artist_name, num_paintings, unit_thb in unique_artists:
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
             f"{len(pnl_lines)} P&L lines + {len(unique_artists)} artists seeded via raw sqlite3."),
        )
        conn.commit()
        logger.info("SHWEDAGON2024 seeded successfully")
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

