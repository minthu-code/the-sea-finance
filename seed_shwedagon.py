import os
from datetime import datetime

from exhibitledger import connect, init_db


EXHIBITION_CODE = "SHWEDAGON2024"


def rate() -> float:
    raw = os.environ.get("SEED_MMK_TO_THB_RATE", "1.0")
    try:
        value = float(raw)
    except ValueError:
        raise RuntimeError("SEED_MMK_TO_THB_RATE must be numeric, for example 0.016")
    if value <= 0:
        raise RuntimeError("SEED_MMK_TO_THB_RATE must be greater than zero")
    return value


def thb(source_mmk: float, conversion_rate: float) -> float:
    return round(float(source_mmk) * conversion_rate, 2)


def reset_exhibition(conn) -> None:
    conn.execute("DELETE FROM artist_payables WHERE exhibition_code = ?", (EXHIBITION_CODE,))
    conn.execute("DELETE FROM pnl_lines WHERE exhibition_code = ?", (EXHIBITION_CODE,))
    conn.execute("DELETE FROM exhibitions WHERE code = ?", (EXHIBITION_CODE,))


def seed() -> None:
    conversion_rate = rate()
    init_db()
    with connect() as conn:
        reset_exhibition(conn)
        seed_note = (
            "Seeded from uploaded Shwe Dagon estimated P&L and artist commission PDF. "
            f"Original source currency was MMK; imported into THB using SEED_MMK_TO_THB_RATE={conversion_rate}. "
            "If this rate is not the correct historical rate, reseed before making real finance decisions."
        )
        conn.execute(
            """
            INSERT INTO exhibitions (code, name, location, start_date, end_date, status, currency, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                EXHIBITION_CODE,
                "Shwe Dagon Platform Exhibition",
                "Bangkok / Yangon logistics",
                "2024-09-01",
                "2024-09-28",
                "prototype",
                "THB",
                seed_note,
            ),
        )

        lines = [
            ("sales_bridge", "Gross artwork sales", "Sales of paintings from estimated P&L", 48_096_000, 10),
            ("gallery_revenue", "Gallery artwork revenue", "Prototype uses estimated P&L revenue line. For consignment-final logic, replace with commission-only revenue after detailed sale split is loaded.", 48_096_000, 20),
            ("direct_cost", "Artists' fees / artist share", "Artist fees from estimated P&L", 25_708_000, 30),
            ("direct_cost", "Rotary commission", "Partner commission from estimated P&L", 432_000, 31),
            ("direct_cost", "Blank canvas", "Artwork preparation from estimated P&L", 1_357_000, 32),
            ("direct_cost", "Catalog printing", "Catalog printing from estimated P&L", 1_200_000, 33),
            ("direct_cost", "Local transportation of paintings in BKK", "Local artwork transport from estimated P&L", 336_150, 34),
            ("operating_expense", "Air cargo YGN-BKK", "Inbound exhibition logistics", 2_355_000, 40),
            ("operating_expense", "Air cargo BKK-YGN", "Return exhibition logistics", 495_000, 41),
            ("operating_expense", "Air tickets for CS", "Travel cost from estimated P&L", 2_889_050, 42),
            ("operating_expense", "Rental of exhibition space", "Venue rental from estimated P&L", 3_932_250, 43),
            ("operating_expense", "Utensil renting", "Event supply rental from estimated P&L", 465_000, 44),
            ("operating_expense", "Coffee & snacks", "Opening/event hospitality from estimated P&L", 2_850_000, 45),
            ("operating_expense", "Photographer", "Photography cost from estimated P&L", 495_000, 46),
        ]
        for section, category, description, amount_mmk, sort_order in lines:
            conn.execute(
                """
                INSERT INTO pnl_lines
                (exhibition_code, section, category, description, amount_thb, source_amount, source_currency, source_ref, sort_order)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    EXHIBITION_CODE,
                    section,
                    category,
                    description,
                    thb(amount_mmk, conversion_rate),
                    amount_mmk,
                    "MMK",
                    "22ShweDagonPlatformEstimatedP&L-Sheet1.pdf",
                    sort_order,
                ),
            )

        # Artist commission example from uploaded commission statement.
        gross_sale = 7_500_000
        gallery_commission = 3_750_000
        artist_payable = 3_750_000
        paid = 0
        conn.execute(
            """
            INSERT INTO artist_payables
            (exhibition_code, artist, invoice_ref, gross_sale_thb, gallery_commission_thb, artist_payable_thb, paid_thb, outstanding_thb, status, source_amount, source_currency, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                EXHIBITION_CODE,
                "Aye Nyein Myint",
                "Commission #22E / Invoice #302E and #303E",
                thb(gross_sale, conversion_rate),
                thb(gallery_commission, conversion_rate),
                thb(artist_payable, conversion_rate),
                thb(paid, conversion_rate),
                thb(artist_payable - paid, conversion_rate),
                "Pending",
                gross_sale,
                "MMK",
                "Imported from 22PlatformofShweDagonCommissionAllArtists-E.pdf. Gallery charge was 50%; balance payable to artist was unpaid in the source file.",
            ),
        )

        conn.execute(
            "INSERT INTO audit_log (timestamp, action, exhibition_code, details) VALUES (?, ?, ?, ?)",
            (datetime.utcnow().isoformat(timespec="seconds") + "Z", "seed_shwedagon", EXHIBITION_CODE, seed_note),
        )

    print(f"Seeded {EXHIBITION_CODE} using SEED_MMK_TO_THB_RATE={conversion_rate}")


if __name__ == "__main__":
    seed()
