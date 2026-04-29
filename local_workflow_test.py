import os
from pathlib import Path

TEST_DB = "/tmp/exhibitledger_workflow_test.db"
os.environ["DB_PATH"] = TEST_DB

from exhibitledger import (  # noqa: E402
    add_artwork,
    calculate_budget_report,
    calculate_inventory_metrics,
    calculate_report,
    confirm_pending_expense,
    create_exhibition,
    create_pending_expense,
    export_report_xlsx,
    format_budget_report_markdown,
    format_expense_report_markdown,
    format_inventory_dashboard_markdown,
    format_readiness_markdown,
    format_report_markdown,
    format_sale_markdown,
    format_split_rules_markdown,
    init_db,
    record_sale,
    set_commission_splits,
    set_expense_budget,
    update_pending_account,
    update_pending_amount,
)


def assert_close(actual: float, expected: float, label: str) -> None:
    if abs(actual - expected) > 0.01:
        raise AssertionError(f"{label}: expected {expected}, got {actual}")


def main() -> None:
    Path(TEST_DB).unlink(missing_ok=True)
    init_db()

    create_exhibition("TEST2026", "Workflow Test Exhibition", "Bangkok", "2026-01-01", "2026-01-31")
    set_commission_splits(
        "TEST2026",
        [
            {"party_type": "gallery", "party_name": "Gallery", "percent": 45},
            {"party_type": "collaborator", "party_name": "Curator", "percent": 10},
            {"party_type": "artist", "party_name": "Artist", "percent": 45},
        ],
    )
    set_expense_budget("TEST2026", "Food & Beverage / Hospitality", 5000)
    set_expense_budget("TEST2026", "Installation & Production", 2000)

    artwork = add_artwork("TEST2026", "Quiet River", "Artist A", 120000)
    add_artwork("TEST2026", "Blue Window", "Artist B", 80000)
    sale_result = record_sale(
        artwork["id"],
        100000,
        buyer_name="Collector One",
        amount_collected_thb=60000,
        payment_method="Bank transfer",
        notes="Balance due after delivery",
    )

    coffee = create_pending_expense("TEST2026", "THB 3500 coffee and snacks opening night")
    confirm_pending_expense(coffee["id"])

    installation = create_pending_expense("TEST2026", "Receipt photo for wall installation")
    update_pending_amount(installation["id"], 2500)
    update_pending_account(installation["id"], "Installation & Production")
    confirm_pending_expense(installation["id"])

    report = calculate_report("TEST2026")
    totals = report["totals"]

    assert_close(totals["gross_sales"], 100000, "gross sales")
    assert_close(totals["gallery_revenue"], 45000, "gallery revenue")
    assert_close(totals["direct_costs"], 57500, "direct costs")
    assert_close(totals["operating_expenses"], 3500, "operating expenses")
    assert_close(totals["net_profit"], -16000, "net profit")
    assert_close(totals["artist_payable_total"], 45000, "artist payable total")

    metrics = calculate_inventory_metrics("TEST2026")
    assert metrics["total_artworks"] == 2, "inventory total artwork count"
    assert metrics["sold_artworks"] == 1, "inventory sold artwork count"
    assert metrics["available_artworks"] == 1, "inventory available artwork count"
    assert_close(metrics["cash_collected_thb"], 60000, "cash collected")
    assert_close(metrics["receivables_thb"], 40000, "sale receivables")
    assert_close(metrics["unsold_asking_value_thb"], 80000, "unsold asking value")

    budget_rows = {row["account_head"]: row for row in calculate_budget_report("TEST2026")}
    assert_close(budget_rows["Food & Beverage / Hospitality"]["variance_thb"], 1500, "hospitality budget variance")
    assert_close(budget_rows["Installation & Production"]["variance_thb"], -500, "installation budget variance")

    print("=== SPLIT RULES ===")
    print(format_split_rules_markdown("TEST2026"))
    print("\n=== SALE RESULT ===")
    print(format_sale_markdown(sale_result))
    print("\n=== INVENTORY DASHBOARD ===")
    print(format_inventory_dashboard_markdown("TEST2026"))
    print("\n=== BUDGET VS ACTUAL ===")
    print(format_budget_report_markdown("TEST2026"))
    print("\n=== EXPENSE REPORT ===")
    print(format_expense_report_markdown("TEST2026"))
    print("\n=== READINESS CHECK ===")
    print(format_readiness_markdown("TEST2026"))
    print("\n=== P&L ===")
    print(format_report_markdown("TEST2026"))

    export_path = export_report_xlsx("TEST2026", "/tmp/exhibitledger_workflow_exports")
    print(f"\nExported: {export_path}")
    print("\nWorkflow test passed.")


if __name__ == "__main__":
    main()
