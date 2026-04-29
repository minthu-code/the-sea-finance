import argparse

from exhibitledger import data_quality_checks, export_report_xlsx, format_artist_payables_markdown, format_report_markdown, init_db


def main() -> None:
    parser = argparse.ArgumentParser(description="Run local ExhibitLedger THB report tests.")
    parser.add_argument("--code", default="SHWEDAGON2024", help="Exhibition code to report")
    parser.add_argument("--export", action="store_true", help="Export XLSX report")
    args = parser.parse_args()

    init_db()

    print(format_report_markdown(args.code))
    print("\n" + "=" * 80 + "\n")
    print(format_artist_payables_markdown(args.code))
    print("\n" + "=" * 80 + "\n")
    print("Data quality checks:")
    for warning in data_quality_checks(args.code):
        print(f"- {warning}")
    if args.export:
        print("\nExported:", export_report_xlsx(args.code))


if __name__ == "__main__":
    main()
