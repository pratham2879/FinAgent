import sys
from pathlib import Path


def _check_env():
    from config import ANTHROPIC_API_KEY, EXCEL_FILE_PATH

    if not ANTHROPIC_API_KEY:
        print("Error: ANTHROPIC_API_KEY is not set.")
        print("  Add it to your .env file:  ANTHROPIC_API_KEY=sk-ant-...")
        sys.exit(1)

    if not Path(EXCEL_FILE_PATH).exists():
        print(f"Error: Excel file not found at:\n  {EXCEL_FILE_PATH}")
        print("  Override with:  EXCEL_PATH=C:\\path\\to\\file.xlsx  in .env")
        sys.exit(1)


def _print_schema_info():
    """Inspect the Excel file and print what was auto-detected."""
    from tools import get_schema
    from config import MONTHLY_BUDGET

    schema = get_schema()
    budget = schema.monthly_budget or MONTHLY_BUDGET

    print(f"  File    : {Path(schema.file_path).name}")
    print(f"  Sheet   : {schema.transaction_sheet}  (header row {schema.header_row + 1})")

    col = schema.columns
    detected = {
        k: v for k, v in {
            "date": col.date, "month": col.month, "description": col.description,
            "category": col.category, "amount": col.amount,
            "payment": col.payment, "notes": col.notes,
        }.items() if v
    }
    print(f"  Columns : {detected}")

    if schema.date_range_start:
        print(f"  Range   : {schema.date_range_start} → {schema.date_range_end}")

    print(f"  Budget  : ${budget:,.0f}/month")
    print(f"  Categories ({len(schema.categories)}): {', '.join(schema.categories)}")

    if schema.has_dashboard:
        d = schema.dashboard
        print(f"  Dashboard: '{d.sheet}' — "
              f"{len(d.category_col_map)} category columns detected")
    else:
        print("  Dashboard: not detected — aggregates computed from transactions")


def main():
    _check_env()

    from agent import run_agent

    print()
    print("=" * 65)
    print("  FinAgent — Personal Expense Analyst")
    print("=" * 65)
    print("  Inspecting Excel file...")

    try:
        _print_schema_info()
    except Exception as exc:
        print(f"  Warning: inspection failed — {exc}")

    print()
    print("  Sample queries:")
    print("    Give me a summary of March 2026")
    print("    How much did I spend on dine out in November 2025?")
    print("    Compare February and March 2026")
    print("    What were my top 5 spending spots in March 2026?")
    print("    Show me all transactions over $30 in March 2026")
    print()
    print("  Type 'exit' to quit  |  'reload' to re-inspect the file")
    print("=" * 65)
    print()

    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye!")
            break

        if not user_input:
            continue

        if user_input.lower() in ("exit", "quit", "bye", "q"):
            print("Goodbye!")
            break

        if user_input.lower() == "reload":
            from tools import clear_schema_cache
            clear_schema_cache()
            print("Re-inspecting file...")
            try:
                _print_schema_info()
            except Exception as exc:
                print(f"  Warning: {exc}")
            print()
            continue

        print()
        answer = run_agent(user_input)
        print(f"FinAgent: {answer}")
        print()


if __name__ == "__main__":
    main()
