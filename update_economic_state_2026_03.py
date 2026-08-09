#!/usr/bin/env python3
"""One-month validation update for March 2026 economic-state data."""

from pathlib import Path

from economic_state_updater import (
    YearMonth,
    collect_economic_state_month,
    update_economic_state_workbook,
)


def main() -> None:
    workbook = Path(__file__).resolve().parent / "data" / "1.economic_state.xlsx"
    month = YearMonth(2026, 3)
    print(f"COLLECT {month}", flush=True)
    values = collect_economic_state_month(month)
    update_economic_state_workbook(workbook, [values])
    print(f"UPDATED {month}: {values.excel_values()}")
    print(workbook)


if __name__ == "__main__":
    main()
