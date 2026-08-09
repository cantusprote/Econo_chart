#!/usr/bin/env python3
"""Validation entrypoint: collect and write only March 2026."""

from pathlib import Path

from econo_updater import YearMonth, collect_month, update_workbook


def main() -> None:
    root = Path(__file__).resolve().parent
    output = root / "data" / "0_general.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    source = output if output.exists() else root / "temp.xlsx"
    values = collect_month(YearMonth(2026, 3))
    update_workbook(source, output, [values])
    print(f"UPDATED {values.month}: {values.excel_values()}")
    print(output)


if __name__ == "__main__":
    main()
