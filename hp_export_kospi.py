#!/usr/bin/env python3
"""Calculate KOSPI valuation gap relative to daily-average exports with an HP filter."""

from __future__ import annotations

import argparse
from datetime import datetime
import json
import math
from pathlib import Path

import numpy as np
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "data" / "0_general.xlsx"
MONTHLY_HP_LAMBDA = 129_600.0


def hp_trend(values: np.ndarray, smoothing: float = MONTHLY_HP_LAMBDA) -> np.ndarray:
    """Return the two-sided Hodrick-Prescott trend using a dense linear solve."""
    if values.ndim != 1 or len(values) < 4:
        raise ValueError("HP 필터에는 최소 4개월의 연속 자료가 필요합니다.")
    if smoothing <= 0:
        raise ValueError("HP 필터 평활계수는 0보다 커야 합니다.")
    n = len(values)
    second_difference = np.zeros((n - 2, n), dtype=float)
    for index in range(n - 2):
        second_difference[index, index : index + 3] = (1.0, -2.0, 1.0)
    system = np.eye(n) + smoothing * second_difference.T @ second_difference
    return np.linalg.solve(system, values)


def calculate(path: Path = DEFAULT_WORKBOOK, smoothing: float = MONTHLY_HP_LAMBDA) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook["Sheet1"]
        rows: list[tuple[datetime, float, float]] = []
        for month, kospi, _, _, daily_export in sheet.iter_rows(
            min_row=2, min_col=1, max_col=5, values_only=True
        ):
            if (
                isinstance(month, datetime)
                and isinstance(kospi, (int, float))
                and isinstance(daily_export, (int, float))
                and kospi > 0
                and daily_export > 0
            ):
                rows.append((month, float(kospi), float(daily_export)))
    finally:
        workbook.close()
    if len(rows) < 4:
        raise ValueError("KOSPI와 일평균 수출의 공통 월별 자료가 부족합니다.")

    # The log KOSPI/export ratio removes their different units.  Its HP cycle
    # measures how far the equity market is above/below the ratio's slow trend.
    log_ratio = np.asarray([math.log(kospi / daily_export) for _, kospi, daily_export in rows])
    trend = hp_trend(log_ratio, smoothing)
    valuation_gap = 100.0 * np.expm1(log_ratio - trend)
    return {
        "workbook": str(path),
        "formula": "100 × {exp[ln(KOSPI/일평균수출) - HP추세] - 1}",
        "interpretation": "0보다 크면 수출 대비 KOSPI 고평가, 0보다 작으면 저평가",
        "hp_lambda": smoothing,
        "frequency": "월별",
        "observations": [
            {
                "date": month.strftime("%Y-%m-%d"),
                "kospi": kospi,
                "daily_export_100m_usd": daily_export,
                "valuation_gap_pct": round(float(gap), 6),
            }
            for (month, kospi, daily_export), gap in zip(rows, valuation_gap)
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--lambda", dest="smoothing", type=float, default=MONTHLY_HP_LAMBDA)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = calculate(args.workbook.resolve(), args.smoothing)
    print(json.dumps(result, ensure_ascii=False, indent=None if args.json else 2))


if __name__ == "__main__":
    main()
