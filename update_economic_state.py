#!/usr/bin/env python3
"""Update economic-state data through the latest completed calendar month."""

from __future__ import annotations

import argparse
from pathlib import Path

from econo_updater import previous_complete_month
from economic_state_updater import (
    ECONOMIC_STATE_FIELD_LABELS,
    YearMonth,
    collect_economic_state_month_partial,
    missing_fields_for_month,
    months_needing_update,
    update_economic_state_workbook,
)


ROOT = Path(__file__).resolve().parent


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=ROOT / "data" / "1.economic_state.xlsx",
        help="지속 갱신할 경제상태 통합문서",
    )
    parser.add_argument(
        "--through",
        type=YearMonth.parse,
        default=previous_complete_month(),
        help="마지막 확인 월(YYYY-MM). 기본값은 실행일 기준 직전 완료 월",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workbook = args.workbook.resolve()
    if not workbook.exists():
        raise SystemExit(f"통합문서를 찾을 수 없습니다: {workbook}")

    months = months_needing_update(workbook, args.through)
    if not months:
        print(f"NO_UPDATE: {args.through}까지 빈 지표가 없습니다.")
        return

    print(f"PLAN {len(months)}개월: {months[0]} ~ {months[-1]}", flush=True)
    updated: list[YearMonth] = []
    pending_count = 0
    for month in months:
        print(f"COLLECT {month}", flush=True)
        needed = missing_fields_for_month(workbook, month)
        values, errors = collect_economic_state_month_partial(month, needed)
        collected = [field for field in needed if getattr(values, field) is not None]
        if collected:
            update_economic_state_workbook(workbook, [values])
            updated.append(month)
            labels = ", ".join(ECONOMIC_STATE_FIELD_LABELS[field] for field in collected)
            print(f"SAVED {month}: {labels}", flush=True)
        for field, message in errors.items():
            pending_count += 1
            print(f"PENDING {month} {ECONOMIC_STATE_FIELD_LABELS[field]}: {message}", flush=True)

    if updated:
        print(f"UPDATED {len(updated)}개월: {updated[0]} ~ {updated[-1]}")
    if pending_count:
        print(f"PENDING_SUMMARY {pending_count}개 지표는 다음 실행 때 다시 확인합니다.")
    print(workbook)


if __name__ == "__main__":
    main()
