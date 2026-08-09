#!/usr/bin/env python3
"""Fill every blank month through the latest fully completed calendar month."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from econo_updater import (
    GENERAL_FIELD_LABELS,
    YearMonth,
    collect_month_partial,
    previous_complete_month,
    update_workbook,
    workbook_missing_fields,
    workbook_months_needing_update,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path(__file__).resolve().parent / "data" / "0_general.xlsx",
        help="지속 갱신할 통합문서. 기본값: 프로젝트 data 폴더의 0_general.xlsx",
    )
    parser.add_argument(
        "--seed",
        type=Path,
        default=Path(__file__).resolve().parent / "temp.xlsx",
        help="0_general.xlsx가 없을 때 한 번만 복사할 초기 통합문서",
    )
    parser.add_argument(
        "--through",
        type=YearMonth.parse,
        default=previous_complete_month(),
        help="마지막 갱신 월(YYYY-MM). 기본값은 직전 완료 월",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        seed_path = args.seed.resolve()
        if not seed_path.exists():
            raise SystemExit(f"초기 통합문서를 찾을 수 없습니다: {seed_path}")
        workbook_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(seed_path, workbook_path)
        print(f"INITIALIZED {workbook_path}")

    months = workbook_months_needing_update(workbook_path, args.through)
    if not months:
        print(f"NO_UPDATE: {args.through}까지 빈 월이 없습니다.")
        return

    print(f"PLAN {len(months)}개월: {months[0]} ~ {months[-1]}", flush=True)
    updated: list[YearMonth] = []
    pending_count = 0
    for month in months:
        print(f"COLLECT {month}", flush=True)
        needed = workbook_missing_fields(workbook_path, month)
        values, errors = collect_month_partial(month, needed)
        collected = [
            field for field in needed if getattr(values, field) is not None
        ]
        if collected:
            update_workbook(workbook_path, workbook_path, [values])
            updated.append(month)
            labels = ", ".join(GENERAL_FIELD_LABELS[field] for field in collected)
            print(f"SAVED {month}: {labels}", flush=True)
        for field, message in errors.items():
            pending_count += 1
            print(
                f"PENDING {month} {GENERAL_FIELD_LABELS[field]}: {message}",
                flush=True,
            )

    if updated:
        print(f"UPDATED {len(updated)}개월: {updated[0]} ~ {updated[-1]}")
    if pending_count:
        print(f"PENDING_SUMMARY {pending_count}개 지표는 다음 실행 때 다시 확인합니다.")
    print(workbook_path)


if __name__ == "__main__":
    main()
