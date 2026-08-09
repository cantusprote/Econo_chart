#!/usr/bin/env python3
"""Update data/2.GDP.xlsx through the latest fully completed calendar year."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from update_gdp_2022 import (
    DEFAULT_WORKBOOK,
    GDPUpdateError,
    existing_inputs,
    fetch_kospi_year_end,
    fetch_nominal_gdp_trillion_krw,
    fetch_treasury_10y_december,
    update_workbook,
    years_needing_update,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument(
        "--through",
        type=int,
        default=date.today().year - 1,
        help="마지막 확인 연도. 기본값은 실행일 기준 직전 완료 연도",
    )
    parser.add_argument("--from-year", type=int, help="확인을 시작할 연도")
    parser.add_argument("--dry-run", action="store_true", help="수집만 하고 저장하지 않음")
    return parser.parse_args()


def collect_one(label: str, collector, year: int):
    try:
        value, source = collector(year)
        print(f"AVAILABLE {year} {label}: {value}", flush=True)
        print(f"SOURCE {year} {label}: {source}", flush=True)
        return value, None
    except Exception as exc:
        print(f"PENDING {year} {label}: {exc}", flush=True)
        return None, str(exc)


def main() -> None:
    args = parse_args()
    workbook = args.workbook.resolve()
    if not workbook.exists():
        raise SystemExit(f"통합문서를 찾을 수 없습니다: {workbook}")
    if args.through >= date.today().year:
        raise SystemExit("완료되지 않은 현재 연도 이후는 --through로 지정할 수 없습니다.")

    years = years_needing_update(workbook, args.through, args.from_year)
    if not years:
        print(f"NO_UPDATE: {args.through}년까지 빈 항목이 없습니다.")
        return
    print(f"PLAN {len(years)}개년: {years[0]} ~ {years[-1]}", flush=True)

    updated_years: list[int] = []
    pending_count = 0
    for year in years:
        print(f"COLLECT {year}", flush=True)
        existing_gdp, existing_kospi, existing_treasury = existing_inputs(workbook, year)
        gdp = kospi = treasury = None
        errors: list[str] = []
        if existing_gdp is None:
            gdp, error = collect_one("명목GDP", fetch_nominal_gdp_trillion_krw, year)
            if error:
                errors.append(error)
        if existing_kospi is None:
            kospi, error = collect_one("KOSPI연말", fetch_kospi_year_end, year)
            if error:
                errors.append(error)
        if existing_treasury is None:
            treasury, error = collect_one("국고채10년", fetch_treasury_10y_december, year)
            if error:
                errors.append(error)
        pending_count += len(errors)

        has_existing_row = any(
            value is not None for value in (existing_gdp, existing_kospi, existing_treasury)
        )
        has_new_value = any(value is not None for value in (gdp, kospi, treasury))
        if not args.dry_run and (has_existing_row or has_new_value):
            row, cells = update_workbook(workbook, year, gdp, kospi, treasury)
            if cells:
                updated_years.append(year)
                print(f"SAVED {year}: {', '.join(cells)}", flush=True)

    if args.dry_run:
        print("DRY_RUN: 파일을 변경하지 않았습니다.")
    elif updated_years:
        print(
            f"UPDATED {len(updated_years)}개년: {updated_years[0]} ~ {updated_years[-1]}",
            flush=True,
        )
    if pending_count:
        print(f"PENDING_SUMMARY {pending_count}개 항목은 다음 실행 때 다시 확인합니다.")
    print(workbook)


if __name__ == "__main__":
    main()
