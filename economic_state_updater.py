#!/usr/bin/env python3
"""Collectors and workbook writer for data/1.economic_state.xlsx."""

from __future__ import annotations

import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from econo_updater import (
    UpdateError,
    YearMonth,
    fetch_inventory_ratio,
    fetch_kospi_month_end,
    fetch_leading_cycle,
    fetch_oecd_cli,
    find_month_row,
    month_from_cell,
)


@dataclass(frozen=True)
class EconomicStateValues:
    month: YearMonth
    kospi: float | None
    leading_cycle: float | None
    oecd_korea_cli: float | None
    inventory_ratio_pct: float | None
    sources: dict[str, str]

    def excel_values(self) -> tuple[float | None, float | None, float | None, float | None]:
        return (
            self.kospi,
            self.leading_cycle,
            self.oecd_korea_cli,
            self.inventory_ratio_pct,
        )


ECONOMIC_STATE_FIELDS = (
    "kospi", "leading_cycle", "oecd_korea_cli", "inventory_ratio_pct"
)

ECONOMIC_STATE_FIELD_LABELS = {
    "kospi": "KOSPI 월말",
    "leading_cycle": "선행지수 순환변동치",
    "oecd_korea_cli": "OECD 한국 선행지수",
    "inventory_ratio_pct": "제조업 재고율",
}


def collect_economic_state_month_partial(
    month: YearMonth, needed_fields: set[str]
) -> tuple[EconomicStateValues, dict[str, str]]:
    collectors = {
        "kospi": fetch_kospi_month_end,
        "leading_cycle": fetch_leading_cycle,
        "oecd_korea_cli": fetch_oecd_cli,
        "inventory_ratio_pct": fetch_inventory_ratio,
    }
    values: dict[str, float | None] = {field: None for field in ECONOMIC_STATE_FIELDS}
    sources: dict[str, str] = {}
    errors: dict[str, str] = {}
    for field in ECONOMIC_STATE_FIELDS:
        if field not in needed_fields:
            continue
        try:
            value, source = collectors[field](month)
            values[field] = value
            sources[field] = source
        except Exception as exc:
            errors[field] = str(exc)
    return EconomicStateValues(month=month, sources=sources, **values), errors


def collect_economic_state_month(month: YearMonth) -> EconomicStateValues:
    values, errors = collect_economic_state_month_partial(month, set(ECONOMIC_STATE_FIELDS))
    if errors:
        field, message = next(iter(errors.items()))
        raise UpdateError(f"{ECONOMIC_STATE_FIELD_LABELS[field]}: {message}")
    return values


def months_needing_update(workbook_path: Path, through: YearMonth) -> list[YearMonth]:
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    sheet = workbook["Sheet1"]
    months = []
    for row in range(2, sheet.max_row + 1):
        month = month_from_cell(sheet.cell(row, 1).value)
        if month and month <= through and any(
            sheet.cell(row, col).value in (None, "") for col in range(2, 6)
        ):
            months.append(month)
    workbook.close()
    return sorted(months)


def update_economic_state_workbook(
    workbook_path: Path, values: Iterable[EconomicStateValues]
) -> None:
    workbook = load_workbook(workbook_path)
    sheet = workbook["Sheet1"]
    for monthly in values:
        row = find_month_row(sheet, monthly.month)
        for col, value in enumerate(monthly.excel_values(), start=2):
            source = sheet.cell(row - 1, col)
            target = sheet.cell(row, col)
            if value is None or target.value not in (None, ""):
                continue
            target._style = copy.copy(source._style)
            target.font = copy.copy(source.font)
            target.fill = copy.copy(source.fill)
            target.border = copy.copy(source.border)
            target.alignment = copy.copy(source.alignment)
            target.number_format = source.number_format
            target.protection = copy.copy(source.protection)
            target.value = value
    workbook.save(workbook_path)


def missing_fields_for_month(workbook_path: Path, month: YearMonth) -> set[str]:
    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    sheet = workbook["Sheet1"]
    row = find_month_row(sheet, month)
    missing = {
        field
        for col, field in enumerate(ECONOMIC_STATE_FIELDS, start=2)
        if sheet.cell(row, col).value in (None, "")
    }
    workbook.close()
    return missing


__all__ = [
    "EconomicStateValues",
    "ECONOMIC_STATE_FIELD_LABELS",
    "UpdateError",
    "YearMonth",
    "collect_economic_state_month",
    "collect_economic_state_month_partial",
    "missing_fields_for_month",
    "months_needing_update",
    "update_economic_state_workbook",
]
