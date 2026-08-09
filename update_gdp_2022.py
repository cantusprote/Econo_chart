#!/usr/bin/env python3
"""Update annual Korean nominal GDP in data/2.GDP.xlsx.

The workbook's GDP series is nominal GDP in KRW trillions. The official source
publishes KRW billions, so the downloaded value is divided by 1,000 before it
is written to column B.
"""

from __future__ import annotations

import argparse
import copy
import html
import math
import os
import re
import statistics
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook

from econo_updater import decode_html, ecos_rows, http_get


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "data" / "2.GDP.xlsx"
SOURCE_BASE = "https://www.index.go.kr/unity/potal/eNara/sub/showStblGams3.do"


class GDPUpdateError(RuntimeError):
    pass


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag == "tr":
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"td", "th"} and self._cell is not None and self._row is not None:
            self._row.append(" ".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def source_url(year: int) -> str:
    query = urllib.parse.urlencode(
        {
            "freq": "Y",
            "idx_cd": "2736",
            "period": f"{year}:{year}",
            "stts_cd": "273601",
        }
    )
    return f"{SOURCE_BASE}?{query}"


def fallback_source_url(year: int) -> str:
    query = urllib.parse.urlencode(
        {
            "freq": "Y",
            "idx_cd": "F0131",
            "period": "N",
            "stts_cd": "F013101",
        }
    )
    return f"{SOURCE_BASE}?{query}"


def parse_nominal_gdp_krw_billion(payload: str, year: int) -> float:
    parser = TableParser()
    parser.feed(payload)
    normalized_rows = [
        [re.sub(r"\s+", "", cell) for cell in row] for row in parser.rows
    ]

    def is_gdp_label(value: str) -> bool:
        return "국내총생산(명목GDP)" in value or "GDP(명목)" in value

    for header in normalized_rows:
        if str(year) not in header:
            continue
        year_index = header.index(str(year))
        for row in normalized_rows:
            if row and is_gdp_label(row[0]) and year_index < len(row):
                match = re.search(r"[0-9][0-9,]*(?:\.[0-9]+)?", row[year_index])
                if match:
                    return float(match.group(0).replace(",", ""))

    for normalized in normalized_rows:
        if any(is_gdp_label(cell) for cell in normalized):
            numbers = [
                float(token.replace(",", ""))
                for cell in normalized[1:]
                for token in re.findall(r"[0-9][0-9,]*(?:\.[0-9]+)?", cell)
            ]
            if numbers:
                return numbers[-1]

    text = re.sub(r"<[^>]+>", " ", html.unescape(payload))
    match = re.search(
        r"(?:국내총생산\s*\(명목GDP\)|GDP\s*\(명목\)).*?([0-9][0-9,]*\.[0-9]+)",
        text,
        flags=re.DOTALL,
    )
    if match:
        return float(match.group(1).replace(",", ""))
    raise GDPUpdateError(f"공식 통계표에서 {year}년 명목 GDP를 찾지 못했습니다.")


def fetch_nominal_gdp_trillion_krw(year: int, timeout: int = 40) -> tuple[float, str]:
    errors: list[str] = []
    for url in (source_url(year), fallback_source_url(year)):
        request = urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0 EconoChart/1.0"}
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = response.read().decode("utf-8", "replace")
            billion_krw = parse_nominal_gdp_krw_billion(payload, year)
            if not 100_000 <= billion_krw <= 10_000_000:
                raise GDPUpdateError(
                    f"{year}년 명목 GDP 값의 단위를 확인할 수 없습니다: {billion_krw}"
                )
            return billion_krw / 1_000.0, url
        except (urllib.error.URLError, TimeoutError, GDPUpdateError) as exc:
            errors.append(str(exc))
    raise GDPUpdateError(f"공식 통계 다운로드 실패: {' / '.join(errors)}")


def fetch_kospi_year_end(year: int) -> tuple[float, str]:
    target_start = f"{year:04d}.12.01"
    target_end = f"{year:04d}.12.31"
    pattern = re.compile(
        r'<td class="date">(\d{4}\.\d{2}\.\d{2})</td>\s*'
        r'<td class="number_1">([\d,]+\.\d+)</td>',
        re.S,
    )

    def page_rows(page: int) -> list[tuple[str, float]]:
        url = f"https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI&page={page}"
        body = decode_html(http_get(url))
        return [(day, float(value.replace(",", ""))) for day, value in pattern.findall(body)]

    low, high = 1, 2_000
    matching_page: int | None = None
    while low <= high:
        page = (low + high) // 2
        rows = page_rows(page)
        if not rows:
            high = page - 1
            continue
        newest = max(day for day, _ in rows)
        oldest = min(day for day, _ in rows)
        if any(target_start <= day <= target_end for day, _ in rows):
            matching_page = page
            break
        if oldest > target_end:
            low = page + 1
        elif newest < target_start:
            high = page - 1
        else:
            matching_page = page
            break
    if matching_page is None:
        raise GDPUpdateError(f"{year}년 12월 KOSPI 시세 페이지를 찾지 못했습니다.")

    observations: list[tuple[str, float]] = []
    # One Naver page contains only a few trading days. Search a generous band
    # around the binary-search hit so the final trading day is never omitted.
    for page in range(max(1, matching_page - 12), matching_page + 13):
        observations.extend(
            (day, value)
            for day, value in page_rows(page)
            if target_start <= day <= target_end
        )
    if not observations:
        raise GDPUpdateError(f"{year}년 12월 KOSPI 일별 시세를 찾지 못했습니다.")
    last_day, value = max(observations, key=lambda item: item[0])
    return value, f"https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI ({last_day})"


def fetch_treasury_10y_december(year: int) -> tuple[float, str]:
    rows = ecos_rows(
        "817Y002", "D", f"{year}1201", f"{year}1231", "010210000"
    )
    values = [float(row["DATA_VALUE"]) for row in rows]
    if not values:
        raise GDPUpdateError(f"{year}년 12월 국고채 10년 금리가 없습니다.")
    return round(statistics.fmean(values), 2), (
        "https://ecos.bok.or.kr (817Y002/010210000, 12월 일평균)"
    )


def copy_row_style(sheet, source_row: int, target_row: int) -> None:
    for column in range(1, sheet.max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        target._style = copy.copy(source._style)
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy.copy(source.protection)


def find_or_append_year_row(sheet, year: int) -> int:
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row, 1).value == year:
            return row
    prior_rows = [
        row
        for row in range(2, sheet.max_row + 1)
        if isinstance(sheet.cell(row, 1).value, int) and sheet.cell(row, 1).value < year
    ]
    if not prior_rows:
        raise GDPUpdateError(f"{year}년 행의 기준이 될 과거 연도를 찾지 못했습니다.")
    source_row = max(prior_rows, key=lambda row: sheet.cell(row, 1).value)
    target_row = sheet.max_row + 1
    copy_row_style(sheet, source_row, target_row)
    sheet.cell(target_row, 1).value = year
    return target_row


def fit_gdp_model(sheet, target_row: int) -> tuple[float, float]:
    samples: list[tuple[float, float]] = []
    for row in range(2, sheet.max_row + 1):
        if row == target_row:
            continue
        gdp = sheet.cell(row, 2).value
        fair_kospi = sheet.cell(row, 4).value
        if isinstance(gdp, (int, float)) and isinstance(fair_kospi, (int, float)):
            samples.append((math.log(float(gdp)), math.log(float(fair_kospi))))
    if len(samples) < 3:
        raise GDPUpdateError("적정 KOSPI 모델을 추정할 과거 자료가 부족합니다.")
    x_mean = statistics.fmean(x for x, _ in samples)
    y_mean = statistics.fmean(y for _, y in samples)
    denominator = sum((x - x_mean) ** 2 for x, _ in samples)
    slope = sum((x - x_mean) * (y - y_mean) for x, y in samples) / denominator
    intercept = y_mean - slope * x_mean
    residual = max(
        abs(math.exp(intercept + slope * x) - math.exp(y)) for x, y in samples
    )
    if residual > 1e-6:
        raise GDPUpdateError(f"기존 GDP 모델과 회귀식이 일치하지 않습니다: 최대 오차 {residual}")
    return intercept, slope


def existing_inputs(path: Path, year: int) -> tuple[float | None, float | None, float | None]:
    workbook = load_workbook(path, data_only=False)
    sheet = workbook["Sheet1"]
    row = next(
        (row for row in range(2, sheet.max_row + 1) if sheet.cell(row, 1).value == year),
        None,
    )
    values = (None, None, None) if row is None else tuple(
        float(sheet.cell(row, col).value)
        if isinstance(sheet.cell(row, col).value, (int, float))
        else None
        for col in (2, 3, 6)
    )
    workbook.close()
    return values


def update_workbook(
    path: Path,
    year: int,
    gdp: float | None,
    kospi: float | None,
    treasury_10y: float | None,
    *,
    force: bool = False,
) -> tuple[int, list[str]]:
    workbook = load_workbook(path)
    sheet = workbook["Sheet1"]
    row = find_or_append_year_row(sheet, year)
    intercept, slope = fit_gdp_model(sheet, row)
    values = {2: gdp, 3: kospi, 6: treasury_10y}
    updated: list[str] = []
    for column, value in values.items():
        target = sheet.cell(row, column)
        if value is not None and (force or target.value in (None, "")):
            target.value = value
            updated.append(target.coordinate)

    effective_gdp = sheet.cell(row, 2).value
    effective_kospi = sheet.cell(row, 3).value
    effective_treasury = sheet.cell(row, 6).value
    if isinstance(effective_gdp, (int, float)):
        fair_kospi = math.exp(intercept + slope * math.log(float(effective_gdp)))
        target = sheet.cell(row, 4)
        if force or target.value in (None, ""):
            target.value = fair_kospi
            updated.append(target.coordinate)

    formulas: dict[int, str] = {}
    if isinstance(effective_gdp, (int, float)):
        formulas.update(
            {8: f"=LN(B{row})", 12: f"=LN(B{row})", 13: f"=L{row}^2"}
        )
    if isinstance(effective_kospi, (int, float)):
        formulas.update({10: f"=LN(C{row})", 15: f"=LN(C{row})"})
    if isinstance(effective_treasury, (int, float)):
        formulas.update(
            {9: f"=LN(1+F{row}/100)", 14: f"=LN(1+F{row}/100)"}
        )
    if isinstance(sheet.cell(row, 4).value, (int, float)) and isinstance(
        effective_kospi, (int, float)
    ):
        formulas[5] = f"=(C{row}/D{row}-1)*100"
    for column, formula in formulas.items():
        target = sheet.cell(row, column)
        if force or target.value in (None, ""):
            target.value = formula
            updated.append(target.coordinate)

    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix=f".{path.stem}.", suffix=".xlsx", dir=path.parent, delete=False
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        workbook.save(temporary_path)
        workbook.close()
        os.replace(temporary_path, path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return row, updated


def years_needing_update(path: Path, through: int, from_year: int | None = None) -> list[int]:
    workbook = load_workbook(path, data_only=False)
    sheet = workbook["Sheet1"]
    rows_by_year = {
        int(sheet.cell(row, 1).value): row
        for row in range(2, sheet.max_row + 1)
        if isinstance(sheet.cell(row, 1).value, int)
    }
    if not rows_by_year:
        workbook.close()
        raise GDPUpdateError("통합문서에서 기준 연도를 찾지 못했습니다.")
    start = from_year if from_year is not None else min(rows_by_year)
    required_columns = (2, 3, 4, 5, 6, 8, 9, 10, 12, 13, 14, 15)
    years = []
    for year in range(start, through + 1):
        row = rows_by_year.get(year)
        if row is None or any(
            sheet.cell(row, column).value in (None, "")
            for column in required_columns
        ):
            years.append(year)
    workbook.close()
    return years


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--year", type=int, default=2022)
    parser.add_argument("--value", type=float, help="검증/오프라인용 조 원 단위 값")
    parser.add_argument("--kospi", type=float, help="검증/오프라인용 KOSPI 연말 종가")
    parser.add_argument("--treasury", type=float, help="검증/오프라인용 국고채 10년 12월 평균(%)")
    parser.add_argument("--force", action="store_true", help="기존 GDP 값을 덮어씀")
    parser.add_argument("--dry-run", action="store_true", help="조회만 하고 저장하지 않음")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workbook = args.workbook.resolve()
    if not workbook.exists():
        raise SystemExit(f"통합문서를 찾을 수 없습니다: {workbook}")
    existing_gdp, existing_kospi, existing_treasury = existing_inputs(workbook, args.year)
    if args.value is not None:
        gdp, gdp_source = args.value, "명령행 --value"
    elif existing_gdp is not None and not args.force:
        gdp, gdp_source = existing_gdp, "통합문서 기존값"
    else:
        gdp, gdp_source = fetch_nominal_gdp_trillion_krw(args.year)
    if args.kospi is not None:
        kospi, kospi_source = args.kospi, "명령행 --kospi"
    elif existing_kospi is not None and not args.force:
        kospi, kospi_source = existing_kospi, "통합문서 기존값"
    else:
        kospi, kospi_source = fetch_kospi_year_end(args.year)
    if args.treasury is not None:
        treasury, treasury_source = args.treasury, "명령행 --treasury"
    elif existing_treasury is not None and not args.force:
        treasury, treasury_source = existing_treasury, "통합문서 기존값"
    else:
        treasury, treasury_source = fetch_treasury_10y_december(args.year)
    print(f"COLLECT {args.year}: GDP={gdp:.4f}조 원", flush=True)
    print(f"COLLECT {args.year}: KOSPI={kospi:.2f}", flush=True)
    print(f"COLLECT {args.year}: 국고채10년={treasury:.2f}%", flush=True)
    print(f"SOURCE GDP {gdp_source}", flush=True)
    print(f"SOURCE KOSPI {kospi_source}", flush=True)
    print(f"SOURCE 국고채10년 {treasury_source}", flush=True)
    if args.dry_run:
        print("DRY_RUN: 파일을 변경하지 않았습니다.")
        return
    row, updated = update_workbook(
        workbook, args.year, gdp, kospi, treasury, force=args.force
    )
    print(f"UPDATED {args.year}: {', '.join(updated) if updated else '변경 없음'}", flush=True)
    print(workbook)


if __name__ == "__main__":
    main()
