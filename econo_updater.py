#!/usr/bin/env python3
"""Monthly updater for the economic time-series workbook.

Only Python's standard library, openpyxl, and pypdf are required.
Network sources:
  - KOSPI month-end: Naver Finance daily KOSPI page (KRX-derived quote)
  - Korean leading index: Ministry of Data and Statistics monthly PDF
  - OECD Korea CLI: OECD SDMX API
  - Daily-average exports: MOTIE monthly export/import PDF
  - Base rate, 3-year treasury yield, USD/KRW: Bank of Korea ECOS API
"""

from __future__ import annotations

import calendar
import copy
import csv
import html
import io
import json
import re
import statistics
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from pypdf import PdfReader


USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) EconoChartUpdater/1.0"
ECOS_BASE = "https://ecos.bok.or.kr/api/StatisticSearch/sample/json/kr"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True, order=True)
class YearMonth:
    year: int
    month: int

    @classmethod
    def parse(cls, value: str) -> "YearMonth":
        match = re.fullmatch(r"(\d{4})-(\d{2})", value)
        if not match:
            raise ValueError(f"월 형식은 YYYY-MM이어야 합니다: {value}")
        year, month = map(int, match.groups())
        if not 1 <= month <= 12:
            raise ValueError(f"유효하지 않은 월입니다: {value}")
        return cls(year, month)

    def __str__(self) -> str:
        return f"{self.year:04d}-{self.month:02d}"

    def next(self) -> "YearMonth":
        return YearMonth(self.year + (self.month == 12), 1 if self.month == 12 else self.month + 1)

    def previous(self) -> "YearMonth":
        return YearMonth(self.year - (self.month == 1), 12 if self.month == 1 else self.month - 1)

    @property
    def compact(self) -> str:
        return f"{self.year:04d}{self.month:02d}"


@dataclass
class MonthlyValues:
    month: str
    kospi: float | None
    leading_cycle: float | None
    oecd_korea_cli: float | None
    daily_average_exports_usd_100m: float | None
    base_rate_pct: float | None
    treasury_3y_monthly_avg_pct: float | None
    usdkrw_monthly_avg: float | None
    sources: dict[str, str]

    def excel_values(self) -> list[float | None]:
        return [
            self.kospi,
            self.leading_cycle,
            self.oecd_korea_cli,
            self.daily_average_exports_usd_100m,
            self.base_rate_pct,
            self.treasury_3y_monthly_avg_pct,
            self.usdkrw_monthly_avg,
        ]


def http_get(url: str, *, timeout: int = 40) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read()
    except (urllib.error.URLError, TimeoutError) as exc:
        raise UpdateError(f"다운로드 실패: {url}: {exc}") from exc


def http_post(url: str, fields: dict[str, str], *, timeout: int = 40) -> bytes:
    request = urllib.request.Request(
        url,
        data=urllib.parse.urlencode(fields).encode("utf-8"),
        headers={"User-Agent": USER_AGENT, "Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read()
    except (urllib.error.URLError, TimeoutError) as exc:
        raise UpdateError(f"다운로드 실패: {url}: {exc}") from exc


def decode_html(raw: bytes) -> str:
    for encoding in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


def pdf_text(raw: bytes) -> str:
    reader = PdfReader(io.BytesIO(raw))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def fetch_kospi_month_end(month: YearMonth) -> tuple[float, str]:
    target_prefix = f"{month.year:04d}.{month.month:02d}."
    pattern = re.compile(
        r'<td class="date">(\d{4}\.\d{2}\.\d{2})</td>\s*'
        r'<td class="number_1">([\d,]+\.\d+)</td>',
        re.S,
    )
    observations: list[tuple[str, float]] = []
    for page in range(1, 80):
        url = f"https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI&page={page}"
        body = decode_html(http_get(url))
        rows = [(d, float(v.replace(",", ""))) for d, v in pattern.findall(body)]
        observations.extend((d, v) for d, v in rows if d.startswith(target_prefix))
        if observations:
            oldest = min(d for d, _ in rows)
            if oldest < target_prefix + "01":
                break
        if rows and max(d for d, _ in rows) < target_prefix + "01":
            break
    if not observations:
        raise UpdateError(f"{month} KOSPI 일별 시세를 찾지 못했습니다.")
    last_date, value = max(observations, key=lambda item: item[0])
    return value, f"https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI ({last_date})"


def find_mods_release(month: YearMonth) -> tuple[str, str]:
    title = f"{month.year}년 {month.month}월 산업활동동향"
    for page in range(1, 6):
        url = (
            "https://www.mods.go.kr/board.es?act=list&bid=216"
            f"&mid=a10301050100&nPage={page}"
        )
        body = decode_html(http_get(url))
        for title_match in re.finditer(re.escape(title), body):
            nearby = body[max(0, title_match.start() - 1500) : title_match.end() + 300]
            list_numbers = re.findall(r"list_no=(\d+)", nearby)
            if list_numbers:
                list_no = list_numbers[-1]
                view_url = (
                    "https://www.mods.go.kr/board.es?act=view&bid=216"
                    f"&list_no={list_no}&mid=a10301050100"
                )
                return view_url, list_no
    raise UpdateError(f"국가데이터처에서 {title} 보도자료를 찾지 못했습니다.")


def fetch_industrial_activity_pdf(month: YearMonth) -> tuple[str, str]:
    view_url, list_no = find_mods_release(month)
    body = decode_html(http_get(view_url))
    pdf_links = re.findall(
        rf'href="([^"]*boardDownload\.es\?[^"#]*list_no={list_no}[^"#]*)"[^>]*>[^<]*.*?\.pdf',
        body,
        re.I | re.S,
    )
    if not pdf_links:
        pdf_links = re.findall(
            rf'(boardDownload\.es\?bid=216&amp;list_no={list_no}&amp;seq=\d+)', body
        )
    if not pdf_links:
        raise UpdateError(f"{month} 산업활동동향 PDF 링크를 찾지 못했습니다.")
    pdf_url = urllib.parse.urljoin("https://www.mods.go.kr", html.unescape(pdf_links[0]))
    text = re.sub(r"\s+", "", pdf_text(http_get(pdf_url)))
    return text, view_url


def fetch_leading_cycle(month: YearMonth) -> tuple[float, str]:
    text, view_url = fetch_industrial_activity_pdf(month)
    sections = list(re.finditer(r"선행지수순환변동치(?:\d+\))?", text))
    candidates: list[float] = []
    for section in sections:
        fragment = text[section.end() : section.end() + 350]
        fragment = fragment.split("전월차", 1)[0]
        numbers = [float(x) for x in re.findall(r"\d{2,3}\.\d", fragment)]
        if numbers:
            candidates.append(numbers[-1])
    if not candidates:
        raise UpdateError(f"{month} 선행지수 순환변동치 수준을 PDF에서 파싱하지 못했습니다.")
    return candidates[-1], view_url


def fetch_inventory_ratio(month: YearMonth) -> tuple[float, str]:
    """Return the manufacturing inventory/shipment ratio from the MODS PDF."""
    text, view_url = fetch_industrial_activity_pdf(month)
    patterns = [
        r"재고/출하비율\(재고율\)은([0-9]+(?:\.[0-9]+)?)%",
        r"재고/출하비율\(재고율\)(?:은)?([0-9]+(?:\.[0-9]+)?)%",
    ]
    for pattern in patterns:
        values = [float(value) for value in re.findall(pattern, text)]
        plausible = [value for value in values if 40 <= value <= 200]
        if plausible:
            return plausible[0], view_url

    # Table fallback: take the last plausible value before the explanatory
    # footnote; the columns are chronological and the requested month is last.
    for match in re.finditer(r"재고/출하비율(?:\d+\))?", text):
        fragment = text[match.end() : match.end() + 240].split("제조업의출하지수", 1)[0]
        values = [float(value) for value in re.findall(r"\d{2,3}\.\d", fragment)]
        plausible = [value for value in values if 40 <= value <= 200]
        if plausible:
            return plausible[-1], view_url
    raise UpdateError(f"{month} 제조업 재고율을 산업활동동향 PDF에서 파싱하지 못했습니다.")


def fetch_oecd_cli(month: YearMonth) -> tuple[float, str]:
    start_month = month
    for _ in range(11):
        start_month = start_month.previous()
    url = (
        "https://sdmx.oecd.org/public/rest/data/"
        "OECD.SDD.STES,DSD_STES@DF_CLI/KOR.M.LI...AA...H"
        f"?startPeriod={start_month}&endPeriod={month}"
        "&dimensionAtObservation=AllDimensions"
    )
    request = urllib.request.Request(
        url, headers={"User-Agent": USER_AGENT, "Accept": "text/csv"}
    )
    try:
        with urllib.request.urlopen(request, timeout=40) as response:
            body = response.read().decode("utf-8-sig")
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise UpdateError(f"OECD에 {month} 한국 CLI가 아직 발표되지 않았습니다.") from exc
        raise UpdateError(f"OECD API 조회 실패: HTTP {exc.code}") from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise UpdateError(f"OECD API 조회 실패: {exc}") from exc
    rows = list(csv.DictReader(io.StringIO(body)))
    target = next((row for row in rows if row.get("TIME_PERIOD") == str(month)), None)
    if target and target.get("OBS_VALUE") not in (None, ""):
        return float(target["OBS_VALUE"]), url
    available = sorted(row["TIME_PERIOD"] for row in rows if row.get("TIME_PERIOD"))
    latest = available[-1] if available else "확인 불가"
    raise UpdateError(
        f"OECD에 {month} 한국 CLI가 아직 발표되지 않았습니다. 최신 발표월: {latest}"
    )


def fetch_daily_average_exports(month: YearMonth) -> tuple[float, str]:
    title = f"{month.year}년 {month.month}월 수출입 동향"
    list_url = "https://www.nlic.go.kr/nlic/logpolLi.action"
    body = decode_html(
        http_post(
            list_url,
            {"fldSrchWord": title, "chkTitle": "T", "tr_page": "1", "command": "LIST"},
        )
    )
    match = re.search(
        rf"fn_viewDetail\('(\d+)'\);\">\s*{re.escape(title)}\s*</a>", body
    )
    if not match:
        # NLIC can lag the ministry release, and ministry titles alternate
        # between "수출입 동향" and "수출입동향".  Fall back to the current
        # ministry board search shared with the Export Trend downloader.
        from motir_export_report import (
            build_opener,
            discover_candidate_ids,
            parse_article,
            parse_report,
            request_bytes,
        )

        opener = build_opener()
        for article_id, _ in discover_candidate_ids(opener):
            article = parse_article(opener, article_id)
            if article.report_month != str(month):
                continue
            request_bytes(opener, article.post_url)
            payload = request_bytes(opener, article.pdf_url, referer=article.post_url)
            report = parse_report(article, payload)
            value = report.get("summary", {}).get("daily_average_exports_100m_usd")
            if isinstance(value, (int, float)):
                return float(value), article.post_url
        raise UpdateError(
            f"국가물류통합정보센터와 산업통상부에서 {title}을 찾지 못했습니다."
        )
    view_url = (
        "https://www.nlic.go.kr/nlic/logpolDt.action?command=VIEW"
        f"&fldLogpolRefSeq={match.group(1)}"
    )
    body = decode_html(http_get(view_url))
    texts = [re.sub(r"\s+", "", html.unescape(re.sub(r"<[^>]+>", " ", body)))]

    # Some NLIC posts contain only a short description, with all figures in the
    # attached official PDF.  Keep the HTML path for older posts and add the PDF
    # as a fallback so a change in the post format does not stop monthly updates.
    file_matches = re.findall(
        r"fileDown\(['\"]([^'\"]+)['\"],\s*['\"]([^'\"]+)['\"]\)", body
    )
    for reference_seq, file_seq in file_matches:
        download_url = (
            "https://www.nlic.go.kr/nlic/logpolDt.action?command=DWLOAD"
            f"&LOGPOL_REF_SEQ={urllib.parse.quote(reference_seq)}"
            f"&LOGPOL_REF_FILE_SEQ={urllib.parse.quote(file_seq)}"
        )
        try:
            texts.append(re.sub(r"\s+", "", pdf_text(http_get(download_url))))
        except (UpdateError, ValueError):
            continue

    patterns = [
        r"일평균수출(?:도)?.{0,100}?증가한([0-9]+(?:\.[0-9]+)?)억달러",
        r"일평균수출(?:액|도)?.{0,100}?([0-9]+(?:\.[0-9]+)?)억달러",
        r"일평균수출액[^\d]{0,80}([0-9]+(?:\.[0-9]+)?)",
    ]
    for text in texts:
        for pattern in patterns:
            matches = re.findall(pattern, text)
            plausible = [float(value) for value in matches if 5 <= float(value) <= 100]
            if plausible:
                return plausible[-1], view_url
    raise UpdateError(f"{month} 일평균 수출액을 공식 본문에서 파싱하지 못했습니다.")


def ecos_rows(
    stat_code: str,
    cycle: str,
    start: str,
    end: str,
    item1: str,
    item2: str | None = None,
) -> list[dict]:
    rows: list[dict] = []
    page_start = 1
    while True:
        page_end = page_start + 9
        parts = [ECOS_BASE, str(page_start), str(page_end), stat_code, cycle, start, end, item1]
        if item2:
            parts.append(item2)
        url = "/".join(parts)
        payload = json.loads(http_get(url).decode("utf-8"))
        block = payload.get("StatisticSearch")
        if not block:
            message = payload.get("RESULT", {}).get("MESSAGE", "자료 없음")
            raise UpdateError(f"ECOS {stat_code}/{item1} 조회 실패: {message}")
        rows.extend(block.get("row", []))
        total = int(block.get("list_total_count", len(rows)))
        if len(rows) >= total:
            return rows
        page_start += 10


def fetch_bok_values(month: YearMonth) -> tuple[float, float, float, dict[str, str]]:
    base_rate, base_source = fetch_base_rate(month)
    treasury_avg, treasury_source = fetch_treasury_3y(month)
    usdkrw_avg, usdkrw_source = fetch_usdkrw(month)
    return base_rate, treasury_avg, usdkrw_avg, {
        "base_rate": base_source,
        "treasury_3y": treasury_source,
        "usdkrw": usdkrw_source,
    }


def fetch_base_rate(month: YearMonth) -> tuple[float, str]:
    rows = ecos_rows("722Y001", "M", month.compact, month.compact, "0101000")
    return float(rows[0]["DATA_VALUE"]), "https://ecos.bok.or.kr (722Y001/0101000)"


def fetch_treasury_3y(month: YearMonth) -> tuple[float, str]:

    last_day = calendar.monthrange(month.year, month.month)[1]
    start = f"{month.compact}01"
    end = f"{month.compact}{last_day:02d}"
    rows = ecos_rows("817Y002", "D", start, end, "010200000")
    value = statistics.fmean(float(row["DATA_VALUE"]) for row in rows)
    return value, "https://ecos.bok.or.kr (817Y002/010200000)"


def fetch_usdkrw(month: YearMonth) -> tuple[float, str]:
    rows = ecos_rows("731Y004", "M", month.compact, month.compact, "0000001", "0000100")
    return float(rows[0]["DATA_VALUE"]), "https://ecos.bok.or.kr (731Y004/0000001/0000100)"


GENERAL_FIELDS = (
    "kospi",
    "leading_cycle",
    "oecd_korea_cli",
    "daily_average_exports_usd_100m",
    "base_rate_pct",
    "treasury_3y_monthly_avg_pct",
    "usdkrw_monthly_avg",
)

GENERAL_FIELD_LABELS = {
    "kospi": "KOSPI 월말",
    "leading_cycle": "선행지수 순환변동치",
    "oecd_korea_cli": "OECD 한국 선행지수",
    "daily_average_exports_usd_100m": "일평균 수출액",
    "base_rate_pct": "기준금리",
    "treasury_3y_monthly_avg_pct": "국고채 3년 월평균",
    "usdkrw_monthly_avg": "원·달러 월평균 환율",
}


def collect_month_partial(
    month: YearMonth, needed_fields: set[str]
) -> tuple[MonthlyValues, dict[str, str]]:
    collectors = {
        "kospi": fetch_kospi_month_end,
        "leading_cycle": fetch_leading_cycle,
        "oecd_korea_cli": fetch_oecd_cli,
        "daily_average_exports_usd_100m": fetch_daily_average_exports,
        "base_rate_pct": fetch_base_rate,
        "treasury_3y_monthly_avg_pct": fetch_treasury_3y,
        "usdkrw_monthly_avg": fetch_usdkrw,
    }
    values: dict[str, float | None] = {field: None for field in GENERAL_FIELDS}
    sources: dict[str, str] = {}
    errors: dict[str, str] = {}
    for field in GENERAL_FIELDS:
        if field not in needed_fields:
            continue
        try:
            value, source = collectors[field](month)
            values[field] = value
            sources[field] = source
        except Exception as exc:
            errors[field] = str(exc)
    return MonthlyValues(month=str(month), sources=sources, **values), errors


def collect_month(month: YearMonth) -> MonthlyValues:
    values, errors = collect_month_partial(month, set(GENERAL_FIELDS))
    if errors:
        field, message = next(iter(errors.items()))
        raise UpdateError(f"{GENERAL_FIELD_LABELS[field]}: {message}")
    return values


def month_from_cell(value: object) -> YearMonth | None:
    if isinstance(value, datetime):
        return YearMonth(value.year, value.month)
    if isinstance(value, date):
        return YearMonth(value.year, value.month)
    return None


def find_month_row(sheet, month: YearMonth) -> int:
    for row in range(2, sheet.max_row + 1):
        if month_from_cell(sheet.cell(row, 1).value) == month:
            return row
    raise UpdateError(f"통합문서 A열에서 {month} 행을 찾지 못했습니다.")


def copy_cell_style(source, target) -> None:
    target._style = copy.copy(source._style)
    if source.has_style:
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy.copy(source.protection)


def update_workbook(
    input_path: Path,
    output_path: Path,
    values: Iterable[MonthlyValues],
    *,
    sheet_name: str = "Sheet1",
) -> None:
    workbook = load_workbook(input_path)
    sheet = workbook[sheet_name]
    for monthly in values:
        month = YearMonth.parse(monthly.month)
        row = find_month_row(sheet, month)
        prior_row = row - 1
        for col, value in enumerate(monthly.excel_values(), start=2):
            target = sheet.cell(row, col)
            if value is not None and target.value in (None, ""):
                copy_cell_style(sheet.cell(prior_row, col), target)
                target.value = value
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def workbook_months_needing_update(input_path: Path, through: YearMonth) -> list[YearMonth]:
    # Some third-party-generated XLSX files omit worksheet dimension metadata,
    # so normal mode is used instead of read_only mode (where max_row can be None).
    workbook = load_workbook(input_path, read_only=False, data_only=False)
    sheet = workbook["Sheet1"]
    candidates: list[YearMonth] = []
    for row in range(2, sheet.max_row + 1):
        month = month_from_cell(sheet.cell(row, 1).value)
        if month and month <= through and any(sheet.cell(row, col).value in (None, "") for col in range(2, 9)):
            candidates.append(month)
    workbook.close()
    return sorted(candidates)


def ensure_workbook_months(input_path: Path, through: YearMonth) -> None:
    """Append styled blank rows when the workbook ends before ``through``."""
    workbook = load_workbook(input_path, read_only=False, data_only=False)
    sheet = workbook["Sheet1"]
    dated_rows = [
        (row, month_from_cell(sheet.cell(row, 1).value))
        for row in range(2, sheet.max_row + 1)
    ]
    dated_rows = [(row, month) for row, month in dated_rows if month is not None]
    if not dated_rows:
        workbook.close()
        raise UpdateError("통합문서 A열에서 기준 월 행을 찾지 못했습니다.")

    last_row, last_month = max(dated_rows, key=lambda item: item[1])
    changed = False
    while last_month < through:
        next_month = last_month.next()
        new_row = last_row + 1
        for col in range(1, sheet.max_column + 1):
            copy_cell_style(sheet.cell(last_row, col), sheet.cell(new_row, col))
            sheet.cell(new_row, col).value = None
        previous_date = sheet.cell(last_row, 1).value
        day = previous_date.day if isinstance(previous_date, (date, datetime)) else 1
        sheet.cell(new_row, 1).value = datetime(next_month.year, next_month.month, day)
        if sheet.row_dimensions[last_row].height is not None:
            sheet.row_dimensions[new_row].height = sheet.row_dimensions[last_row].height
        last_row, last_month = new_row, next_month
        changed = True

    if changed:
        workbook.save(input_path)
    workbook.close()


def workbook_missing_fields(input_path: Path, month: YearMonth) -> set[str]:
    workbook = load_workbook(input_path, read_only=False, data_only=False)
    sheet = workbook["Sheet1"]
    row = find_month_row(sheet, month)
    missing = {
        field
        for col, field in enumerate(GENERAL_FIELDS, start=2)
        if sheet.cell(row, col).value in (None, "")
    }
    workbook.close()
    return missing


def previous_complete_month(today: date | None = None) -> YearMonth:
    today = today or date.today()
    return YearMonth(today.year, today.month).previous()


def save_audit(path: Path, values: Iterable[MonthlyValues]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps([asdict(value) for value in values], ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
