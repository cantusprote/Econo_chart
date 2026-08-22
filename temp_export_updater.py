#!/usr/bin/env python3
"""Update preliminary 1-10/1-20 day export averages from Korea Customs PDFs."""

from __future__ import annotations

import argparse
import copy
import hashlib
import html
import io
import json
import re
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
WORKBOOK_PATH = ROOT / "data" / "3.temp_export.xlsx"
GENERAL_WORKBOOK = ROOT / "data" / "0_general.xlsx"
OUTPUT_DIR = ROOT / "Temp_Export_trend"
HISTORY_PATH = OUTPUT_DIR / "temp_export_history.json"
BASE_URL = "https://www.customs.go.kr"
LIST_PATH = "/kcs/na/ntt/selectNttList.do"
DETAIL_PATH = "/kcs/na/ntt/selectNttInfo.do"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 EconoChartUpdater/1.0"
)


class TempExportError(RuntimeError):
    pass


@dataclass(frozen=True, order=True)
class YearMonth:
    year: int
    month: int

    @classmethod
    def parse(cls, value: str) -> "YearMonth":
        match = re.fullmatch(r"(20\d{2})-(\d{2})", value)
        if not match:
            raise argparse.ArgumentTypeError("월은 YYYY-MM 형식이어야 합니다.")
        year, month = map(int, match.groups())
        if not 1 <= month <= 12:
            raise argparse.ArgumentTypeError(f"유효하지 않은 월입니다: {value}")
        return cls(year, month)

    def __str__(self) -> str:
        return f"{self.year:04d}-{self.month:02d}"

    def next(self) -> "YearMonth":
        return YearMonth(self.year + (self.month == 12), 1 if self.month == 12 else self.month + 1)


@dataclass(frozen=True)
class Article:
    month: YearMonth
    period_days: int
    title: str
    ntt_sn: str
    ntt_url_token: str
    published: str


def request_bytes(
    url: str,
    *,
    fields: dict[str, str] | None = None,
    referer: str | None = None,
    timeout: int = 60,
) -> bytes:
    data = urllib.parse.urlencode(fields).encode() if fields else None
    headers = {"User-Agent": USER_AGENT}
    if data is not None:
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    if referer:
        headers["Referer"] = referer
    request = urllib.request.Request(url, data=data, headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read()
    except Exception as exc:
        raise TempExportError(f"다운로드 실패: {url}: {exc}") from exc


def decode_html(payload: bytes) -> str:
    return payload.decode("utf-8", errors="replace")


def clean_html(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", value))).strip()


def months_between(start: YearMonth, end: YearMonth) -> list[YearMonth]:
    result: list[YearMonth] = []
    current = start
    while current <= end:
        result.append(current)
        current = current.next()
    return result


def parse_title(title: str) -> tuple[YearMonth, int] | None:
    normalized = re.sub(r"\s+", " ", title).strip()
    match = re.fullmatch(
        r"(20\d{2})년\s*(\d{1,2})월\s*1일\s*[~∼-]\s*"
        r"(?:\d{1,2}월\s*)?(10|20)일\s*수출입\s*현황\s*\[잠정치\]",
        normalized,
    )
    if not match:
        return None
    year, month, period = map(int, match.groups())
    return YearMonth(year, month), period


def discover_articles(start: YearMonth, end: YearMonth) -> dict[tuple[str, int], Article]:
    fields = {
        "mi": "2891",
        "bbsId": "1362",
        "currPage": "1",
        "listCo": "50",
        "searchType": "sj",
        "searchValue": "수출입 현황 [잠정치]",
    }
    body = decode_html(request_bytes(BASE_URL + LIST_PATH, fields=fields))
    pattern = re.compile(
        r'<a\s+(?P<attrs>[^>]*class="nttInfoBtn"[^>]*)>(?P<title>.*?)</a>'
        r".*?<td[^>]*data-table=\"date\"[^>]*>(?P<date>.*?)</td>",
        re.I | re.S,
    )
    found: dict[tuple[str, int], Article] = {}
    for match in pattern.finditer(body):
        title = clean_html(match.group("title"))
        parsed = parse_title(title)
        if not parsed:
            continue
        month, period_days = parsed
        if not start <= month <= end:
            continue
        attrs = match.group("attrs")
        id_match = re.search(r'data-id="(\d+)"', attrs)
        token_match = re.search(r'data-url="([^"]+)"', attrs)
        if not id_match or not token_match:
            continue
        found[(str(month), period_days)] = Article(
            month=month,
            period_days=period_days,
            title=title,
            ntt_sn=id_match.group(1),
            ntt_url_token=token_match.group(1),
            published=clean_html(match.group("date")),
        )
    return found


def article_pdf(article: Article) -> tuple[str, str, bytes]:
    fields = {
        "mi": "2891",
        "bbsId": "1362",
        "currPage": "1",
        "nttSn": article.ntt_sn,
        "nttSnUrl": article.ntt_url_token,
    }
    detail_url = BASE_URL + DETAIL_PATH
    body = decode_html(request_bytes(detail_url, fields=fields))
    attachments = re.findall(
        r'<a\s+href="(?P<url>[^"]*nttFileDownload\.do\?fileKey=[^"]+)"'
        r'\s+title="(?P<name>[^"]+\.pdf)\s+다운로드"',
        body,
        re.I,
    )
    if not attachments:
        raise TempExportError(f"PDF 첨부파일을 찾지 못했습니다: {article.title}")
    relative_url, filename = attachments[-1]
    pdf_url = urllib.parse.urljoin(BASE_URL, html.unescape(relative_url))
    payload = request_bytes(pdf_url, referer=detail_url)
    if not payload.startswith(b"%PDF-"):
        raise TempExportError(f"PDF가 아닌 응답입니다: {article.title}")
    return clean_html(filename), pdf_url, payload


def parse_daily_average(article: Article, payload: bytes) -> tuple[float, float | None, str]:
    reader = PdfReader(io.BytesIO(payload))
    if not reader.pages:
        raise TempExportError(f"빈 PDF입니다: {article.title}")
    text = "".join(page.extract_text() or "" for page in reader.pages[:2])
    compact = re.sub(r"\s+", "", text).replace("‘", "’").replace("'", "’")
    section_match = re.search(r"일평균수출액\[(?P<section>[^\]]+)\]", compact)
    if not section_match:
        raise TempExportError(f"일평균 수출액 표기를 찾지 못했습니다: {article.title}")
    section = section_match.group("section")
    observations = re.findall(
        r"\(’?(\d{2,4})\.(\d{1,2})\.\)([\d.]+)(?:억달러)?",
        section,
    )
    target_year = article.month.year % 100
    values = [
        float(value)
        for year, month, value in observations
        if int(year) % 100 == target_year and int(month) == article.month.month
    ]
    if not values:
        raise TempExportError(f"해당 연월의 일평균 수출액을 찾지 못했습니다: {article.title}")
    days_match = re.search(
        rf"조업일수\[.*?\(’?{target_year:02d}\)([\d.]+)일\]",
        compact,
    )
    working_days = float(days_match.group(1)) if days_match else None
    return values[-1], working_days, compact[:1200]


def monthly_final_values() -> dict[str, float]:
    if not GENERAL_WORKBOOK.exists():
        return {}
    workbook = load_workbook(GENERAL_WORKBOOK, read_only=True, data_only=True)
    sheet = workbook.active
    result: dict[str, float] = {}
    for month_value, _, _, _, daily_average in sheet.iter_rows(
        min_row=2, min_col=1, max_col=5, values_only=True
    ):
        if isinstance(month_value, (date, datetime)) and isinstance(daily_average, (int, float)):
            result[f"{month_value.year:04d}-{month_value.month:02d}"] = float(daily_average)
    workbook.close()
    return result


def load_history() -> dict:
    if not HISTORY_PATH.exists():
        return {"updated_at": None, "observations": {}}
    try:
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"updated_at": None, "observations": {}}


def save_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def workbook_update_months(workbook_path: Path, end: YearMonth) -> list[YearMonth]:
    """Return only incomplete and newly reached months for the normal UI update."""
    if not workbook_path.exists():
        return months_between(YearMonth(2026, 1), end)

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    sheet = workbook["Data"] if "Data" in workbook.sheetnames else workbook.active
    known: dict[YearMonth, tuple[object, object, object]] = {}
    for month_value, days_10, days_20, final in sheet.iter_rows(
        min_row=2, min_col=1, max_col=4, values_only=True
    ):
        if isinstance(month_value, (date, datetime)):
            month = YearMonth(month_value.year, month_value.month)
        elif isinstance(month_value, str) and re.match(r"\d{4}-\d{2}", month_value):
            month = YearMonth.parse(month_value[:7])
        else:
            continue
        if month <= end:
            known[month] = (days_10, days_20, final)
    workbook.close()

    if not known:
        return months_between(YearMonth(2026, 1), end)

    targets = {
        month
        for month, values in known.items()
        if any(value in (None, "") for value in values)
    }
    latest = max(known)
    if latest < end:
        targets.update(months_between(latest.next(), end))
    # The current/latest month is cheap to recheck and may receive a revised PDF.
    targets.add(end)
    return sorted(targets)


def collect(target_months: list[YearMonth]) -> dict:
    if not target_months:
        return load_history()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    articles = discover_articles(min(target_months), max(target_months))
    history = load_history()
    observations: dict[str, dict] = history.setdefault("observations", {})
    final_values = monthly_final_values()
    month_list = ", ".join(str(month) for month in target_months)
    print(f"PLAN {len(target_months)}개월: {month_list}", flush=True)
    for month in target_months:
        month_key = str(month)
        month_entry = observations.setdefault(month_key, {"month": month_key})
        month_entry["final_daily_average_100m_usd"] = final_values.get(month_key)
        for period_days in (10, 20):
            key = (month_key, period_days)
            article = articles.get(key)
            field = f"days_{period_days}"
            if not article:
                print(f"PENDING {month_key} 1-{period_days}일 잠정치: 공식 게시물 미발표", flush=True)
                continue
            print(f"COLLECT {month_key} 1-{period_days}일", flush=True)
            filename, pdf_url, payload = article_pdf(article)
            pdf_path = OUTPUT_DIR / f"{month_key}_01-{period_days}_잠정수출입.pdf"
            digest = hashlib.sha256(payload).hexdigest()
            if not pdf_path.exists() or hashlib.sha256(pdf_path.read_bytes()).hexdigest() != digest:
                pdf_path.write_bytes(payload)
            daily_average, working_days, _ = parse_daily_average(article, payload)
            month_entry[field] = {
                "daily_average_exports_100m_usd": daily_average,
                "working_days": working_days,
                "title": article.title,
                "published": article.published,
                "post_url": BASE_URL + DETAIL_PATH,
                "pdf_url": pdf_url,
                "local_pdf": str(pdf_path),
                "pdf_filename_original": filename,
                "sha256": digest,
            }
            print(f"SAVED {month_key} 1-{period_days}일: {daily_average:.1f}억 달러", flush=True)
    history["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    history["source"] = BASE_URL + LIST_PATH + "?mi=2891&bbsId=1362"
    save_json(HISTORY_PATH, history)
    return history


def workbook_month_row(sheet, month: str) -> int | None:
    for row in range(2, sheet.max_row + 1):
        value = sheet.cell(row, 1).value
        if isinstance(value, (date, datetime)) and f"{value.year:04d}-{value.month:02d}" == month:
            return row
        if isinstance(value, str) and value[:7] == month:
            return row
    return None


def copy_row_style(sheet, source_row: int, target_row: int) -> None:
    for column in range(1, sheet.max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        target._style = copy.copy(source._style)
        target.number_format = source.number_format
        target.alignment = copy.copy(source.alignment)
        target.border = copy.copy(source.border)
        target.fill = copy.copy(source.fill)
        target.font = copy.copy(source.font)


def update_workbook(
    history: dict,
    workbook_path: Path = WORKBOOK_PATH,
    target_months: list[YearMonth] | None = None,
) -> int:
    if not workbook_path.exists():
        raise TempExportError(
            f"통합문서가 없습니다: {workbook_path}. 초기 data/3.temp_export.xlsx를 먼저 생성하세요."
        )
    workbook = load_workbook(workbook_path)
    sheet = workbook["Data"] if "Data" in workbook.sheetnames else workbook.active
    changed = 0
    target_keys = {str(month) for month in target_months} if target_months is not None else None
    for month, entry in sorted(history.get("observations", {}).items()):
        if target_keys is not None and month not in target_keys:
            continue
        row = workbook_month_row(sheet, month)
        if row is None:
            row = sheet.max_row + 1
            copy_row_style(sheet, max(2, row - 1), row)
            sheet.cell(row, 1).value = datetime.strptime(month, "%Y-%m")
            sheet.cell(row, 5).value = f'=IF(OR(B{row}="",D{row}=""),"",B{row}-D{row})'
            sheet.cell(row, 6).value = f'=IF(OR(C{row}="",D{row}=""),"",C{row}-D{row})'
        values = {
            2: entry.get("days_10", {}).get("daily_average_exports_100m_usd"),
            3: entry.get("days_20", {}).get("daily_average_exports_100m_usd"),
            4: entry.get("final_daily_average_100m_usd"),
            7: entry.get("days_10", {}).get("local_pdf"),
            8: entry.get("days_20", {}).get("local_pdf"),
            9: entry.get("days_10", {}).get("pdf_url"),
            10: entry.get("days_20", {}).get("pdf_url"),
            11: history.get("updated_at"),
        }
        for column, value in values.items():
            if value is not None and sheet.cell(row, column).value != value:
                sheet.cell(row, column).value = value
                changed += 1
    if sheet.tables:
        sheet.tables[next(iter(sheet.tables))].ref = f"A1:K{sheet.max_row}"

    if "Chart" in workbook.sheetnames:
        chart_sheet = workbook["Chart"]
        data_rows = [
            row
            for row in range(2, sheet.max_row + 1)
            if sheet.cell(row, 1).value not in (None, "")
        ]
        for index, source_row in enumerate(data_rows, start=4):
            chart_sheet.cell(index, 1).value = f'=TEXT(\'Data\'!A{source_row},"yyyy-mm")'
            for target_column, source_column in ((2, "B"), (3, "C")):
                chart_sheet.cell(index, target_column).value = f"='Data'!{source_column}{source_row}"
            chart_sheet.cell(index, 4).value = (
                f"='Data'!D{source_row}" if sheet.cell(source_row, 4).value not in (None, "") else None
            )
        last_chart_row = 3 + len(data_rows)
        if chart_sheet._charts and data_rows:
            chart = chart_sheet._charts[0]
            for index, series in enumerate(chart.ser, start=2):
                series_end_row = last_chart_row
                if index == 4:
                    final_rows = [
                        offset
                        for offset, source_row in enumerate(data_rows, start=4)
                        if sheet.cell(source_row, 4).value not in (None, "")
                    ]
                    if final_rows:
                        series_end_row = final_rows[-1]
                series.val.numRef.f = (
                    f"'Chart'!${chr(64 + index)}$4:${chr(64 + index)}${series_end_row}"
                )
                if series.cat.strRef is not None:
                    series.cat.strRef.f = f"'Chart'!$A$4:$A${series_end_row}"
                elif series.cat.numRef is not None:
                    series.cat.numRef.f = f"'Chart'!$A$4:$A${series_end_row}"
    workbook.save(workbook_path)
    workbook.close()
    return changed


def current_month() -> YearMonth:
    today = date.today()
    return YearMonth(today.year, today.month)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--from-month",
        type=YearMonth.parse,
        help="지정하면 해당 월부터 강제로 다시 확인합니다. 생략하면 빈 값과 최신 월만 확인합니다.",
    )
    parser.add_argument("--through", type=YearMonth.parse, default=current_month())
    parser.add_argument("--workbook", type=Path, default=WORKBOOK_PATH)
    parser.add_argument("--download-only", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    if args.from_month is not None and args.through < args.from_month:
        raise SystemExit("종료월은 시작월보다 빠를 수 없습니다.")
    try:
        target_months = (
            months_between(args.from_month, args.through)
            if args.from_month is not None
            else workbook_update_months(args.workbook.resolve(), args.through)
        )
        history = collect(target_months)
        if not args.download_only:
            changed = update_workbook(history, args.workbook.resolve(), target_months)
            print(f"UPDATED {changed}개 셀: {args.workbook.resolve()}", flush=True)
    except Exception as exc:
        print(f"FAILED {exc}", flush=True)
        raise SystemExit(1) from exc
    if args.json:
        print("RESULT_JSON " + json.dumps(history, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
