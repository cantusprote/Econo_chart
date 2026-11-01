#!/usr/bin/env python3
"""Download and parse the latest monthly export/import report from MOTIR."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import hashlib
import html
import io
import json
import re
from pathlib import Path
import shutil
import urllib.parse
import urllib.request

from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "Export_trend"
BASE_URL = "https://www.motir.go.kr"
BOARD_PATH = "/kor/article/ATCL3f49a5a8c"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36"
)
PRODUCT_GROUPS = (
    ("반도체", "디스플레이", "무선통신기기", "컴퓨터", "자동차", "자동차부품", "선박"),
    ("석유제품", "석유화학", "이차전지", "일반기계", "전기기기", "철강", "비철금속"),
    ("농수산식품", "화장품", "생활용품", "섬유", "가전", "바이오헬스", "전체"),
)
REGIONS = ("중국", "미국", "일본", "인도", "아세안", "EU(27)", "중남미", "중동", "CIS", "전체")


class ExportReportError(RuntimeError):
    pass


@dataclass(frozen=True)
class Article:
    report_month: str
    article_id: str
    title: str
    published: str
    department: str
    post_url: str
    pdf_url: str
    pdf_filename: str


def build_opener() -> urllib.request.OpenerDirector:
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor())


def request_bytes(
    opener: urllib.request.OpenerDirector,
    url: str,
    *,
    referer: str | None = None,
    timeout: int = 60,
) -> bytes:
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/pdf,application/octet-stream,text/html;q=0.9,*/*;q=0.8",
    }
    if referer:
        headers["Referer"] = referer
    request = urllib.request.Request(url, headers=headers)
    try:
        with opener.open(request, timeout=timeout) as response:
            return response.read()
    except Exception as exc:
        raise ExportReportError(f"다운로드 실패: {url}: {exc}") from exc


def decode_html(payload: bytes) -> str:
    return payload.decode("utf-8", errors="replace")


def normalized_text(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", value))).strip()


def report_month_from_title(title: str) -> str | None:
    match = re.fullmatch(r"\s*(20\d{2})년\s*(\d{1,2})월\s*수출입\s*동향\s*", title)
    if not match:
        return None
    year, month = map(int, match.groups())
    if not 1 <= month <= 12:
        return None
    return f"{year:04d}-{month:02d}"


def discover_candidate_ids(opener: urllib.request.OpenerDirector) -> list[tuple[str, str]]:
    found: dict[str, str] = {}
    pattern = re.compile(
        r"article\.view\(['\"](?P<id>\d+)['\"]\).*?<i>(?P<title>.*?)</i>", re.S
    )
    for keyword in ("수출입 동향", "수출입동향"):
        keyword_found = False
        for page in range(1, 4):
            query = urllib.parse.urlencode(
                {
                    "pageIndex": page,
                    "searchCondition": 1,
                    "searchKeyword": keyword,
                }
            )
            body = decode_html(request_bytes(opener, f"{BASE_URL}{BOARD_PATH}?{query}"))
            for match in pattern.finditer(body):
                title = normalized_text(match.group("title"))
                if report_month_from_title(title):
                    found[match.group("id")] = title
                    keyword_found = True
            if keyword_found:
                break
    if not found:
        raise ExportReportError("산업통상부에서 월별 수출입 동향 게시물을 찾지 못했습니다.")
    return list(found.items())


def parse_article(opener: urllib.request.OpenerDirector, article_id: str) -> Article:
    post_url = f"{BASE_URL}{BOARD_PATH}/{article_id}/view"
    body = decode_html(request_bytes(opener, post_url))
    title_match = re.search(r'<div class="detail-tit"><span>(.*?)</span>', body, re.S)
    if not title_match:
        raise ExportReportError(f"게시물 제목을 찾지 못했습니다: {post_url}")
    title = normalized_text(title_match.group(1))
    report_month = report_month_from_title(title)
    if not report_month:
        raise ExportReportError(f"월별 수출입 동향 게시물이 아닙니다: {title}")

    def detail(label: str) -> str:
        match = re.search(
            rf"<li><em>{re.escape(label)}</em><span>(.*?)</span></li>", body, re.S
        )
        return normalized_text(match.group(1)) if match else ""

    department = detail("담당부서")
    published = detail("등록일")
    # MOTIR has used both normal hrefs and javascript:location.href wrappers
    # for attachment downloads. Keep both patterns so a site-template change
    # does not break monthly updates again.
    attachments = re.findall(
        r"<a\s+[^>]*href=[\"'](?:javascript:location\.href=[\"'])?"
        r"(?P<path>/attach/down/[^\"']+)(?:[\"'])?[^>]*>"
        r"(?P<name>.*?)</a>",
        body,
        re.S | re.I,
    )
    pdfs = [
        (path, normalized_text(name).rsplit(" [", 1)[0])
        for path, name in attachments
        if ".pdf" in normalized_text(name).lower()
    ]
    if not pdfs:
        raise ExportReportError(f"PDF 첨부파일을 찾지 못했습니다: {post_url}")
    pdf_path, pdf_filename = pdfs[-1]
    return Article(
        report_month=report_month,
        article_id=article_id,
        title=title,
        published=published,
        department=department,
        post_url=post_url,
        pdf_url=urllib.parse.urljoin(BASE_URL, pdf_path),
        pdf_filename=pdf_filename,
    )


def find_latest_article(opener: urllib.request.OpenerDirector) -> Article:
    articles: list[Article] = []
    errors: list[str] = []
    for article_id, _ in discover_candidate_ids(opener):
        try:
            article = parse_article(opener, article_id)
        except ExportReportError as exc:
            errors.append(str(exc))
            continue
        if article.department == "수출입과":
            articles.append(article)
    if not articles:
        raise ExportReportError("최신 수출입과 PDF를 확인하지 못했습니다: " + " / ".join(errors))
    return max(articles, key=lambda item: (item.report_month, item.published, item.article_id))


def extract_pdf_text(payload: bytes) -> tuple[PdfReader, list[str]]:
    if not payload.startswith(b"%PDF-"):
        preview = payload[:120].decode("utf-8", errors="replace").replace("\n", " ")
        raise ExportReportError(f"PDF 대신 다른 응답을 받았습니다: {preview}")
    reader = PdfReader(io.BytesIO(payload))
    if len(reader.pages) < 4:
        raise ExportReportError(f"PDF 페이지 수가 비정상적입니다: {len(reader.pages)}")
    pages = [page.extract_text(extraction_mode="layout") or "" for page in reader.pages]
    if sum(len(page) for page in pages) < 5_000:
        raise ExportReportError("PDF 텍스트 추출량이 너무 적습니다.")
    return reader, pages


def number_tokens(line: str, *, signed: bool = False) -> list[float]:
    pattern = r"(?:[+△-])?\d[\d,]*(?:\.\d+)?" if signed else r"\d[\d,]*(?:\.\d+)?"
    values: list[float] = []
    for token in re.findall(pattern, line):
        negative = token.startswith(("△", "-"))
        cleaned = token.lstrip("+△-").replace(",", "")
        value = float(cleaned)
        values.append(-value if negative else value)
    return values


def parse_table_groups(section: str, groups: tuple[tuple[str, ...], ...]) -> list[dict]:
    lines = [line.strip() for line in section.splitlines() if line.strip()]
    amount_lines = [line for line in lines if re.match(r"^수출액\s", line)]
    growth_lines = [line for line in lines if re.match(r"^증감률\s", line)]
    rank_lines = [line for line in lines if re.match(r"^역대순위\s", line)]
    if len(amount_lines) < len(groups) or len(growth_lines) < len(groups):
        raise ExportReportError("품목별 수출 표의 행을 충분히 찾지 못했습니다.")
    result: list[dict] = []
    for index, names in enumerate(groups):
        amounts = number_tokens(amount_lines[index])
        growths = number_tokens(growth_lines[index], signed=True)
        ranks = re.findall(r"\d+위|-", rank_lines[index]) if index < len(rank_lines) else []
        if len(amounts) < len(names) or len(growths) < len(names):
            raise ExportReportError(f"품목 표 {index + 1}번째 묶음의 열 수가 맞지 않습니다.")
        for position, name in enumerate(names):
            result.append(
                {
                    "name": name,
                    "export_100m_usd": amounts[position],
                    "yoy_pct": growths[position],
                    "historical_rank": ranks[position] if position < len(ranks) else None,
                }
            )
    return result


def parse_region_table(section: str) -> list[dict]:
    lines = [line.strip() for line in section.splitlines() if line.strip()]
    amount_line = next((line for line in lines if re.match(r"^수출액\s", line)), "")
    growth_line = next((line for line in lines if re.match(r"^증감률\s", line)), "")
    rank_line = next((line for line in lines if re.match(r"^역대순위\s", line)), "")
    amounts = number_tokens(amount_line)
    growths = number_tokens(growth_line, signed=True)
    ranks = re.findall(r"\d+위|-", rank_line)
    if len(amounts) < len(REGIONS) or len(growths) < len(REGIONS):
        raise ExportReportError("지역별 수출 표의 열 수가 맞지 않습니다.")
    return [
        {
            "name": name,
            "export_100m_usd": amounts[index],
            "yoy_pct": growths[index],
            "historical_rank": ranks[index] if index < len(ranks) else None,
        }
        for index, name in enumerate(REGIONS)
    ]


def parse_monthly_trade(full_text: str) -> list[dict]:
    match = re.search(
        r"연도별\s*수출입\s*실적\s*통계(?P<section>.*?)(?:참고\s*3|세부\s*품목\s*예시)",
        full_text,
        re.S,
    )
    if not match:
        return []
    current_year: int | None = None
    rows: list[dict] = []
    for line in match.group("section").splitlines():
        year_match = re.search(r"(20\d{2})년", line)
        if year_match:
            current_year = int(year_match.group(1))
        month_match = re.match(
            r"\s*(\d{1,2})\s+([\d,]+)\s+([△+\-]?\d+(?:\.\d+)?)\s+"
            r"([\d,]+)\s+([△+\-]?\d+(?:\.\d+)?)\s+([△+\-]?[\d,]+)\s*$",
            line,
        )
        if not month_match or current_year is None:
            continue
        month = int(month_match.group(1))
        if not 1 <= month <= 12:
            continue

        def signed_number(token: str) -> float:
            negative = token.startswith(("△", "-"))
            value = float(token.lstrip("△+-").replace(",", ""))
            return -value if negative else value

        rows.append(
            {
                "month": f"{current_year:04d}-{month:02d}",
                "exports_100m_usd": round(float(month_match.group(2).replace(",", "")) / 100, 2),
                "exports_yoy_pct": signed_number(month_match.group(3)),
                "imports_100m_usd": round(float(month_match.group(4).replace(",", "")) / 100, 2),
                "imports_yoy_pct": signed_number(month_match.group(5)),
                "trade_balance_100m_usd": round(signed_number(month_match.group(6)) / 100, 2),
            }
        )
    unique = {row["month"]: row for row in rows}
    return [unique[key] for key in sorted(unique)]


def parse_report(article: Article, payload: bytes) -> dict:
    reader, pages = extract_pdf_text(payload)
    full_text = "\n".join(pages)
    compact = re.sub(r"\s+", " ", full_text)
    if article.title.replace(" ", "") not in re.sub(r"\s+", "", full_text):
        raise ExportReportError("게시물 제목과 PDF 본문의 기준월이 일치하지 않습니다.")

    headline = re.search(
        r"수출\s*([\d,.]+)억\s*달러\s*\(([+△-]?[\d.]+)%\).*?"
        r"수입\s*([\d,.]+)억\s*달러\s*\(([+△-]?[\d.]+)%\).*?"
        r"(?:무역)?수지\s*([+△-]?[\d,.]+)억\s*달러\s*(흑자|적자)",
        compact,
    )
    if not headline:
        raise ExportReportError("PDF 첫 부분에서 수출·수입·무역수지를 파싱하지 못했습니다.")

    def signed(value: str) -> float:
        negative = value.startswith(("△", "-"))
        number = float(value.lstrip("+△-").replace(",", ""))
        return -number if negative else number

    exports = float(headline.group(1).replace(",", ""))
    export_yoy = signed(headline.group(2))
    imports = float(headline.group(3).replace(",", ""))
    import_yoy = signed(headline.group(4))
    balance = signed(headline.group(5))
    if headline.group(6) == "적자" and balance > 0:
        balance = -balance

    daily = re.search(
        r"일평균\s*수출(?:액|은|도)?\s*(?:전년\s*동월\s*대비\s*)?"
        r"([\d.]+)%\s*(증가|감소)한\s*([\d.]+)억\s*달러",
        compact,
    )
    if not daily:
        raise ExportReportError("PDF에서 일평균 수출액을 파싱하지 못했습니다.")
    daily_yoy = float(daily.group(1)) * (-1 if daily.group(2) == "감소" else 1)
    daily_value = float(daily.group(3))

    table_page = next(
        (
            page
            for page in pages
            if re.search(r"20\s*대", page)
            and re.search(r"9\s*대\s*주요\s*지역별", page)
        ),
        None,
    )
    if not table_page:
        raise ExportReportError("20대 품목·9대 지역 표가 있는 페이지를 찾지 못했습니다.")
    sections = re.split(r"□\s*\(\s*지역\s*\)", table_page, maxsplit=1)
    if len(sections) != 2:
        raise ExportReportError("품목·지역 표에서 지역 구분 제목을 찾지 못했습니다.")
    product_section, region_section = sections
    products = [row for row in parse_table_groups(product_section, PRODUCT_GROUPS) if row["name"] != "전체"]
    regions = [row for row in parse_region_table(region_section) if row["name"] != "전체"]
    monthly_trade = parse_monthly_trade(full_text)

    checks = {
        "trade_balance_difference": round(exports - imports - balance, 3),
        "product_total_difference": round(
            next(row["export_100m_usd"] for row in parse_table_groups(product_section, PRODUCT_GROUPS) if row["name"] == "전체") - exports,
            3,
        ),
        "region_total_difference": round(
            next(row["export_100m_usd"] for row in parse_region_table(region_section) if row["name"] == "전체") - exports,
            3,
        ),
    }
    warnings = [name for name, difference in checks.items() if abs(difference) > 1.0]
    return {
        "report_month": article.report_month,
        "title": article.title,
        "published": article.published,
        "department": article.department,
        "post_url": article.post_url,
        "pdf_url": article.pdf_url,
        "pdf_filename_original": article.pdf_filename,
        "pages": len(reader.pages),
        "summary": {
            "exports_100m_usd": exports,
            "exports_yoy_pct": export_yoy,
            "imports_100m_usd": imports,
            "imports_yoy_pct": import_yoy,
            "trade_balance_100m_usd": balance,
            "daily_average_exports_100m_usd": daily_value,
            "daily_average_exports_yoy_pct": daily_yoy,
        },
        "products": products,
        "regions": regions,
        "monthly_trade": monthly_trade,
        "validation": {
            "status": "warning" if warnings else "valid",
            "checks": checks,
            "warnings": warnings,
        },
    }


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def atomic_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def save_report(output_dir: Path, article: Article, payload: bytes, report: dict) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    parsed_dir = output_dir / "parsed"
    parsed_dir.mkdir(parents=True, exist_ok=True)
    year, month = article.report_month.split("-")
    pdf_path = output_dir / f"{year}년_{month}월_수출입동향.pdf"
    digest = hashlib.sha256(payload).hexdigest()
    status = "downloaded"
    if pdf_path.exists():
        existing_digest = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
        if existing_digest == digest:
            status = "already_downloaded"
        else:
            archive = output_dir / f"{pdf_path.stem}_revision_{existing_digest[:8]}.pdf"
            if not archive.exists():
                shutil.copy2(pdf_path, archive)
            status = "updated_revision"
    if status != "already_downloaded":
        pdf_path.write_bytes(payload)

    report.update(
        {
            "local_pdf": str(pdf_path),
            "sha256": digest,
            "file_size": len(payload),
            "downloaded_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "download_status": status,
        }
    )
    parsed_path = parsed_dir / f"{article.report_month}.json"
    atomic_json(parsed_path, report)

    history_path = output_dir / "export_history.json"
    history = load_json(history_path, {"reports": []})
    reports = [item for item in history.get("reports", []) if item.get("report_month") != article.report_month]
    reports.append(report)
    reports.sort(key=lambda item: item["report_month"])
    history = {"updated_at": report["downloaded_at"], "reports": reports}
    atomic_json(history_path, history)

    manifest_path = output_dir / "manifest.json"
    manifest = load_json(manifest_path, {"reports": {}})
    manifest.setdefault("reports", {})[article.report_month] = {
        "title": article.title,
        "published": article.published,
        "post_url": article.post_url,
        "pdf_url": article.pdf_url,
        "local_pdf": str(pdf_path),
        "sha256": digest,
        "downloaded_at": report["downloaded_at"],
        "status": status,
    }
    manifest["latest_month"] = article.report_month
    manifest["updated_at"] = report["downloaded_at"]
    atomic_json(manifest_path, manifest)
    return report


def run(output_dir: Path = OUTPUT_DIR) -> dict:
    opener = build_opener()
    print("SEARCH 최신 수출입 동향", flush=True)
    article = find_latest_article(opener)
    print(f"FOUND {article.title} ({article.published})", flush=True)
    # Open the article in the same cookie session immediately before the file.
    request_bytes(opener, article.post_url)
    print(f"DOWNLOAD {article.pdf_filename}", flush=True)
    payload = request_bytes(opener, article.pdf_url, referer=article.post_url)
    print(f"VALIDATE {len(payload):,} bytes", flush=True)
    report = parse_report(article, payload)
    print(
        f"EXTRACT 수출={report['summary']['exports_100m_usd']} "
        f"일평균={report['summary']['daily_average_exports_100m_usd']} "
        f"품목={len(report['products'])} 지역={len(report['regions'])}",
        flush=True,
    )
    saved = save_report(output_dir, article, payload, report)
    print(f"SAVED {saved['local_pdf']}", flush=True)
    return saved


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = run(args.output_dir.resolve())
    except Exception as exc:
        print(f"FAILED {exc}", flush=True)
        raise SystemExit(1) from exc
    if args.json:
        print("RESULT_JSON " + json.dumps(result, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
