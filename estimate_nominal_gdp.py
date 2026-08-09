#!/usr/bin/env python3
"""Estimate the current year's Korean nominal GDP from the latest NABO outlook."""

from __future__ import annotations

import argparse
import io
import json
import re
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "data" / "2.GDP.xlsx"
NABO_BASE = "https://www.nabo.go.kr"
NABO_LIST = f"{NABO_BASE}/ko/report/analysisList.do"
USER_AGENT = "Mozilla/5.0 EconoChart/1.0"


class NominalGDPForecastError(RuntimeError):
    pass


def http_get(url: str, timeout: int = 60) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def latest_actual_gdp(path: Path, target_year: int) -> tuple[int, float]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["Sheet1"]
    candidates = [
        (int(year), float(gdp))
        for year, gdp in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True)
        if isinstance(year, (int, float))
        and isinstance(gdp, (int, float))
        and int(year) < target_year
    ]
    workbook.close()
    if not candidates:
        raise NominalGDPForecastError("기준이 될 과거 명목 GDP 실적이 없습니다.")
    return max(candidates, key=lambda item: item[0])


def find_latest_nabo_report(target_year: int) -> tuple[str, str, str]:
    query = urllib.parse.urlencode(
        {
            "key": "2509250001",
            "sc": "title",
            "sw": f"{target_year}년 NABO 경제전망",
            "pageUnit": 30,
            "orderBy": "pubDate_desc",
        }
    )
    list_url = f"{NABO_LIST}?{query}"
    html = http_get(list_url).decode("utf-8", "replace")
    pattern = re.compile(
        r'<a href="(?P<href>analysisView\.do\?idx=(?P<idx>\d+)[^"]*)">'
        r'.*?<div class="text_wrap">\s*(?P<title>.*?)\s*</div>'
        r'.*?<div class="date">(?P<date>\d{4}-\d{2}-\d{2})</div>',
        re.DOTALL,
    )
    reports: list[tuple[str, str, str]] = []
    for match in pattern.finditer(html):
        title = re.sub(r"\s+", " ", match.group("title")).strip()
        if f"{target_year}년 NABO 경제전망" in title:
            reports.append((match.group("date"), match.group("idx"), title))
    if not reports:
        raise NominalGDPForecastError(
            f"NABO에서 {target_year}년 경제전망 보고서를 찾지 못했습니다."
        )
    published, index, title = max(reports)
    return f"{NABO_BASE}/ko/report/analysisView.do?idx={index}&key=2509250001", published, title


def report_pdf_url(report_url: str) -> str:
    html = http_get(report_url).decode("utf-8", "replace")
    match = re.search(
        r"href=['\"](?P<href>/board/file/down\.do\?fid=\d+)['\"][^>]*>.*?"
        r"<span class=\"text_wrap\">[^<]*경제전망[^<]*\.pdf</span>",
        html,
        re.DOTALL | re.IGNORECASE,
    )
    if not match:
        raise NominalGDPForecastError("NABO 보고서의 PDF 첨부파일을 찾지 못했습니다.")
    return urllib.parse.urljoin(NABO_BASE, match.group("href"))


def extract_rates(pdf_payload: bytes, target_year: int) -> tuple[float, float, float, float, int]:
    reader = PdfReader(io.BytesIO(pdf_payload))
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if "명목GDP 성장률 전망" not in text:
            continue

        def last_rate(label_pattern: str) -> float:
            match = re.search(
                rf"{label_pattern}\s+((?:-?\d+(?:\.\d+)?\s+){{2,}}-?\d+(?:\.\d+)?)",
                text,
            )
            if not match:
                raise NominalGDPForecastError(
                    f"NABO 표에서 {label_pattern} 전망값을 읽지 못했습니다."
                )
            values = [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", match.group(1))]
            return values[-1]

        nominal = last_rate(r"명목\s*(?:국내총생산|GDP)")
        real = last_rate(r"실질\s*(?:국내총생산|GDP)")
        deflator = last_rate(r"GDP\s*디플레이터")
        cpi = last_rate(r"소비자물가")
        if not all(-20 < value < 30 for value in (nominal, real, deflator, cpi)):
            raise NominalGDPForecastError("추출된 성장률 값의 범위를 확인할 수 없습니다.")
        return nominal, real, deflator, cpi, page_number
    raise NominalGDPForecastError(
        f"NABO PDF에서 {target_year}년 명목GDP 성장률 전망 표를 찾지 못했습니다."
    )


def estimate(path: Path = DEFAULT_WORKBOOK, target_year: int | None = None) -> dict:
    year = target_year or date.today().year
    base_year, base_gdp = latest_actual_gdp(path, year)
    if base_year != year - 1:
        raise NominalGDPForecastError(
            f"{year}년 전망을 계산하려면 {year - 1}년 명목 GDP 실적이 필요하지만 "
            f"현재 최신 실적은 {base_year}년입니다."
        )
    report_url, published, title = find_latest_nabo_report(year)
    pdf_url = report_pdf_url(report_url)
    nominal_rate, real_rate, deflator_rate, cpi_rate, page = extract_rates(http_get(pdf_url), year)
    forecast_gdp = base_gdp * (1.0 + nominal_rate / 100.0)
    identity_rate = ((1.0 + real_rate / 100.0) * (1.0 + deflator_rate / 100.0) - 1.0) * 100.0
    return {
        "target_year": year,
        "calculated_at": date.today().isoformat(),
        "base_year": base_year,
        "base_nominal_gdp_trillion_krw": round(base_gdp, 4),
        "nominal_growth_forecast_pct": nominal_rate,
        "real_growth_forecast_pct": real_rate,
        "gdp_deflator_forecast_pct": deflator_rate,
        "cpi_forecast_pct": cpi_rate,
        "identity_nominal_growth_pct": round(identity_rate, 4),
        "forecast_nominal_gdp_trillion_krw": round(forecast_gdp, 4),
        "formula": f"{base_gdp:.4f} × (1 + {nominal_rate:.1f}/100) = {forecast_gdp:.4f}조원",
        "classification": "국회예산정책처 공식 성장률을 적용한 프로그램 계산값",
        "report_title": title,
        "report_published": published,
        "report_page": page,
        "report_url": report_url,
        "pdf_url": pdf_url,
        "actual_source": "data/2.GDP.xlsx (한국은행·국가지표체계 명목 GDP 실적)",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--year", type=int, default=date.today().year)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = estimate(args.workbook.resolve(), args.year)
    if args.json:
        print(json.dumps(result, ensure_ascii=False))
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
