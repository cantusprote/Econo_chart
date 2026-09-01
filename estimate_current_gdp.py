#!/usr/bin/env python3
"""Nowcast current-year nominal GDP from BOK actuals and official forecasts."""

from __future__ import annotations

import argparse
import html
import io
import json
import re
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

from pypdf import PdfReader

from estimate_nominal_gdp import DEFAULT_WORKBOOK, estimate as nabo_estimate, http_get


BOK_RSS_STATS = "https://www.bok.or.kr/portal/bbs/B0000501/news.rss?menuNo=201264"
BOK_RSS_OUTLOOK = "https://www.bok.or.kr/portal/bbs/B0000502/news.rss?menuNo=201265"


class CurrentGDPError(RuntimeError):
    pass


def rss_items(url: str) -> list[dict[str, str]]:
    root = ET.fromstring(http_get(url))
    return [
        {
            "title": item.findtext("title") or "",
            "url": item.findtext("link") or "",
            "published": item.findtext("pubDate") or "",
            "description": item.findtext("description") or "",
        }
        for item in root.findall(".//item")
    ]


def compact_text(payload: bytes) -> str:
    decoded = payload.decode("utf-8", "replace")
    return re.sub(r"\s+", "", html.unescape(re.sub(r"<[^>]+>", " ", decoded)))


def page_pdf_url(page_url: str, title_fragment: str) -> str:
    page = http_get(page_url).decode("utf-8", "replace")
    candidates = re.findall(r'href="([^"]+\.pdf(?:\?[^"]*)?)"[^>]*title="([^"]*)"', page, re.I)
    for href, title in candidates:
        if title_fragment in title:
            return urllib.parse.urljoin("https://www.bok.or.kr", html.unescape(href))
    if candidates:
        return urllib.parse.urljoin("https://www.bok.or.kr", html.unescape(candidates[0][0]))
    raise CurrentGDPError(f"한국은행 발표 페이지에서 PDF를 찾지 못했습니다: {page_url}")


def latest_income_release(year: int) -> dict[str, str | int]:
    candidates: list[tuple[int, dict[str, str]]] = []
    for item in rss_items(BOK_RSS_STATS):
        match = re.search(rf"{year}년\s*([1-4])/4분기(?:\s*및\s*연간)?\s*국민소득\(잠정\)", item["title"])
        if match:
            candidates.append((int(match.group(1)), item))
    if not candidates:
        raise CurrentGDPError(f"한국은행 RSS에서 {year}년 국민소득 잠정치를 찾지 못했습니다.")
    quarter, item = max(candidates, key=lambda entry: entry[0])
    return {**item, "quarter": quarter}


def extract_income_pdf(pdf: bytes, actual_quarter: int) -> tuple[list[float], list[float], list[float]]:
    pages = [page.extract_text() or "" for page in PdfReader(io.BytesIO(pdf)).pages]
    normalized = [re.sub(r"\s+", "", text) for text in pages]

    income_page = next((text for text in normalized if "명목소득지표(당해년가격기준)" in text), None)
    if not income_page:
        raise CurrentGDPError("국민소득 PDF에서 명목소득지표 표를 찾지 못했습니다.")
    level_match = re.search(r"명목국내총생산(?P<values>.*?)\(명목GDP\)", income_page)
    if not level_match:
        raise CurrentGDPError("국민소득 PDF에서 분기 명목 GDP 금액을 읽지 못했습니다.")
    # PDF extraction removes table-cell separators (for example 635.8 639.5
    # becomes 635.8639.5), while every published level has one decimal place.
    levels = [float(value) for value in re.findall(r"\d{3,4}\.\d", level_match.group("values"))]
    needed = 8 + actual_quarter
    if len(levels) < needed:
        raise CurrentGDPError(f"분기 명목 GDP 관측치가 부족합니다: {len(levels)}/{needed}")
    levels = levels[-needed:]
    previous_levels = levels[4:8]
    current_levels = levels[8:]

    deflator_page = next(
        (text for text in normalized if "부가가치항목별명목국내총생산및디플레이터" in text),
        None,
    )
    if not deflator_page:
        raise CurrentGDPError("국민소득 PDF에서 GDP 디플레이터 표를 찾지 못했습니다.")
    deflator_match = re.search(r"GDP디플레이터\d*\)(?P<values>.*?)내수", deflator_page)
    if not deflator_match:
        raise CurrentGDPError("GDP 디플레이터 분기 실적을 읽지 못했습니다.")
    deflators = [float(value) for value in re.findall(r"-?\d{1,2}\.\d", deflator_match.group("values"))]
    if len(deflators) < needed:
        raise CurrentGDPError("GDP 디플레이터 관측치가 부족합니다.")
    current_deflators = deflators[-needed:][8:]
    return previous_levels, current_levels, current_deflators


def real_gdp_actuals(year: int) -> tuple[dict[int, float], list[dict[str, str]]]:
    values: dict[int, float] = {}
    sources: list[dict[str, str]] = []
    for item in rss_items(BOK_RSS_STATS):
        match = re.search(rf"{year}년\s*([1-4])/4분기\s*실질\s*국내총생산\(속보\)", item["title"])
        if not match:
            continue
        quarter = int(match.group(1))
        page = compact_text(http_get(item["url"]))
        rate_match = re.search(r"전년동기대비(-?\d+(?:\.\d+)?)%성장", page)
        if rate_match:
            values[quarter] = float(rate_match.group(1))
            sources.append(item)
    return values, sources


def latest_bok_outlook(year: int) -> dict[str, str | float]:
    candidates = [
        item
        for item in rss_items(BOK_RSS_OUTLOOK)
        if re.search(rf"경제전망\({year}년\s*\d+월\)", item["title"])
    ]
    if not candidates:
        raise CurrentGDPError(f"한국은행 RSS에서 {year}년 경제전망을 찾지 못했습니다.")
    item = candidates[0]
    page = compact_text(http_get(item["url"]))
    real_patterns = (
        r"올해성장률전망[,，:]?(-?\d+(?:\.\d+)?)%",
        r"국내경제는올해(-?\d+(?:\.\d+)?)%",
        r"올해경제성장률(?:은|을)?(-?\d+(?:\.\d+)?)%",
        r"올해성장률(?:은|을)?(-?\d+(?:\.\d+)?)%",
    )
    cpi_patterns = (
        r"소비자물가는올해(-?\d+(?:\.\d+)?)%",
        r"소비자물가상승률(?:은|을)?올해(-?\d+(?:\.\d+)?)%",
        r"소비자물가.{0,80}?(-?\d+(?:\.\d+)?)%상승전망",
    )
    real_match = next((match for pattern in real_patterns if (match := re.search(pattern, page))), None)
    cpi_match = next((match for pattern in cpi_patterns if (match := re.search(pattern, page))), None)
    if not real_match or not cpi_match:
        raise CurrentGDPError("한국은행 경제전망에서 성장률 또는 물가 전망을 읽지 못했습니다.")
    return {
        **item,
        "real_growth_pct": float(real_match.group(1)),
        "cpi_pct": float(cpi_match.group(1)),
    }


def nowcast(path: Path = DEFAULT_WORKBOOK, target_year: int | None = None) -> dict:
    year = target_year or date.today().year
    nabo = nabo_estimate(path, year)
    income = latest_income_release(year)
    actual_quarter = int(income["quarter"])
    income_pdf_url = page_pdf_url(str(income["url"]), "국민소득")
    previous_levels, actual_levels, actual_deflators = extract_income_pdf(
        http_get(income_pdf_url), actual_quarter
    )
    real_actuals, real_sources = real_gdp_actuals(year)
    income_page = compact_text(http_get(str(income["url"])))
    provisional_real = re.search(r"실질국내총생산\(GDP\).*?전년동기대비(-?\d+(?:\.\d+)?)%성장", income_page)
    if provisional_real:
        real_actuals[actual_quarter] = float(provisional_real.group(1))
    outlook = latest_bok_outlook(year)

    known_real = [real_actuals[q] for q in sorted(real_actuals)]
    missing_real_count = 4 - len(known_real)
    implied_missing_real = (
        (4 * float(outlook["real_growth_pct"]) - sum(known_real)) / missing_real_count
        if missing_real_count
        else float(outlook["real_growth_pct"])
    )
    official_deflator_anchor = float(nabo["gdp_deflator_forecast_pct"]) + (
        float(outlook["cpi_pct"]) - float(nabo["cpi_forecast_pct"])
    )
    latest_actual_deflator = actual_deflators[-1]

    quarters: list[dict] = []
    forecast_levels: list[float] = []
    low_levels: list[float] = []
    high_levels: list[float] = []
    for quarter in range(1, 5):
        prior = previous_levels[quarter - 1]
        if quarter <= actual_quarter:
            level = actual_levels[quarter - 1]
            deflator = actual_deflators[quarter - 1]
            real = real_actuals.get(quarter)
            status = "한국은행 명목 GDP 잠정 실적"
            low_level = high_level = level
        else:
            real = real_actuals.get(quarter, implied_missing_real)
            steps_ahead = quarter - actual_quarter
            recent_weight = max(0.0, 0.75 - 0.25 * steps_ahead)
            deflator = recent_weight * latest_actual_deflator + (1 - recent_weight) * official_deflator_anchor
            level = prior * (1 + real / 100) * (1 + deflator / 100)
            status = (
                "한국은행 실질 GDP 속보 + 디플레이터 추정"
                if quarter in real_actuals
                else "실질성장·디플레이터 기반 프로그램 추정"
            )
            real_margin = 0.0 if quarter in real_actuals else 0.5
            low_level = prior * (1 + (real - real_margin) / 100) * (1 + (deflator - 2.0) / 100)
            high_level = prior * (1 + (real + real_margin) / 100) * (1 + (deflator + 2.0) / 100)
        forecast_levels.append(level)
        low_levels.append(low_level)
        high_levels.append(high_level)
        quarters.append(
            {
                "quarter": quarter,
                "status": status,
                "prior_year_nominal_gdp": round(prior, 1),
                "real_growth_yoy_pct": None if real is None else round(real, 3),
                "deflator_yoy_pct": round(deflator, 3),
                "nominal_gdp_trillion_krw": round(level, 1),
                "formula": (
                    "공식 명목 GDP 잠정 실적"
                    if quarter <= actual_quarter
                    else f"{prior:.1f} × (1+{real:.3f}/100) × (1+{deflator:.3f}/100)"
                ),
            }
        )

    annual = sum(forecast_levels)
    annual_low = sum(low_levels)
    annual_high = sum(high_levels)
    base = float(nabo["base_nominal_gdp_trillion_krw"])
    annual_growth = (annual / base - 1) * 100
    estimated_quarters = 4 - actual_quarter
    return {
        "target_year": year,
        "calculated_at": date.today().isoformat(),
        "forecast_nominal_gdp_trillion_krw": round(annual, 1),
        "forecast_nominal_growth_pct": round(annual_growth, 2),
        "forecast_range_low_trillion_krw": round(annual_low, 1),
        "forecast_range_high_trillion_krw": round(annual_high, 1),
        "actual_quarters": actual_quarter,
        "estimated_quarters": estimated_quarters,
        "actual_share_pct": actual_quarter * 25,
        "method": "발표 분기 명목 GDP 실적 + 미발표 분기 실질성장률·GDP 디플레이터 추정",
        "quarters": quarters,
        "assumptions": {
            "bok_annual_real_growth_pct": outlook["real_growth_pct"],
            "bok_cpi_forecast_pct": outlook["cpi_pct"],
            "nabo_deflator_forecast_pct": nabo["gdp_deflator_forecast_pct"],
            "nabo_cpi_forecast_pct": nabo["cpi_forecast_pct"],
            "adjusted_deflator_anchor_pct": round(official_deflator_anchor, 3),
            "latest_actual_deflator_pct": latest_actual_deflator,
            "missing_quarter_real_growth_pct": round(implied_missing_real, 3),
        },
        "sources": [
            {"label": str(income["title"]), "published": str(income["published"]), "url": str(income["url"])},
            *[
                {"label": source["title"], "published": source["published"], "url": source["url"]}
                for source in real_sources
                if (match := re.search(r"([1-4])/4분기", source["title"]))
                and int(match.group(1)) > actual_quarter
            ],
            {"label": str(outlook["title"]), "published": str(outlook["published"]), "url": str(outlook["url"])},
            {"label": str(nabo["report_title"]), "published": str(nabo["report_published"]), "url": str(nabo["report_url"])},
        ],
        "quality_note": (
            f"{actual_quarter}개 분기는 한국은행 명목 GDP 잠정 실적이며, "
            f"나머지 {estimated_quarters}개 분기는 공식자료 기반 프로그램 추정입니다. "
            "범위는 미발표 분기 디플레이터 ±2%p와 미발표 실질성장률 ±0.5%p를 적용했습니다."
        ),
        "confidence": "낮음" if actual_quarter <= 1 else "보통" if actual_quarter < 4 else "높음",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--year", type=int, default=date.today().year)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = nowcast(args.workbook.resolve(), args.year)
    print(json.dumps(result, ensure_ascii=False, indent=None if args.json else 2))


if __name__ == "__main__":
    main()
