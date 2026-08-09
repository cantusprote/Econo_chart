#!/usr/bin/env python3
"""Forecast year-end KOSPI with three GDP/rate regression models.

The historical regression is fitted from data/2.GDP.xlsx:

    M1: ln(KOSPI) = alpha + beta*ln(GDP)
    M2: ln(KOSPI) = alpha + beta*ln(GDP) + gamma*ln(1+r)
    M3: ln(KOSPI) = alpha + beta1*ln(GDP) + beta2*ln(GDP)^2
                     + gamma*ln(1+r)

Only rows before the forecast year are used, so a future workbook row cannot
leak into the regression.  By default the two GDP inputs are collected from
estimate_nominal_gdp.py and estimate_current_gdp.py.
"""

from __future__ import annotations

import argparse
from datetime import date
import json
import math
from pathlib import Path
import statistics
from typing import Any

import numpy as np
from openpyxl import load_workbook

from estimate_current_gdp import nowcast
from estimate_nominal_gdp import estimate as estimate_nabo_gdp


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "data" / "2.GDP.xlsx"


class KospiForecastError(RuntimeError):
    """Raised when the regression or forecast inputs are invalid."""


def load_history(path: Path, forecast_year: int) -> list[tuple[int, float, float, float]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        if "Sheet1" not in workbook.sheetnames:
            raise KospiForecastError("2.GDP.xlsx에서 Sheet1을 찾지 못했습니다.")
        sheet = workbook["Sheet1"]
        rows: list[tuple[int, float, float, float]] = []
        for year, gdp, kospi, _, _, rate in sheet.iter_rows(
            min_row=2, min_col=1, max_col=6, values_only=True
        ):
            if not isinstance(year, int) or year >= forecast_year:
                continue
            if not all(isinstance(value, (int, float)) for value in (gdp, kospi, rate)):
                continue
            if gdp <= 0 or kospi <= 0 or rate <= -100:
                continue
            rows.append((year, float(gdp), float(kospi), float(rate)))
    finally:
        workbook.close()
    rows.sort(key=lambda row: row[0])
    if len(rows) < 3:
        raise KospiForecastError("로그 회귀분석에 사용할 과거 GDP·KOSPI 자료가 부족합니다.")
    return rows


def fit_log_regression(rows: list[tuple[int, float, float, float]]) -> dict[str, Any]:
    """Fit model 1 manually; retained as a public helper."""
    x = [math.log(gdp) for _, gdp, _, _ in rows]
    y = [math.log(kospi) for _, _, kospi, _ in rows]
    n = len(rows)
    x_mean = statistics.fmean(x)
    y_mean = statistics.fmean(y)
    sxx = sum((value - x_mean) ** 2 for value in x)
    if sxx == 0:
        raise KospiForecastError("GDP 값의 분산이 없어 회귀계수를 계산할 수 없습니다.")
    beta = sum((xi - x_mean) * (yi - y_mean) for xi, yi in zip(x, y)) / sxx
    alpha = y_mean - beta * x_mean
    fitted = [alpha + beta * xi for xi in x]
    residuals = [yi - fitted_i for yi, fitted_i in zip(y, fitted)]
    sse = sum(value**2 for value in residuals)
    ssr = sum((value - y_mean) ** 2 for value in fitted)
    sst = sum((value - y_mean) ** 2 for value in y)
    degrees_of_freedom = n - 2
    residual_standard_error = math.sqrt(sse / degrees_of_freedom)
    standard_error_beta = residual_standard_error / math.sqrt(sxx)
    standard_error_alpha = residual_standard_error * math.sqrt(1 / n + x_mean**2 / sxx)
    r_squared = 1 - sse / sst
    f_value = ssr / (sse / degrees_of_freedom)
    return {
        "alpha": alpha,
        "beta": beta,
        "standard_error_alpha": standard_error_alpha,
        "standard_error_beta": standard_error_beta,
        "r_squared": r_squared,
        "residual_standard_error": residual_standard_error,
        "f_value": f_value,
        "degrees_of_freedom": degrees_of_freedom,
        "regression_sum_squares": ssr,
        "residual_sum_squares": sse,
        "sample_size": n,
        "sample_start_year": rows[0][0],
        "sample_end_year": rows[-1][0],
    }


def fit_ols_models(rows: list[tuple[int, float, float, float]]) -> dict[str, dict[str, Any]]:
    """Fit all three models with least squares and conventional OLS diagnostics."""
    log_gdp = np.asarray([math.log(gdp) for _, gdp, _, _ in rows], dtype=float)
    log_rate = np.asarray([math.log1p(rate / 100.0) for _, _, _, rate in rows], dtype=float)
    y = np.asarray([math.log(kospi) for _, _, kospi, _ in rows], dtype=float)
    designs = {
        "gdp_only": (np.column_stack([np.ones(len(rows)), log_gdp]), ["alpha", "beta"]),
        "gdp_and_rate": (
            np.column_stack([np.ones(len(rows)), log_gdp, log_rate]),
            ["alpha", "beta", "gamma"],
        ),
        "gdp_squared_and_rate": (
            np.column_stack([np.ones(len(rows)), log_gdp, log_gdp**2, log_rate]),
            ["alpha", "beta1", "beta2", "gamma"],
        ),
    }
    results: dict[str, dict[str, Any]] = {}
    y_mean = float(np.mean(y))
    sst = float(np.sum((y - y_mean) ** 2))
    for model_name, (design, names) in designs.items():
        coefficients, _, rank, _ = np.linalg.lstsq(design, y, rcond=None)
        if rank != design.shape[1]:
            raise KospiForecastError(f"{model_name} 회귀행렬의 순위가 부족합니다.")
        fitted = design @ coefficients
        residuals = y - fitted
        sse = float(residuals @ residuals)
        ssr = float(np.sum((fitted - y_mean) ** 2))
        df_residual = len(rows) - design.shape[1]
        df_model = design.shape[1] - 1
        sigma2 = sse / df_residual
        covariance = sigma2 * np.linalg.inv(design.T @ design)
        standard_errors = np.sqrt(np.diag(covariance))
        result: dict[str, Any] = {
            name: float(value) for name, value in zip(names, coefficients)
        }
        result["standard_errors"] = {
            name: float(value) for name, value in zip(names, standard_errors)
        }
        result.update(
            {
                "r_squared": 1 - sse / sst,
                "adjusted_r_squared": 1 - (sse / df_residual) / (sst / (len(rows) - 1)),
                "residual_standard_error": math.sqrt(sigma2),
                "f_value": (ssr / df_model) / (sse / df_residual),
                "degrees_of_freedom": df_residual,
                "regression_sum_squares": ssr,
                "residual_sum_squares": sse,
                "sample_size": len(rows),
                "sample_start_year": rows[0][0],
                "sample_end_year": rows[-1][0],
                "condition_number": float(np.linalg.cond(design)),
            }
        )
        results[model_name] = result
    return results


def forecast_kospi(alpha: float, beta: float, nominal_gdp: float) -> dict[str, float | str]:
    if nominal_gdp <= 0:
        raise KospiForecastError("명목 GDP 예측치는 0보다 커야 합니다.")
    log_gdp = math.log(nominal_gdp)
    log_kospi = alpha + beta * log_gdp
    kospi = math.exp(log_kospi)
    return {
        "nominal_gdp_trillion_krw": round(nominal_gdp, 4),
        "log_gdp": round(log_gdp, 9),
        "log_kospi": round(log_kospi, 9),
        "forecast_year_end_kospi": round(kospi, 2),
        "calculation": (
            f"exp({alpha:.9f} + {beta:.9f} × ln({nominal_gdp:.4f})) = {kospi:.2f}"
        ),
    }


def forecast_all_models(
    models: dict[str, dict[str, Any]], nominal_gdp: float, rate_pct: float
) -> dict[str, dict[str, float | str]]:
    if nominal_gdp <= 0 or rate_pct <= -100:
        raise KospiForecastError("GDP는 양수이고 금리는 -100%보다 커야 합니다.")
    x = math.log(nominal_gdp)
    z = math.log1p(rate_pct / 100.0)
    outputs: dict[str, dict[str, float | str]] = {}
    expressions = {
        "gdp_only": lambda m: m["alpha"] + m["beta"] * x,
        "gdp_and_rate": lambda m: m["alpha"] + m["beta"] * x + m["gamma"] * z,
        "gdp_squared_and_rate": lambda m: (
            m["alpha"] + m["beta1"] * x + m["beta2"] * x**2 + m["gamma"] * z
        ),
    }
    for name, expression in expressions.items():
        log_kospi = float(expression(models[name]))
        outputs[name] = {
            "forecast_year_end_kospi": round(math.exp(log_kospi), 2),
            "log_kospi": round(log_kospi, 9),
        }
    return outputs


def predict(
    path: Path = DEFAULT_WORKBOOK,
    forecast_year: int | None = None,
    nabo_gdp: float | None = None,
    current_gdp: float | None = None,
    rate_pct: float | None = None,
) -> dict[str, Any]:
    year = forecast_year or date.today().year
    rows = load_history(path, year)
    models = fit_ols_models(rows)
    latest_rate_year, latest_rate = rows[-1][0], rows[-1][3]
    forecast_rate = latest_rate if rate_pct is None else rate_pct
    rate_method = (
        f"{latest_rate_year}년 12월 최신 확정값을 예측연도 대용값으로 사용"
        if rate_pct is None
        else "명령행 --rate 입력값"
    )

    if nabo_gdp is None:
        nabo_result = estimate_nabo_gdp(path, year)
        nabo_gdp = float(nabo_result["forecast_nominal_gdp_trillion_krw"])
        nabo_metadata: dict[str, Any] = {
            "method": nabo_result["classification"],
            "growth_forecast_pct": nabo_result["nominal_growth_forecast_pct"],
            "source_title": nabo_result["report_title"],
            "source_published": nabo_result["report_published"],
            "source_url": nabo_result["report_url"],
        }
    else:
        nabo_metadata = {"method": "명령행에서 입력한 검증용 GDP 값"}

    if current_gdp is None:
        current_result = nowcast(path, year)
        current_gdp = float(current_result["forecast_nominal_gdp_trillion_krw"])
        current_metadata: dict[str, Any] = {
            "method": current_result["method"],
            "actual_quarters": current_result["actual_quarters"],
            "estimated_quarters": current_result["estimated_quarters"],
            "confidence": current_result["confidence"],
            "forecast_range_low_trillion_krw": current_result["forecast_range_low_trillion_krw"],
            "forecast_range_high_trillion_krw": current_result["forecast_range_high_trillion_krw"],
            "sources": current_result["sources"],
        }
    else:
        current_metadata = {"method": "명령행에서 입력한 검증용 GDP 값"}

    regression = models["gdp_only"]
    alpha = float(regression["alpha"])
    beta = float(regression["beta"])
    nabo_forecast = forecast_kospi(alpha, beta, nabo_gdp)
    current_forecast = forecast_kospi(alpha, beta, current_gdp)
    if "forecast_range_low_trillion_krw" in current_metadata:
        current_forecast["forecast_year_end_kospi_low"] = forecast_kospi(
            alpha, beta, float(current_metadata["forecast_range_low_trillion_krw"])
        )["forecast_year_end_kospi"]
        current_forecast["forecast_year_end_kospi_high"] = forecast_kospi(
            alpha, beta, float(current_metadata["forecast_range_high_trillion_krw"])
        )["forecast_year_end_kospi"]
    nabo_forecast.update(nabo_metadata)
    current_forecast.update(current_metadata)
    nabo_forecast["model_forecasts"] = forecast_all_models(models, nabo_gdp, forecast_rate)
    current_forecast["model_forecasts"] = forecast_all_models(models, current_gdp, forecast_rate)

    return {
        "forecast_year": year,
        "calculated_at": date.today().isoformat(),
        "workbook": str(path),
        "models": {
            "gdp_only": {
                "formula": "ln(KOSPI) = alpha + beta*ln(GDP)",
                "fitted_equation": (
                    f"ln(KOSPI) = {models['gdp_only']['alpha']:.9f} "
                    f"+ {models['gdp_only']['beta']:.9f}*ln(GDP)"
                ),
                "regression": models["gdp_only"],
            },
            "gdp_and_rate": {
                "formula": "ln(KOSPI) = alpha + beta*ln(GDP) + gamma*ln(1+r)",
                "fitted_equation": (
                    f"ln(KOSPI) = {models['gdp_and_rate']['alpha']:.9f} "
                    f"+ {models['gdp_and_rate']['beta']:.9f}*ln(GDP) "
                    f"+ {models['gdp_and_rate']['gamma']:.9f}*ln(1+r)"
                ),
                "regression": models["gdp_and_rate"],
            },
            "gdp_squared_and_rate": {
                "formula": "ln(KOSPI) = alpha + beta1*ln(GDP) + beta2*ln(GDP)^2 + gamma*ln(1+r)",
                "fitted_equation": (
                    f"ln(KOSPI) = {models['gdp_squared_and_rate']['alpha']:.9f} "
                    f"+ {models['gdp_squared_and_rate']['beta1']:.9f}*ln(GDP) "
                    f"+ {models['gdp_squared_and_rate']['beta2']:.9f}*ln(GDP)^2 "
                    f"+ {models['gdp_squared_and_rate']['gamma']:.9f}*ln(1+r)"
                ),
                "regression": models["gdp_squared_and_rate"],
            },
        },
        "dependent_variable": "KOSPI 연말 종가",
        "independent_variables": "명목 GDP(조원), 한국 10년 국채금리(12월, %)",
        "rate_assumption": {
            "rate_pct": forecast_rate,
            "method": rate_method,
        },
        "scenarios": {
            "nominal_gdp_forecast": nabo_forecast,
            "estimate_current_gdp": current_forecast,
        },
        "interpretation_note": (
            f"beta={beta:.4f}는 GDP가 1% 변할 때 KOSPI가 평균적으로 약 {beta:.4f}% "
            "변하는 로그 탄력성입니다. "
            "금리 계수의 양수 추정은 인과관계를 의미하지 않으며, 제곱항 모형은 높은 조건수로 "
            "다중공선성과 표본외 불안정성에 특히 주의해야 합니다. 예측치는 실제 연말 지수를 보장하지 않습니다."
        ),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--year", type=int, default=date.today().year)
    parser.add_argument("--nabo-gdp", type=float, help="수집 대신 사용할 명목GDP 예측치(조원)")
    parser.add_argument("--current-gdp", type=float, help="수집 대신 사용할 현재연도 GDP 추정치(조원)")
    parser.add_argument("--rate", type=float, help="예측연도 12월 한국 10년 국채금리 가정(%)")
    parser.add_argument("--json", action="store_true", help="한 줄 JSON으로 출력")
    args = parser.parse_args()
    result = predict(
        args.workbook.resolve(),
        args.year,
        nabo_gdp=args.nabo_gdp,
        current_gdp=args.current_gdp,
        rate_pct=args.rate,
    )
    print(json.dumps(result, ensure_ascii=False, indent=None if args.json else 2))


if __name__ == "__main__":
    main()
