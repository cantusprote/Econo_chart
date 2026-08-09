#!/usr/bin/env python3
"""Desktop control panel for the monthly Econo Chart updater."""

from __future__ import annotations

import argparse
import json
import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from math import ceil, floor
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


ROOT = Path(__file__).resolve().parent
GENERAL_UPDATE_SCRIPT = ROOT / "update_to_current_month.py"
ECONOMIC_STATE_UPDATE_SCRIPT = ROOT / "update_economic_state.py"
GDP_UPDATE_SCRIPT = ROOT / "update_gdp.py"
GDP_FORECAST_SCRIPT = ROOT / "estimate_nominal_gdp.py"
CURRENT_GDP_ESTIMATE_SCRIPT = ROOT / "estimate_current_gdp.py"
KOSPI_GDP_MODEL_SCRIPT = ROOT / "predict_kospi_from_gdp.py"
HP_EXPORT_KOSPI_SCRIPT = ROOT / "hp_export_kospi.py"
GENERAL_WORKBOOK = ROOT / "data" / "0_general.xlsx"
ECONOMIC_STATE_WORKBOOK = ROOT / "data" / "1.economic_state.xlsx"
GDP_WORKBOOK = ROOT / "data" / "2.GDP.xlsx"

COLORS = {
    "window": "#16182B",
    "toolbar": "#20243A",
    "panel": "#0D0F13",
    "panel_alt": "#12151B",
    "border": "#064C7B",
    "accent": "#00D7B0",
    "accent_hover": "#00BFA0",
    "blue": "#0877D1",
    "blue_hover": "#0868B5",
    "text": "#F0F3F6",
    "muted": "#9AA4AF",
    "success": "#16D98A",
    "warning": "#FFC857",
    "danger": "#FF5F57",
}

GENERAL_SOURCES = [
    ("KOSPI 월말", "네이버 증권", "KRX 기반 일별 시세의 해당 월 마지막 거래일 종가", "https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI"),
    ("선행지수 순환변동치", "국가데이터처", "월별 산업활동동향 공식 PDF", "https://www.mods.go.kr/board.es?act=list&bid=216&mid=a10301050100"),
    ("OECD 한국선행지수", "OECD", "SDMX Composite Leading Indicators API", "https://data-explorer.oecd.org/"),
    ("일평균 수출", "산업통상자원부", "NLIC에 게시된 월별 수출입 동향", "https://www.nlic.go.kr/nlic/logpolLi.action"),
    ("기준금리", "한국은행 ECOS", "722Y001 / 0101000", "https://ecos.bok.or.kr/"),
    ("국고채 3년 월평균", "한국은행 ECOS", "817Y002 / 010200000 일별 평균", "https://ecos.bok.or.kr/"),
    ("원/달러 월평균", "한국은행 ECOS", "731Y004 / 평균자료", "https://ecos.bok.or.kr/"),
]

ECONOMIC_STATE_SOURCES = [
    GENERAL_SOURCES[0],
    GENERAL_SOURCES[1],
    GENERAL_SOURCES[2],
    ("제조업 재고율", "국가데이터처", "산업활동동향의 재고/출하 비율", "https://www.mods.go.kr/board.es?act=list&bid=216&mid=a10301050100"),
]

GDP_SOURCES = [
    ("명목 GDP(조원)", "국가지표체계·한국은행", "국민소득 명목 GDP(십억원)를 조원으로 변환", "https://www.index.go.kr/unity/potal/eNara/sub/showStblGams3.do?freq=Y&idx_cd=F0131&period=N&stts_cd=F013101"),
    ("KOSPI 연말", "네이버 증권", "해당 연도 마지막 거래일 KOSPI 종가", "https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI"),
    ("국고채 10년", "한국은행 ECOS", "817Y002 / 010210000, 12월 일별 금리 평균", "https://ecos.bok.or.kr/"),
    ("적정 KOSPI·평가율", "GDP 모델", "기존 ln(GDP)-ln(KOSPI) 모형과 통합문서 수식", "https://www.index.go.kr/unity/potal/eNara/sub/showStblGams3.do?freq=Y&idx_cd=F0131&period=N&stts_cd=F013101"),
]


class ColorButton(tk.Label):
    """Flat label-backed button whose colors are stable on macOS Tk/Aqua."""

    def __init__(
        self,
        parent: tk.Misc,
        *,
        text: str,
        command,
        background: str,
        foreground: str,
        active_background: str,
    ) -> None:
        super().__init__(
            parent,
            text=text,
            bg=background,
            fg=foreground,
            disabledforeground="#68717A",
            padx=18,
            pady=10,
            cursor="pointinghand",
            font=("Helvetica Neue", 13, "bold"),
            relief="flat",
            bd=0,
        )
        self.command = command
        self.normal_background = background
        self.active_background = active_background
        self.bind("<Button-1>", self._invoke)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    def _invoke(self, _event=None) -> None:
        if str(self.cget("state")) != "disabled":
            self.command()

    def _enter(self, _event=None) -> None:
        if str(self.cget("state")) != "disabled":
            super().configure(bg=self.active_background)

    def _leave(self, _event=None) -> None:
        super().configure(bg=self.normal_background)


class EconoUpdaterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Econo Chart Updater v1.0.0")
        self.geometry("1460x820")
        self.minsize(1120, 650)
        self.configure(bg=COLORS["window"])
        self.process: subprocess.Popen[str] | None = None
        self.pending_data = False
        self.progress_total = 0
        self.progress_completed = 0
        self.progress_unit = "개월"
        self.active_script = GENERAL_UPDATE_SCRIPT
        self.active_workbook = GENERAL_WORKBOOK
        self.active_update_name = "update_general"
        self.current_sources = GENERAL_SOURCES
        self.chart_range_months: int | None = None
        self.gdp_forecast_running = False
        self.events: queue.Queue[tuple[str, str | int]] = queue.Queue()
        self._configure_styles()
        self._build_ui()
        self.after(100, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Source.Treeview",
            background=COLORS["panel_alt"],
            foreground=COLORS["text"],
            fieldbackground=COLORS["panel_alt"],
            bordercolor=COLORS["border"],
            rowheight=40,
            font=("Apple SD Gothic Neo", 12),
        )
        style.configure(
            "Source.Treeview.Heading",
            background="#17334A",
            foreground=COLORS["accent"],
            relief="flat",
            font=("Apple SD Gothic Neo", 12, "bold"),
        )
        style.configure(
            "Update.Horizontal.TProgressbar",
            troughcolor="#0A0D12",
            background=COLORS["accent"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
            thickness=16,
        )
        style.map("Source.Treeview", background=[("selected", "#14567C")])
        style.configure(
            "Vertical.TScrollbar",
            background="#17334A",
            troughcolor=COLORS["panel"],
            bordercolor=COLORS["panel"],
            arrowcolor=COLORS["muted"],
        )

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg="#34393E", height=44)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="Econo Chart Updater v1.0.0",
            bg="#34393E",
            fg="#C8CDD2",
            font=("Helvetica Neue", 15, "bold"),
        ).pack(expand=True)

        self.tab_bar = tk.Frame(self, bg="#111827", padx=16, pady=8)
        self.tab_bar.pack(fill="x")
        self.update_tab_button = self._button(
            self.tab_bar,
            "Update",
            lambda: self._show_tab("update"),
            COLORS["accent"],
            "#071B18",
            COLORS["accent_hover"],
        )
        self.update_tab_button.pack(side="left", padx=(0, 6))
        self.graph_tab_button = self._button(
            self.tab_bar,
            "Graph",
            lambda: self._show_tab("graph"),
            "#17334A",
            COLORS["text"],
            "#214B69",
        )
        self.graph_tab_button.pack(side="left")

        self.toolbar_host = tk.Frame(self, bg=COLORS["toolbar"])
        self.toolbar_host.pack(fill="x")
        toolbar = tk.Frame(self.toolbar_host, bg=COLORS["toolbar"], padx=16, pady=14)
        self.update_toolbar = toolbar
        self.update_toolbar.pack(fill="x")
        self.update_button = self._button(
            toolbar,
            "▶  update_general",
            lambda: self.run_update(
                GENERAL_UPDATE_SCRIPT, GENERAL_WORKBOOK, "update_general"
            ),
            COLORS["accent"],
            "#071B18",
            COLORS["accent_hover"],
        )
        self.update_button.pack(side="left", padx=(0, 10))
        self.economic_state_button = self._button(
            toolbar,
            "▶  update_economic_state",
            lambda: self.run_update(
                ECONOMIC_STATE_UPDATE_SCRIPT,
                ECONOMIC_STATE_WORKBOOK,
                "update_economic_state",
            ),
            COLORS["blue"],
            COLORS["text"],
            COLORS["blue_hover"],
        )
        self.economic_state_button.pack(side="left", padx=(0, 10))
        self.gdp_button = self._button(
            toolbar,
            "▶  update_gdp",
            lambda: self.run_update(GDP_UPDATE_SCRIPT, GDP_WORKBOOK, "update_gdp"),
            "#7B4BB7",
            COLORS["text"],
            "#9060CC",
        )
        self.gdp_button.pack(side="left", padx=(0, 10))
        self.stop_button = self._button(
            toolbar, "■  Stop", self.stop_update, "#07507D", COLORS["muted"], "#0A608F"
        )
        self.stop_button.configure(state="disabled")
        self.stop_button.pack(side="left", padx=(0, 10))
        self.clear_button = self._button(
            toolbar, "▣  Clear", self.clear_log, "#1D2940", COLORS["text"], "#263754"
        )
        self.clear_button.pack(side="left")

        self.status_label = tk.Label(
            toolbar,
            text="● READY",
            bg=COLORS["toolbar"],
            fg=COLORS["success"],
            font=("Menlo", 12, "bold"),
        )
        self.status_label.pack(side="right", padx=8)
        self.workbook_label = tk.Label(
            toolbar,
            text=str(GENERAL_WORKBOOK),
            bg=COLORS["toolbar"],
            fg=COLORS["muted"],
            font=("Menlo", 10),
        )
        self.workbook_label.pack(side="right", padx=16)

        self.graph_toolbar = tk.Frame(
            self.toolbar_host, bg=COLORS["toolbar"], padx=16, pady=14
        )
        self.graph_buttons: list[ColorButton] = []
        for number in range(1, 7):
            if number == 1:
                graph_command = self.plot_graph_1
            elif number == 2:
                graph_command = self.plot_graph_2
            elif number == 3:
                graph_command = self.plot_graph_3
            elif number == 4:
                graph_command = self.plot_graph_4
            elif number == 5:
                graph_command = self.plot_graph_5
            else:
                graph_command = self.plot_graph_6
            button = self._button(
                self.graph_toolbar,
                f"{number}번 버튼",
                graph_command,
                COLORS["blue"],
                COLORS["text"],
                COLORS["blue_hover"],
            )
            button.pack(side="left", padx=(0, 10))
            self.graph_buttons.append(button)

        progress_frame = tk.Frame(self, bg=COLORS["toolbar"], padx=16)
        self.progress_frame = progress_frame
        self.progress_frame.pack(fill="x", pady=(0, 12))
        self.progress_label = tk.Label(
            progress_frame,
            text="진행 대기",
            bg=COLORS["toolbar"],
            fg=COLORS["muted"],
            font=("Apple SD Gothic Neo", 11, "bold"),
            width=25,
            anchor="w",
        )
        self.progress_label.pack(side="left", padx=(0, 12))
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            style="Update.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress_bar.pack(side="left", fill="x", expand=True)
        self.progress_percent = tk.Label(
            progress_frame,
            text="0%",
            bg=COLORS["toolbar"],
            fg=COLORS["text"],
            font=("Menlo", 11, "bold"),
            width=6,
            anchor="e",
        )
        self.progress_percent.pack(side="right", padx=(12, 0))

        content = tk.PanedWindow(
            self,
            orient="horizontal",
            bg=COLORS["window"],
            sashwidth=10,
            sashrelief="flat",
            bd=0,
        )
        self.content = content
        content.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.graph_content = tk.Frame(
            self,
            bg=COLORS["panel"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        graph_header = tk.Frame(self.graph_content, bg=COLORS["panel"], padx=16, pady=12)
        graph_header.pack(fill="x")
        self.graph_title_label = tk.Label(
            graph_header,
            text="Graph 1 · KOSPI와 선행지수 순환변동치",
            bg=COLORS["panel"],
            fg=COLORS["accent"],
            font=("Apple SD Gothic Neo", 18, "bold"),
            anchor="w",
        )
        self.graph_title_label.pack(fill="x")
        self.graph_subtitle_label = tk.Label(
            graph_header,
            text="1번 버튼을 누르면 data/0_general.xlsx의 최신 데이터로 다시 그립니다.",
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=("Apple SD Gothic Neo", 11),
            anchor="w",
        )
        self.graph_subtitle_label.pack(fill="x", pady=(3, 0))

        range_frame = tk.Frame(self.graph_content, bg=COLORS["panel"], padx=16)
        range_frame.pack(fill="x", pady=(0, 10))
        tk.Label(
            range_frame,
            text="표시 기간",
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=("Apple SD Gothic Neo", 11, "bold"),
        ).pack(side="left", padx=(0, 10))
        self.range_buttons: dict[int | None, ColorButton] = {}
        for label, months in (("전체", None), ("최근 10년", 120), ("최근 5년", 60)):
            button = self._button(
                range_frame,
                label,
                lambda selected=months: self._set_chart_range(selected),
                COLORS["blue"],
                COLORS["text"],
                COLORS["blue_hover"],
            )
            button.pack(side="left", padx=(0, 8))
            self.range_buttons[months] = button
        self.gdp_forecast_button = self._button(
            range_frame,
            "명목GDP 예측치",
            self.show_nominal_gdp_forecast,
            "#7B4BB7",
            COLORS["text"],
            "#9060CC",
        )
        self.current_gdp_button = self._button(
            range_frame,
            "estimate_current_gdp",
            self.show_current_gdp_estimate,
            "#B45309",
            COLORS["text"],
            "#D97706",
        )
        self.kospi_model_button = self._button(
            range_frame,
            "KOSPI 예측모형",
            self.show_kospi_model_forecast,
            "#BE123C",
            COLORS["text"],
            "#E11D48",
        )
        self.hp_export_button = self._button(
            range_frame,
            "HP 수출평가",
            self.plot_graph_3_hp,
            "#047857",
            COLORS["text"],
            "#059669",
        )
        self._update_range_button_styles()

        self.chart_container = tk.Frame(self.graph_content, bg="#FFFFFF")
        self.chart_container.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.chart_canvas = None
        self.chart_figure = None
        self.chart_points = 0
        self.active_graph_number = 0
        self.active_graph_variant = "default"

        left = self._panel(content)
        right = self._panel(content)
        content.add(left, minsize=470, stretch="always", width=700)
        content.add(right, minsize=520, stretch="always", width=720)

        self._section_title(left, "실행 로그", "업데이트 진행 상황과 오류 메시지")
        log_frame = tk.Frame(left, bg=COLORS["panel"])
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical")
        log_scroll.pack(side="right", fill="y")
        self.log = tk.Text(
            log_frame,
            bg="#080A0D",
            fg="#D9E1E8",
            insertbackground=COLORS["accent"],
            selectbackground="#14567C",
            relief="flat",
            padx=14,
            pady=14,
            font=("Menlo", 12),
            wrap="word",
            yscrollcommand=log_scroll.set,
        )
        self.log.pack(fill="both", expand=True)
        log_scroll.configure(command=self.log.yview)
        self.log.tag_configure("time", foreground="#6D8494")
        self.log.tag_configure("info", foreground="#D9E1E8")
        self.log.tag_configure("success", foreground=COLORS["success"])
        self.log.tag_configure("warning", foreground=COLORS["warning"])
        self.log.tag_configure("error", foreground="#FF6B6B")
        self._append_log("update_general 버튼을 누르면 직전 완료 월까지 갱신합니다.", "info")
        self._append_log("결과는 data/0_general.xlsx 한 파일에 계속 저장됩니다.", "info")
        self._append_log("update_economic_state 버튼은 data/1.economic_state.xlsx를 갱신합니다.", "info")
        self._append_log("update_gdp 버튼은 직전 완료 연도까지 data/2.GDP.xlsx를 갱신합니다.", "info")

        self.source_title_label, self.source_subtitle_label = self._section_title(
            right,
            "자료 출처 · update_general",
            "data/0_general.xlsx의 7개 지표 · 행을 더블클릭하면 제공처를 엽니다",
        )
        table_frame = tk.Frame(right, bg=COLORS["panel"])
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        columns = ("indicator", "provider", "method")
        self.source_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            style="Source.Treeview",
            selectmode="browse",
        )
        self.source_table.heading("indicator", text="지표")
        self.source_table.heading("provider", text="제공처")
        self.source_table.heading("method", text="수집 방식 / 통계코드")
        self.source_table.column("indicator", width=175, minwidth=145, stretch=False)
        self.source_table.column("provider", width=145, minwidth=120, stretch=False)
        self.source_table.column("method", width=330, minwidth=220, stretch=True)
        source_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.source_table.yview)
        self.source_table.configure(yscrollcommand=source_scroll.set)
        source_scroll.pack(side="right", fill="y")
        self.source_table.pack(fill="both", expand=True)
        self.source_table.bind("<Double-1>", self._open_selected_source)

        self.source_note = tk.Label(
            right,
            text="※ 월중 실행 시 당월 확정 자료가 없어 기본 종료월은 직전 달입니다.",
            bg=COLORS["panel"],
            fg=COLORS["warning"],
            anchor="w",
            padx=15,
            pady=12,
            font=("Apple SD Gothic Neo", 11),
        )
        self.source_note.pack(fill="x", side="bottom")
        self._show_sources("update_general")

    def _show_tab(self, tab_name: str) -> None:
        if tab_name == "graph":
            self.update_toolbar.pack_forget()
            self.graph_toolbar.pack(fill="x")
            self.progress_frame.pack_forget()
            self.content.pack_forget()
            self.graph_content.pack(fill="both", expand=True, padx=16, pady=(0, 16))
            self.update_tab_button.normal_background = "#17334A"
            self.update_tab_button.active_background = "#214B69"
            self.update_tab_button.configure(bg="#17334A", fg=COLORS["text"])
            self.graph_tab_button.normal_background = COLORS["accent"]
            self.graph_tab_button.active_background = COLORS["accent_hover"]
            self.graph_tab_button.configure(bg=COLORS["accent"], fg="#071B18")
        else:
            self.graph_toolbar.pack_forget()
            self.update_toolbar.pack(fill="x")
            self.graph_content.pack_forget()
            self.content.pack(fill="both", expand=True, padx=16, pady=(0, 16))
            self.progress_frame.pack(
                fill="x", pady=(0, 12), before=self.content
            )
            self.update_tab_button.normal_background = COLORS["accent"]
            self.update_tab_button.active_background = COLORS["accent_hover"]
            self.update_tab_button.configure(bg=COLORS["accent"], fg="#071B18")
            self.graph_tab_button.normal_background = "#17334A"
            self.graph_tab_button.active_background = "#214B69"
            self.graph_tab_button.configure(bg="#17334A", fg=COLORS["text"])

    def _graph_placeholder(self, number: int) -> None:
        self._append_log(f"Graph {number}번 버튼: 그래프 기능 준비 중입니다.", "info")

    def _set_chart_range(self, months: int | None) -> None:
        self.chart_range_months = months
        self._update_range_button_styles()
        self._refresh_active_graph()

    def _update_range_button_styles(self) -> None:
        for months, button in self.range_buttons.items():
            selected = months == self.chart_range_months
            normal = COLORS["accent"] if selected else COLORS["blue"]
            active = COLORS["accent_hover"] if selected else COLORS["blue_hover"]
            foreground = "#071B18" if selected else COLORS["text"]
            button.normal_background = normal
            button.active_background = active
            button.configure(bg=normal, fg=foreground)

    def _refresh_active_graph(self) -> None:
        graph_number = self.active_graph_number or 1
        if graph_number == 3 and self.active_graph_variant == "hp_export":
            self.plot_graph_3_hp()
            return
        graph_method = getattr(self, f"plot_graph_{graph_number}")
        graph_method()

    def _show_graph3_submenu(self, visible: bool) -> None:
        if visible:
            if not self.hp_export_button.winfo_manager():
                self.hp_export_button.pack(side="left", padx=(12, 0))
        else:
            self.hp_export_button.pack_forget()

    def _show_graph6_submenu(self, visible: bool) -> None:
        if visible:
            if not self.gdp_forecast_button.winfo_manager():
                self.gdp_forecast_button.pack(side="left", padx=(12, 0))
                self.current_gdp_button.pack(side="left", padx=(8, 0))
                self.kospi_model_button.pack(side="left", padx=(8, 0))
        else:
            self.gdp_forecast_button.pack_forget()
            self.current_gdp_button.pack_forget()
            self.kospi_model_button.pack_forget()

    def _clear_chart_display(self) -> None:
        for child in self.chart_container.winfo_children():
            child.destroy()
        self.chart_canvas = None
        self.chart_figure = None

    @staticmethod
    def _filter_chart_period(
        dates: list[datetime],
        kospi: list[float],
        comparison: list[float],
        months: int | None,
        *,
        execution_date: datetime | None = None,
    ) -> tuple[list[datetime], list[float], list[float]]:
        if months is None:
            return dates, kospi, comparison

        current = execution_date or datetime.now()
        end_month_index = current.year * 12 + current.month - 1
        start_month_index = end_month_index - months + 1
        start = datetime(start_month_index // 12, start_month_index % 12 + 1, 1)
        end = datetime(current.year, current.month, 1)
        selected = [
            index for index, month in enumerate(dates) if start <= month <= end
        ]
        return (
            [dates[index] for index in selected],
            [kospi[index] for index in selected],
            [comparison[index] for index in selected],
        )

    def plot_graph_1(self) -> None:
        self._plot_general_dual_axis(
            graph_number=1,
            right_column=3,
            right_label="선행지수 순환변동치",
            chart_title="KOSPI와 선행지수 순환변동치",
            legend_label="선행지수순환변동치(우)",
        )

    def plot_graph_2(self) -> None:
        self._plot_general_dual_axis(
            graph_number=2,
            right_column=4,
            right_label="OECD 한국 선행지수",
            chart_title="KOSPI와 OECD 한국 선행지수",
            legend_label="OECD 한국선행지수(우)",
        )

    def plot_graph_3(self) -> None:
        self.active_graph_variant = "default"
        self._plot_general_dual_axis(
            graph_number=3,
            right_column=5,
            right_label="일평균 수출액(억 달러)",
            chart_title="KOSPI와 일평균 수출액",
            legend_label="일평균 수출(억 달러, 우)",
        )

    def plot_graph_3_hp(self) -> None:
        self._show_graph6_submenu(False)
        self._show_graph3_submenu(True)
        try:
            import matplotlib as mpl
            from matplotlib import dates as mdates
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure

            from hp_export_kospi import calculate

            mpl.rcParams["font.family"] = "Apple SD Gothic Neo"
            mpl.rcParams["axes.unicode_minus"] = False
            result = calculate(GENERAL_WORKBOOK)
            dates = [datetime.strptime(row["date"], "%Y-%m-%d") for row in result["observations"]]
            gaps = [float(row["valuation_gap_pct"]) for row in result["observations"]]
            dates, gaps, _ = self._filter_chart_period(
                dates, gaps, gaps, self.chart_range_months
            )
            if not dates:
                raise ValueError(f"{self._chart_range_label()} 범위에 발표된 데이터가 없습니다.")

            self._clear_chart_display()
            figure = Figure(figsize=(12, 6.4), dpi=100, facecolor="#FFFFFF")
            axis = figure.add_subplot(111)
            green = "#16A765"
            axis.plot(dates, gaps, color=green, linewidth=2.5, label="수출 대비 KOSPI 평가갭")
            axis.axhline(0, color="#374151", linewidth=1.1)
            axis.fill_between(dates, gaps, 0, where=[value >= 0 for value in gaps], color="#DCFCE7", alpha=0.55)
            axis.fill_between(dates, gaps, 0, where=[value < 0 for value in gaps], color="#DBEAFE", alpha=0.45)
            low, high = min(gaps), max(gaps)
            padding = max(5.0, (high - low) * 0.08)
            axis.set_ylim(low - padding, high + padding)
            axis.set_ylabel("수출 대비 KOSPI 평가갭(%)", color="#4B5563", fontsize=10)
            axis.grid(axis="y", color="#D1D5DB", linewidth=0.8)
            axis.grid(axis="x", visible=False)
            axis.spines["top"].set_visible(False)
            axis.spines["right"].set_visible(False)
            axis.spines["bottom"].set_color("#CBD5E1")
            axis.spines["left"].set_color("#CBD5E1")
            axis.tick_params(colors="#4B5563", labelsize=9)
            tick_interval = max(1, round(len(dates) / 18))
            axis.xaxis.set_major_locator(mdates.MonthLocator(interval=tick_interval))
            axis.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
            axis.tick_params(axis="x", rotation=60)
            axis.set_xlim(dates[0], dates[-1])
            axis.annotate(
                f"최근 {gaps[-1]:+.1f}%",
                xy=(dates[-1], gaps[-1]),
                xytext=(-12, 12),
                textcoords="offset points",
                ha="right",
                color=green,
                fontsize=10,
                fontweight="bold",
            )
            figure.suptitle(
                "HP 필터 기반 수출 대비 KOSPI 평가갭",
                fontsize=15,
                fontweight="bold",
                color="#1F2937",
                y=0.98,
            )
            axis.set_title(
                "0% 초과: 수출 대비 고평가 · 0% 미만: 수출 대비 저평가",
                fontsize=10,
                color="#6B7280",
                pad=12,
            )
            axis.text(
                0.015,
                0.975,
                "Rₜ = ln(KOSPIₜ / 일평균수출ₜ)\n"
                "평가갭ₜ = 100 × {exp(Rₜ − HP추세ₜ) − 1}",
                transform=axis.transAxes,
                ha="left",
                va="top",
                fontsize=10,
                color="#374151",
                bbox={
                    "boxstyle": "round,pad=0.45",
                    "facecolor": "#FFFFFF",
                    "edgecolor": "#CBD5E1",
                    "alpha": 0.88,
                },
            )
            axis.legend(loc="lower center", bbox_to_anchor=(0.5, -0.24), frameon=False)
            figure.subplots_adjust(left=0.09, right=0.97, top=0.86, bottom=0.23)

            self.chart_figure = figure
            self.chart_canvas = FigureCanvasTkAgg(figure, master=self.chart_container)
            self.chart_canvas.draw()
            self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
            self.chart_points = len(dates)
            self.active_graph_number = 3
            self.active_graph_variant = "hp_export"
            self.graph_title_label.configure(text="Graph 3 · HP 필터 수출 대비 KOSPI 평가갭")
            self.graph_subtitle_label.configure(
                text=(
                    f"data/0_general.xlsx · {dates[0]:%Y-%m}~{dates[-1]:%Y-%m} · "
                    f"{len(dates)}개월 · {self._chart_range_label()} · 월별 λ={result['hp_lambda']:,.0f} · 양방향 필터"
                )
            )
            self._append_log(
                f"Graph 3 HP 수출평가 생성 완료: {len(dates)}개월 · 최근 {gaps[-1]:+.1f}%",
                "success",
            )
        except Exception as exc:
            self._append_log(f"Graph 3 HP 수출평가 생성 실패: {exc}", "error")
            messagebox.showerror("HP 수출평가 그래프 실패", str(exc))

    def plot_graph_4(self) -> None:
        self._plot_general_dual_axis(
            graph_number=4,
            right_column=7,
            right_label="국고채 3년 월평균 금리(%)",
            chart_title="KOSPI와 국고채 3년 금리",
            legend_label="국고채3년(월평균, %, 우)",
            right_axis_zero=True,
        )

    def plot_graph_5(self) -> None:
        self._plot_general_dual_axis(
            graph_number=5,
            right_column=8,
            right_label="원·달러 월평균 환율",
            chart_title="KOSPI와 원·달러 환율",
            legend_label="원-달러 환율(월평균,우)",
            right_axis_step=100,
        )

    @staticmethod
    def _filter_annual_chart_period(
        years: list[int],
        kospi: list[float],
        fair_kospi: list[float],
        months: int | None,
        *,
        execution_date: datetime | None = None,
    ) -> tuple[list[int], list[float], list[float]]:
        if months is None:
            return years, kospi, fair_kospi

        current_year = (execution_date or datetime.now()).year
        period_years = max(1, months // 12)
        start_year = current_year - period_years + 1
        selected = [
            index
            for index, year in enumerate(years)
            if start_year <= year <= current_year
        ]
        return (
            [years[index] for index in selected],
            [kospi[index] for index in selected],
            [fair_kospi[index] for index in selected],
        )

    def plot_graph_6(self) -> None:
        self._show_graph3_submenu(False)
        self._show_graph6_submenu(True)
        try:
            import matplotlib as mpl
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
            from openpyxl import load_workbook

            mpl.rcParams["font.family"] = "Apple SD Gothic Neo"
            mpl.rcParams["axes.unicode_minus"] = False

            workbook = load_workbook(GDP_WORKBOOK, read_only=True, data_only=True)
            sheet = workbook["Sheet1"]
            years: list[int] = []
            kospi: list[float] = []
            fair_kospi: list[float] = []
            for year_value, _, kospi_value, fair_value in sheet.iter_rows(
                min_row=2, min_col=1, max_col=4, values_only=True
            ):
                if (
                    isinstance(year_value, (int, float))
                    and isinstance(kospi_value, (int, float))
                    and isinstance(fair_value, (int, float))
                ):
                    years.append(int(year_value))
                    kospi.append(float(kospi_value))
                    fair_kospi.append(float(fair_value))
            workbook.close()
            if not years:
                raise ValueError("그래프에 사용할 완전한 연간 데이터가 없습니다.")

            years, kospi, fair_kospi = self._filter_annual_chart_period(
                years, kospi, fair_kospi, self.chart_range_months
            )
            if not years:
                raise ValueError(
                    f"{self._chart_range_label(annual=True)} 범위에 발표된 데이터가 없습니다."
                )

            self._clear_chart_display()

            figure = Figure(figsize=(12, 6.4), dpi=100, facecolor="#FFFFFF")
            axis = figure.add_subplot(111)
            kospi_line = axis.plot(
                years,
                kospi,
                color="#4472C4",
                linewidth=2.6,
                label="KOSPI(연말)",
            )[0]
            fair_line = axis.plot(
                years,
                fair_kospi,
                color="#ED7D31",
                linewidth=2.6,
                label="적정 KOSPI (GDP모델)",
            )[0]

            axis.set_ylim(bottom=0)
            axis.set_ylabel("KOSPI 지수", color="#4B5563", fontsize=10)
            axis.grid(axis="y", color="#D1D5DB", linewidth=0.8)
            axis.grid(axis="x", visible=False)
            axis.spines["top"].set_visible(False)
            axis.spines["right"].set_visible(False)
            axis.spines["bottom"].set_color("#CBD5E1")
            axis.spines["left"].set_color("#CBD5E1")
            axis.tick_params(colors="#4B5563", labelsize=9)
            tick_step = max(1, ceil(len(years) / 12))
            tick_years = years[::tick_step]
            if tick_years[-1] != years[-1]:
                tick_years.append(years[-1])
            axis.set_xticks(tick_years)
            axis.set_xlim(years[0], years[-1])

            figure.suptitle(
                "KOSPI와 적정 KOSPI (GDP 모델)",
                fontsize=15,
                fontweight="bold",
                color="#1F2937",
                y=0.98,
            )
            axis.set_title(
                f"연간 자료 · {years[0]} ~ {years[-1]}",
                fontsize=10,
                color="#6B7280",
                pad=12,
            )
            figure.legend(
                [kospi_line, fair_line],
                [kospi_line.get_label(), fair_line.get_label()],
                loc="lower center",
                ncol=2,
                frameon=False,
                fontsize=10,
                bbox_to_anchor=(0.5, 0.01),
            )
            figure.subplots_adjust(left=0.075, right=0.96, top=0.87, bottom=0.18)

            self.chart_figure = figure
            self.chart_canvas = FigureCanvasTkAgg(figure, master=self.chart_container)
            self.chart_canvas.draw()
            self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
            self.chart_points = len(years)
            self.active_graph_number = 6
            self.graph_title_label.configure(
                text="Graph 6 · KOSPI와 적정 KOSPI (GDP 모델)"
            )
            self.graph_subtitle_label.configure(
                text=(
                    f"data/2.GDP.xlsx · {years[0]}~{years[-1]} "
                    f"· {len(years)}개년 · {self._chart_range_label(annual=True)} "
                    "· KOSPI 연말 / 적정 KOSPI"
                )
            )
            self._append_log(
                f"Graph 6 생성 완료: {len(years)}개년", "success"
            )
        except Exception as exc:
            self._append_log(f"Graph 6 생성 실패: {exc}", "error")
            messagebox.showerror("그래프 생성 실패", str(exc))

    def show_nominal_gdp_forecast(self) -> None:
        if self.gdp_forecast_running:
            return
        self.gdp_forecast_running = True
        self.gdp_forecast_button.configure(state="disabled")
        self.current_gdp_button.configure(state="disabled")
        self.kospi_model_button.configure(state="disabled")
        self.graph_title_label.configure(text="Graph 6 · 명목GDP 예측치")
        self.graph_subtitle_label.configure(
            text="최신 NABO 공식 경제전망과 명목 GDP 실적을 수집하고 있습니다."
        )
        self._clear_chart_display()
        waiting = tk.Label(
            self.chart_container,
            text="공식 자료 수집 및 계산 중…",
            bg="#FFFFFF",
            fg="#334155",
            font=("Apple SD Gothic Neo", 18, "bold"),
        )
        waiting.pack(expand=True)
        self._append_log("최신 명목GDP 전망 자료 수집 시작", "info")
        threading.Thread(target=self._forecast_worker, daemon=True).start()

    def show_current_gdp_estimate(self) -> None:
        if self.gdp_forecast_running:
            return
        self.gdp_forecast_running = True
        self.gdp_forecast_button.configure(state="disabled")
        self.current_gdp_button.configure(state="disabled")
        self.kospi_model_button.configure(state="disabled")
        self.graph_title_label.configure(text="Graph 6 · 실행일 기준 명목GDP 추정")
        self.graph_subtitle_label.configure(
            text="한국은행 분기 실적·속보·최신 경제전망과 NABO 기준 전망을 수집하고 있습니다."
        )
        self._clear_chart_display()
        tk.Label(
            self.chart_container,
            text="분기 실적과 최신 공식 전망 수집 중…",
            bg="#FFFFFF",
            fg="#334155",
            font=("Apple SD Gothic Neo", 18, "bold"),
        ).pack(expand=True)
        self._append_log("실행일 기준 명목GDP 혼합 추정 시작", "info")
        threading.Thread(target=self._current_gdp_worker, daemon=True).start()

    def show_kospi_model_forecast(self) -> None:
        if self.gdp_forecast_running:
            return
        self.gdp_forecast_running = True
        self.gdp_forecast_button.configure(state="disabled")
        self.current_gdp_button.configure(state="disabled")
        self.kospi_model_button.configure(state="disabled")
        self.graph_title_label.configure(text="Graph 6 · 명목GDP 기반 KOSPI 예측모형")
        self.graph_subtitle_label.configure(
            text="두 GDP 시나리오를 수집하고 세 회귀모형을 다시 추정하고 있습니다."
        )
        self._clear_chart_display()
        tk.Label(
            self.chart_container,
            text="GDP 전망 수집 및 KOSPI 회귀분석 중…",
            bg="#FFFFFF",
            fg="#334155",
            font=("Apple SD Gothic Neo", 18, "bold"),
        ).pack(expand=True)
        self._append_log("명목GDP 기반 KOSPI 3개 예측모형 계산 시작", "info")
        threading.Thread(target=self._kospi_model_worker, daemon=True).start()

    def _kospi_model_worker(self) -> None:
        try:
            completed = subprocess.run(
                [sys.executable, "-u", str(KOSPI_GDP_MODEL_SCRIPT), "--json"],
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=360,
            )
            if completed.returncode != 0:
                error = completed.stderr.strip() or completed.stdout.strip()
                raise RuntimeError(error or f"종료 코드 {completed.returncode}")
            self.events.put(("kospi_model_done", completed.stdout.strip()))
        except Exception as exc:
            self.events.put(("kospi_model_error", str(exc)))

    def _current_gdp_worker(self) -> None:
        try:
            completed = subprocess.run(
                [sys.executable, "-u", str(CURRENT_GDP_ESTIMATE_SCRIPT), "--json"],
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=240,
            )
            if completed.returncode != 0:
                error = completed.stderr.strip() or completed.stdout.strip()
                raise RuntimeError(error or f"종료 코드 {completed.returncode}")
            self.events.put(("current_gdp_done", completed.stdout.strip()))
        except Exception as exc:
            self.events.put(("current_gdp_error", str(exc)))

    def _forecast_worker(self) -> None:
        try:
            completed = subprocess.run(
                [sys.executable, "-u", str(GDP_FORECAST_SCRIPT), "--json"],
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=180,
            )
            if completed.returncode != 0:
                error = completed.stderr.strip() or completed.stdout.strip()
                raise RuntimeError(error or f"종료 코드 {completed.returncode}")
            self.events.put(("forecast_done", completed.stdout.strip()))
        except Exception as exc:
            self.events.put(("forecast_error", str(exc)))

    def _display_nominal_gdp_forecast(self, result: dict) -> None:
        self._clear_chart_display()
        outer = tk.Frame(self.chart_container, bg="#FFFFFF", padx=28, pady=20)
        outer.pack(fill="both", expand=True)
        text = tk.Text(
            outer,
            bg="#FFFFFF",
            fg="#1F2937",
            relief="flat",
            wrap="word",
            font=("Apple SD Gothic Neo", 13),
            padx=8,
            pady=8,
        )
        scroll = ttk.Scrollbar(outer, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.tag_configure("title", font=("Apple SD Gothic Neo", 23, "bold"), foreground="#0F766E")
        text.tag_configure("value", font=("Apple SD Gothic Neo", 20, "bold"), foreground="#7C3AED")
        text.tag_configure("heading", font=("Apple SD Gothic Neo", 15, "bold"), foreground="#1D4ED8", spacing1=14, spacing3=5)
        text.tag_configure("note", foreground="#64748B")
        text.insert("end", f"{result['target_year']}년 명목 GDP 예측\n", "title")
        text.insert("end", f"{result['forecast_nominal_gdp_trillion_krw']:,.1f}조 원\n", "value")
        text.insert("end", f"계산 기준일 {result['calculated_at']} · {result['classification']}\n\n", "note")
        text.insert("end", "수집된 수치와 예측치\n", "heading")
        rows = (
            (f"{result['base_year']}년 명목 GDP 실적", f"{result['base_nominal_gdp_trillion_krw']:,.4f}조 원"),
            (f"{result['target_year']}년 명목 GDP 성장률 전망", f"{result['nominal_growth_forecast_pct']:.1f}%"),
            ("실질 GDP 성장률 전망", f"{result['real_growth_forecast_pct']:.1f}%"),
            ("GDP 디플레이터 상승률 전망", f"{result['gdp_deflator_forecast_pct']:.1f}%"),
            ("실질성장률·디플레이터 결합값", f"{result['identity_nominal_growth_pct']:.2f}%"),
        )
        for label, value in rows:
            text.insert("end", f"• {label}: {value}\n")
        text.insert("end", "\nGDP 예측치 계산식\n", "heading")
        text.insert("end", "예상 명목 GDP = 직전연도 명목 GDP × (1 + 명목 GDP 성장률 전망/100)\n")
        text.insert("end", f"{result['formula']}\n")
        text.insert("end", "검산식: (1 + 실질성장률) × (1 + GDP 디플레이터 상승률) − 1\n")
        text.insert("end", "공식 표의 수치는 소수 첫째 자리에서 반올림되므로 검산값과 작은 차이가 날 수 있습니다.\n", "note")
        text.insert("end", "\n정보 출처\n", "heading")
        text.insert("end", f"• 전망: {result['report_title']}\n")
        text.insert("end", f"  발표일 {result['report_published']} · PDF {result['report_page']}쪽\n")
        text.insert("end", f"  {result['report_url']}\n")
        text.insert("end", f"• 실적: {result['actual_source']}\n")
        text.insert("end", "\n※ 이 값은 공식 성장률 전망을 실제 GDP에 적용한 프로그램 계산값이며, 한국은행의 확정 GDP가 아닙니다.\n", "note")
        text.configure(state="disabled")
        self.graph_subtitle_label.configure(
            text=(
                f"{result['report_published']} 발표 전망 · 기준 실적 {result['base_year']}년 "
                f"· 계산일 {result['calculated_at']}"
            )
        )
        self.active_graph_number = 6
        self._append_log(
            f"{result['target_year']}년 명목GDP 예측 계산 완료: "
            f"{result['forecast_nominal_gdp_trillion_krw']:,.1f}조 원",
            "success",
        )

    def _display_current_gdp_estimate(self, result: dict) -> None:
        self._clear_chart_display()
        outer = tk.Frame(self.chart_container, bg="#FFFFFF", padx=26, pady=18)
        outer.pack(fill="both", expand=True)
        text = tk.Text(
            outer,
            bg="#FFFFFF",
            fg="#1F2937",
            relief="flat",
            wrap="word",
            font=("Apple SD Gothic Neo", 12),
            padx=8,
            pady=8,
        )
        scroll = ttk.Scrollbar(outer, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.tag_configure("title", font=("Apple SD Gothic Neo", 22, "bold"), foreground="#0F766E")
        text.tag_configure("value", font=("Apple SD Gothic Neo", 20, "bold"), foreground="#B45309")
        text.tag_configure("heading", font=("Apple SD Gothic Neo", 14, "bold"), foreground="#1D4ED8", spacing1=12, spacing3=5)
        text.tag_configure("actual", foreground="#047857")
        text.tag_configure("estimate", foreground="#B45309")
        text.tag_configure("note", foreground="#64748B")
        text.insert("end", f"{result['target_year']}년 실행일 기준 명목 GDP 추정\n", "title")
        text.insert("end", f"{result['forecast_nominal_gdp_trillion_krw']:,.1f}조 원", "value")
        text.insert("end", f"  (전년 대비 {result['forecast_nominal_growth_pct']:.2f}%)\n")
        text.insert(
            "end",
            f"추정 범위 {result['forecast_range_low_trillion_krw']:,.1f}~"
            f"{result['forecast_range_high_trillion_krw']:,.1f}조 원 · 신뢰도 {result['confidence']}\n",
            "note",
        )
        text.insert(
            "end",
            f"계산 기준일 {result['calculated_at']} · 실적 {result['actual_quarters']}개 분기 / 추정 {result['estimated_quarters']}개 분기\n\n",
            "note",
        )
        text.insert("end", "분기별 계산 내역\n", "heading")
        for quarter in result["quarters"]:
            tag = "actual" if "실적" in quarter["status"] else "estimate"
            text.insert(
                "end",
                f"{quarter['quarter']}분기  {quarter['nominal_gdp_trillion_krw']:>7,.1f}조 원  "
                f"[{quarter['status']}]\n",
                tag,
            )
            text.insert(
                "end",
                f"   전년동기 실질성장 {quarter['real_growth_yoy_pct']}% · "
                f"GDP 디플레이터 {quarter['deflator_yoy_pct']}%\n"
                f"   {quarter['formula']}\n",
            )
        assumptions = result["assumptions"]
        text.insert("end", "\n미발표 분기 가정\n", "heading")
        text.insert("end", f"• 한국은행 최신 연간 실질성장 전망: {assumptions['bok_annual_real_growth_pct']}%\n")
        text.insert("end", f"• 한국은행 최신 소비자물가 전망: {assumptions['bok_cpi_forecast_pct']}%\n")
        text.insert("end", f"• 최근 발표 GDP 디플레이터: {assumptions['latest_actual_deflator_pct']}%\n")
        text.insert("end", f"• 물가전망 보정 후 디플레이터 기준점: {assumptions['adjusted_deflator_anchor_pct']}%\n")
        text.insert("end", f"• 아직 실질 속보가 없는 분기 성장률: {assumptions['missing_quarter_real_growth_pct']}%\n")
        text.insert("end", "\n계산식\n", "heading")
        text.insert("end", "연간 명목 GDP = 발표된 분기 명목 GDP 합계 + 미발표 분기 명목 GDP 추정 합계\n")
        text.insert("end", "미발표 분기 = 전년동기 명목 GDP × (1+실질성장률) × (1+GDP 디플레이터 상승률)\n")
        text.insert("end", "디플레이터는 최근 실적의 가중치를 분기마다 낮추면서 공식 전망 기준점으로 수렴시킵니다.\n", "note")
        text.insert("end", "\n정보 출처\n", "heading")
        seen: set[str] = set()
        for source in result["sources"]:
            if source["url"] in seen:
                continue
            seen.add(source["url"])
            text.insert("end", f"• {source['label']}\n  {source['published']}\n  {source['url']}\n")
        text.insert("end", f"\n데이터 품질: {result['quality_note']}\n", "note")
        text.insert("end", "※ 공식기관 자체 전망치가 아니라 공식 실적과 전망을 결합한 프로그램 추정치입니다.\n", "note")
        text.configure(state="disabled")
        self.graph_subtitle_label.configure(
            text=(
                f"실적 반영 {result['actual_share_pct']}% · "
                f"추정 {100-result['actual_share_pct']}% · 계산일 {result['calculated_at']}"
            )
        )
        self.active_graph_number = 6
        self._append_log(
            f"실행일 기준 {result['target_year']}년 명목GDP 추정 완료: "
            f"{result['forecast_nominal_gdp_trillion_krw']:,.1f}조 원",
            "success",
        )

    def _display_kospi_model_forecast(self, result: dict) -> None:
        self._clear_chart_display()
        outer = tk.Frame(self.chart_container, bg="#FFFFFF", padx=22, pady=16)
        outer.pack(fill="both", expand=True)
        text = tk.Text(
            outer,
            bg="#FFFFFF",
            fg="#1F2937",
            relief="flat",
            wrap="word",
            font=("Apple SD Gothic Neo", 11),
            padx=10,
            pady=8,
        )
        scroll = ttk.Scrollbar(outer, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.tag_configure("title", font=("Apple SD Gothic Neo", 21, "bold"), foreground="#9F1239")
        text.tag_configure("value", font=("Menlo", 14, "bold"), foreground="#7C3AED")
        text.tag_configure("heading", font=("Apple SD Gothic Neo", 14, "bold"), foreground="#1D4ED8", spacing1=14, spacing3=5)
        text.tag_configure("model", font=("Apple SD Gothic Neo", 13, "bold"), foreground="#0F766E", spacing1=10)
        text.tag_configure("formula", font=("Menlo", 10), foreground="#111827")
        text.tag_configure("warning", foreground="#B45309")
        text.tag_configure("danger", foreground="#BE123C")
        text.tag_configure("note", foreground="#64748B")

        scenarios = result["scenarios"]
        nabo = scenarios["nominal_gdp_forecast"]
        current = scenarios["estimate_current_gdp"]
        rate = result["rate_assumption"]
        text.insert("end", f"{result['forecast_year']}년 연말 KOSPI 회귀 예측\n", "title")
        text.insert(
            "end",
            f"GDP 시나리오: NABO {nabo['nominal_gdp_trillion_krw']:,.1f}조 원 / "
            f"현재추정 {current['nominal_gdp_trillion_krw']:,.1f}조 원\n",
            "value",
        )
        text.insert(
            "end",
            f"금리 가정 {rate['rate_pct']:.2f}% · {rate['method']} · 계산일 {result['calculated_at']}\n\n",
            "note",
        )

        labels = {
            "gdp_only": "모형 1 · GDP 단일 로그모형",
            "gdp_and_rate": "모형 2 · GDP + 10년 국채금리",
            "gdp_squared_and_rate": "모형 3 · GDP 제곱항 + 10년 국채금리",
        }
        text.insert("end", "예측값 비교\n", "heading")
        text.insert("end", f"{'모형':<28} {'NABO GDP':>12} {'현재 GDP 추정':>14}\n", "formula")
        text.insert("end", "─" * 62 + "\n", "note")
        for key, label in labels.items():
            left = nabo["model_forecasts"][key]["forecast_year_end_kospi"]
            right = current["model_forecasts"][key]["forecast_year_end_kospi"]
            text.insert("end", f"{label:<28} {left:>12,.2f} {right:>14,.2f}\n", "formula")

        for key, label in labels.items():
            model = result["models"][key]
            stats = model["regression"]
            nabo_kospi = nabo["model_forecasts"][key]["forecast_year_end_kospi"]
            current_kospi = current["model_forecasts"][key]["forecast_year_end_kospi"]
            text.insert("end", f"\n{label}\n", "model")
            text.insert("end", f"NABO GDP 기준 KOSPI  {nabo_kospi:,.2f}\n", "value")
            text.insert("end", f"현재 GDP 추정 기준 KOSPI  {current_kospi:,.2f}\n", "value")
            text.insert("end", model["fitted_equation"] + "\n", "formula")
            text.insert(
                "end",
                f"표본 {stats['sample_start_year']}~{stats['sample_end_year']}년, n={stats['sample_size']}  |  "
                f"R²={stats['r_squared']:.4f}, 조정 R²={stats['adjusted_r_squared']:.4f}\n",
            )
            text.insert(
                "end",
                f"회귀 표준오차={stats['residual_standard_error']:.4f}  |  "
                f"F={stats['f_value']:.4f}  |  잔차 자유도={stats['degrees_of_freedom']}\n",
            )
            text.insert(
                "end",
                f"SSR={stats['regression_sum_squares']:.6f}  |  "
                f"SSE={stats['residual_sum_squares']:.6f}  |  조건수={stats['condition_number']:.1f}\n",
            )
            coefficient_names = [
                name for name in ("alpha", "beta", "beta1", "beta2", "gamma") if name in stats
            ]
            text.insert("end", "계수(표준오차): ", "note")
            text.insert(
                "end",
                ", ".join(
                    f"{name}={stats[name]:.6f} ({stats['standard_errors'][name]:.6f})"
                    for name in coefficient_names
                ) + "\n",
            )
            if key == "gdp_only":
                text.insert("end", "해석: 설명력은 비교적 높지만 GDP 한 변수만으로 시장 변동을 모두 설명할 수 없어 신뢰도는 보통입니다.\n", "warning")
            elif key == "gdp_and_rate":
                text.insert("end", "해석: R² 개선이 작고 금리계수 표준오차가 크며 부호도 통상적 기대와 반대이므로 신뢰도는 낮음~보통입니다.\n", "warning")
            else:
                text.insert("end", "해석: 표본 내 R²는 가장 높지만 조건수가 매우 커 다중공선성이 강합니다. 표본외 예측 신뢰도는 낮습니다.\n", "danger")

        text.insert("end", "\n통계값 읽는 법\n", "heading")
        text.insert("end", "• R²: 과거 ln(KOSPI) 변동 중 모형이 설명한 비율입니다. 높다고 미래예측이 자동으로 정확해지는 것은 아닙니다.\n")
        text.insert("end", "• 조정 R²: 변수 수가 늘어날 때의 과적합 벌점을 반영한 설명력입니다.\n")
        text.insert("end", "• 계수 표준오차: 계수 추정의 불확실성입니다. 계수보다 표준오차가 크면 방향과 크기를 신뢰하기 어렵습니다.\n")
        text.insert("end", "• 회귀 표준오차: 로그 지수 기준의 전형적인 잔차 크기이며 작을수록 과거 적합이 좋습니다.\n")
        text.insert("end", "• F값: 모든 설명변수의 계수가 동시에 0인 모형과 비교한 전체 설명력 지표입니다.\n")
        text.insert("end", "• SSR/SSE: 각각 모형이 설명한 변동과 설명하지 못한 잔차 변동입니다.\n")
        text.insert("end", "• 조건수: 회귀행렬의 불안정성 지표입니다. 매우 크면 계수가 작은 데이터 변화에도 크게 달라질 수 있습니다.\n")

        text.insert("end", "\n전체 판단\n", "heading")
        text.insert("end", "세 값은 구조적 시나리오 범위로 비교하는 것이 적절합니다. 단일 숫자를 목표지수로 해석하기보다는 GDP 추정 방식과 금리 가정에 따른 민감도를 확인해야 합니다. ", "warning")
        text.insert("end", "현재는 예측연도 12월 금리가 발표되지 않아 최신 확정금리를 대용하며, 정책·유동성·기업이익·위험선호는 별도로 반영되지 않았습니다. 투자 판단을 보장하는 수치가 아닙니다.\n", "danger")
        text.insert("end", f"\n원자료: {result['workbook']}\n", "note")
        text.configure(state="disabled")
        self.graph_subtitle_label.configure(
            text=(
                f"표본 {result['models']['gdp_only']['regression']['sample_start_year']}~"
                f"{result['models']['gdp_only']['regression']['sample_end_year']}년 · "
                f"금리 가정 {rate['rate_pct']:.2f}% · 계산일 {result['calculated_at']}"
            )
        )
        self.active_graph_number = 6
        self._append_log("KOSPI 3개 회귀모형과 통계 해설 표시 완료", "success")

    def _plot_general_dual_axis(
        self,
        *,
        graph_number: int,
        right_column: int,
        right_label: str,
        chart_title: str,
        legend_label: str,
        right_axis_zero: bool = False,
        right_axis_step: int | None = None,
    ) -> None:
        self._show_graph6_submenu(False)
        self._show_graph3_submenu(graph_number == 3)
        try:
            import matplotlib as mpl
            from matplotlib import dates as mdates
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
            from openpyxl import load_workbook

            mpl.rcParams["font.family"] = "Apple SD Gothic Neo"
            mpl.rcParams["axes.unicode_minus"] = False

            workbook = load_workbook(GENERAL_WORKBOOK, read_only=True, data_only=True)
            sheet = workbook["Sheet1"]
            dates: list[datetime] = []
            kospi: list[float] = []
            comparison: list[float] = []
            for row in sheet.iter_rows(
                min_row=2, min_col=1, max_col=right_column, values_only=True
            ):
                month = row[0]
                kospi_value = row[1]
                comparison_value = row[right_column - 1]
                if (
                    isinstance(month, datetime)
                    and isinstance(kospi_value, (int, float))
                    and isinstance(comparison_value, (int, float))
                ):
                    dates.append(month)
                    kospi.append(float(kospi_value))
                    comparison.append(float(comparison_value))
            workbook.close()
            if not dates:
                raise ValueError("그래프에 사용할 완전한 월별 데이터가 없습니다.")

            dates, kospi, comparison = self._filter_chart_period(
                dates, kospi, comparison, self.chart_range_months
            )
            if not dates:
                raise ValueError(
                    f"{self._chart_range_label()} 범위에 발표된 데이터가 없습니다."
                )

            self._clear_chart_display()

            figure = Figure(figsize=(12, 6.4), dpi=100, facecolor="#FFFFFF")
            axis_left = figure.add_subplot(111)
            axis_right = axis_left.twinx()
            blue = "#4472C4"
            orange = "#ED7D31"
            kospi_line = axis_left.plot(
                dates, kospi, color=blue, linewidth=2.4, label="KOSPI(월말,좌)"
            )[0]
            comparison_line = axis_right.plot(
                dates,
                comparison,
                color=orange,
                linewidth=2.4,
                label=legend_label,
            )[0]

            axis_left.set_ylim(bottom=0)
            if right_axis_step:
                right_min = (
                    floor(min(comparison) / right_axis_step) * right_axis_step
                    - right_axis_step
                )
                right_max = ceil(max(comparison) / right_axis_step) * right_axis_step
            else:
                right_min = 0 if right_axis_zero else floor(min(comparison)) - 1
                right_max = ceil(max(comparison)) + 1
            axis_right.set_ylim(right_min, right_max)
            axis_left.set_ylabel("KOSPI 월말 지수", color="#4B5563", fontsize=10)
            axis_right.set_ylabel(right_label, color="#4B5563", fontsize=10)
            axis_left.grid(axis="y", color="#D1D5DB", linewidth=0.8)
            axis_left.grid(axis="x", visible=False)
            axis_left.spines["top"].set_visible(False)
            axis_right.spines["top"].set_visible(False)
            axis_left.spines["right"].set_visible(False)
            axis_right.spines["left"].set_visible(False)
            axis_left.spines["bottom"].set_color("#CBD5E1")
            axis_left.spines["left"].set_color("#CBD5E1")
            axis_right.spines["right"].set_color("#CBD5E1")
            axis_left.tick_params(colors="#4B5563", labelsize=9)
            axis_right.tick_params(colors="#4B5563", labelsize=9)

            tick_interval = max(1, round(len(dates) / 36))
            axis_left.xaxis.set_major_locator(mdates.MonthLocator(interval=tick_interval))
            axis_left.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
            axis_left.tick_params(axis="x", rotation=90)
            axis_left.set_xlim(dates[0], dates[-1])

            figure.suptitle(
                chart_title,
                fontsize=15,
                fontweight="bold",
                color="#1F2937",
                y=0.98,
            )
            axis_left.set_title(
                f"월별 자료 · {dates[0]:%Y-%m} ~ {dates[-1]:%Y-%m}",
                fontsize=10,
                color="#6B7280",
                pad=12,
            )
            figure.legend(
                [kospi_line, comparison_line],
                [kospi_line.get_label(), comparison_line.get_label()],
                loc="lower center",
                ncol=2,
                frameon=False,
                fontsize=10,
                bbox_to_anchor=(0.5, 0.01),
            )
            figure.subplots_adjust(left=0.075, right=0.925, top=0.87, bottom=0.23)

            self.chart_figure = figure
            self.chart_canvas = FigureCanvasTkAgg(figure, master=self.chart_container)
            self.chart_canvas.draw()
            self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
            self.chart_points = len(dates)
            self.active_graph_number = graph_number
            self.graph_title_label.configure(text=f"Graph {graph_number} · {chart_title}")
            self.graph_subtitle_label.configure(
                text=(
                    f"data/0_general.xlsx · {dates[0]:%Y-%m}~{dates[-1]:%Y-%m} "
                    f"· {len(dates)}개월 · {self._chart_range_label()} "
                    f"· 좌축 KOSPI / 우축 {right_label}"
                )
            )
            self._append_log(
                f"Graph {graph_number} 생성 완료: {len(dates)}개월", "success"
            )
        except Exception as exc:
            self._append_log(f"Graph {graph_number} 생성 실패: {exc}", "error")
            messagebox.showerror("그래프 생성 실패", str(exc))

    def _chart_range_label(self, *, annual: bool = False) -> str:
        if self.chart_range_months == 120:
            return "실행연도 기준 최근 10년" if annual else "실행월 기준 최근 10년"
        if self.chart_range_months == 60:
            return "실행연도 기준 최근 5년" if annual else "실행월 기준 최근 5년"
        return "전체"

    def _panel(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["panel"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )

    def _section_title(
        self, parent: tk.Misc, title: str, subtitle: str
    ) -> tuple[tk.Label, tk.Label]:
        box = tk.Frame(parent, bg=COLORS["panel"], padx=14, pady=12)
        box.pack(fill="x")
        title_label = tk.Label(
            box,
            text=title,
            bg=COLORS["panel"],
            fg=COLORS["accent"],
            font=("Apple SD Gothic Neo", 19, "bold"),
            anchor="w",
        )
        title_label.pack(fill="x")
        subtitle_label = tk.Label(
            box,
            text=subtitle,
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=("Apple SD Gothic Neo", 11),
            anchor="w",
        )
        subtitle_label.pack(fill="x", pady=(3, 0))
        return title_label, subtitle_label

    def _button(
        self,
        parent: tk.Misc,
        text: str,
        command,
        background: str,
        foreground: str,
        active_background: str,
    ) -> ColorButton:
        return ColorButton(
            parent,
            text=text,
            command=command,
            background=background,
            foreground=foreground,
            active_background=active_background,
        )

    def _append_log(self, message: str, tag: str = "info") -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{stamp}] ", "time")
        self.log.insert("end", message.rstrip() + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def run_update(self, script: Path, workbook: Path, update_name: str) -> None:
        if self.process is not None:
            return
        self._show_sources(update_name)
        if not script.exists():
            messagebox.showerror("스크립트 없음", f"{script}을 찾을 수 없습니다.")
            return
        self.active_script = script
        self.active_workbook = workbook
        self.active_update_name = update_name
        self.pending_data = False
        self.progress_total = 0
        self.progress_completed = 0
        self.progress_unit = "개월"
        self._set_progress(0, "업데이트 대상 확인 중")
        self.update_button.configure(state="disabled")
        self.economic_state_button.configure(state="disabled")
        self.gdp_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_label.configure(text="● RUNNING", fg=COLORS["warning"])
        self.workbook_label.configure(text=str(workbook))
        self._append_log(f"{update_name} 실행 시작", "info")
        threading.Thread(target=self._worker, daemon=True).start()

    def _show_sources(self, update_name: str) -> None:
        if update_name == "update_gdp":
            self.current_sources = GDP_SOURCES
            self.source_title_label.configure(text="자료 출처 · update_gdp")
            self.source_subtitle_label.configure(
                text="data/2.GDP.xlsx의 연간 원자료와 GDP 모델 · 행을 더블클릭하면 제공처를 엽니다"
            )
            self.source_note.configure(
                text="※ 실행일 기준 직전 완료 연도까지 확인하며, 미발표 항목은 다음 실행 때 다시 확인합니다."
            )
        elif update_name == "update_economic_state":
            self.current_sources = ECONOMIC_STATE_SOURCES
            self.source_title_label.configure(text="자료 출처 · update_economic_state")
            self.source_subtitle_label.configure(
                text="data/1.economic_state.xlsx의 4개 지표 · 행을 더블클릭하면 제공처를 엽니다"
            )
            self.source_note.configure(
                text="※ KOSPI, 선행지수, OECD CLI, 제조업 재고율이 모두 확보된 월만 저장합니다."
            )
        else:
            self.current_sources = GENERAL_SOURCES
            self.source_title_label.configure(text="자료 출처 · update_general")
            self.source_subtitle_label.configure(
                text="data/0_general.xlsx의 7개 지표 · 행을 더블클릭하면 제공처를 엽니다"
            )
            self.source_note.configure(
                text="※ 월중 실행 시 당월 확정 자료가 없어 기본 종료월은 직전 달입니다."
            )

        for item in self.source_table.get_children():
            self.source_table.delete(item)
        for index, (indicator, provider, method, _) in enumerate(self.current_sources):
            self.source_table.insert(
                "", "end", iid=str(index), values=(indicator, provider, method)
            )

    def _set_progress(self, percent: float, label: str) -> None:
        value = max(0.0, min(100.0, percent))
        self.progress_bar.configure(value=value)
        self.progress_percent.configure(text=f"{value:.0f}%")
        self.progress_label.configure(text=label)

    def _worker(self) -> None:
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        try:
            self.process = subprocess.Popen(
                [sys.executable, "-u", str(self.active_script)],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.events.put(("line", line.rstrip()))
            code = self.process.wait()
            self.events.put(("done", code))
        except Exception as exc:  # UI boundary: report unexpected launch failures.
            self.events.put(("fatal", str(exc)))

    def stop_update(self) -> None:
        if self.process and self.process.poll() is None:
            self._append_log("중지 요청을 보냈습니다.", "error")
            self.process.terminate()

    def _drain_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "line":
                    line = str(payload)
                    if line.startswith("PLAN "):
                        match = re.match(r"PLAN\s+(\d+)(개월|개년):", line)
                        if match:
                            self.progress_total = int(match.group(1))
                            self.progress_unit = match.group(2)
                        else:
                            self.progress_total = 0
                        self._set_progress(
                            0, f"0/{self.progress_total}{self.progress_unit} 완료"
                        )
                        tag = "info"
                    elif line.startswith("COLLECT "):
                        month = line.split(maxsplit=1)[1]
                        base = 100 * self.progress_completed / self.progress_total if self.progress_total else 0
                        self._set_progress(base, f"{month} 수집 중")
                        tag = "info"
                    elif line.startswith("SAVED "):
                        self.progress_completed += 1
                        percent = 100 * self.progress_completed / self.progress_total if self.progress_total else 100
                        self._set_progress(
                            percent,
                            f"{self.progress_completed}/{self.progress_total}{self.progress_unit} 완료",
                        )
                        tag = "success"
                    elif line.startswith("NO_UPDATE"):
                        self._set_progress(100, "추가 업데이트 없음")
                        tag = "success"
                    elif line.startswith("PENDING"):
                        self.pending_data = True
                        percent = 100 * self.progress_completed / self.progress_total if self.progress_total else 0
                        self._set_progress(
                            percent,
                            f"{self.progress_completed}/{self.progress_total}{self.progress_unit} 완료 · 발표 대기",
                        )
                        tag = "warning"
                    else:
                        tag = "error" if line.startswith("FAILED") else "success" if line.startswith("UPDATED") else "info"
                    self._append_log(line, tag)
                elif kind == "done":
                    self._finish(int(payload))
                elif kind == "fatal":
                    self._append_log(str(payload), "error")
                    self._finish(1)
                elif kind == "forecast_done":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    try:
                        self._display_nominal_gdp_forecast(json.loads(str(payload)))
                    except Exception as exc:
                        self._append_log(f"명목GDP 예측 결과 표시 실패: {exc}", "error")
                        messagebox.showerror("명목GDP 예측 실패", str(exc))
                elif kind == "forecast_error":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    self._append_log(f"명목GDP 예측 실패: {payload}", "error")
                    self.graph_subtitle_label.configure(
                        text="공식 자료를 수집하지 못했습니다. 실행 로그를 확인하세요."
                    )
                    messagebox.showerror("명목GDP 예측 실패", str(payload))
                elif kind == "current_gdp_done":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    try:
                        self._display_current_gdp_estimate(json.loads(str(payload)))
                    except Exception as exc:
                        self._append_log(f"현재 GDP 추정 결과 표시 실패: {exc}", "error")
                        messagebox.showerror("현재 GDP 추정 실패", str(exc))
                elif kind == "current_gdp_error":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    self._append_log(f"현재 GDP 추정 실패: {payload}", "error")
                    self.graph_subtitle_label.configure(
                        text="공식 자료를 수집하지 못했습니다. 실행 로그를 확인하세요."
                    )
                    messagebox.showerror("현재 GDP 추정 실패", str(payload))
                elif kind == "kospi_model_done":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    try:
                        self._display_kospi_model_forecast(json.loads(str(payload)))
                    except Exception as exc:
                        self._append_log(f"KOSPI 예측모형 표시 실패: {exc}", "error")
                        messagebox.showerror("KOSPI 예측모형 실패", str(exc))
                elif kind == "kospi_model_error":
                    self.gdp_forecast_running = False
                    self.gdp_forecast_button.configure(state="normal")
                    self.current_gdp_button.configure(state="normal")
                    self.kospi_model_button.configure(state="normal")
                    self._append_log(f"KOSPI 예측모형 계산 실패: {payload}", "error")
                    self.graph_subtitle_label.configure(
                        text="KOSPI 예측모형을 계산하지 못했습니다. 실행 로그를 확인하세요."
                    )
                    messagebox.showerror("KOSPI 예측모형 실패", str(payload))
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _finish(self, return_code: int) -> None:
        self.process = None
        self.update_button.configure(state="normal")
        self.economic_state_button.configure(state="normal")
        self.gdp_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        if return_code == 0:
            if self.pending_data:
                self.status_label.configure(text="● WAITING FOR RELEASE", fg=COLORS["warning"])
                self._append_log("공식 발표된 자료까지 저장했습니다. 다음 자료 발표 후 다시 실행하세요.", "warning")
            else:
                self.status_label.configure(text="● COMPLETE", fg=COLORS["success"])
                self._append_log(f"완료: {self.active_workbook}", "success")
            refresh_general = (
                self.active_update_name == "update_general"
                and self.active_graph_number in {1, 2, 3, 4, 5}
            )
            refresh_gdp = (
                self.active_update_name == "update_gdp"
                and self.active_graph_number == 6
            )
            if self.chart_canvas is not None and (refresh_general or refresh_gdp):
                self._append_log(
                    f"추가된 데이터로 Graph {self.active_graph_number}을 자동 갱신합니다.",
                    "info",
                )
                self._refresh_active_graph()
        else:
            self.status_label.configure(text=f"● FAILED ({return_code})", fg=COLORS["danger"])
            self._append_log(f"업데이트가 종료 코드 {return_code}로 끝났습니다.", "error")

    def _open_selected_source(self, _event=None) -> None:
        selected = self.source_table.selection()
        if selected:
            webbrowser.open(self.current_sources[int(selected[0])][3])

    def _on_close(self) -> None:
        if self.process and self.process.poll() is None:
            if not messagebox.askyesno("업데이트 실행 중", "업데이트를 중지하고 창을 닫을까요?"):
                return
            self.process.terminate()
        self.destroy()


def self_check() -> int:
    required = (
        GENERAL_UPDATE_SCRIPT,
        ECONOMIC_STATE_UPDATE_SCRIPT,
        GDP_UPDATE_SCRIPT,
        GDP_FORECAST_SCRIPT,
        CURRENT_GDP_ESTIMATE_SCRIPT,
        KOSPI_GDP_MODEL_SCRIPT,
        HP_EXPORT_KOSPI_SCRIPT,
        GENERAL_WORKBOOK,
        ECONOMIC_STATE_WORKBOOK,
        GDP_WORKBOOK,
    )
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        print("MISSING", *missing, sep="\n")
        return 1
    if len(GENERAL_SOURCES) != 7 or len(ECONOMIC_STATE_SOURCES) != 4 or len(GDP_SOURCES) != 4:
        print(
            "INVALID_SOURCE_COUNT "
            f"general={len(GENERAL_SOURCES)} economic={len(ECONOMIC_STATE_SOURCES)} gdp={len(GDP_SOURCES)}"
        )
        return 1
    print(
        "UI_CHECK_OK workbooks=3 scripts=7 sources=general:7,economic:4,gdp:4 "
        f"python={sys.executable}"
    )
    return 0


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="창을 열지 않고 구성만 검증")
    args = parser.parse_args()
    if args.check:
        raise SystemExit(self_check())
    EconoUpdaterApp().mainloop()


if __name__ == "__main__":
    main()
