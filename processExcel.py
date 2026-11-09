"""Read dated Excel workbooks and build Extract.xlsx plus Plotly charts.

Input files are named YYYYMMDD.xlsx. Each workbook is read until column A
is "?". Output is written to an "output" folder inside the selected folder.
"""

from __future__ import annotations

import datetime as dt
from fractions import Fraction
import html
import math
import os
import re
import subprocess
import sys
import threading
import urllib.parse
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw, ImageFont
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText as ChartRichText
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawFont,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichTextProperties,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

FILE_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})\.xlsx$", re.IGNORECASE)
QUESTION_MARKS = {"?", "？"}
CLOCK_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$")
METER_RE = re.compile(r"(\d+(?:\.\d+)?)\s*米")
NUMBER_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
DATE_RE = re.compile(r"(\d{4})[-/]?(\d{2})[-/]?(\d{2})")
FONT_NAME = "Microsoft JhengHei"
PALETTE = [
    "1F77B4",
    "FF7F0E",
    "2CA02C",
    "D62728",
    "9467BD",
    "8C564B",
    "E377C2",
    "7F7F7F",
    "BCBD22",
    "17BECF",
]
HEADER_COLORS = ["1F4E79", "196F3D", "B9770E", "6C3483", "922B21", "1A5276"]
THIN = Border(
    left=Side(style="thin", color="D5D8DC"),
    right=Side(style="thin", color="D5D8DC"),
    top=Side(style="thin", color="D5D8DC"),
    bottom=Side(style="thin", color="D5D8DC"),
)

EXERCISE_STRIDE = 4
BP_STRIDE = 5
SLEEP_STRIDE = 5
WATER_STRIDE = 4
WATER_HEADERS = ["喝水", "喝水日期", "喝水時間", "喝水量(杯)"]
CUP_FRACTION_RE = re.compile(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)")


@dataclass
class PlotStyle:
    width_cm: float = 22.0
    height_cm: float = 12.0
    title_size: float = 18.0
    xlabel_size: float = 14.0
    ylabel_size: float = 14.0
    legend_size: float = 12.0
    tick_size: float = 11.0

    @property
    def width_px(self) -> int:
        return int(round(self.width_cm / 2.54 * 96))

    @property
    def height_px(self) -> int:
        return int(round(self.height_cm / 2.54 * 96))


@dataclass
class ProcessResult:
    excel_path: Path
    plot_path: Path
    messages: list[str] = field(default_factory=list)
    pdf_path: Path | None = None


def is_question(value) -> bool:
    return isinstance(value, str) and value.strip() in QUESTION_MARKS


def is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        text = value.strip()
        if text == "" or text in QUESTION_MARKS or set(text) <= QUESTION_MARKS:
            return True
        if text.startswith("="):
            return True
    return False


def parse_date(value) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if is_missing(value):
        return None
    match = DATE_RE.search(str(value).strip())
    if not match:
        return None
    try:
        return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def parse_clock_seconds(value) -> float | None:
    """Return seconds from midnight for a clock value."""
    if isinstance(value, dt.datetime):
        return value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6
    if isinstance(value, dt.time):
        return value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6
    if isinstance(value, dt.timedelta):
        return value.total_seconds() % (24 * 3600)
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if 0 <= number < 1:
            return number * 86400
        return None
    if is_missing(value):
        return None
    text = str(value).strip()
    match = CLOCK_RE.fullmatch(text)
    if not match:
        return None
    hour, minute, second = int(match.group(1)), int(match.group(2)), int(match.group(3) or 0)
    if minute > 59 or second > 59 or hour > 48:
        return None
    return hour * 3600 + minute * 60 + second


def duration_seconds(start, end, fallback=None) -> float | None:
    """End minus start, wrapping past midnight. A stored duration is used when a clock is missing."""
    start_s = parse_clock_seconds(start)
    end_s = parse_clock_seconds(end)
    if start_s is not None and end_s is not None:
        diff = end_s - start_s
        if diff < 0:
            diff += 24 * 3600
        return diff
    return parse_clock_seconds(fallback)


def format_clock(value) -> str | None:
    seconds = parse_clock_seconds(value)
    if seconds is None:
        return None
    whole = int(round(seconds)) % (24 * 3600)
    hour, rem = divmod(whole, 3600)
    minute, second = divmod(rem, 60)
    return f"{hour:02d}:{minute:02d}:{second:02d}"


def parse_meters(value) -> float | None:
    """Parse 10 from '10米: 10 meter'."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if is_missing(value):
        return None
    text = str(value).strip()
    match = METER_RE.search(text)
    if match:
        return float(match.group(1))
    if NUMBER_RE.fullmatch(text):
        return float(text)
    return None


def parse_number(value) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if is_missing(value):
        return None
    text = str(value).strip().replace(",", "")
    if not NUMBER_RE.fullmatch(text):
        return None
    return float(text)


def as_number(value: float | None):
    if value is None:
        return None
    if math.isfinite(value) and abs(value - round(value)) < 1e-9:
        return int(round(value))
    return value


def header_text(header_row: list, column: int, fallback: str) -> str:
    index = column - 1
    if 0 <= index < len(header_row):
        value = header_row[index]
        if isinstance(value, str) and value.strip():
            return value
    return fallback


def with_unit(title: str, unit: str) -> str:
    text = title.strip()
    if unit in text:
        return text
    return f"{text}\n({unit})"


def list_dated_workbooks(folder: Path) -> list[Path]:
    """Excel files named YYYYMMDD.xlsx, sorted by year, month, date."""
    found = []
    for path in Path(folder).iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        match = FILE_RE.match(path.name)
        if not match:
            continue
        year, month, day = (int(part) for part in match.groups())
        try:
            stamp = dt.date(year, month, day)
        except ValueError:
            continue
        found.append((stamp, path.name.lower(), path))
    found.sort()
    return [item[-1] for item in found]


def load_table(worksheet, max_col: int, cached=None) -> tuple[list, list[list]]:
    """Read headers and data rows. Stop at '?' in column A and skip blank rows."""
    header: list = []
    rows: list[list] = []
    limit = min(worksheet.max_row or 1, 5000)
    cached_rows = []
    if cached is not None:
        cached_rows = list(
            cached.iter_rows(min_row=1, max_row=limit, max_col=max_col, values_only=True)
        )
    for index, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=limit, max_col=max_col, values_only=True),
        start=1,
    ):
        values = list(row)
        if cached_rows and index - 1 < len(cached_rows):
            values = _prefer_cached(values, list(cached_rows[index - 1]))
        if index == 1:
            header = values
            continue
        label = values[0] if values else None
        if is_question(label):
            break
        if not values or all(is_missing(value) for value in values):
            continue
        rows.append(values)
    return header, rows


def _prefer_cached(values: list, cached: list) -> list:
    merged = []
    for index, value in enumerate(values):
        if isinstance(value, str) and value.startswith("="):
            merged.append(cached[index] if index < len(cached) else None)
        else:
            merged.append(value)
    return merged


def cell_at(row: list, column: int):
    index = column - 1
    if index < 0 or index >= len(row):
        return None
    return row[index]


def file_date(path: Path) -> dt.date | None:
    match = FILE_RE.match(path.name)
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None


def date_label(value: dt.date | None, fallback: str) -> str:
    if value is None:
        return fallback
    return value.strftime("%Y-%m-%d")


def legend_for(blocks: list[dict]) -> None:
    seen: dict[str, int] = {}
    for block in blocks:
        label = block["legend"]
        seen[label] = seen.get(label, 0) + 1
        if seen[label] > 1:
            block["legend"] = f"{label} ({seen[label]})"


def extract_exercise(path: Path, worksheet, cached=None) -> dict | None:
    header, rows = load_table(worksheet, 11, cached)
    parsed = []
    for row in rows:
        minutes_s = duration_seconds(cell_at(row, 5), cell_at(row, 7), cell_at(row, 9))
        minutes = None if minutes_s is None else minutes_s / 60
        parsed.append(
            {
                "sport": None if is_missing(cell_at(row, 1)) else str(cell_at(row, 1)).strip(),
                "date": parse_date(cell_at(row, 3)) or file_date(path),
                "minutes": minutes,
                "meters": parse_meters(cell_at(row, 11)),
            }
        )
    parsed = [row for row in parsed if row["sport"] or row["minutes"] is not None or row["meters"] is not None]
    if not parsed:
        return None
    stamp = file_date(path) or next((row["date"] for row in parsed if row["date"]), None)
    return {
        "legend": date_label(stamp, path.stem),
        "date_label": date_label(stamp, path.stem),
        "headers": [
            header_text(header, 1, "選擇運動\nPilih olahraga"),
            header_text(header, 3, "運動日期\nTanggal Olahraga"),
            with_unit(header_text(header, 9, "運動時間\nWaktu berolahraga"), "分鐘"),
            with_unit(header_text(header, 11, "運動距離\nJarak pergerakan"), "米"),
        ],
        "rows": parsed,
    }


def measurement_label(when: dt.date | None, clock, fallback: str) -> str:
    base = date_label(when, fallback)
    text = format_clock(clock)
    if text is None:
        return base
    return f"{base} {text[:5]}"


def extract_bp(path: Path, worksheet, cached=None) -> dict | None:
    header, rows = load_table(worksheet, 7, cached)
    parsed = []
    for row in rows:
        when = parse_date(cell_at(row, 3)) or file_date(path)
        parsed.append(
            {
                "label": None if is_missing(cell_at(row, 1)) else str(cell_at(row, 1)).strip(),
                "date": when,
                "when": measurement_label(when, cell_at(row, 4), path.stem),
                "high": parse_number(cell_at(row, 5)),
                "low": parse_number(cell_at(row, 6)),
                "heart": parse_number(cell_at(row, 7)),
            }
        )
    parsed = [
        row for row in parsed
        if row["label"] or row["high"] is not None or row["low"] is not None or row["heart"] is not None
    ]
    if not parsed:
        return None
    stamp = file_date(path) or next((row["date"] for row in parsed if row["date"]), None)
    return {
        "legend": date_label(stamp, path.stem),
        "date_label": date_label(stamp, path.stem),
        "headers": [
            header_text(header, 1, "健康\nSehat"),
            header_text(header, 3, "量血壓日期\nTanggal pengukuran tekanan darah"),
            header_text(header, 5, "收縮壓, 高壓\nTekanan darah sistolik\ntekanan tinggi"),
            header_text(header, 6, "舒張壓, 低壓\nTekanan darah diastolik\ntekanan rendah"),
            header_text(header, 7, "心跳\nDetak jantung"),
        ],
        "rows": parsed,
    }


def extract_sleep(path: Path, worksheet, cached=None) -> dict | None:
    header, rows = load_table(worksheet, 9, cached)
    parsed = []
    for row in rows:
        hours_s = duration_seconds(cell_at(row, 5), cell_at(row, 7), cell_at(row, 9))
        hours = None if hours_s is None else hours_s / 3600
        parsed.append(
            {
                "label": None if is_missing(cell_at(row, 1)) else str(cell_at(row, 1)).strip(),
                "date": parse_date(cell_at(row, 3)) or file_date(path),
                "start": format_clock(cell_at(row, 5)),
                "end": format_clock(cell_at(row, 7)),
                "hours": hours,
            }
        )
    parsed = [row for row in parsed if row["label"] or row["hours"] is not None]
    if not parsed:
        return None
    stamp = file_date(path) or next((row["date"] for row in parsed if row["date"]), None)
    return {
        "legend": date_label(stamp, path.stem),
        "date_label": date_label(stamp, path.stem),
        "headers": [
            header_text(header, 1, "睡眠\nTidur"),
            header_text(header, 3, "日期\nTanggal"),
            header_text(header, 5, "開始時間\nWaktu mulai"),
            header_text(header, 7, "結束時間\nWaktu berakhir"),
            with_unit(header_text(header, 9, "睡眠時間\nWaktu berolahraga"), "小時"),
        ],
        "rows": parsed,
    }


def parse_cups(value) -> float | None:
    """Parse a cup amount. '1/6' is one sixth of a cup. '5 1/2' is five and a half. '?' is skipped."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if is_missing(value):
        return None
    text = str(value).strip().replace("／", "/").replace("⁄", "/")
    mixed = re.match(r"^(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
    if mixed:
        whole, numerator, denominator = (float(part) for part in mixed.groups())
        if denominator == 0:
            return None
        return whole + numerator / denominator
    fraction = CUP_FRACTION_RE.search(text)
    if fraction:
        denominator = float(fraction.group(2))
        if denominator == 0:
            return None
        return float(fraction.group(1)) / denominator
    return parse_number(text)


def format_hhmm(value) -> str | None:
    text = format_clock(value)
    if text is None:
        return None
    return text[:5]


def _water_columns(header: list) -> tuple[int, int, int]:
    """清潔日期, 清潔時間, 喝水量(杯). Fall back to C, E, G when the titles are absent."""
    date_col = _column_for_title(header, "清潔日期 Tanggal pembersihan")
    time_col = _column_for_title(header, "清潔時間 Waktu bersih-bersih")
    cup_col = _column_for_title(header, "喝水量(杯) Asupan air (cangkir)")
    if date_col is None and time_col is None and cup_col is None:
        return 3, 5, 7
    return date_col or 3, time_col or 5, cup_col or 7


def extract_water(path: Path, worksheet, cached=None) -> dict | None:
    """Rows on 清潔 above the first blank or '?' cup.

    喝水日期 is 清潔日期, 喝水時間 is 清潔時間, and 喝水量(杯) is the cup column.
    Template rows under a blank cup are left out. A missing date uses the file date.
    """
    values_rows = _read_sheet_values(worksheet, cached, 11)
    header = values_rows[0] if values_rows else []
    date_col, time_col, cup_col = _water_columns(header)
    parsed = []
    for row in _rows_until_cup_stop(values_rows, cup_col):
        cups = parse_cups(cell_at(row, cup_col))
        when = parse_date(cell_at(row, date_col)) or file_date(path)
        clock = format_hhmm(cell_at(row, time_col))
        if when is None and clock is None and cups is None:
            continue
        parsed.append({"label": "喝水", "date": when, "time": clock, "cups": cups})
    if not parsed:
        return None
    stamp = file_date(path) or next((row["date"] for row in parsed if row["date"]), None)
    return {
        "legend": date_label(stamp, path.stem),
        "date_label": date_label(stamp, path.stem),
        "headers": list(WATER_HEADERS),
        "rows": parsed,
    }


def read_water(path: Path) -> tuple[dict | None, list[str]]:
    messages = []
    block = None
    workbook = load_workbook(path, data_only=False, read_only=True)
    cached_book = load_workbook(path, data_only=True, read_only=True)
    try:
        if "清潔" not in workbook.sheetnames:
            messages.append(f"{path.name}: sheet 清潔 not found, skipped")
        else:
            cached = cached_book["清潔"] if "清潔" in cached_book.sheetnames else None
            block = extract_water(path, workbook["清潔"], cached)
    finally:
        workbook.close()
        cached_book.close()
    count = 0 if block is None else len(block["rows"])
    messages.append(f"{path.name}: 喝水量 {count}")
    return block, messages


def water_points(blocks: list[dict]) -> list[tuple[str, float]]:
    """One plot point per 喝水量 row, in the same order as the stacked sheet."""
    points = []
    for block in blocks:
        for row in block["rows"]:
            if row.get("cups") is None:
                continue
            label = date_label(row.get("date"), block["date_label"])
            if row.get("time"):
                label = f"{label} {row['time']}"
            points.append((label, float(row["cups"])))
    return points


def add_water_sheets(workbook: Workbook, plots: list[dict], blocks: list[dict], style: PlotStyle) -> None:
    """Add 喝水量 as one table and 喝水量 Plot for those rows. Existing sheets are left as they are."""
    if not blocks:
        return
    if workbook.sheetnames == ["Extract"]:
        sheet = workbook.active
        sheet.title = "喝水量"
    else:
        sheet = workbook.create_sheet("喝水量")
    write_record_sheet(
        sheet,
        _append_blocks(blocks),
        WATER_STRIDE,
        [(0, "label", "text"), (1, "date", "date"), (2, "time", "text"), (3, "cups", "decimal")],
    )
    plot_sheet = workbook.create_sheet("喝水量 Plot")
    plot_sheet._chart_writer = SheetWriter(plot_sheet)
    totals = daily_totals(blocks, "cups")
    counts = daily_counts(blocks, "cups")
    specs = [
        ("喝水量-喝水日期", "喝水日期", "喝水量(杯)", *spread(water_points(blocks), "喝水量(杯)"), True),
        ("全部喝水量-日期", "日期", "全部喝水量", *spread(totals, "全部喝水量"), True),
        ("全部喝水量次數-日期", "日期", "全部喝水量次數", *spread(counts, "全部喝水量次數"), True),
    ]
    anchors = chart_anchors(sum(1 for spec in specs if spec[3] and spec[4]), style.height_cm)
    anchor_index = 0
    for title, xlabel, ylabel, categories, series, rotate in specs:
        if not categories or not series:
            continue
        add_line_chart(plot_sheet, anchors[anchor_index], title, xlabel, ylabel, categories, series, style, rotate)
        plots.append(plot_spec(title, xlabel, ylabel, categories, series, rotate))
        anchor_index += 1


# Column titles on input sheet 清潔, or a sheet of that name, and the output sheet for each.
# kind is how the value is read. write_kind is how write_record_sheet stores it.
CLEAN_COPIES = (
    ("清潔日期 Tanggal pembersihan", "清潔", "date", "date"),
    ("清潔時間 Waktu bersih-bersih", "喝水時間", "time", "text"),
    ("喝水量(杯) Asupan air (cangkir)", "喝水量(杯)", "cups", "decimal"),
)


def _flat_title(value) -> str:
    """One line, with the line break between Chinese and Indonesian turned into a space."""
    if value is None:
        return ""
    return " ".join(str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ").split())


def _column_for_title(header: list, title: str) -> int | None:
    target = _flat_title(title)
    for index, value in enumerate(header, start=1):
        if _flat_title(value) == target:
            return index
    return None


def _sheet_for_title(workbook, title: str):
    target = _flat_title(title)
    for name in workbook.sheetnames:
        if _flat_title(name) == target:
            return workbook[name]
    return None


def _copy_value(value, kind: str):
    if kind == "date":
        return parse_date(value)
    if kind == "time":
        return format_hhmm(value)
    if kind == "cups":
        return parse_cups(value)
    return None


def _copy_block(path: Path, title: str, kind: str, values: list) -> dict | None:
    parsed = [_copy_value(value, kind) for value in values]
    if not any(value is not None for value in parsed):
        return None
    return _record_block(path, [{"value": value} for value in parsed], [_flat_title(title)])


def _rows_until_cup_stop(values_rows: list[list], stop_column: int) -> list[list]:
    """Data rows above the first '?' or blank cup cell. A '?' in column A also stops."""
    kept = []
    for index, values in enumerate(values_rows, start=1):
        if index == 1:
            continue
        if is_question(cell_at(values, 1)):
            break
        if _page_cell_stops(cell_at(values, stop_column)):
            break
        kept.append(values)
    return kept


def _column_a_values(worksheet, cached) -> list:
    """Column A below the header, above the first '?' or blank cell."""
    values_rows = _read_sheet_values(worksheet, cached, 1)
    values = []
    for index, row in enumerate(values_rows, start=1):
        if index == 1:
            continue
        cell = cell_at(row, 1)
        if _page_cell_stops(cell):
            break
        values.append(cell)
    return values


def read_clean_copies(path: Path) -> tuple[dict | None, dict | None, dict | None, list[str]]:
    """Copy 清潔日期, 清潔時間, and 喝水量(杯) into three blocks.

    A sheet whose name is the column title is copied from column A.
    Otherwise the matching column on sheet 清潔 is copied. Those three columns
    share one row window: the rows above the first '?' or blank cup cell.
    When 清潔 has no such titles, columns C, E, and G are used.
    """
    messages = []
    workbook = load_workbook(path, data_only=False, read_only=True)
    cached_book = load_workbook(path, data_only=True, read_only=True)
    try:
        blocks: list[dict | None] = [None, None, None]
        named = []
        for spec in CLEAN_COPIES:
            sheet = _sheet_for_title(workbook, spec[0])
            cached = None
            if sheet is not None and sheet.title in cached_book.sheetnames:
                cached = cached_book[sheet.title]
            named.append((sheet, cached))

        value_rows: list[list] = []
        columns: list[int | None] = [None, None, None]
        if "清潔" in workbook.sheetnames:
            cached = cached_book["清潔"] if "清潔" in cached_book.sheetnames else None
            values_rows = _read_sheet_values(workbook["清潔"], cached, 11)
            header = values_rows[0] if values_rows else []
            columns = [_column_for_title(header, spec[0]) for spec in CLEAN_COPIES]
            if not any(column is not None for column in columns):
                columns = [3, 5, 7]
            stop_column = columns[2]
            if stop_column is not None:
                value_rows = _rows_until_cup_stop(values_rows, stop_column)

        seen = "清潔" in workbook.sheetnames or any(sheet is not None for sheet, _cached in named)
        for index, spec in enumerate(CLEAN_COPIES):
            title, _sheet_name, kind, _write_kind = spec
            sheet, cached = named[index]
            if sheet is not None:
                blocks[index] = _copy_block(path, title, kind, _column_a_values(sheet, cached))
                continue
            column = columns[index]
            if column is None:
                continue
            blocks[index] = _copy_block(
                path,
                title,
                kind,
                [cell_at(row, column) for row in value_rows],
            )
        if not seen:
            messages.append(f"{path.name}: sheet 清潔 not found, skipped")
        else:
            messages.append(
                f"{path.name}: 清潔 {0 if blocks[0] is None else len(blocks[0]['rows'])}, "
                f"喝水時間 {0 if blocks[1] is None else len(blocks[1]['rows'])}, "
                f"喝水量(杯) {0 if blocks[2] is None else len(blocks[2]['rows'])}"
            )
    finally:
        workbook.close()
        cached_book.close()
    return blocks[0], blocks[1], blocks[2], messages


def _append_blocks(blocks: list[dict]) -> list[dict]:
    """Stack every file's rows in the order the files were read."""
    if not blocks:
        return []
    rows = []
    for block in blocks:
        rows.extend(block["rows"])
    first = blocks[0]
    return [{
        "legend": first["legend"],
        "date_label": first["date_label"],
        "headers": first["headers"],
        "rows": rows,
    }]


def add_clean_copy_sheets(
    workbook: Workbook,
    dates: list[dict],
    times: list[dict],
    cups: list[dict],
) -> None:
    """Add 清潔, 喝水時間, and 喝水量(杯). Later files are appended under the first. No plot."""
    fields = (
        ("清潔", dates, "date"),
        ("喝水時間", times, "text"),
        ("喝水量(杯)", cups, "decimal"),
    )
    for name, blocks, kind in fields:
        stacked = _append_blocks(blocks)
        if not stacked:
            continue
        sheet = _take_output_sheet(workbook, name)
        write_record_sheet(sheet, stacked, 1, [(0, "value", kind)])


def plain_text(value) -> str | None:
    if is_missing(value):
        return None
    return str(value).strip()


def _record_block(path: Path, rows: list[dict], headers: list[str]) -> dict | None:
    if not rows:
        return None
    stamp = file_date(path) or next((row["date"] for row in rows if row["date"]), None)
    return {
        "legend": date_label(stamp, path.stem),
        "date_label": date_label(stamp, path.stem),
        "headers": headers,
        "rows": rows,
    }


def extract_excretion(path: Path, worksheet, cached=None) -> dict | None:
    """排洩 rows above '?'. Date and time are kept; '?' measurements are left blank.

    A row is kept when any measurement is present. Template rows that only
    repeat the default date and time are skipped.
    """
    header, rows = load_table(worksheet, 13, cached)
    parsed = []
    for row in rows:
        record = {
            "date": parse_date(cell_at(row, 1)) or file_date(path),
            "time": format_hhmm(cell_at(row, 3)),
            "stool": plain_text(cell_at(row, 5)),
            "urine": plain_text(cell_at(row, 7)),
            "color": plain_text(cell_at(row, 9)),
            "diarrhea": plain_text(cell_at(row, 11)),
            "pain": plain_text(cell_at(row, 13)),
        }
        if not any(record[key] is not None for key in ("stool", "urine", "color", "diarrhea", "pain")):
            continue
        parsed.append(record)
    return _record_block(
        path,
        parsed,
        [
            header_text(header, 1, "排洩日期\nTanggal ekskres"),
            header_text(header, 3, "排洩時間\nWaktu ekskresi"),
            header_text(header, 5, "大便量\nVolume feses)"),
            header_text(header, 7, "小便量\nVolume urin"),
            header_text(header, 9, "小便顏色改變\nPerubahan warna urin"),
            header_text(header, 11, "拉肚子\nDiare"),
            header_text(header, 13, "肚子痛\nSakit perut"),
        ],
    )


def extract_body(path: Path, worksheet, cached=None) -> dict | None:
    """身體情況 rows above '?'. '?' observations are left blank and do not keep a row."""
    header, rows = load_table(worksheet, 11, cached)
    parsed = []
    for row in rows:
        record = {
            "date": parse_date(cell_at(row, 1)) or file_date(path),
            "sleep": plain_text(cell_at(row, 3)),
            "cry": plain_text(cell_at(row, 5)),
            "mood": plain_text(cell_at(row, 7)),
            "walk": plain_text(cell_at(row, 9)),
            "voice": plain_text(cell_at(row, 11)),
        }
        if not any(record[key] is not None for key in ("sleep", "cry", "mood", "walk", "voice")):
            continue
        parsed.append(record)
    return _record_block(
        path,
        parsed,
        [
            header_text(header, 1, "身體情況日期\nTanggal kondisi fisik"),
            header_text(header, 3, "睡眠\nTidur"),
            header_text(header, 5, "哀叫\nRatapan"),
            header_text(header, 7, "情緒\nemosi"),
            header_text(header, 9, "走路力量\nkekuatan berjalan"),
            header_text(header, 11, "說話聲音\nsuara"),
        ],
    )


def read_condition(path: Path) -> tuple[dict | None, dict | None, list[str]]:
    """Read 排洩 and 身體情況. Missing sheets are skipped."""
    messages = []
    excretion = body = None
    workbook = load_workbook(path, data_only=False, read_only=True)
    cached_book = load_workbook(path, data_only=True, read_only=True)
    try:
        sheets = set(workbook.sheetnames)

        def cached_sheet(name: str):
            if name in cached_book.sheetnames:
                return cached_book[name]
            return None

        if "排洩" in sheets:
            excretion = extract_excretion(path, workbook["排洩"], cached_sheet("排洩"))
        else:
            messages.append(f"{path.name}: sheet 排洩 not found, skipped")
        if "身體情況" in sheets:
            body = extract_body(path, workbook["身體情況"], cached_sheet("身體情況"))
        else:
            messages.append(f"{path.name}: sheet 身體情況 not found, skipped")
    finally:
        workbook.close()
        cached_book.close()
    messages.append(
        f"{path.name}: 排洩 {0 if excretion is None else len(excretion['rows'])}, "
        f"身體情況 {0 if body is None else len(body['rows'])}"
    )
    return excretion, body, messages


def _take_output_sheet(workbook: Workbook, name: str) -> Worksheet:
    if workbook.sheetnames == ["Extract"]:
        sheet = workbook.active
        sheet.title = name
        return sheet
    return workbook.create_sheet(name)


def _one_line_header(text: str) -> str:
    """Turn a two-line Chinese / Indonesian title into one line."""
    return text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")


def _write_condition_sheet(worksheet: Worksheet, blocks: list[dict], fields: list[tuple]) -> None:
    """One header row, then every file's rows stacked in those columns. No plot data."""
    worksheet.freeze_panes = "A2"
    headers = [_one_line_header(blocks[0]["headers"][slot]) for slot, _key, _kind in fields]
    fill = PatternFill("solid", fgColor=HEADER_COLORS[0])
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    for index, title in enumerate(headers, start=1):
        cell = worksheet.cell(1, index, title)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=False, vertical="center", horizontal="center")
        cell.border = THIN
        width = sum(2 if ord(char) > 127 else 1 for char in title) + 2
        worksheet.column_dimensions[get_column_letter(index)].width = min(max(width, 14), 48)
    row_index = 2
    body_font = Font(name=FONT_NAME, size=11)
    for block in blocks:
        for record in block["rows"]:
            for index, (_slot, key, kind) in enumerate(fields, start=1):
                value = record[key]
                if value is None:
                    continue
                if kind == "date":
                    cell = worksheet.cell(row_index, index, date_label(value, ""))
                else:
                    cell = worksheet.cell(row_index, index, value)
                cell.font = body_font
                cell.border = THIN
                cell.alignment = Alignment(vertical="center", wrap_text=False)
            row_index += 1


def add_condition_sheets(workbook: Workbook, excretion: list[dict], body: list[dict]) -> None:
    """Add 排洩 and 身體情況 as one table each. No plot sheet and no chart is created."""
    if excretion:
        sheet = _take_output_sheet(workbook, "排洩")
        _write_condition_sheet(
            sheet,
            excretion,
            [
                (0, "date", "date"),
                (1, "time", "text"),
                (2, "stool", "text"),
                (3, "urine", "text"),
                (4, "color", "text"),
                (5, "diarrhea", "text"),
                (6, "pain", "text"),
            ],
        )
    if body:
        sheet = _take_output_sheet(workbook, "身體情況")
        _write_condition_sheet(
            sheet,
            body,
            [
                (0, "date", "date"),
                (1, "sleep", "text"),
                (2, "cry", "text"),
                (3, "mood", "text"),
                (4, "walk", "text"),
                (5, "voice", "text"),
            ],
        )


def read_workbook(path: Path) -> tuple[dict | None, dict | None, dict | None, list[str]]:
    messages = []
    exercise = bp = sleep = None
    workbook = load_workbook(path, data_only=False, read_only=True)
    cached_book = load_workbook(path, data_only=True, read_only=True)
    try:
        sheets = set(workbook.sheetnames)

        def cached_sheet(name: str):
            if name in cached_book.sheetnames:
                return cached_book[name]
            return None

        if "運動" in sheets:
            exercise = extract_exercise(path, workbook["運動"], cached_sheet("運動"))
        else:
            messages.append(f"{path.name}: sheet 運動 not found, skipped")
        if "血壓" in sheets:
            bp = extract_bp(path, workbook["血壓"], cached_sheet("血壓"))
        else:
            messages.append(f"{path.name}: sheet 血壓 not found, skipped")
        if "睡眠" in sheets:
            sleep = extract_sleep(path, workbook["睡眠"], cached_sheet("睡眠"))
        else:
            messages.append(f"{path.name}: sheet 睡眠 not found, skipped")
    finally:
        workbook.close()
        cached_book.close()
    counts = (
        f"運動 {0 if exercise is None else len(exercise['rows'])}, "
        f"血壓 {0 if bp is None else len(bp['rows'])}, "
        f"睡眠 {0 if sleep is None else len(sleep['rows'])}"
    )
    messages.insert(0, f"{path.name}: {counts}")
    return exercise, bp, sleep, messages


def activity_matrix(blocks: list[dict], key: str) -> tuple[list[int], list[tuple[str, list]]]:
    included = []
    for block in blocks:
        values = [row[key] for row in block["rows"]]
        if any(value is not None for value in values):
            included.append((block["legend"], values))
    if not included:
        return [], []
    width = max(len(values) for _, values in included)
    xs = list(range(1, width + 1))
    columns = []
    for legend, values in included:
        padded = list(values) + [None] * (width - len(values))
        columns.append((legend, padded))
    return xs, columns


def daily_totals(blocks: list[dict], key: str) -> list[tuple[str, float]]:
    totals: dict[str, float] = {}
    for block in blocks:
        for row in block["rows"]:
            value = row.get(key)
            if value is None:
                continue
            label = date_label(row.get("date"), block["date_label"])
            totals[label] = totals.get(label, 0.0) + float(value)
    return [(label, totals[label]) for label in sorted(totals)]


def daily_counts(blocks: list[dict], key: str) -> list[tuple[str, int]]:
    counts: dict[str, int] = {}
    for block in blocks:
        for row in block["rows"]:
            if row.get(key) is None:
                continue
            label = date_label(row.get("date"), block["date_label"])
            counts[label] = counts.get(label, 0) + 1
    return [(label, counts[label]) for label in sorted(counts)]


def single_series(points: list[tuple], name: str) -> tuple[list[str], list[tuple[str, list]]]:
    if not points:
        return [], []
    return [label for label, _ in points], [(name, [value for _, value in points])]


def _draw_font(size_pt: float) -> CharacterProperties:
    size = int(round(size_pt * 100))
    face = DrawFont(typeface=FONT_NAME)
    return CharacterProperties(sz=size, latin=face, ea=face)


def make_title(text: str, size_pt: float) -> Title:
    props = _draw_font(size_pt)
    paragraphs = []
    for line in str(text).split("\n"):
        run = RegularTextRun(t=line)
        run.rPr = props
        paragraphs.append(Paragraph(r=[run], pPr=ParagraphProperties(defRPr=props)))
    title = Title()
    title.tx.rich.paragraphs = paragraphs
    title.overlay = False
    return title


def text_body(size_pt: float, rotation: int | None = None) -> ChartRichText:
    props = _draw_font(size_pt)
    body = RichTextProperties(rot=rotation) if rotation else RichTextProperties()
    return ChartRichText(
        bodyPr=body,
        p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)],
    )


def paint_series(chart: LineChart) -> None:
    for index, series in enumerate(chart.series):
        color = PALETTE[index % len(PALETTE)]
        series.graphicalProperties.line = LineProperties(solidFill=color, w=22000)
        series.marker = Marker(symbol="circle", size=7)
        series.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
        series.smooth = False


def style_chart(chart: LineChart, title: str, xlabel: str, ylabel: str, style: PlotStyle, rotate_x: bool) -> None:
    chart.title = make_title(title, style.title_size)
    chart.x_axis.title = make_title(xlabel, style.xlabel_size)
    chart.y_axis.title = make_title(ylabel, style.ylabel_size)
    rotation = -2700000 if rotate_x else None
    chart.x_axis.txPr = text_body(style.tick_size, rotation)
    chart.y_axis.txPr = text_body(style.tick_size)
    if chart.legend is not None:
        chart.legend.txPr = text_body(style.legend_size)
        chart.legend.position = "b"
        chart.legend.overlay = False
    chart.x_axis.crosses = "min"
    chart.y_axis.crosses = "min"
    chart.x_axis.tickLblPos = "low"
    chart.display_blanks = "gap"
    chart.style = 10
    chart.width = style.width_cm
    chart.height = style.height_cm


class SheetWriter:
    def __init__(self, worksheet: Worksheet):
        self.worksheet = worksheet
        self.row = 1
        self.col = 40

    def write_table(self, category_header: str, series: list[tuple[str, list]], categories: list) -> tuple[int, int, int]:
        start = self.row
        self.worksheet.cell(start, self.col, category_header)
        for offset, (name, _values) in enumerate(series, start=1):
            self.worksheet.cell(start, self.col + offset, name)
        for row_offset, category in enumerate(categories):
            self.worksheet.cell(start + 1 + row_offset, self.col, category)
            for series_offset, (_name, values) in enumerate(series, start=1):
                value = values[row_offset] if row_offset < len(values) else None
                if value is not None:
                    cell = self.worksheet.cell(start + 1 + row_offset, self.col + series_offset, as_number(value))
                    cell.number_format = "0.00" if isinstance(value, float) else "0"
        end_row = start + len(categories)
        self.row = end_row + 3
        return start, end_row, self.col


def add_line_chart(
    worksheet: Worksheet,
    anchor: str,
    title: str,
    xlabel: str,
    ylabel: str,
    categories: list,
    series: list[tuple[str, list]],
    style: PlotStyle,
    rotate_x: bool,
) -> LineChart | None:
    if not categories or not series:
        return None
    writer = getattr(worksheet, "_chart_writer")
    header_row, last_row, origin = writer.write_table(xlabel, series, categories)
    chart = LineChart()
    data = Reference(
        worksheet,
        min_col=origin + 1,
        max_col=origin + len(series),
        min_row=header_row,
        max_row=last_row,
    )
    cats = Reference(worksheet, min_col=origin, min_row=header_row + 1, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    paint_series(chart)
    style_chart(chart, title, xlabel, ylabel, style, rotate_x)
    worksheet.add_chart(chart, anchor)
    return chart


def chart_anchors(count: int, height_cm: float) -> list[str]:
    span = max(int(math.ceil(height_cm / 0.53)) + 2, 8)
    return [f"A{2 + index * span}" for index in range(count)]


def write_record_sheet(worksheet: Worksheet, blocks: list[dict], stride: int, fields: list[tuple[str, str]]) -> None:
    """Write one column group per input file. fields are (header index, row key)."""
    worksheet.sheet_view.showGridLines = False
    worksheet.row_dimensions[1].height = 48
    worksheet.freeze_panes = "A2"
    for block_index, block in enumerate(blocks):
        fill = PatternFill("solid", fgColor=HEADER_COLORS[block_index % len(HEADER_COLORS)])
        font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
        base = block_index * stride
        for field_index, (header_slot, _key, kind) in enumerate(fields):
            column = base + field_index + 1
            cell = worksheet.cell(1, column, block["headers"][header_slot])
            cell.font = font
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            cell.border = THIN
            width = 36 if kind == "text" else 20
            worksheet.column_dimensions[get_column_letter(column)].width = width
        for row_index, record in enumerate(block["rows"], start=2):
            for field_index, (_header_slot, key, kind) in enumerate(fields):
                value = record[key]
                if value is None:
                    continue
                column = base + field_index + 1
                if kind == "date":
                    cell = worksheet.cell(row_index, column, value)
                    cell.number_format = "yyyy-mm-dd"
                elif kind == "text":
                    cell = worksheet.cell(row_index, column, value)
                else:
                    cell = worksheet.cell(row_index, column, as_number(value))
                    cell.number_format = "0.00" if kind == "decimal" else "0"
                cell.font = Font(name=FONT_NAME, size=11)
                cell.border = THIN
                cell.alignment = Alignment(vertical="center", wrap_text=(kind == "text"))


def build_workbook(
    exercise: list[dict],
    bp: list[dict],
    sleep: list[dict],
    style: PlotStyle,
) -> tuple[Workbook, list[dict]]:
    """Return the workbook and a list of plot specs for Plotly."""
    workbook = Workbook()
    default = workbook.active
    created = False
    plots: list[dict] = []

    def take_sheet(name: str) -> Worksheet:
        nonlocal created, default
        if not created:
            default.title = name
            created = True
            return default
        return workbook.create_sheet(name)

    if exercise:
        sheet = take_sheet("運動")
        write_record_sheet(
            sheet,
            exercise,
            EXERCISE_STRIDE,
            [(0, "sport", "text"), (1, "date", "date"), (2, "minutes", "decimal"), (3, "meters", "decimal")],
        )
        plot_sheet = take_sheet("運動 Plot")
        plot_sheet._chart_writer = SheetWriter(plot_sheet)
        time_x, time_series = activity_matrix(exercise, "minutes")
        dist_x, dist_series = activity_matrix(exercise, "meters")
        total_time = daily_totals(exercise, "minutes")
        total_dist = daily_totals(exercise, "meters")
        specs = [
            ("運動時間", "1:n", "運動時間", time_x, time_series, False),
            ("運動距離", "1:n", "運動距離", dist_x, dist_series, False),
            ("全部運動時間-運動日期", "運動日期", "全部運動時間", *spread(total_time, "全部運動時間"), True),
            ("全部運動距離-運動日期", "運動日期", "全部運動距離", *spread(total_dist, "全部運動距離"), True),
        ]
        anchors = chart_anchors(sum(1 for spec in specs if spec[3] and spec[4]), style.height_cm)
        anchor_index = 0
        for title, xlabel, ylabel, categories, series, rotate in specs:
            if not categories or not series:
                continue
            add_line_chart(plot_sheet, anchors[anchor_index], title, xlabel, ylabel, categories, series, style, rotate)
            plots.append(plot_spec(title, xlabel, ylabel, categories, series, rotate))
            anchor_index += 1

    if bp:
        sheet = take_sheet("血壓")
        write_record_sheet(
            sheet,
            bp,
            BP_STRIDE,
            [
                (0, "label", "text"),
                (1, "date", "date"),
                (2, "high", "int"),
                (3, "low", "int"),
                (4, "heart", "int"),
            ],
        )
        plot_sheet = take_sheet("血壓 Plot")
        plot_sheet._chart_writer = SheetWriter(plot_sheet)
        specs = []
        for key, title, ylabel in (
            ("high", "高血壓-量血壓日期", "高血壓"),
            ("low", "低血壓-量血壓日期", "低血壓"),
            ("heart", "心跳-量血壓日期", "心跳"),
        ):
            points = []
            for block in bp:
                for row in block["rows"]:
                    if row[key] is not None:
                        points.append((row["when"], row[key]))
            points.sort(key=lambda item: item[0])
            categories, series = single_series(points, ylabel)
            specs.append((title, "量血壓日期", ylabel, categories, series, True))
        anchors = chart_anchors(sum(1 for spec in specs if spec[3]), style.height_cm)
        anchor_index = 0
        for title, xlabel, ylabel, categories, series, rotate in specs:
            if not categories:
                continue
            add_line_chart(plot_sheet, anchors[anchor_index], title, xlabel, ylabel, categories, series, style, rotate)
            plots.append(plot_spec(title, xlabel, ylabel, categories, series, rotate))
            anchor_index += 1

    if sleep:
        sheet = take_sheet("睡眠")
        write_record_sheet(
            sheet,
            sleep,
            SLEEP_STRIDE,
            [
                (0, "label", "text"),
                (1, "date", "date"),
                (2, "start", "text"),
                (3, "end", "text"),
                (4, "hours", "decimal"),
            ],
        )
        plot_sheet = take_sheet("睡眠 Plot")
        plot_sheet._chart_writer = SheetWriter(plot_sheet)
        total_hours = daily_totals(sleep, "hours")
        counts = daily_counts(sleep, "hours")
        specs = [
            ("全部睡眠時間-日期", "日期", "全部睡眠時間", *spread(total_hours, "全部睡眠時間"), True),
            ("全部睡眠次數-日期", "日期", "全部睡眠次數", *spread(counts, "全部睡眠次數"), True),
        ]
        anchors = chart_anchors(sum(1 for spec in specs if spec[3] and spec[4]), style.height_cm)
        anchor_index = 0
        for title, xlabel, ylabel, categories, series, rotate in specs:
            if not categories or not series:
                continue
            add_line_chart(plot_sheet, anchors[anchor_index], title, xlabel, ylabel, categories, series, style, rotate)
            plots.append(plot_spec(title, xlabel, ylabel, categories, series, rotate))
            anchor_index += 1

    if not created:
        default.title = "Extract"
    return workbook, plots


def spread(points: list[tuple], name: str):
    categories, series = single_series(points, name)
    return categories, series


def plot_spec(title, xlabel, ylabel, categories, series, rotate_x) -> dict:
    return {
        "title": title,
        "xlabel": xlabel,
        "ylabel": ylabel,
        "categories": [str(item) for item in categories],
        "series": [(name, [None if value is None else float(value) for value in values]) for name, values in series],
        "rotate_x": rotate_x,
    }


def write_plotly(path: Path, plots: list[dict], style: PlotStyle, intro: str) -> None:
    import plotly.graph_objects as go

    parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'>",
        "<title>Extract plots</title>",
        f"<style>body{{font-family:'{FONT_NAME}',Segoe UI,sans-serif;margin:24px;color:#1c2833;}}",
        "h1{font-size:22px;} h2{margin-top:28px;font-size:18px;} p{max-width:920px;line-height:1.45;}</style>",
        "</head><body>",
        "<h1>Extract plots</h1>",
        f"<p>{html.escape(intro)}</p>",
    ]
    include_js = True
    for spec in plots:
        figure = go.Figure()
        for index, (name, values) in enumerate(spec["series"]):
            color = "#" + PALETTE[index % len(PALETTE)]
            figure.add_trace(
                go.Scatter(
                    x=spec["categories"],
                    y=values,
                    mode="lines+markers",
                    name=name,
                    connectgaps=False,
                    line={"color": color, "width": 2},
                    marker={"size": 8, "color": color},
                )
            )
        figure.update_layout(
            title={"text": spec["title"], "font": {"size": style.title_size, "family": FONT_NAME}},
            legend={"font": {"size": style.legend_size, "family": FONT_NAME}, "orientation": "h", "y": -0.25},
            width=style.width_px,
            height=style.height_px,
            margin={"l": 80, "r": 40, "t": 70, "b": 90},
            template="plotly_white",
            paper_bgcolor="white",
            plot_bgcolor="white",
        )
        figure.update_xaxes(
            title={"text": spec["xlabel"], "font": {"size": style.xlabel_size, "family": FONT_NAME}},
            tickfont={"size": style.tick_size, "family": FONT_NAME},
            tickangle=-40 if spec["rotate_x"] else 0,
            type="category",
        )
        figure.update_yaxes(
            title={"text": spec["ylabel"], "font": {"size": style.ylabel_size, "family": FONT_NAME}},
            tickfont={"size": style.tick_size, "family": FONT_NAME},
        )
        parts.append(f"<h2>{html.escape(spec['title'])}</h2>")
        parts.append(figure.to_html(full_html=False, include_plotlyjs=include_js))
        include_js = False
    if not plots:
        parts.append("<p>No plots. The selected files had no usable rows.</p>")
    parts.append("</body></html>")
    path.write_text("".join(parts), encoding="utf-8")


def validate_style(style: PlotStyle) -> None:
    checks = (
        ("Chart width (cm)", style.width_cm, 5, 60),
        ("Chart height (cm)", style.height_cm, 5, 40),
        ("Title font size", style.title_size, 6, 72),
        ("X label font size", style.xlabel_size, 6, 72),
        ("Y label font size", style.ylabel_size, 6, 72),
        ("Legend font size", style.legend_size, 6, 72),
        ("Tick font size", style.tick_size, 6, 72),
    )
    for name, value, low, high in checks:
        if not math.isfinite(value) or not (low <= value <= high):
            raise ValueError(f"{name} must be between {low} and {high}.")


def process_folder(folder, style: PlotStyle | None = None, output_dir=None) -> ProcessResult:
    folder = Path(folder)
    if not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")
    style = style or PlotStyle()
    validate_style(style)
    files = list_dated_workbooks(folder)
    if not files:
        raise FileNotFoundError(f"No Excel files named YYYYMMDD.xlsx in {folder}")

    exercise: list[dict] = []
    bp: list[dict] = []
    sleep: list[dict] = []
    water: list[dict] = []
    excretion: list[dict] = []
    body: list[dict] = []
    clean_dates: list[dict] = []
    clean_times: list[dict] = []
    clean_cups: list[dict] = []
    messages = [f"Folder: {folder}", f"Files: {', '.join(path.name for path in files)}"]
    for path in files:
        try:
            one_exercise, one_bp, one_sleep, notes = read_workbook(path)
            one_water, water_notes = read_water(path)
            one_excretion, one_body, condition_notes = read_condition(path)
            one_date, one_time, one_cups, copy_notes = read_clean_copies(path)
        except Exception as exc:
            messages.append(f"{path.name}: skipped ({exc})")
            continue
        messages.extend(notes)
        messages.extend(water_notes)
        messages.extend(condition_notes)
        messages.extend(copy_notes)
        if one_exercise:
            exercise.append(one_exercise)
        if one_bp:
            bp.append(one_bp)
        if one_sleep:
            sleep.append(one_sleep)
        if one_water:
            water.append(one_water)
        if one_excretion:
            excretion.append(one_excretion)
        if one_body:
            body.append(one_body)
        if one_date:
            clean_dates.append(one_date)
        if one_time:
            clean_times.append(one_time)
        if one_cups:
            clean_cups.append(one_cups)

    if not exercise and not bp and not sleep and not water and not excretion and not body and not clean_dates and not clean_times and not clean_cups:
        raise ValueError("No 運動, 血壓, 睡眠, 喝水量, 排洩, or 身體情況 data was found before the '?' rows.")

    legend_for(exercise)
    legend_for(bp)
    legend_for(sleep)
    legend_for(water)
    legend_for(excretion)
    legend_for(body)
    workbook, plots = build_workbook(exercise, bp, sleep, style)
    add_water_sheets(workbook, plots, water, style)
    add_condition_sheets(workbook, excretion, body)
    add_clean_copy_sheets(workbook, clean_dates, clean_times, clean_cups)
    add_food_sheet(workbook, files)
    destination = Path(output_dir) if output_dir else folder / "output"
    destination.mkdir(parents=True, exist_ok=True)
    excel_path = destination / "Extract.xlsx"
    plot_path = destination / "Extract_plots.html"
    try:
        workbook.save(excel_path)
    except PermissionError as exc:
        raise PermissionError(f"Cannot write {excel_path}. Close the file if it is open in Excel.") from exc
    intro = (
        "運動時間 is minutes. 運動距離 is meters. 睡眠時間 is hours. 喝水量 is cups. "
        "Rows at and below '?' are omitted. Missing values are left blank and are not drawn."
    )
    food_plot = food_plot_from_sheet(workbook)
    if food_plot is not None:
        plots.append(food_plot)
    write_plotly(plot_path, plots, style, intro)
    messages.append("運動時間 is stored in minutes. 運動距離 is stored in meters. 睡眠時間 is stored in hours. 喝水量 is stored in cups.")
    messages.append(f"Saved {excel_path}")
    messages.append(f"Saved {plot_path}")
    page_notes, pdf_path = export_page_files(files, destination, excel_path.stem)
    messages.extend(page_notes)
    return ProcessResult(excel_path, plot_path, messages, pdf_path)


# Landscape A4 pages. Each page repeats the column letters and the header.
# A later page is written only when the rows do not fit on the page above.
PAGE_W = 1600
PAGE_H = 1132
PAGE_LEFT = 28
PAGE_RIGHT = 1572
PAGE_TOP = 63
PAGE_BOTTOM = 1104
PAGE_INDEX_W = 58
PAGE_GAP = 22
PAGE_BG = (251, 255, 253)
PAGE_HEADER_BG = (215, 236, 230)
PAGE_INDEX_HEADER_BG = (231, 243, 239)
PAGE_INDEX_BG = (247, 251, 250)
PAGE_GRID = (155, 184, 178)
PAGE_INK = (16, 36, 40)
PAGE_LINE = 20
PAGE_LETTER_H = 26
PAGE_MIN_HEADER_H = 52
PAGE_MIN_ROW_H = 33
_PAGE_FONTS = (
    Path(r"C:\Windows\Fonts\msjh.ttc"),
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\mingliu.ttc"),
    Path(r"C:\Windows\Fonts\simsun.ttc"),
    Path(r"C:\Windows\Fonts\arial.ttf"),
)

# stem is the file name. title is the label drawn on the page.
# parts are (sheet, excel columns). 飲食 keeps columns A, E, and G only.
# 睡眠 columns follow the sample page: 日期, 開始時間, 結束時間, 睡眠時間.
PAGE_SPECS = (
    {"stem": "運動", "title": "運動", "parts": (("運動", (3, 5, 11)),)},
    {"stem": "清潔", "title": "清潔", "parts": (("清潔", (2, 4, 6)),)},
    {"stem": "飲食", "title": "飲食", "parts": (("飲食", (1, 5, 7)),)},
    {"stem": "睡眠", "title": "睡眠", "parts": (("睡眠", (3, 5, 7, 9)),)},
    {"stem": "排洩1", "title": "排洩", "parts": (("排洩", (1, 5, 7)),)},
    {"stem": "排洩2", "title": "排洩", "parts": (("排洩", (1, 9, 11, 13)),)},
    {"stem": "血壓1", "title": "血壓", "parts": (("血壓", (1, 3, 5, 6)),)},
    {"stem": "血壓2", "title": "血壓", "parts": (("血壓", (1, 3, 7)),)},
    {"stem": "身體情況1", "title": "身體情況", "parts": (("身體情況", (1, 3, 5, 7)),)},
    {"stem": "身體情況2", "title": "身體情況", "parts": (("身體情況", (1, 9, 11)),)},
)
# A page keeps the rows above the first '?' or blank cell in this column.
PAGE_STOP_COLUMN = {
    "運動": 11,
    "清潔": 6,
    "飲食": 7,
    "睡眠": 9,
    "排洩": 5,
    "血壓": 5,
    "身體情況": 3,
}


def _page_font(size: int) -> ImageFont.FreeTypeFont:
    for path in _PAGE_FONTS:
        if path.is_file():
            return ImageFont.truetype(str(path), size=size)
    raise FileNotFoundError("No font found for the landscape pages. Install Microsoft JhengHei.")


def _page_text(value) -> str:
    """Cell text for a page. Blank and '?' stay blank. Dates use YYYY-MM-DD."""
    if isinstance(value, str) and value.startswith("="):
        return ""
    if is_missing(value):
        return ""
    if isinstance(value, dt.datetime):
        if value.hour == value.minute == value.second == value.microsecond == 0:
            return value.strftime("%Y-%m-%d")
        if value.second == 0 and value.microsecond == 0:
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.time):
        if value.second == 0 and value.microsecond == 0:
            return f"{value.hour:02d}:{value.minute:02d}"
        return f"{value.hour:02d}:{value.minute:02d}:{value.second:02d}"
    if isinstance(value, float) and math.isfinite(value):
        return str(as_number(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def _page_header(value) -> str:
    text = _page_text(value)
    if not text:
        return ""
    return "\n".join(line.strip() for line in text.split("\n"))


def _page_cell_stops(value) -> bool:
    """True when this cell is '?' or empty. Rows above it are the page rows."""
    if value is None:
        return True
    if isinstance(value, str):
        text = value.strip()
        return text == "" or text in QUESTION_MARKS or set(text) <= QUESTION_MARKS
    return False


def _prefer_page_value(value, cached):
    if isinstance(value, str) and value.startswith("="):
        return cached
    return value


def _read_sheet_values(worksheet, cached, max_col: int) -> list[list]:
    """One pass over a sheet. A read-only worksheet cannot be scanned twice."""
    limit = min(worksheet.max_row or 1, 5000)
    cached_rows = []
    if cached is not None:
        cached_rows = list(cached.iter_rows(min_row=1, max_row=limit, max_col=max_col, values_only=True))
    values_rows = []
    for index, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=limit, max_col=max_col, values_only=True),
        start=1,
    ):
        values = list(row)
        if cached_rows and index - 1 < len(cached_rows):
            cached_values = list(cached_rows[index - 1])
            values = [
                _prefer_page_value(value, cached_values[offset] if offset < len(cached_values) else None)
                for offset, value in enumerate(values)
            ]
        values_rows.append(values)
    return values_rows


def extract_page_table(
    values_rows: list[list],
    columns: tuple[int, ...],
    stop_column: int,
) -> tuple[list[str], list[list[str]]]:
    """Header and data cells above the first '?' or blank in stop_column.

    A '?' in column A ends the page as well. Rows at and below that cell are
    left off the page, including later rows that have values again. A calculated
    duration such as 00:00:00 does not keep those rows.
    """
    header = [""] * len(columns)
    rows: list[list[str]] = []
    for index, values in enumerate(values_rows, start=1):
        if index == 1:
            header = [_page_header(cell_at(values, column)) for column in columns]
            continue
        if is_question(cell_at(values, 1)) or _page_cell_stops(cell_at(values, stop_column)):
            break
        rows.append([_page_text(cell_at(values, column)) for column in columns])
    return header, rows


def collect_page_groups(files: list[Path]) -> list[dict]:
    """Read every dated workbook and stack each page table in that file order."""
    grouped = [
        {
            "stem": spec["stem"],
            "title": spec["title"],
            "tables": [
                {
                    "sheet": part[0],
                    "columns": part[1],
                    "letters": [get_column_letter(column) for column in part[1]],
                    "header": [],
                    "rows": [],
                }
                for part in spec["parts"]
            ],
        }
        for spec in PAGE_SPECS
    ]
    needed: dict[str, list[tuple[int, int]]] = {}
    for group_index, group in enumerate(grouped):
        for table_index, table in enumerate(group["tables"]):
            needed.setdefault(table["sheet"], []).append((group_index, table_index))

    for path in files:
        workbook = load_workbook(path, data_only=False, read_only=True)
        cached_book = load_workbook(path, data_only=True, read_only=True)
        try:
            for sheet_name, targets in needed.items():
                if sheet_name not in workbook.sheetnames:
                    continue
                cached = cached_book[sheet_name] if sheet_name in cached_book.sheetnames else None
                stop_column = PAGE_STOP_COLUMN[sheet_name]
                max_col = stop_column
                for group_index, table_index in targets:
                    max_col = max(max_col, max(grouped[group_index]["tables"][table_index]["columns"]))
                values_rows = _read_sheet_values(workbook[sheet_name], cached, max_col)
                for group_index, table_index in targets:
                    table = grouped[group_index]["tables"][table_index]
                    header, rows = extract_page_table(values_rows, table["columns"], stop_column)
                    if any(header) and not any(table["header"]):
                        table["header"] = header
                    table["rows"].extend(rows)
        finally:
            workbook.close()
            cached_book.close()

    result = []
    for group in grouped:
        if group["stem"] == "飲食":
            for table in group["tables"]:
                pivot_diet_table(table)
        if group["stem"] == "睡眠":
            for table in group["tables"]:
                add_sleep_day_total(table)
        if group["stem"] == "運動":
            for table in group["tables"]:
                add_exercise_day_total(table)
        tables = [table for table in group["tables"] if table["rows"]]
        if tables:
            result.append({"stem": group["stem"], "title": group["title"], "tables": tables})
    return result


def _sleep_seconds(value: str) -> int | None:
    """Duration text such as 03:00:00 or 9:00, not a clock past midnight."""
    match = re.fullmatch(r"(\d+):(\d{2})(?::(\d{2}))?", value.strip())
    if not match:
        return None
    hour, minute, second = int(match.group(1)), int(match.group(2)), int(match.group(3) or 0)
    if minute > 59 or second > 59:
        return None
    return hour * 3600 + minute * 60 + second


def _format_hms(seconds: int) -> str:
    hours, rem = divmod(seconds, 3600)
    minutes, secs = divmod(rem, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def add_sleep_day_total(table: dict) -> None:
    """Add 一天睡眠時間 on the first row of each date. Later rows of that date stay blank.

    The total is the sum of 睡眠時間 from the input rows, not from a page image.
    """
    rows = table.get("rows") or []
    if not rows:
        return
    header = list(table.get("header") or [])
    while len(header) < 4:
        header.append("")
    totals: dict[str, int] = {}
    counted: dict[str, bool] = {}
    for row in rows:
        when = row[0].strip() if row and row[0] else ""
        totals.setdefault(when, 0)
        counted.setdefault(when, False)
        seconds = _sleep_seconds(row[3] if len(row) > 3 and row[3] else "")
        if seconds is None:
            continue
        totals[when] += seconds
        counted[when] = True
    seen: set[str] = set()
    widened = []
    for row in rows:
        padded = list(row[:4]) + [""] * (4 - len(row))
        when = padded[0].strip()
        if when not in seen and counted.get(when):
            padded.append(_format_hms(totals[when]))
        else:
            padded.append("")
        seen.add(when)
        widened.append(padded)
    letters = list(table.get("letters") or ["C", "E", "G", "I"])
    while len(letters) < 4:
        letters.append("")
    table["letters"] = letters[:4] + [""]
    table["header"] = header[:4] + ["一天睡眠時間\nWaktu tidur per hari"]
    table["rows"] = widened


def _format_meters(meters: float) -> str:
    number = as_number(meters)
    return f"{number}米: {number} meter"


def add_exercise_day_total(table: dict) -> None:
    """Add 一天運動距離 on the first row of each date. Later rows of that date stay blank.

    The total is the sum of 運動距離 from the input rows, not from a page image.
    """
    rows = table.get("rows") or []
    if not rows:
        return
    header = list(table.get("header") or [])
    while len(header) < 3:
        header.append("")
    totals: dict[str, float] = {}
    counted: dict[str, bool] = {}
    for row in rows:
        when = row[0].strip() if row and row[0] else ""
        totals.setdefault(when, 0.0)
        counted.setdefault(when, False)
        meters = parse_meters(row[2] if len(row) > 2 and row[2] else None)
        if meters is None:
            continue
        totals[when] += meters
        counted[when] = True
    seen: set[str] = set()
    widened = []
    for row in rows:
        padded = list(row[:3]) + [""] * (3 - len(row))
        when = padded[0].strip()
        if when not in seen and counted.get(when):
            padded.append(_format_meters(totals[when]))
        else:
            padded.append("")
        seen.add(when)
        widened.append(padded)
    letters = list(table.get("letters") or ["C", "E", "K"])
    while len(letters) < 3:
        letters.append("")
    table["letters"] = letters[:3] + [""]
    table["header"] = header[:3] + ["一天運動距離\nJarak olahraga harian"]
    table["rows"] = widened


def _diet_meal_order(meal: str, seen: dict[str, int]) -> tuple[int, int]:
    if "早餐" in meal:
        group = 0
    elif "午餐" in meal:
        group = 1
    elif "晚餐" in meal:
        group = 2
    else:
        group = 3
    return group, seen[meal]


def _one_line(value: str) -> str:
    return " ".join(part.strip() for part in value.split("\n") if part.strip())


DAY_FOOD = "一天食物 Makanan untuk satu hari"
_MIXED_PORTION_RE = re.compile(r"^(\d+)\s+(\d+)\s*/\s*(\d+)$")
_FRACTION_PORTION_RE = re.compile(r"^(\d+)\s*/\s*(\d+)$")
_DECIMAL_PORTION_RE = re.compile(r"^(\d+(?:\.\d+)?)$")


def _parse_portion(value: str) -> Fraction | None:
    """Read a bowl amount such as '2', '1/2', or '1 1/2'."""
    text = value.replace("／", "/").replace("⁄", "/").strip()
    if not text:
        return None
    mixed = _MIXED_PORTION_RE.fullmatch(text)
    if mixed:
        whole, numerator, denominator = (int(part) for part in mixed.groups())
        if denominator == 0:
            return None
        return Fraction(whole) + Fraction(numerator, denominator)
    fraction = _FRACTION_PORTION_RE.fullmatch(text)
    if fraction:
        numerator, denominator = (int(part) for part in fraction.groups())
        if denominator == 0:
            return None
        return Fraction(numerator, denominator)
    number = _DECIMAL_PORTION_RE.fullmatch(text)
    if number:
        return Fraction(number.group(1))
    return None


def _format_portion(value: Fraction) -> str:
    whole = value.numerator // value.denominator
    rest = value - whole
    if rest == 0:
        return str(whole)
    piece = f"{rest.numerator}/{rest.denominator}"
    if whole:
        return f"{whole} {piece}"
    return piece


def _portion_total(amounts: list[str]) -> str:
    total = Fraction(0)
    found = False
    for amount in amounts:
        for part in amount.split("\n"):
            value = _parse_portion(part)
            if value is None:
                continue
            total += value
            found = True
    if not found:
        return ""
    return _format_portion(total)


def pivot_diet_table(table: dict) -> None:
    """Turn 飲食 rows into one row per date. Meal names become the columns.

    The source rows are 飲食健康, 日期, 食物份量. The page keeps the date header,
    places each meal above that meal's portion, and adds the day's total.
    """
    header = table.get("header") or []
    rows = table.get("rows") or []
    if len(header) < 3 or not rows:
        return
    date_header = header[1]
    portion_label = _one_line(header[2])
    seen: dict[str, int] = {}
    meals: list[str] = []
    by_date: dict[str, dict[str, str]] = {}
    order: list[str] = []
    for row in rows:
        meal = row[0].strip() if row and row[0] else ""
        when = row[1].strip() if len(row) > 1 and row[1] else ""
        amount = row[2].strip() if len(row) > 2 and row[2] else ""
        if not when:
            continue
        if meal and meal not in seen:
            seen[meal] = len(seen)
            meals.append(meal)
        if when not in by_date:
            by_date[when] = {}
            order.append(when)
        if not meal:
            continue
        current = by_date[when].get(meal, "")
        if not current:
            by_date[when][meal] = amount
        elif amount and amount not in current.split("\n"):
            by_date[when][meal] = f"{current}\n{amount}"
    if not meals:
        return
    meals.sort(key=lambda meal: _diet_meal_order(meal, seen))
    labels = meals + [DAY_FOOD]
    table["letters"] = [""] * (1 + len(labels))
    table["show_letters"] = False
    table["show_index"] = False
    table["header"] = [date_header] + [
        label if not portion_label else f"{label}\n{portion_label}" for label in labels
    ]
    table["rows"] = [[""] + labels]
    for when in order:
        amounts = [by_date[when].get(meal, "") for meal in meals]
        table["rows"].append([when] + amounts + [_portion_total(amounts)])


def add_food_sheet(workbook: Workbook, files: list[Path]) -> None:
    """Add 食物 from the 飲食 rows already read out of the input workbooks."""
    diet = None
    for group in collect_page_groups(list(files)):
        if group["stem"] == "飲食" and group["tables"]:
            diet = group["tables"][0]
            break
    if not diet or not diet["rows"]:
        return
    _write_food_sheet(_take_output_sheet(workbook, "食物"), diet)


def _write_food_sheet(worksheet: Worksheet, table: dict) -> None:
    """One header row, the meal names, then one row per date including the daily total."""
    worksheet.freeze_panes = "A3"
    worksheet.row_dimensions[1].height = 36
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    fill = PatternFill("solid", fgColor=HEADER_COLORS[0])
    body_font = Font(name=FONT_NAME, size=11)
    for index, title in enumerate(table["header"], start=1):
        cell = worksheet.cell(1, index, title)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
        width = sum(2 if ord(char) > 127 else 1 for char in _one_line(title)) + 2
        worksheet.column_dimensions[get_column_letter(index)].width = min(max(width, 16), 42)
    for row_offset, row in enumerate(table["rows"]):
        excel_row = row_offset + 2
        for index in range(len(table["header"])):
            value = row[index] if index < len(row) and row[index] else None
            cell = worksheet.cell(excel_row, index + 1, value)
            cell.font = body_font
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center")


FOOD_PLOT_SERIES = (
    "早餐食物: Makanan sarapan",
    "午餐食物: Makanan makan siang",
    "晚餐食物: Makanan makan malam",
    "一天食物 Makanan untuk satu hari",
)


def food_plot_from_sheet(workbook: Workbook) -> dict | None:
    """Line plot of the 食物 sheet: each meal and the daily total against 日期."""
    if "食物" not in workbook.sheetnames:
        return None
    sheet = workbook["食物"]
    header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_row:
        return None
    headers = ["" if value is None else str(value).split("\n")[0].strip() for value in header_row]
    if not headers or not headers[0].startswith("日期"):
        return None
    column_of = {name: headers.index(name) if name in headers else None for name in FOOD_PLOT_SERIES}
    categories: list[str] = []
    values = {name: [] for name in FOOD_PLOT_SERIES}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        when = row[0] if row else None
        if when is None or str(when).strip() == "":
            continue
        categories.append(str(when).strip())
        for name in FOOD_PLOT_SERIES:
            column = column_of[name]
            raw = None if column is None or column >= len(row) else row[column]
            parsed = None if raw is None or str(raw).strip() == "" else _parse_portion(str(raw))
            values[name].append(None if parsed is None else float(parsed))
    if not categories:
        return None
    return plot_spec(
        "食物-日期",
        "日期 Tanggal",
        "食物份量(碗)",
        categories,
        [(name, values[name]) for name in FOOD_PLOT_SERIES],
        True,
    )


def pack_page_chunks(blocks: list[dict], body_height: int, gap: int = PAGE_GAP) -> list[list[dict]]:
    """Split tables into page chunks. Each chunk repeats that table's header."""
    pages: list[list[dict]] = []
    current: list[dict] = []
    used = 0

    def close() -> None:
        nonlocal current, used
        if current:
            pages.append(current)
        current = []
        used = 0

    for block_index, block in enumerate(blocks):
        heights = block["row_heights"]
        if not heights:
            continue
        head = block["letter_h"] + block["header_h"]
        start = 0
        count = len(heights)
        while start < count:
            prefix = gap if current else 0
            if current and used + prefix + head + heights[start] > body_height:
                close()
                prefix = 0
            room = body_height - used - prefix - head
            consumed = 0
            end = start
            while end < count and consumed + heights[end] <= room:
                consumed += heights[end]
                end += 1
            if end == start:
                end = start + 1
                consumed = heights[start]
            current.append({"block_index": block_index, "row_start": start, "row_end": end})
            used += prefix + head + consumed
            start = end
    close()
    return pages


def _wrap_line(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, width: int) -> list[str]:
    if text == "":
        return [""]
    lines: list[str] = []
    current = ""
    for char in text:
        trial = current + char
        if current and draw.textlength(trial, font=font) > width:
            lines.append(current)
            current = char
        else:
            current = trial
    if current:
        lines.append(current)
    return lines or [""]


def _wrap_cell(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, width: int) -> list[str]:
    lines: list[str] = []
    for paragraph in text.split("\n"):
        lines.extend(_wrap_line(draw, paragraph, font, width))
    return lines or [""]


def _text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> float:
    widest = 0.0
    for line in text.split("\n"):
        widest = max(widest, draw.textlength(line, font=font))
    return widest


def _column_widths(
    draw: ImageDraw.ImageDraw,
    letters: list[str],
    header: list[str],
    rows: list[list[str]],
    font: ImageFont.FreeTypeFont,
    letter_font: ImageFont.FreeTypeFont,
    index_width: int = PAGE_INDEX_W,
) -> list[int]:
    usable = PAGE_RIGHT - PAGE_LEFT - index_width
    prefs = []
    for index, letter in enumerate(letters):
        widest = draw.textlength(letter, font=letter_font)
        widest = max(widest, _text_width(draw, header[index] if index < len(header) else "", font))
        for row in rows:
            if index < len(row):
                widest = max(widest, _text_width(draw, row[index], font))
        prefs.append(max(int(math.ceil(widest)) + 16, 80))
    total = sum(prefs)
    if total <= 0:
        return [usable // max(len(letters), 1)] * len(letters)
    if total <= usable:
        extra = usable - total
        widths = []
        used = 0
        for index, pref in enumerate(prefs):
            if index == len(prefs) - 1:
                widths.append(usable - used)
            else:
                width = int(round(pref + extra * (pref / total)))
                widths.append(width)
                used += width
        return widths
    scale = usable / total
    widths = [max(64, int(pref * scale)) for pref in prefs]
    overflow = sum(widths) - usable
    while overflow > 0:
        index = max(range(len(widths)), key=lambda item: widths[item])
        if widths[index] <= 64:
            break
        widths[index] -= 1
        overflow -= 1
    short = usable - sum(widths)
    if short > 0 and widths:
        widths[-1] += short
    return widths


def _block_metrics(draw: ImageDraw.ImageDraw, table: dict, font: ImageFont.FreeTypeFont, letter_font: ImageFont.FreeTypeFont) -> dict:
    letters = table["letters"]
    show_letters = table.get("show_letters", True)
    show_index = table.get("show_index", True)
    header = list(table["header"]) + [""] * (len(letters) - len(table["header"]))
    rows = table["rows"]
    index_width = PAGE_INDEX_W if show_index else 0
    widths = _column_widths(draw, letters, header, rows, font, letter_font, index_width)
    inner = [max(width - 16, 8) for width in widths]
    header_lines = [_wrap_cell(draw, header[index], font, inner[index]) for index in range(len(letters))]
    wrapped_rows = []
    row_heights = []
    for row in rows:
        cells = []
        for index in range(len(letters)):
            text = row[index] if index < len(row) else ""
            cells.append(_wrap_cell(draw, text, font, inner[index]))
        wrapped_rows.append(cells)
        row_heights.append(max(PAGE_MIN_ROW_H, max(len(lines) for lines in cells) * PAGE_LINE + 12))
    header_h = max(PAGE_MIN_HEADER_H, max(len(lines) for lines in header_lines) * PAGE_LINE + 12)
    return {
        "letters": letters,
        "widths": widths,
        "header_lines": header_lines,
        "wrapped_rows": wrapped_rows,
        "letter_h": PAGE_LETTER_H if show_letters else 0,
        "header_h": header_h,
        "row_heights": row_heights,
        "show_index": show_index,
    }


def _draw_text_block(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    lines: list[str],
    font: ImageFont.FreeTypeFont,
    align: str,
) -> None:
    x0, y0, x1, y1 = box
    block_h = PAGE_LINE * len(lines)
    y = y0 + max(0, (y1 - y0 - block_h) / 2)
    for line in lines:
        if align == "center":
            width = draw.textlength(line, font=font)
            x = x0 + (x1 - x0 - width) / 2
        else:
            x = x0 + 8
        draw.text((x, y), line, font=font, fill=PAGE_INK)
        y += PAGE_LINE


def _draw_table_chunk(
    draw: ImageDraw.ImageDraw,
    top: int,
    block: dict,
    row_start: int,
    row_end: int,
    font: ImageFont.FreeTypeFont,
    letter_font: ImageFont.FreeTypeFont,
) -> int:
    widths = block["widths"]
    letter_h = block["letter_h"]
    header_h = block["header_h"]
    show_index = block.get("show_index", True)
    index_w = PAGE_INDEX_W if show_index else 0
    xs = [PAGE_LEFT, PAGE_LEFT + index_w]
    for width in widths:
        xs.append(xs[-1] + width)
    y_letter = top
    y_header = y_letter + letter_h
    y_data = y_header + header_h
    row_tops = [y_data]
    for height in block["row_heights"][row_start:row_end]:
        row_tops.append(row_tops[-1] + height)
    bottom = row_tops[-1]

    for index in range(len(xs) - 1):
        draw.rectangle([xs[index], y_letter, xs[index + 1], y_header], fill=PAGE_HEADER_BG)
        header_fill = PAGE_INDEX_HEADER_BG if index == 0 else PAGE_HEADER_BG
        draw.rectangle([xs[index], y_header, xs[index + 1], y_data], fill=header_fill)
    for row_index, y0 in enumerate(row_tops[:-1]):
        y1 = row_tops[row_index + 1]
        draw.rectangle([xs[0], y0, xs[1], y1], fill=PAGE_INDEX_BG)
        for column in range(1, len(xs) - 1):
            draw.rectangle([xs[column], y0, xs[column + 1], y1], fill=PAGE_BG)

    for x in xs:
        draw.line([(x, y_letter), (x, bottom)], fill=PAGE_GRID)
    for y in (y_letter, y_header, y_data, *row_tops[1:]):
        draw.line([(xs[0], y), (xs[-1], y)], fill=PAGE_GRID)

    if letter_h:
        for index, letter in enumerate(block["letters"]):
            _draw_text_block(draw, (xs[index + 1], y_letter, xs[index + 2], y_header), [letter], letter_font, "center")
    if show_index:
        _draw_text_block(draw, (xs[0], y_header, xs[1], y_data), ["1"], font, "center")
    for index, lines in enumerate(block["header_lines"]):
        _draw_text_block(draw, (xs[index + 1], y_header, xs[index + 2], y_data), lines, font, "left")
    for offset, lines_row in enumerate(block["wrapped_rows"][row_start:row_end]):
        y0 = row_tops[offset]
        y1 = row_tops[offset + 1]
        if show_index:
            _draw_text_block(draw, (xs[0], y0, xs[1], y1), [str(row_start + offset + 2)], font, "center")
        for index, lines in enumerate(lines_row):
            _draw_text_block(draw, (xs[index + 1], y0, xs[index + 2], y1), lines, font, "left")
    return bottom


def render_group_pages(group: dict) -> list[Image.Image]:
    """Draw one output name. Tables share a page until the next row does not fit."""
    font = _page_font(16)
    letter_font = _page_font(15)
    title_font = _page_font(22)
    probe = Image.new("RGB", (PAGE_W, PAGE_H), PAGE_BG)
    draw = ImageDraw.Draw(probe)
    blocks = [_block_metrics(draw, table, font, letter_font) for table in group["tables"]]
    body_height = PAGE_BOTTOM - PAGE_TOP
    chunks = pack_page_chunks(blocks, body_height, PAGE_GAP)
    pages = []
    total = max(len(chunks), 1)
    for page_index, page in enumerate(chunks, start=1):
        image = Image.new("RGB", (PAGE_W, PAGE_H), PAGE_BG)
        page_draw = ImageDraw.Draw(image)
        page_draw.text((30, 28), f"{group['title']}  {page_index}/{total}", font=title_font, fill=PAGE_INK)
        cursor = PAGE_TOP
        for chunk_index, chunk in enumerate(page):
            if chunk_index:
                cursor += PAGE_GAP
            cursor = _draw_table_chunk(
                page_draw,
                cursor,
                blocks[chunk["block_index"]],
                chunk["row_start"],
                chunk["row_end"],
                font,
                letter_font,
            )
        pages.append(image)
    return pages


def _clear_page_pngs(destination: Path) -> None:
    stems = {spec["stem"] for spec in PAGE_SPECS}
    for path in destination.glob("*.png"):
        match = re.match(r"^(.+)_(\d+)$", path.stem)
        if match and match.group(1) in stems:
            path.unlink()


def _write_pdf(path: Path, images: list[Image.Image]) -> None:
    if not images:
        return
    resolution = 1600 / (297 / 25.4)
    images[0].save(
        path,
        "PDF",
        resolution=resolution,
        save_all=True,
        append_images=images[1:],
    )


def export_page_files(files: list[Path], destination: Path, pdf_stem: str) -> tuple[list[str], Path | None]:
    """Write landscape PNG pages and one PDF named like the output workbook."""
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    groups = collect_page_groups(list(files))
    groups = _insert_water_page(groups, destination / f"{pdf_stem}.xlsx")
    _clear_page_pngs(destination)
    for path in destination.glob("喝水量_*.png"):
        path.unlink()
    notes = []
    images: list[Image.Image] = []
    if not groups:
        notes.append("No landscape pages. The selected columns had no rows above '?'.")
        return notes, None
    for group in groups:
        pages = render_group_pages(group)
        images.extend(pages)
        for index, image in enumerate(pages, start=1):
            path = destination / f"{group['stem']}_{index}.png"
            image.save(path, "PNG")
            notes.append(f"Saved {path}")
    pdf_path = destination / f"{pdf_stem}.pdf"
    _write_pdf(pdf_path, images)
    notes.append(f"Saved {pdf_path}")
    return notes, pdf_path


def _water_sheet_text(value, cups: bool) -> str:
    if value is None:
        return ""
    if cups and isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(float(value)):
        return f"{float(value):.2f}"
    return _page_text(value)


def _water_page_group(excel_path: Path) -> dict | None:
    """Read Extract.xlsx sheet 喝水量 into one landscape page table."""
    if not excel_path.is_file():
        return None
    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if "喝水量" not in workbook.sheetnames:
            return None
        rows = list(workbook["喝水量"].iter_rows(values_only=True))
    finally:
        workbook.close()
    if not rows:
        return None
    header = [_page_header(value) for value in rows[0]]
    while header and not header[-1]:
        header.pop()
    if not any(header):
        return None
    cups_at = {index for index, title in enumerate(header) if "喝水量" in title}
    data = []
    for row in rows[1:]:
        values = [
            _water_sheet_text(row[index] if index < len(row) else None, index in cups_at)
            for index in range(len(header))
        ]
        if any(values):
            data.append(values)
    if not data:
        return None
    return {
        "stem": "喝水量",
        "title": "喝水量",
        "tables": [{
            "letters": [""] * len(header),
            "header": header,
            "rows": data,
            "show_letters": False,
            "show_index": False,
        }],
    }


def _insert_water_page(groups: list[dict], excel_path: Path) -> list[dict]:
    """Place the 喝水量 page immediately after 飲食."""
    water = _water_page_group(excel_path)
    if water is None:
        return groups
    placed = []
    inserted = False
    for group in groups:
        placed.append(group)
        if group["stem"] == "飲食":
            placed.append(water)
            inserted = True
    if not inserted:
        placed.append(water)
    return placed


# Same receivers as readExcel.py "Gmail 傳送 Excel" in app7.
GMAIL_TO = "ghpiu0661@gmail.com,ghjane5168@gmail.com,a290772@gmail.com"
GMAIL_SUBJECT = "Extract"


def _ps_quote(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def gmail_output_files(folder: Path) -> list[Path]:
    """Extract.xlsx, Extract.pdf, and Extract.html in the output folder."""
    import shutil

    output = Path(folder) / "output"
    excel = output / "Extract.xlsx"
    pdf = output / "Extract.pdf"
    html_path = output / "Extract.html"
    plots = output / "Extract_plots.html"
    if not html_path.is_file() and plots.is_file():
        shutil.copyfile(plots, html_path)
    chosen = [excel, pdf, html_path]
    missing = [path.name for path in chosen if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "Run Process first. These files are missing from output:\n" + "\n".join(missing)
        )
    return [path.resolve() for path in chosen]


def gmail_compose_url(to: str, subject: str, body: str) -> str:
    query = urllib.parse.urlencode(
        {"view": "cm", "fs": "1", "tf": "cm", "to": to, "su": subject, "body": body}
    )
    return "https://mail.google.com/mail/u/0/?" + query


def send_gmail_files(files: list[Path], to: str = GMAIL_TO, runner=None) -> str:
    """Open Gmail compose and put the files on the clipboard, then paste once."""
    paths = [Path(item).resolve() for item in files]
    missing = [str(item) for item in paths if not item.is_file()]
    if missing:
        raise FileNotFoundError("Missing files:\n" + "\n".join(missing))
    names = "\n".join(item.name for item in paths)
    body = "Extract.xlsx, Extract.pdf, and Extract.html attached."
    url = gmail_compose_url(to, GMAIL_SUBJECT, body)
    adds = "\n".join(
        "[void]$files.Add((Resolve-Path -LiteralPath " + _ps_quote(str(item)) + ").Path)" for item in paths
    )
    script = """
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class GmailWin {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@
$files = New-Object System.Collections.Specialized.StringCollection
__ADDS__
[System.Windows.Forms.Clipboard]::SetFileDropList($files)
$before = @(Get-Process | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object { [int64]$_.MainWindowHandle })
Start-Process __URL__
$target = $null
for ($i = 0; $i -lt 40; $i++) {
  Start-Sleep -Milliseconds 500
  $windows = @(Get-Process | Where-Object {
    $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -match 'Gmail|Compose|Mail'
  })
  $fresh = @($windows | Where-Object { $before -notcontains [int64]$_.MainWindowHandle })
  if ($fresh.Count -gt 0) { $target = $fresh[0]; break }
  $compose = @($windows | Where-Object { $_.MainWindowTitle -match 'Compose|Mail' })
  if ($compose.Count -gt 0) { $target = $compose[0]; break }
}
if ($null -eq $target) { throw 'Gmail compose window was not found. Files were not attached.' }
$hwnd = $target.MainWindowHandle
[GmailWin]::ShowWindow($hwnd, 9) | Out-Null
Start-Sleep -Milliseconds 400
[GmailWin]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 800
$rect = New-Object GmailWin+RECT
[GmailWin]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top
if ($width -lt 200 -or $height -lt 200) { throw 'Gmail window is too small to attach files.' }
$x = [int]($rect.Left + ($width * 0.45))
$y = [int]($rect.Top + ($height * 0.58))
[GmailWin]::SetForegroundWindow($hwnd) | Out-Null
[GmailWin]::SetCursorPos($x, $y) | Out-Null
Start-Sleep -Milliseconds 200
[GmailWin]::mouse_event(0x0002, 0, 0, 0, 0)
[GmailWin]::mouse_event(0x0004, 0, 0, 0, 0)
Start-Sleep -Milliseconds 300
[GmailWin]::SetCursorPos($x, $y) | Out-Null
[GmailWin]::mouse_event(0x0002, 0, 0, 0, 0)
[GmailWin]::mouse_event(0x0004, 0, 0, 0, 0)
Start-Sleep -Milliseconds 400
[GmailWin]::SetForegroundWindow($hwnd) | Out-Null
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 800
Write-Output 'GMAIL_OK'
""".replace("__ADDS__", adds).replace("__URL__", _ps_quote(url))
    if runner is None:
        if os.name != "nt":
            raise OSError("Windows required for Gmail attach.")
        import ctypes

        ctypes.windll.user32.AllowSetForegroundWindow(-1)
        runner = subprocess.run
    script_path = Path(os.path.join(os.environ.get("TEMP", "."), "processExcel_gmail.ps1"))
    script_path.write_text(script, encoding="utf-8-sig")
    proc = runner(
        ["powershell", "-NoProfile", "-NoLogo", "-STA", "-ExecutionPolicy", "Bypass", "-File", str(script_path)],
        capture_output=True,
        text=True,
        timeout=70,
        check=False,
    )
    out = (proc.stdout or "") + (proc.stderr or "")
    if "GMAIL_OK" not in out:
        detail = out.strip() or "Gmail open failed."
        raise RuntimeError(detail)
    return names


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("processExcel")
        self.geometry("760x640")
        self.minsize(680, 560)
        self.option_add("*Font", (FONT_NAME, 11))
        self.folder = tk.StringVar()
        self.width_cm = tk.StringVar(value="22")
        self.height_cm = tk.StringVar(value="12")
        self.title_size = tk.StringVar(value="18")
        self.xlabel_size = tk.StringVar(value="14")
        self.ylabel_size = tk.StringVar(value="14")
        self.legend_size = tk.StringVar(value="12")
        self.tick_size = tk.StringVar(value="11")
        self._build()

    def _build(self) -> None:
        pad = {"padx": 12, "pady": 6}
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True)

        folder_row = ttk.Frame(frame)
        folder_row.pack(fill="x", **pad)
        ttk.Label(folder_row, text="Excel folder").pack(side="left")
        ttk.Entry(folder_row, textvariable=self.folder).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(folder_row, text="Browse", command=self.browse).pack(side="left")

        style_box = ttk.LabelFrame(frame, text="Plot size and font size")
        style_box.pack(fill="x", padx=12, pady=8)
        fields = (
            ("Chart width (cm)", self.width_cm),
            ("Chart height (cm)", self.height_cm),
            ("Title font", self.title_size),
            ("X label font", self.xlabel_size),
            ("Y label font", self.ylabel_size),
            ("Legend font", self.legend_size),
            ("Tick font", self.tick_size),
        )
        for index, (label, variable) in enumerate(fields):
            row, col = divmod(index, 2)
            ttk.Label(style_box, text=label).grid(row=row, column=col * 2, sticky="w", padx=10, pady=6)
            ttk.Entry(style_box, textvariable=variable, width=10).grid(row=row, column=col * 2 + 1, sticky="w", pady=6)

        hint = (
            "Reads YYYYMMDD.xlsx, sorted by date. Sheets 運動, 血壓, 睡眠, and 清潔 (喝水量) are copied "
            "until the row above '?'. Output: output\\Extract.xlsx and output\\Extract_plots.html."
        )
        ttk.Label(frame, text=hint, wraplength=700, justify="left").pack(fill="x", padx=12, pady=4)

        button_row = ttk.Frame(frame)
        button_row.pack(anchor="w", padx=12, pady=6)
        self.run_button = ttk.Button(button_row, text="Process", command=self.on_run)
        self.run_button.pack(side="left")
        self.gmail_button = ttk.Button(button_row, text="Gmail 傳送 Excel", command=self.on_gmail)
        self.gmail_button.pack(side="left", padx=(8, 0))
        self.log = ScrolledText(frame, height=18, font=(FONT_NAME, 10))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def browse(self) -> None:
        current = self.folder.get().strip()
        initial = current if current and Path(current).is_dir() else str(Path.home() / "Downloads")
        selected = filedialog.askdirectory(title="Select folder with Excel files", initialdir=initial)
        if selected:
            self.folder.set(selected)

    def read_style(self) -> PlotStyle:
        def number(label: str, variable: tk.StringVar) -> float:
            text = variable.get().strip()
            try:
                return float(text)
            except ValueError as exc:
                raise ValueError(f"{label} must be a number.") from exc

        style = PlotStyle(
            width_cm=number("Chart width", self.width_cm),
            height_cm=number("Chart height", self.height_cm),
            title_size=number("Title font", self.title_size),
            xlabel_size=number("X label font", self.xlabel_size),
            ylabel_size=number("Y label font", self.ylabel_size),
            legend_size=number("Legend font", self.legend_size),
            tick_size=number("Tick font", self.tick_size),
        )
        validate_style(style)
        return style

    def append_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def on_run(self) -> None:
        folder = self.folder.get().strip()
        if not folder:
            messagebox.showerror("processExcel", "Choose a folder first.")
            return
        try:
            style = self.read_style()
        except ValueError as exc:
            messagebox.showerror("processExcel", str(exc))
            return
        self.run_button.state(["disabled"])
        self.append_log("Working...")
        thread = threading.Thread(target=self._work, args=(folder, style), daemon=True)
        thread.start()

    def _work(self, folder: str, style: PlotStyle) -> None:
        try:
            result = process_folder(folder, style)
        except Exception as exc:
            self.after(0, self._finish, None, str(exc))
            return
        self.after(0, self._finish, result, None)

    def _finish(self, result: ProcessResult | None, error: str | None) -> None:
        self.run_button.state(["!disabled"])
        if error:
            self.append_log(error)
            messagebox.showerror("processExcel", error)
            return
        assert result is not None
        for line in result.messages:
            self.append_log(line)
        saved = [str(result.excel_path), str(result.plot_path)]
        if result.pdf_path is not None:
            saved.append(str(result.pdf_path))
        messagebox.showinfo("processExcel", "Saved\n" + "\n".join(saved))

    def on_gmail(self) -> None:
        folder = self.folder.get().strip()
        if not folder:
            messagebox.showerror("Gmail 傳送 Excel", "Choose a folder first.")
            return
        try:
            files = gmail_output_files(Path(folder))
        except FileNotFoundError as exc:
            messagebox.showerror("Gmail 傳送 Excel", str(exc))
            return
        self.gmail_button.state(["disabled"])
        self.append_log("Opening Gmail...")
        thread = threading.Thread(target=self._gmail_work, args=(files,), daemon=True)
        thread.start()

    def _gmail_work(self, files: list[Path]) -> None:
        try:
            names = send_gmail_files(files)
        except Exception as exc:
            self.after(0, self._gmail_finish, None, str(exc))
            return
        self.after(0, self._gmail_finish, names, None)

    def _gmail_finish(self, names: str | None, error: str | None) -> None:
        self.gmail_button.state(["!disabled"])
        if error:
            self.append_log(error)
            messagebox.showerror("Gmail 傳送 Excel", error)
            return
        assert names is not None
        self.append_log("Gmail opened.\n" + names)
        self.append_log("To: " + GMAIL_TO)
        messagebox.showinfo(
            "Gmail 傳送 Excel",
            "Gmail compose is open.\n\n"
            "These files were pasted into the message:\n"
            + names
            + "\n\nCheck that all three are attached, then click Send.\n"
            "If they are missing, click the message body and press Ctrl+V.\n\nReceivers:\n"
            + GMAIL_TO,
        )


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if args:
        result = process_folder(args[0])
        print("\n".join(result.messages))
        return 0
    App().mainloop()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(exc, file=sys.stderr)
        raise
