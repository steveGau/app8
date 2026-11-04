"""Read dated Excel workbooks and build Extract.xlsx plus Plotly charts.

Input files are named YYYYMMDD.xlsx. Each workbook is read until column A
is "?". Output is written to an "output" folder inside the selected folder.
"""

from __future__ import annotations

import datetime as dt
import html
import math
import re
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from openpyxl import Workbook, load_workbook
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
    """Parse a cup amount. '1/6' is one sixth of a cup. '?' is skipped."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if is_missing(value):
        return None
    text = str(value).strip()
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


def extract_water(path: Path, worksheet, cached=None) -> dict | None:
    """Rows on 清潔 up to '?', mapped to 喝水 / 喝水日期 / 喝水時間 / 喝水量(杯)."""
    _header, rows = load_table(worksheet, 7, cached)
    parsed = []
    for row in rows:
        cups = parse_cups(cell_at(row, 7))
        when = parse_date(cell_at(row, 3)) or file_date(path)
        clock = format_hhmm(cell_at(row, 5))
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


def add_water_sheets(workbook: Workbook, plots: list[dict], blocks: list[dict], style: PlotStyle) -> None:
    """Add 喝水量 and 喝水量 Plot. Existing sheets are left as they are."""
    if not blocks:
        return
    if workbook.sheetnames == ["Extract"]:
        sheet = workbook.active
        sheet.title = "喝水量"
    else:
        sheet = workbook.create_sheet("喝水量")
    write_record_sheet(
        sheet,
        blocks,
        WATER_STRIDE,
        [(0, "label", "text"), (1, "date", "date"), (2, "time", "text"), (3, "cups", "decimal")],
    )
    plot_sheet = workbook.create_sheet("喝水量 Plot")
    plot_sheet._chart_writer = SheetWriter(plot_sheet)
    totals = daily_totals(blocks, "cups")
    counts = daily_counts(blocks, "cups")
    specs = [
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
    messages = [f"Folder: {folder}", f"Files: {', '.join(path.name for path in files)}"]
    for path in files:
        try:
            one_exercise, one_bp, one_sleep, notes = read_workbook(path)
            one_water, water_notes = read_water(path)
            one_excretion, one_body, condition_notes = read_condition(path)
        except Exception as exc:
            messages.append(f"{path.name}: skipped ({exc})")
            continue
        messages.extend(notes)
        messages.extend(water_notes)
        messages.extend(condition_notes)
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

    if not exercise and not bp and not sleep and not water and not excretion and not body:
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
    write_plotly(plot_path, plots, style, intro)
    messages.append("運動時間 is stored in minutes. 運動距離 is stored in meters. 睡眠時間 is stored in hours. 喝水量 is stored in cups.")
    messages.append(f"Saved {excel_path}")
    messages.append(f"Saved {plot_path}")
    return ProcessResult(excel_path, plot_path, messages)


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

        self.run_button = ttk.Button(frame, text="Process", command=self.on_run)
        self.run_button.pack(anchor="w", padx=12, pady=6)
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
        messagebox.showinfo("processExcel", f"Saved\n{result.excel_path}\n{result.plot_path}")


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
