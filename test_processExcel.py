"""Tests for processExcel extraction, workbook layout, and charts."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook, load_workbook

import processExcel as app


def _headers(ws, pairs):
    for column, text in pairs.items():
        ws.cell(1, column, text)


def _exercise_sheet(wb, rows, stop="?"):
    ws = wb.active
    ws.title = "運動"
    _headers(
        ws,
        {
            1: "選擇運動 Pilih olahraga",
            3: "運動日期 Tanggal Olahraga",
            5: "開始時間 Waktu mulai",
            7: "結束時間 Waktu berakhir",
            9: "運動時間 Waktu berolahraga",
            11: "運動距離 Jarak pergerakan",
        },
    )
    for offset, row in enumerate(rows, start=2):
        if row is None:
            continue
        ws.cell(offset, 1, row.get("sport"))
        ws.cell(offset, 3, row.get("date"))
        ws.cell(offset, 5, row.get("start"))
        ws.cell(offset, 7, row.get("end"))
        if "duration" in row:
            ws.cell(offset, 9, row["duration"])
        ws.cell(offset, 11, row.get("distance"))
    if stop is not None:
        ws.cell(len(rows) + 2, 1, stop)
    return ws


class ParseTests(unittest.TestCase):
    def test_distance_labels(self):
        for meters in (5, 10, 15, 20, 25, 30, 35, 40):
            self.assertEqual(app.parse_meters(f"{meters}米: {meters} meter"), meters)
        self.assertEqual(app.parse_meters("10米: 12 meter"), 10)
        self.assertEqual(app.parse_meters(10), 10)
        self.assertIsNone(app.parse_meters(""))
        self.assertIsNone(app.parse_meters(None))
        self.assertIsNone(app.parse_meters("沒有數字"))

    def test_duration_minutes_and_overnight(self):
        self.assertEqual(app.duration_seconds(None, None, "01:30:00"), 90 * 60)
        self.assertEqual(app.duration_seconds(time(8, 0), time(8, 30), None), 30 * 60)
        self.assertEqual(app.duration_seconds(time(23, 30), time(0, 10), None), 40 * 60)
        self.assertEqual(app.duration_seconds(time(21, 0), time(5, 0), None), 8 * 3600)
        self.assertEqual(app.duration_seconds(time(0, 0), time(2, 0), "01:00:00"), 2 * 3600)
        self.assertAlmostEqual(app.duration_seconds(None, None, 0.5) / 60, 720)

    def test_dates_and_file_order(self):
        self.assertEqual(app.parse_date("2026/10/04"), date(2026, 10, 4))
        self.assertEqual(app.parse_date("20261004"), date(2026, 10, 4))
        self.assertEqual(app.parse_date(datetime(2026, 10, 4, 8, 0)), date(2026, 10, 4))
        self.assertIsNone(app.parse_date("?"))
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            for name in ("20261005.xlsx", "notes.xlsx", "20260230.xlsx", "20261004.xlsx", "~$20261004.xlsx"):
                (root / name).write_bytes(b"")
            names = [path.name for path in app.list_dated_workbooks(root)]
            self.assertEqual(names, ["20261004.xlsx", "20261005.xlsx"])

    def test_question_row_and_blank_rows(self):
        wb = Workbook()
        _exercise_sheet(
            wb,
            [
                {"sport": "走路", "start": time(8, 0), "end": time(8, 30), "distance": "10米: 10 meter"},
                None,
                {"sport": None, "start": time(9, 0), "end": time(9, 20), "distance": "20米: 20 meter"},
                {"sport": "不該讀", "distance": "40米: 40 meter"},
            ],
            stop="？",
        )
        # The helper places '?' on len(rows)+2. Rebuild the stop so the ignored
        # row is after the marker: rows are 2, blank 3, data 4, marker must be 5.
        ws = wb["運動"]
        ws.cell(5, 1, "？")
        ws.cell(6, 1, "不該讀")
        ws.cell(6, 11, "40米: 40 meter")
        block = app.extract_exercise(Path("20261004.xlsx"), ws)
        self.assertEqual([row["sport"] for row in block["rows"]], ["走路", None])
        self.assertEqual([row["meters"] for row in block["rows"]], [10, 20])
        self.assertEqual([row["minutes"] for row in block["rows"]], [30, 20])
        self.assertTrue(all(row["date"] == date(2026, 10, 4) for row in block["rows"]))

    def test_formula_duration_uses_clocks(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "運動"
        ws.cell(1, 1, "選擇運動")
        ws.cell(2, 1, "騎車")
        ws.cell(2, 5, time(7, 0))
        ws.cell(2, 7, time(7, 15))
        ws.cell(2, 9, "=G2-E2")
        ws.cell(2, 11, "15米: 15 meter")
        block = app.extract_exercise(Path("20261005.xlsx"), ws)
        self.assertEqual(block["rows"][0]["minutes"], 15)
        self.assertEqual(block["rows"][0]["meters"], 15)
        self.assertEqual(block["rows"][0]["date"], date(2026, 10, 5))


class WorkbookTests(unittest.TestCase):
    def _build_folder(self, root: Path) -> None:
        empty = Workbook()
        empty.active.title = "運動"
        empty.active.cell(1, 1, "選擇運動 Pilih olahraga")
        empty.active.cell(2, 1, "?")
        empty.save(root / "20261003.xlsx")
        empty.close()

        first = Workbook()
        _exercise_sheet(
            first,
            [
                {
                    "sport": "走路",
                    "date": date(2026, 10, 4),
                    "start": time(8, 0),
                    "end": time(8, 30),
                    "duration": "00:30:00",
                    "distance": "10米: 10 meter",
                },
                None,
                {
                    "sport": "跑步",
                    "date": date(2026, 10, 4),
                    "start": time(23, 30),
                    "end": time(0, 10),
                    "distance": "5米: 5 meter",
                },
            ],
        )
        # Insert a kept row with an empty label between the blank row and the stop.
        exercise = first["運動"]
        exercise.cell(3, 3, date(2026, 10, 4))
        exercise.cell(3, 5, time(9, 0))
        exercise.cell(3, 7, time(9, 20))
        exercise.cell(3, 11, "20米: 20 meter")
        exercise.cell(6, 1, "不該讀")
        exercise.cell(6, 11, "40米: 40 meter")

        bp = first.create_sheet("血壓")
        _headers(
            bp,
            {
                1: "健康 Sehat",
                3: "量血壓日期 Tanggal pengukuran tekanan darah",
                4: "量血壓時間 Waktu pengukuran tekanan darah",
                5: "收縮壓, 高壓 Tekanan darah sistolik tekanan tinggi",
                6: "舒張壓, 低壓 Tekanan darah diastolik tekanan rendah",
                7: "心跳 Detak jantung",
            },
        )
        bp.cell(2, 1, "早上")
        bp.cell(2, 3, date(2026, 10, 4))
        bp.cell(2, 4, time(10, 0))
        bp.cell(2, 5, 130)
        bp.cell(2, 6, 85)
        bp.cell(2, 7, 72)
        bp.cell(3, 1, "更早")
        bp.cell(3, 3, date(2026, 10, 4))
        bp.cell(3, 4, time(8, 0))
        bp.cell(3, 5, 120)
        bp.cell(3, 6, 80)
        bp.cell(3, 7, 70)
        bp.cell(4, 1, "?")

        sleep = first.create_sheet("睡眠")
        _headers(
            sleep,
            {
                1: "睡眠 Tidur",
                3: "日期 Tanggal",
                5: "開始時間 Waktu mulai",
                7: "結束時間 Waktu berakhir",
                9: "睡眠時間 Waktu berolahraga",
            },
        )
        sleep.cell(2, 1, "午睡")
        sleep.cell(2, 3, date(2026, 10, 4))
        sleep.cell(2, 5, time(23, 0))
        sleep.cell(2, 7, time(1, 0))
        sleep.cell(3, 1, "夜睡")
        sleep.cell(3, 3, date(2026, 10, 4))
        sleep.cell(3, 5, time(1, 30))
        sleep.cell(3, 7, time(5, 0))
        sleep.cell(4, 1, "?")
        first.save(root / "20261004.xlsx")
        first.close()

        second = Workbook()
        ws = second.active
        ws.title = "運動"
        ws.cell(1, 1, "選擇運動 Pilih olahraga")
        ws.cell(1, 9, "運動時間 Waktu berolahraga")
        ws.cell(1, 11, "運動距離 Jarak pergerakan")
        ws.cell(2, 1, "騎車")
        ws.cell(2, 5, time(7, 0))
        ws.cell(2, 7, time(7, 15))
        ws.cell(2, 9, "=G2-E2")
        ws.cell(2, 11, "15米: 15 meter")
        bp2 = second.create_sheet("血壓")
        bp2.cell(1, 1, "健康 Sehat")
        bp2.cell(2, 1, "上午")
        bp2.cell(2, 3, date(2026, 10, 5))
        bp2.cell(2, 4, time(9, 0))
        bp2.cell(2, 5, 140)
        bp2.cell(2, 6, 90)
        bp2.cell(2, 7, 80)
        sleep2 = second.create_sheet("睡眠")
        sleep2.cell(1, 1, "睡眠 Tidur")
        sleep2.cell(2, 1, "睡眠")
        sleep2.cell(2, 3, date(2026, 10, 5))
        sleep2.cell(2, 5, time(22, 0))
        sleep2.cell(2, 7, time(6, 0))
        sleep2.cell(2, 9, "07:30:00")
        second.save(root / "20261005.xlsx")
        second.close()

        junk = Workbook()
        junk.active["A1"] = "JUNK"
        junk.save(root / "notes.xlsx")
        junk.close()
        (root / "20260230.xlsx").write_text("not a workbook", encoding="utf-8")
        (root / "~$20261004.xlsx").write_text("lock", encoding="utf-8")

    def test_extract_layout_charts_and_plotly(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._build_folder(root)
            style = app.PlotStyle(
                width_cm=16,
                height_cm=9,
                title_size=21,
                xlabel_size=15,
                ylabel_size=14,
                legend_size=13,
                tick_size=12,
            )
            result = app.process_folder(root, style)
            self.assertEqual(result.excel_path, root / "output" / "Extract.xlsx")
            self.assertEqual(result.plot_path, root / "output" / "Extract_plots.html")
            app.process_folder(root, style)

            book = load_workbook(result.excel_path)
            self.assertEqual(
                book.sheetnames,
                ["運動", "運動 Plot", "血壓", "血壓 Plot", "睡眠", "睡眠 Plot"],
            )
            exercise = book["運動"]
            self.assertIn("選擇運動", str(exercise["A1"].value))
            self.assertIn("分鐘", str(exercise["C1"].value))
            self.assertIn("米", str(exercise["D1"].value))
            self.assertEqual(
                [exercise.cell(row, 1).value for row in (2, 3, 4)],
                ["走路", None, "跑步"],
            )
            self.assertEqual([exercise.cell(row, 3).value for row in (2, 3, 4)], [30, 20, 40])
            self.assertEqual([exercise.cell(row, 4).value for row in (2, 3, 4)], [10, 20, 5])
            self.assertEqual(exercise["E2"].value, "騎車")
            self.assertEqual(self._as_date(exercise["F2"].value), date(2026, 10, 5))
            self.assertEqual(exercise["G2"].value, 15)
            self.assertEqual(exercise["H2"].value, 15)
            self.assertNotIn("不該讀", self._values(exercise))
            self.assertNotIn("2026-10-03", self._values(exercise))
            self.assertNotIn("JUNK", self._values(exercise))

            bp = book["血壓"]
            self.assertEqual([bp.cell(row, 3).value for row in (2, 3)], [130, 120])
            self.assertEqual(bp["F2"].value, "上午")
            self.assertEqual(bp["H2"].value, 140)
            self.assertEqual(bp["J2"].value, 80)

            sleep = book["睡眠"]
            self.assertEqual(sleep["C2"].value, "23:00:00")
            self.assertEqual(sleep["E2"].value, 2)
            self.assertEqual(sleep["E3"].value, 3.5)
            self.assertEqual(sleep["J2"].value, 8)
            self.assertEqual(sleep["H2"].value, "22:00:00")

            exercise_plot = book["運動 Plot"]
            self.assertEqual(
                [self._chart_title(chart) for chart in exercise_plot._charts],
                ["運動時間", "運動距離", "全部運動時間-運動日期", "全部運動距離-運動日期"],
            )
            self.assertEqual(self._title_size(exercise_plot._charts[0]), 2100)
            self.assertIn('cx="5760000"', self._drawing_xml(result.excel_path))
            self.assertIn('cy="3240000"', self._drawing_xml(result.excel_path))
            self.assertIn("2026-10-04", self._values(exercise_plot))
            self.assertIn(90, self._values(exercise_plot))
            self.assertEqual(
                [self._chart_title(chart) for chart in book["血壓 Plot"]._charts],
                ["高血壓-量血壓日期", "低血壓-量血壓日期", "心跳-量血壓日期"],
            )
            bp_values = self._values(book["血壓 Plot"])
            self.assertLess(bp_values.index("2026-10-04 08:00"), bp_values.index("2026-10-04 10:00"))
            self.assertEqual(
                [self._chart_title(chart) for chart in book["睡眠 Plot"]._charts],
                ["全部睡眠時間-日期", "全部睡眠次數-日期"],
            )
            self.assertIn(5.5, self._values(book["睡眠 Plot"]))
            self.assertIn(2, self._values(book["睡眠 Plot"]))
            book.close()

            html = result.plot_path.read_text(encoding="utf-8")
            for title in ("運動時間", "全部運動距離-運動日期", "高血壓-量血壓日期", "全部睡眠次數-日期"):
                self.assertIn(title, html)
            self.assertIn(str(style.width_px), html)

    def test_missing_folder_and_empty_data(self):
        with self.assertRaises(FileNotFoundError):
            app.process_folder(Path("Z:/folder-that-does-not-exist"))
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "notes.txt").write_text("x", encoding="utf-8")
            with self.assertRaises(FileNotFoundError):
                app.process_folder(root)
            wb = Workbook()
            wb.active.title = "運動"
            wb.active["A1"] = "選擇運動"
            wb.active["A2"] = "?"
            wb.save(root / "20260101.xlsx")
            wb.close()
            with self.assertRaises(ValueError):
                app.process_folder(root)
        with self.assertRaises(ValueError):
            app.validate_style(app.PlotStyle(width_cm=1))
        with self.assertRaises(ValueError):
            app.validate_style(app.PlotStyle(title_size=100))

    def _values(self, ws):
        values = []
        for row in ws.iter_rows(values_only=True):
            values.extend(value for value in row if value is not None)
        return values

    def _as_date(self, value):
        if isinstance(value, datetime):
            return value.date()
        return value

    def _chart_title(self, chart):
        return chart.title.tx.rich.p[0].r[0].t

    def _title_size(self, chart):
        return chart.title.tx.rich.p[0].r[0].rPr.sz

    def _drawing_xml(self, path):
        with zipfile.ZipFile(path) as archive:
            names = [name for name in archive.namelist() if name.startswith("xl/drawings/drawing")]
            return "\n".join(archive.read(name).decode("utf-8") for name in names)


class WaterTests(unittest.TestCase):
    def test_cup_fractions(self):
        self.assertAlmostEqual(app.parse_cups("1/6"), 1 / 6)
        self.assertAlmostEqual(app.parse_cups("1/2"), 0.5)
        self.assertAlmostEqual(app.parse_cups("5 1/2"), 5.5)
        self.assertAlmostEqual(app.parse_cups("1 1/2"), 1.5)
        self.assertEqual(app.parse_cups("1"), 1)
        self.assertEqual(app.parse_cups(2), 2)
        self.assertIsNone(app.parse_cups("?"))
        self.assertIsNone(app.parse_cups(None))

    def test_water_layout_and_plots(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first = Workbook()
            sheet = first.active
            sheet.title = "清潔"
            sheet["A1"] = "清潔\nMembersihkan"
            sheet["C1"] = "清潔日期\nTanggal pembersihan"
            sheet["E1"] = "清潔時間\nWaktu bersih-bersih"
            sheet["G1"] = "喝水量(杯)\nAsupan air (cangkir)"
            sheet["A2"] = "起床"
            sheet["C2"] = date(2026, 10, 4)
            sheet["E2"] = time(5, 32)
            sheet["G2"] = "1/6"
            sheet["A3"] = "喝水"
            sheet["C3"] = date(2026, 10, 4)
            sheet["E3"] = "09:24"
            sheet["G3"] = "1"
            sheet["A4"] = "喝水"
            sheet["C4"] = date(2026, 10, 4)
            sheet["E4"] = "12:00"
            sheet["G4"] = "1/2"
            sheet["A5"] = "洗澡"
            sheet["C5"] = date(2026, 10, 4)
            sheet["E5"] = "14:25"
            sheet["G5"] = "?"
            sheet["A6"] = "?"
            sheet["G7"] = "9"
            first.save(root / "20261004.xlsx")
            first.close()

            second = Workbook()
            clean = second.active
            clean.title = "清潔"
            clean["A1"] = "清潔"
            clean["A2"] = "喝水"
            clean["C2"] = date(2026, 10, 5)
            clean["E2"] = time(7, 0)
            clean["G2"] = "2"
            clean["A3"] = "?"
            second.save(root / "20261005.xlsx")
            second.close()

            result = app.process_folder(root, output_dir=root / "output")
            book = load_workbook(result.excel_path)
            self.assertEqual(
                book.sheetnames,
                ["喝水量", "喝水量 Plot", "清潔", "喝水時間", "喝水量(杯)"],
            )
            water = book["喝水量"]
            self.assertEqual([water.cell(1, col).value for col in range(1, 5)], ["喝水", "喝水日期", "喝水時間", "喝水量(杯)"])
            self.assertEqual([water.cell(row, 1).value for row in range(2, 6)], ["喝水", "喝水", "喝水", "喝水"])
            self.assertIsNone(water["A6"].value)
            self.assertIsNone(water["E1"].value)
            self.assertEqual(water["C2"].value, "05:32")
            self.assertEqual(water["C3"].value, "09:24")
            self.assertEqual(water["C4"].value, "12:00")
            self.assertEqual(water["C5"].value, "07:00")
            self.assertEqual(self._as_date(water["B5"].value), date(2026, 10, 5))
            self.assertAlmostEqual(water["D2"].value, 1 / 6)
            self.assertEqual(water["D3"].value, 1)
            self.assertAlmostEqual(water["D4"].value, 0.5)
            self.assertEqual(water["D5"].value, 2)
            self.assertNotIn(9, [cell.value for row in water.iter_rows() for cell in row])

            copied_dates = book["清潔"]
            copied_times = book["喝水時間"]
            copied_cups = book["喝水量(杯)"]
            self.assertEqual(copied_dates["A1"].value, "清潔日期 Tanggal pembersihan")
            self.assertIsNone(copied_dates["B1"].value)
            self.assertEqual(
                [self._as_date(copied_dates.cell(row, 1).value) for row in range(2, 6)],
                [date(2026, 10, 4), date(2026, 10, 4), date(2026, 10, 4), date(2026, 10, 5)],
            )
            self.assertIsNone(copied_dates["A6"].value)
            self.assertEqual([copied_times.cell(row, 1).value for row in range(2, 6)], ["05:32", "09:24", "12:00", "07:00"])
            self.assertAlmostEqual(copied_cups["A2"].value, 1 / 6)
            self.assertEqual(copied_cups["A3"].value, 1)
            self.assertAlmostEqual(copied_cups["A4"].value, 0.5)
            self.assertEqual(copied_cups["A5"].value, 2)
            self.assertNotIn(9, [cell.value for row in copied_cups.iter_rows() for cell in row])
            self.assertEqual(len(copied_dates._charts), 0)

            titles = [chart.title.tx.rich.p[0].r[0].t for chart in book["喝水量 Plot"]._charts]
            self.assertEqual(titles, ["喝水量-喝水日期", "全部喝水量-日期", "全部喝水量次數-日期"])
            plot_values = [value for row in book["喝水量 Plot"].iter_rows(values_only=True) for value in row if value is not None]
            self.assertEqual(
                [value for value in plot_values if isinstance(value, str) and value.startswith("2026-10-")],
                ["2026-10-04 05:32", "2026-10-04 09:24", "2026-10-04 12:00", "2026-10-05 07:00", "2026-10-04", "2026-10-05", "2026-10-04", "2026-10-05"],
            )
            self.assertIn(3, plot_values)
            self.assertIn(2, plot_values)
            self.assertTrue(any(isinstance(value, float) and abs(value - (1 / 6 + 1 + 0.5)) < 1e-9 for value in plot_values))
            book.close()

            html = result.plot_path.read_text(encoding="utf-8")
            self.assertIn("喝水量-喝水日期", html)
            self.assertIn("全部喝水量-日期", html)
            self.assertIn("全部喝水量次數-日期", html)

    def _as_date(self, value):
        if isinstance(value, datetime):
            return value.date()
        return value

    def test_columns_stop_above_blank_cup_and_named_sheets(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            first = Workbook()
            sheet = first.active
            sheet.title = "清潔"
            sheet["B1"] = "清潔日期\nTanggal pembersihan"
            sheet["D1"] = "清潔時間\nWaktu bersih-bersih"
            sheet["F1"] = "喝水量(杯)\nAsupan air (cangkir)"
            sheet["B2"] = date(2026, 11, 4)
            sheet["D2"] = "23:37"
            sheet["F2"] = "2.5"
            sheet["B3"] = date(2026, 9, 11)
            sheet["D3"] = "00:00"
            sheet["F4"] = "9"
            sheet["D4"] = "03:00"
            first.save(root / "20261016.xlsx")
            first.close()

            second = Workbook()
            clean = second.active
            clean.title = "清潔"
            clean["B1"] = "清潔日期\nTanggal pembersihan"
            clean["D1"] = "清潔時間\nWaktu bersih-bersih"
            clean["F1"] = "喝水量(杯)\nAsupan air (cangkir)"
            clean["B2"] = date(2026, 11, 6)
            clean["D2"] = time(8, 5)
            clean["F2"] = "1/4"
            clean["F3"] = "?"
            clean["B3"] = date(2026, 11, 7)
            second.save(root / "20261017.xlsx")
            second.close()

            result = app.process_folder(root, output_dir=root / "output")
            book = load_workbook(result.excel_path)
            dates = book["清潔"]
            times = book["喝水時間"]
            cups = book["喝水量(杯)"]
            self.assertEqual(self._as_date(dates["A2"].value), date(2026, 11, 4))
            self.assertEqual(self._as_date(dates["A3"].value), date(2026, 11, 6))
            self.assertIsNone(dates["A4"].value)
            self.assertIsNone(dates["B1"].value)
            self.assertEqual(times["A2"].value, "23:37")
            self.assertEqual(times["A3"].value, "08:05")
            self.assertEqual(cups["A2"].value, 2.5)
            self.assertAlmostEqual(cups["A3"].value, 0.25)
            water = book["喝水量"]
            self.assertEqual(water["A2"].value, "喝水")
            self.assertEqual(self._as_date(water["B2"].value), date(2026, 11, 4))
            self.assertEqual(water["C2"].value, "23:37")
            self.assertEqual(water["D2"].value, 2.5)
            self.assertEqual(water["A3"].value, "喝水")
            self.assertEqual(self._as_date(water["B3"].value), date(2026, 11, 6))
            self.assertEqual(water["C3"].value, "08:05")
            self.assertAlmostEqual(water["D3"].value, 0.25)
            self.assertIsNone(water["A4"].value)
            self.assertIsNone(water["E1"].value)
            plot_values = [value for row in book["喝水量 Plot"].iter_rows(values_only=True) for value in row if value is not None]
            self.assertEqual(
                [value for value in plot_values if isinstance(value, str)],
                ["喝水日期", "喝水量(杯)", "2026-11-04 23:37", "2026-11-06 08:05", "日期", "全部喝水量", "2026-11-04", "2026-11-06", "日期", "全部喝水量次數", "2026-11-04", "2026-11-06"],
            )
            self.assertNotIn(9, [cell.value for row in cups.iter_rows() for cell in row])
            self.assertNotIn("00:00", [cell.value for row in times.iter_rows() for cell in row])
            self.assertNotIn(date(2026, 9, 11), [self._as_date(cell.value) for row in dates.iter_rows() for cell in row])
            self.assertNotIn(date(2026, 10, 16), [self._as_date(cell.value) for row in water.iter_rows() for cell in row])
            book.close()

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            book = Workbook()
            dates = book.active
            dates.title = "清潔日期 Tanggal pembersihan"
            dates["A1"] = "清潔日期 Tanggal pembersihan"
            dates["A2"] = date(2026, 10, 16)
            dates["A3"] = date(2026, 10, 17)
            dates["A4"] = "?"
            dates["A5"] = date(2026, 1, 1)
            times = book.create_sheet("清潔時間 Waktu bersih-bersih")
            times["A1"] = "清潔時間 Waktu bersih-bersih"
            times["A2"] = "06:45"
            times["A3"] = ""
            times["A4"] = "18:00"
            cups = book.create_sheet("喝水量(杯) Asupan air (cangkir)")
            cups["A1"] = "喝水量(杯) Asupan air (cangkir)"
            cups["A2"] = "1/3"
            cups["A3"] = 2
            cups["A4"] = "?"
            book.save(root / "20261016.xlsx")
            book.close()

            result = app.process_folder(root, output_dir=root / "output")
            saved = load_workbook(result.excel_path)
            self.assertEqual(saved.sheetnames, ["清潔", "喝水時間", "喝水量(杯)"])
            self.assertEqual(self._as_date(saved["清潔"]["A2"].value), date(2026, 10, 16))
            self.assertEqual(self._as_date(saved["清潔"]["A3"].value), date(2026, 10, 17))
            self.assertIsNone(saved["清潔"]["A4"].value)
            self.assertEqual(saved["喝水時間"]["A2"].value, "06:45")
            self.assertIsNone(saved["喝水時間"]["A3"].value)
            self.assertAlmostEqual(saved["喝水量(杯)"]["A2"].value, 1 / 3)
            self.assertEqual(saved["喝水量(杯)"]["A3"].value, 2)
            self.assertNotIn(date(2026, 1, 1), [self._as_date(cell.value) for row in saved["清潔"].iter_rows() for cell in row])
            saved.close()


class ConditionSheetTests(unittest.TestCase):
    def test_excretion_and_body_without_plots(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._save(root / "20261004.xlsx", date(2026, 10, 4), time(3, 0), "少: Kurang", "好: Bagus")
            self._save(root / "20261005.xlsx", date(2026, 10, 5), time(8, 15), "多: lebih", "不好: tidak bagus")

            result = app.process_folder(root)
            book = load_workbook(result.excel_path)
            self.assertEqual(book.sheetnames, ["排洩", "身體情況"])
            self.assertFalse(any(name.endswith("Plot") for name in book.sheetnames))

            excretion = book["排洩"]
            self.assertEqual(
                [excretion.cell(1, col).value for col in range(1, 8)],
                [
                    "排洩日期  Tanggal ekskres",
                    "排洩時間  Waktu ekskresi",
                    "大便量 Volume feses)",
                    "小便量  Volume urin",
                    "小便顏色改變 Perubahan warna urin",
                    "拉肚子 Diare",
                    "肚子痛   Sakit perut",
                ],
            )
            self.assertEqual(
                [excretion.cell(2, col).value for col in range(1, 8)],
                ["2026-10-04", "03:00", "少: Kurang", "普通: Normal", "沒有: Tidak", "沒有: Tidak", "有: Ya"],
            )
            self.assertEqual(
                [excretion.cell(3, col).value for col in range(1, 8)],
                ["2026-10-05", "08:15", "多: lebih", "普通: Normal", "沒有: Tidak", "沒有: Tidak", "有: Ya"],
            )
            self.assertIsNone(excretion["A4"].value)
            self.assertIsNone(excretion["H1"].value)
            self.assertNotIn("不該讀", self._values(excretion))
            self.assertNotIn("\n", "".join(str(excretion.cell(1, col).value) for col in range(1, 8)))

            body = book["身體情況"]
            self.assertEqual(
                [body.cell(1, col).value for col in range(1, 7)],
                [
                    "身體情況日期 Tanggal kondisi fisik",
                    "睡眠 Tidur",
                    "哀叫 Ratapan",
                    " 情緒 emosi ",
                    "走路力量 kekuatan berjalan",
                    "說話聲音 suara",
                ],
            )
            self.assertEqual(
                [body.cell(2, col).value for col in range(1, 7)],
                ["2026-10-04", "好: Bagus", "沒有: Tidak", "好: Bagus", "有: Ya", "普通: Normal"],
            )
            self.assertEqual(
                [body.cell(3, col).value for col in range(1, 7)],
                ["2026-10-05", "不好: tidak bagus", "沒有: Tidak", "好: Bagus", "有: Ya", "普通: Normal"],
            )
            self.assertIsNone(body["A4"].value)
            self.assertIsNone(body["G1"].value)
            self.assertEqual(len(body._charts), 0)
            self.assertEqual(len(excretion._charts), 0)
            book.close()

            html = result.plot_path.read_text(encoding="utf-8")
            self.assertNotIn("排洩", html)
            self.assertNotIn("身體情況", html)
            self.assertIn("No plots", html)

    def test_condition_sheets_do_not_change_existing_plots(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            wb = Workbook()
            ws = wb.active
            ws.title = "運動"
            ws["A1"] = "選擇運動"
            ws["A2"] = "走路"
            ws["E2"] = time(8, 0)
            ws["G2"] = time(8, 30)
            ws["K2"] = "10米: 10 meter"
            ws["A3"] = "?"
            excretion = wb.create_sheet("排洩")
            excretion["A1"] = "排洩日期"
            excretion["E1"] = "大便量"
            excretion["A2"] = date(2026, 10, 6)
            excretion["E2"] = "普通: Normal"
            excretion["A3"] = "?"
            body = wb.create_sheet("身體情況")
            body["A1"] = "身體情況日期"
            body["C1"] = "睡眠"
            body["A2"] = date(2026, 10, 6)
            body["C2"] = "好: Bagus"
            wb.save(root / "20261006.xlsx")
            wb.close()

            result = app.process_folder(root)
            book = load_workbook(result.excel_path)
            self.assertEqual(book.sheetnames, ["運動", "運動 Plot", "排洩", "身體情況"])
            self.assertEqual(
                [chart.title.tx.rich.p[0].r[0].t for chart in book["運動 Plot"]._charts],
                ["運動時間", "運動距離", "全部運動時間-運動日期", "全部運動距離-運動日期"],
            )
            book.close()
            html = result.plot_path.read_text(encoding="utf-8")
            self.assertIn("運動時間", html)
            self.assertNotIn("排洩", html)
            self.assertNotIn("身體情況", html)

    def _save(self, path, when, clock, stool, sleep):
        wb = Workbook()
        excretion = wb.active
        excretion.title = "排洩"
        excretion["A1"] = "排洩日期 \nTanggal ekskres"
        excretion["C1"] = "排洩時間 \nWaktu ekskresi"
        excretion["E1"] = "大便量\nVolume feses)"
        excretion["G1"] = "小便量 \nVolume urin"
        excretion["I1"] = "小便顏色改變\nPerubahan warna urin"
        excretion["K1"] = "拉肚子\nDiare"
        excretion["M1"] = "肚子痛  \nSakit perut"
        excretion["A2"] = when
        excretion["C2"] = time(3, 0)
        excretion["E2"] = "?"
        excretion["G2"] = "?"
        excretion["A3"] = when
        excretion["C3"] = clock
        excretion["E3"] = stool
        excretion["G3"] = "普通: Normal"
        excretion["I3"] = "沒有: Tidak"
        excretion["K3"] = "沒有: Tidak"
        excretion["M3"] = "有: Ya"
        excretion["A4"] = "?"
        excretion["E4"] = "不該讀"
        body = wb.create_sheet("身體情況")
        body["A1"] = "身體情況日期\nTanggal kondisi fisik"
        body["C1"] = "睡眠\nTidur"
        body["E1"] = "哀叫\nRatapan"
        body["G1"] = " 情緒\nemosi "
        body["I1"] = "走路力量\nkekuatan berjalan"
        body["K1"] = "說話聲音\nsuara"
        body["A2"] = when
        body["C2"] = "?"
        body["A3"] = when
        body["C3"] = sleep
        body["E3"] = "沒有: Tidak"
        body["G3"] = "好: Bagus"
        body["I3"] = "有: Ya"
        body["K3"] = "普通: Normal"
        body["A4"] = "?"
        body["C4"] = "不該讀"
        wb.save(path)
        wb.close()

    def _values(self, ws):
        values = []
        for row in ws.iter_rows(values_only=True):
            values.extend(value for value in row if value is not None)
        return values

    def _as_date(self, value):
        if isinstance(value, datetime):
            return value.date()
        return value


class PageImageTests(unittest.TestCase):
    def test_rows_stop_above_question_or_blank(self):
        rows = [
            ["選擇運動", None, "運動日期", None, "開始時間", None, None, None, None, None, "運動距離"],
            ["走路", None, date(2026, 10, 16), None, "08:00:00", None, None, None, None, None, "10米: 10 meter"],
            ["跑步", None, date(2026, 10, 16), None, "09:00:00", None, None, None, None, None, "20米: 20 meter"],
            ["模板", None, date(2026, 10, 16), None, "10:00:00", None, None, None, None, None, "?"],
            ["不該讀", None, date(2026, 10, 16), None, "11:00:00", None, None, None, None, None, "30米: 30 meter"],
        ]
        header, kept = app.extract_page_table(rows, (3, 5, 11), 11)
        self.assertEqual(header[0].split("\n")[0], "運動日期")
        self.assertEqual(
            kept,
            [
                ["2026-10-16", "08:00:00", "10米: 10 meter"],
                ["2026-10-16", "09:00:00", "20米: 20 meter"],
            ],
        )

        clean = [
            [None, "清潔日期", None, "清潔時間", None, "喝水量(杯)"],
            [None, date(2026, 10, 16), None, "23:37", None, "1.5"],
            [None, date(2026, 10, 16), None, "07:15", None, "1/2"],
            [None, date(2026, 9, 11), None, "00:00", None, None],
            [None, date(2026, 9, 11), None, "03:00", None, "9"],
            [None, date(2026, 9, 11), None, "00:00:00", None, "?"],
        ]
        _header, water = app.extract_page_table(clean, (2, 4, 6), 6)
        self.assertEqual(water, [["2026-10-16", "23:37", "1.5"], ["2026-10-16", "07:15", "1/2"]])

        excretion = [
            ["排洩日期", None, None, None, "大便量", None, "小便量", None, "顏色", None, "拉肚子", None, "肚子痛"],
            [date(2026, 10, 16), None, None, None, "多: lebih", None, "普通: Normal", None, "沒有: Tidak", None, "沒有: Tidak", None, "沒有: Tidak"],
            [date(2026, 9, 16), None, None, None, "?", None, "?", None, "不該讀", None, "不該讀", None, "不該讀"],
            [date(2026, 9, 17), None, None, None, "", None, "", None, "也不該讀", None, "也不該讀", None, "也不該讀"],
        ]
        _header, body = app.extract_page_table(excretion, (1, 5, 7), 5)
        self.assertEqual(body, [["2026-10-16", "多: lebih", "普通: Normal"]])
        _header, second = app.extract_page_table(excretion, (1, 9, 11, 13), 5)
        self.assertEqual(second, [["2026-10-16", "沒有: Tidak", "沒有: Tidak", "沒有: Tidak"]])
        self.assertNotIn("不該讀", str(second))

        # Template rows keep a calculated 00:00:00 in the stop column. '?' in column A ends the page.
        sleep = [
            [None, None, "日期", None, "開始時間", None, "結束時間", None, "睡眠時間"],
            ["半夜: Kerja malam", None, date(2026, 11, 7), None, "19:10:00", None, "19:55:00", None, "00:45:00"],
            ["半夜: Kerja malam", None, date(2026, 11, 7), None, "05:40:00", None, "07:00:00", None, "00:00:00"],
            ["?", None, date(2026, 9, 17), None, "01:03", None, "01:03", None, "00:00:00"],
            ["?", None, date(2026, 9, 17), None, "01:03", None, "01:03", None, "01:00:00"],
        ]
        _header, sleep_rows = app.extract_page_table(sleep, (3, 5, 7, 9), 9)
        self.assertEqual(
            sleep_rows,
            [
                ["2026-11-07", "19:10:00", "19:55:00", "00:45:00"],
                ["2026-11-07", "05:40:00", "07:00:00", "00:00:00"],
            ],
        )

    def test_files_stack_in_date_order_and_diet_table(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._save(root / "20261017.xlsx", date(2026, 10, 17), "晚餐食物: Makanan makan malam", "1/2")
            self._save(root / "20261016.xlsx", date(2026, 10, 16), "早餐食物: Makanan sarapan", "1 1/4")
            groups = {group["stem"]: group for group in app.collect_page_groups(app.list_dated_workbooks(root))}
            exercise = groups["運動"]["tables"][0]["rows"]
            self.assertEqual([row[0] for row in exercise], ["2026-10-16", "2026-10-17"])
            diet = groups["飲食"]["tables"]
            self.assertEqual(len(diet), 1)
            self.assertFalse(diet[0]["show_letters"])
            self.assertFalse(diet[0]["show_index"])
            self.assertEqual(diet[0]["header"][0].split("\n")[0], "日期")
            self.assertEqual(
                [column.split("\n")[0] for column in diet[0]["header"][1:]],
                [
                    "早餐食物: Makanan sarapan",
                    "晚餐食物: Makanan makan malam",
                    "一天食物 Makanan untuk satu hari",
                ],
            )
            self.assertEqual(
                diet[0]["rows"],
                [
                    [
                        "",
                        "早餐食物: Makanan sarapan",
                        "晚餐食物: Makanan makan malam",
                        "一天食物 Makanan untuk satu hari",
                    ],
                    ["2026-10-16", "1 1/4", "", "1 1/4"],
                    ["2026-10-17", "", "1/2", "1/2"],
                ],
            )
            self.assertEqual(groups["睡眠"]["tables"][0]["letters"], ["C", "E", "G", "I", ""])
            self.assertEqual(groups["睡眠"]["tables"][0]["header"][-1].split("\n")[0], "一天睡眠時間")
            self.assertEqual(groups["睡眠"]["tables"][0]["rows"][0][-1], "00:01:01")
            self.assertEqual(
                groups["排洩1"]["tables"][0]["rows"],
                [
                    ["2026-10-16", "多: lebih", "普通: Normal"],
                    ["2026-10-17", "多: lebih", "普通: Normal"],
                ],
            )
            self.assertEqual(groups["排洩2"]["tables"][0]["letters"], ["A", "I", "K", "M"])
            self.assertEqual(groups["血壓1"]["tables"][0]["letters"], ["A", "C", "E", "F"])
            self.assertEqual(groups["血壓2"]["tables"][0]["rows"][0][2], "121")
            self.assertEqual(groups["身體情況1"]["tables"][0]["letters"], ["A", "C", "E", "G"])
            self.assertEqual(groups["身體情況2"]["tables"][0]["rows"][0][1:], ["有: Ya", "普通: Normal"])
            self.assertNotIn("不該讀", str(groups["運動"]["tables"][0]["rows"]))

    def test_diet_pivot_is_one_row_per_date(self):
        table = {
            "header": [
                "飲食健康\nMakan sehat",
                "日期\nTanggal",
                "食物份量(碗)\nPorsi makanan (mangkuk)",
            ],
            "rows": [
                ["晚餐食物: Makanan makan malam", "2026-11-05", "1 1/2"],
                ["早餐食物: Makanan sarapan", "2026-11-05", "2"],
                ["午餐食物: Makanan makan siang", "2026-11-05", "2"],
                ["早餐食物: Makanan sarapan", "2026-11-06", "1 1/2"],
                ["午餐食物: Makanan makan siang", "2026-11-06", "2"],
                ["晚餐食物: Makanan makan malam", "2026-11-06", "2"],
                ["早餐食物: Makanan sarapan", "2026-11-07", "1 1/2"],
                ["午餐食物: Makanan makan siang", "2026-11-07", "1 1/2"],
                ["晚餐食物: Makanan makan malam", "2026-11-07", "2"],
            ],
            "letters": ["A", "E", "G"],
        }
        app.pivot_diet_table(table)
        self.assertEqual(
            [column.split("\n")[0] for column in table["header"]],
            [
                "日期",
                "早餐食物: Makanan sarapan",
                "午餐食物: Makanan makan siang",
                "晚餐食物: Makanan makan malam",
                "一天食物 Makanan untuk satu hari",
            ],
        )
        self.assertEqual(table["header"][1].split("\n")[1], "食物份量(碗) Porsi makanan (mangkuk)")
        self.assertEqual(table["header"][-1].split("\n")[1], "食物份量(碗) Porsi makanan (mangkuk)")
        self.assertEqual(
            table["rows"],
            [
                [
                    "",
                    "早餐食物: Makanan sarapan",
                    "午餐食物: Makanan makan siang",
                    "晚餐食物: Makanan makan malam",
                    "一天食物 Makanan untuk satu hari",
                ],
                ["2026-11-05", "2", "2", "1 1/2", "5 1/2"],
                ["2026-11-06", "1 1/2", "2", "2", "5 1/2"],
                ["2026-11-07", "1 1/2", "1 1/2", "2", "5"],
            ],
        )

    def test_food_sheet_uses_extracted_diet_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._save(root / "20261017.xlsx", date(2026, 10, 17), "晚餐食物: Makanan makan malam", "1/2")
            self._save(root / "20261016.xlsx", date(2026, 10, 16), "早餐食物: Makanan sarapan", "1 1/4")
            book = Workbook()
            book.active.title = "運動"
            app.add_food_sheet(book, app.list_dated_workbooks(root))
            self.assertEqual(book.sheetnames, ["運動", "食物"])
            sheet = book["食物"]
            self.assertEqual(sheet["A1"].value.split("\n")[0], "日期")
            self.assertEqual(sheet["D1"].value.split("\n")[0], "一天食物 Makanan untuk satu hari")
            self.assertEqual(sheet["B2"].value, "早餐食物: Makanan sarapan")
            self.assertEqual(sheet["D2"].value, "一天食物 Makanan untuk satu hari")
            self.assertEqual(sheet["A3"].value, "2026-10-16")
            self.assertEqual(sheet["B3"].value, "1 1/4")
            self.assertEqual(sheet["C3"].value, None)
            self.assertEqual(sheet["D3"].value, "1 1/4")
            self.assertEqual(sheet["A4"].value, "2026-10-17")
            self.assertEqual(sheet["C4"].value, "1/2")
            self.assertEqual(sheet["D4"].value, "1/2")
            book.close()

    def test_food_plot_reads_the_food_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._save(root / "20261017.xlsx", date(2026, 10, 17), "晚餐食物: Makanan makan malam", "1/2")
            self._save(root / "20261016.xlsx", date(2026, 10, 16), "早餐食物: Makanan sarapan", "1 1/4")
            book = Workbook()
            book.active.title = "運動"
            app.add_food_sheet(book, app.list_dated_workbooks(root))
            spec = app.food_plot_from_sheet(book)
            book.close()
            self.assertEqual(spec["title"], "食物-日期")
            self.assertEqual(spec["xlabel"], "日期 Tanggal")
            self.assertEqual(spec["ylabel"], "食物份量(碗)")
            self.assertEqual(spec["categories"], ["2026-10-16", "2026-10-17"])
            self.assertEqual(
                spec["series"],
                [
                    ("早餐食物: Makanan sarapan", [1.25, None]),
                    ("午餐食物: Makanan makan siang", [None, None]),
                    ("晚餐食物: Makanan makan malam", [None, 0.5]),
                    ("一天食物 Makanan untuk satu hari", [1.25, 0.5]),
                ],
            )

    def test_sleep_page_adds_daily_total_on_the_first_row(self):
        table = {
            "letters": ["C", "E", "G", "I"],
            "header": [
                "日期\nTanggal",
                "開始時間\nWaktu mulai",
                "結束時間\nWaktu berakhir",
                "睡眠時間\nWaktu berolahraga",
            ],
            "rows": [
                ["2026-11-05", "20:30:00", "23:30:00", "03:00:00"],
                ["2026-11-05", "23:30:00", "01:30:00", "02:00:00"],
                ["2026-11-05", "01:40:00", "06:40:00", "05:00:00"],
                ["2026-11-06", "20:00:00", "20:40:00", "00:40:00"],
                ["2026-11-06", "21:10:00", "23:20:00", "02:10:00"],
                ["2026-11-06", "23:30:00", "01:30:00", "02:00:00"],
                ["2026-11-06", "01:45:00", "02:20:00", "00:35:00"],
                ["2026-11-06", "02:45:00", "06:20:00", "03:35:00"],
                ["2026-11-07", "19:10:00", "19:55:00", "00:45:00"],
                ["2026-11-07", "20:00:00", "20:40:00", "00:40:00"],
                ["2026-11-07", "21:10:00", "21:22:00", "00:12:00"],
                ["2026-11-07", "21:32:00", "01:00:00", "03:28:00"],
                ["2026-11-07", "01:10:00", "05:20:00", "04:10:00"],
                ["2026-11-07", "05:40:00", "07:00:00", "01:20:00"],
            ],
        }
        app.add_sleep_day_total(table)
        self.assertEqual(table["header"][-1], "一天睡眠時間\nWaktu tidur per hari")
        self.assertEqual(table["rows"][0][-1], "10:00:00")
        self.assertEqual(table["rows"][1][-1], "")
        self.assertEqual(table["rows"][2][-1], "")
        self.assertEqual(table["rows"][3][-1], "09:00:00")
        self.assertEqual(table["rows"][7][-1], "")
        self.assertEqual(table["rows"][8][-1], "10:35:00")
        self.assertEqual(table["rows"][13][-1], "")
        self.assertEqual([row[3] for row in table["rows"][:3]], ["03:00:00", "02:00:00", "05:00:00"])

    def test_exercise_page_adds_daily_distance_on_the_first_row(self):
        table = {
            "letters": ["C", "E", "K"],
            "header": [
                "運動日期\nTanggal Olahraga",
                "開始時間\nWaktu mulai",
                "運動距離\nJarak pergerakan",
            ],
            "rows": [
                ["2026-11-05", "07:50:00", "15米: 15 meter"],
                ["2026-11-05", "09:20:00", "15米: 15 meter"],
                ["2026-11-05", "13:16:00", "15米: 15 meter"],
                ["2026-11-05", "16:16:00", "15米: 15 meter"],
                ["2026-11-06", "07:50:00", "15米: 15 meter"],
                ["2026-11-06", "11:50:00", "15米: 15 meter"],
                ["2026-11-06", "12:10:00", "15米: 15 meter"],
                ["2026-11-06", "13:40:00", "15米: 15 meter"],
                ["2026-11-06", "15:30:00", "15米: 15 meter"],
                ["2026-11-07", "08:20:00", "15米: 15 meter"],
                ["2026-11-07", "09:00:00", "15米: 15 meter"],
                ["2026-11-07", "12:52:00", "15米: 15 meter"],
                ["2026-11-07", "14:52:00", "15米: 15 meter"],
                ["2026-11-07", "16:53:00", "15米: 15 meter"],
                ["2026-11-07", "18:40:00", "15米: 15 meter"],
            ],
        }
        app.add_exercise_day_total(table)
        self.assertEqual(table["header"][-1], "一天運動距離\nJarak olahraga harian")
        self.assertEqual(table["rows"][0][-1], "60米: 60 meter")
        self.assertEqual(table["rows"][1][-1], "")
        self.assertEqual(table["rows"][3][-1], "")
        self.assertEqual(table["rows"][4][-1], "75米: 75 meter")
        self.assertEqual(table["rows"][8][-1], "")
        self.assertEqual(table["rows"][9][-1], "90米: 90 meter")
        self.assertEqual(table["rows"][14][-1], "")
        self.assertEqual(table["rows"][0][2], "15米: 15 meter")

    def test_water_sheet_page_follows_diet(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            book = Workbook()
            sheet = book.active
            sheet.title = "喝水量"
            sheet.append(["喝水", "喝水日期", "喝水時間", "喝水量(杯)"])
            sheet.append(["喝水", date(2026, 11, 5), "20:00", 5.5])
            sheet.append(["喝水", date(2026, 11, 6), "20:00", 5])
            sheet.append(["喝水", date(2026, 11, 7), "20:00", 5.5])
            path = root / "Extract.xlsx"
            book.save(path)
            book.close()
            groups = [
                {"stem": "飲食", "title": "飲食", "tables": []},
                {"stem": "睡眠", "title": "睡眠", "tables": []},
            ]
            placed = app._insert_water_page(groups, path)
            self.assertEqual([group["stem"] for group in placed], ["飲食", "喝水量", "睡眠"])
            rows = placed[1]["tables"][0]["rows"]
            self.assertEqual(rows[0], ["喝水", "2026-11-05", "20:00", "5.50"])
            self.assertEqual(rows[1][3], "5.00")
            self.assertEqual(rows[2], ["喝水", "2026-11-07", "20:00", "5.50"])

    def test_pack_splits_when_the_next_row_does_not_fit(self):
        block = {"letter_h": 26, "header_h": 52, "row_heights": [33] * 5}
        pages = app.pack_page_chunks([block], body_height=26 + 52 + 33 * 3, gap=22)
        self.assertEqual(len(pages), 2)
        self.assertEqual([(chunk["row_start"], chunk["row_end"]) for chunk in pages[0]], [(0, 3)])
        self.assertEqual([(chunk["row_start"], chunk["row_end"]) for chunk in pages[1]], [(3, 5)])

        short = {"letter_h": 26, "header_h": 52, "row_heights": [33]}
        shared = app.pack_page_chunks([short, short], body_height=400, gap=22)
        self.assertEqual(len(shared), 1)
        self.assertEqual(len(shared[0]), 2)
        tight = app.pack_page_chunks([short, short], body_height=26 + 52 + 33, gap=22)
        self.assertEqual(len(tight), 2)

    def test_png_pages_are_landscape_and_pdf_uses_excel_name(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self._save(root / "20261016.xlsx", date(2026, 10, 16), "早餐食物: Makanan sarapan", "1 1/4")
            result = app.process_folder(root)
            self.assertEqual(result.pdf_path, root / "output" / "Extract.pdf")
            self.assertTrue(result.pdf_path.is_file())
            exercise_pages = sorted((root / "output").glob("運動_*.png"))
            self.assertEqual([path.name for path in exercise_pages], ["運動_1.png"])
            from PIL import Image

            image = Image.open(exercise_pages[0])
            self.assertEqual(image.size, (1600, 1132))
            self.assertGreater(image.size[0], image.size[1])
            image.close()
            self.assertFalse((root / "output" / "清潔_1.png").is_file())
            for stem in ("飲食_1.png", "睡眠_1.png", "排洩1_1.png", "排洩2_1.png", "血壓1_1.png", "血壓2_1.png", "身體情況1_1.png", "身體情況2_1.png"):
                self.assertTrue((root / "output" / stem).is_file(), stem)
            self.assertIn(str(result.pdf_path), "\n".join(result.messages))

            wide = Workbook()
            sheet = wide.active
            sheet.title = "運動"
            sheet["C1"] = "運動日期"
            sheet["E1"] = "開始時間"
            sheet["K1"] = "運動距離"
            for row in range(2, 42):
                sheet.cell(row, 3, date(2026, 10, 16))
                sheet.cell(row, 5, f"{row:02d}:00:00")
                sheet.cell(row, 11, f"{row}米: {row} meter")
            sheet.cell(42, 1, "?")
            sheet.cell(43, 5, "不該讀")
            wide.save(root / "20261018.xlsx")
            wide.close()
            notes, pdf_path = app.export_page_files(app.list_dated_workbooks(root), root / "output", "Extract")
            pages = sorted((root / "output").glob("運動_*.png"))
            self.assertGreaterEqual(len(pages), 2)
            self.assertTrue(all(page.name.startswith("運動_") for page in pages))
            self.assertTrue(pdf_path.is_file())
            self.assertTrue(any("運動_2.png" in note for note in notes))
            book = load_workbook(result.excel_path)
            self.assertIn("運動", book.sheetnames)
            book.close()

    def _save(self, path, when, meal, amount):
        wb = Workbook()
        exercise = wb.active
        exercise.title = "運動"
        exercise["C1"] = "運動日期\nTanggal Olahraga"
        exercise["E1"] = "開始時間\nWaktu mulai"
        exercise["K1"] = "運動距離\nJarak pergerakan"
        exercise["C2"] = when
        exercise["E2"] = "23:36:57"
        exercise["K2"] = "10米: 10 meter"
        exercise["A3"] = "?"
        exercise["K4"] = "不該讀"

        clean = wb.create_sheet("清潔")
        clean["B1"] = "清潔日期\nTanggal pembersihan"
        clean["D1"] = "清潔時間\nWaktu bersih-bersih"
        clean["F1"] = "喝水量(杯)\nAsupan air (cangkir)"
        clean["B2"] = when
        clean["D2"] = "23:37"
        clean["F2"] = "1.5"
        clean["B3"] = date(2026, 9, 11)
        clean["D3"] = "00:00"

        diet = wb.create_sheet("飲食")
        diet["A1"] = "飲食健康\nMakan sehat"
        diet["C1"] = "食物\nMakanan"
        diet["E1"] = "日期\nTanggal"
        diet["G1"] = "食物份量(碗)\nPorsi makanan (mangkuk)"
        diet["I1"] = "2 主食: makanan pokok"
        diet["A2"] = meal
        diet["E2"] = when
        diet["G2"] = amount
        if amount == "1/2":
            diet["C2"] = "麵: mi"
            diet["I2"] = "白米: Beras"
        diet["A3"] = "?"

        sleep = wb.create_sheet("睡眠")
        sleep["C1"] = "日期\nTanggal"
        sleep["E1"] = "開始時間\nWaktu mulai"
        sleep["G1"] = "結束時間\nWaktu berakhir"
        sleep["I1"] = "睡眠時間\nWaktu berolahraga"
        sleep["C2"] = when
        sleep["E2"] = "23:38:38"
        sleep["G2"] = "23:39:39"
        sleep["I2"] = "00:01:01"
        sleep["A3"] = "?"

        excretion = wb.create_sheet("排洩")
        excretion["A1"] = "排洩日期\nTanggal ekskres"
        excretion["E1"] = "大便量\nVolume feses)"
        excretion["G1"] = "小便量\nVolume urin"
        excretion["I1"] = "小便顏色改變\nPerubahan warna urin"
        excretion["K1"] = "拉肚子\nDiare"
        excretion["M1"] = "肚子痛\nSakit perut"
        excretion["A2"] = when
        excretion["E2"] = "多: lebih"
        excretion["G2"] = "普通: Normal"
        excretion["I2"] = "沒有: Tidak"
        excretion["K2"] = "沒有: Tidak"
        excretion["M2"] = "沒有: Tidak"
        excretion["A3"] = date(2026, 9, 16)
        excretion["E3"] = "?"
        excretion["G3"] = "?"

        pressure = wb.create_sheet("血壓")
        pressure["A1"] = "健康\nSehat"
        pressure["C1"] = "量血壓日期\nTanggal pengukuran tekanan darah"
        pressure["E1"] = "收縮壓, 高壓\nTekanan darah sistolik\ntekanan tinggi"
        pressure["F1"] = "舒張壓, 低壓\nTekanan darah diastolik\ntekanan rendah"
        pressure["G1"] = "心跳\nDetak jantung"
        pressure["A2"] = "早餐前量, 量血壓: Ukur tekanan darah"
        pressure["C2"] = when
        pressure["E2"] = 113
        pressure["F2"] = 119
        pressure["G2"] = 121
        pressure["A3"] = "?"

        body = wb.create_sheet("身體情況")
        body["A1"] = "身體情況日期\nTanggal kondisi fisik"
        body["C1"] = "睡眠\nTidur"
        body["E1"] = "哀叫\nRatapan"
        body["G1"] = "情緒\nemosi"
        body["I1"] = "走路力量\nkekuatan berjalan"
        body["K1"] = "說話聲音\nsuara"
        body["A2"] = when
        body["C2"] = "好: Bagus"
        body["E2"] = "沒有: Tidak"
        body["G2"] = "好: Bagus"
        body["I2"] = "有: Ya"
        body["K2"] = "普通: Normal"
        body["A3"] = date(2026, 9, 16)
        body["C3"] = "?"
        body["E3"] = "?"
        body["G3"] = "?"
        body["I3"] = "?"
        body["K3"] = "?"
        wb.save(path)
        wb.close()


class GmailSendTests(unittest.TestCase):
    def test_receivers_match_app7(self):
        self.assertEqual(
            app.GMAIL_TO,
            "ghpiu0661@gmail.com,ghjane5168@gmail.com,a290772@gmail.com",
        )
        url = app.gmail_compose_url(app.GMAIL_TO, app.GMAIL_SUBJECT, "body")
        self.assertIn("ghpiu0661%40gmail.com", url)
        self.assertIn("ghjane5168%40gmail.com", url)
        self.assertIn("a290772%40gmail.com", url)
        self.assertIn("mail.google.com/mail/u/0/", url)

    def test_output_files_and_compose_script(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            output = root / "output"
            output.mkdir()
            excel = output / "Extract.xlsx"
            pdf = output / "Extract.pdf"
            html_path = output / "Extract_plots.html"
            excel.write_bytes(b"xlsx")
            pdf.write_bytes(b"pdf")
            html_path.write_bytes(b"html")
            files = app.gmail_output_files(root)
            self.assertEqual([path.name for path in files], ["Extract.xlsx", "Extract.pdf", "Extract.html"])
            self.assertEqual((output / "Extract.html").read_bytes(), b"html")

            (output / "Extract.html").write_bytes(b"named")
            files = app.gmail_output_files(root)
            self.assertEqual(files[2].read_bytes(), b"named")

            captured = {}

            class _Proc:
                stdout = "GMAIL_OK"
                stderr = ""

            def runner(command, **kwargs):
                captured["command"] = command
                captured["kwargs"] = kwargs
                captured["script"] = Path(command[-1]).read_text(encoding="utf-8-sig")
                return _Proc()

            names = app.send_gmail_files(files, runner=runner)
            script = captured["script"]
            self.assertEqual(captured["command"][-2], "-File")
            self.assertGreaterEqual(captured["kwargs"]["timeout"], 60)
            self.assertIn("ghpiu0661%40gmail.com", script)
            self.assertIn("ghjane5168%40gmail.com", script)
            self.assertIn("a290772%40gmail.com", script)
            self.assertIn("Extract.xlsx", script)
            self.assertIn("Extract.pdf", script)
            self.assertIn("Extract.html", script)
            self.assertIn("SetFileDropList", script)
            self.assertIn("SetForegroundWindow", script)
            self.assertIn("mouse_event", script)
            self.assertIn("SendWait('^v')", script)
            self.assertIn("Gmail compose window was not found", script)
            self.assertEqual(names.splitlines(), ["Extract.xlsx", "Extract.pdf", "Extract.html"])

            (output / "Extract.pdf").unlink()
            with self.assertRaises(FileNotFoundError):
                app.gmail_output_files(root)

    def test_gmail_button_is_right_of_process(self):
        window = app.App()
        try:
            self.assertEqual(str(window.run_button.cget("text")), "Process")
            self.assertEqual(str(window.gmail_button.cget("text")), "Gmail 傳送 Excel")
            self.assertEqual(str(window.gmail_button.master), str(window.run_button.master))
            self.assertEqual(str(window.run_button.pack_info()["side"]), "left")
            self.assertEqual(str(window.gmail_button.pack_info()["side"]), "left")
        finally:
            window.destroy()


if __name__ == "__main__":
    unittest.main()
