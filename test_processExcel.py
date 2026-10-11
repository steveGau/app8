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


if __name__ == "__main__":
    unittest.main()
