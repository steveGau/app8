# processExcel (app8)

Windows desktop app that reads dated Excel workbooks and writes one summary workbook plus Plotly charts. Sheet titles and source labels are Traditional Chinese and Indonesian.

This README describes the **current** app8 tree on `main`. The previous release stays on GitHub as tag **`v1`**. That tag is not moved, and `main` is not force-pushed, so the old tree remains reachable.

| Tag | Where |
|---|---|
| **`v1`** | Previous processExcel release (commit `ff2a2d3`). Sheets 運動, 血壓, and 睡眠 only. Checkout this tag to open the old version. |
| **`v2`** | This tree: 喝水量 from 清潔, plus 排洩 and 身體情況 tables with no plots. `main` points here. |

```bat
git fetch origin tag v1
git checkout v1
```

Return to the new version with `git checkout main`.

| | |
|---|---|
| Program | `processExcel.py` |
| Launch | Double-click `processExcel.bat` |
| Input | Excel files named `YYYYMMDD.xlsx` |
| Output | `output\Extract.xlsx` and `output\Extract_plots.html` inside the selected folder |

## Requirements

- Windows 10/11
- Python 3.12 or newer, with tkinter (`py -3.12` or `python` on PATH)
- Packages in `requirements.txt`: `openpyxl`, `plotly`

`processExcel.bat` creates `.venv` in this folder and installs those packages the first time it runs. A `.venv` copied from another computer is rebuilt automatically, because that environment still points at the other PC’s Python.

## Run

Double-click **`processExcel.bat`**.

1. Click **Browse** and choose the folder that contains the dated workbooks.
2. Set chart width, height, and font sizes if you want sizes other than the defaults.
3. Click **Process**.

From a command prompt, you can skip the window and pass the folder:

```bat
processExcel.bat "C:\path\to\folder"
```

## What it reads

Only files named like `20261005.xlsx` are used. They are sorted by date. Other `.xlsx` names are skipped. In each workbook the app reads from row 2 down to the row above `?` or `？`. Empty cells and missing sheets are skipped. `?` cells are left blank.

| Sheet | Output sheet | Kept columns | Notes |
|---|---|---|---|
| 運動 | 運動, plus 運動 Plot | 選擇運動, 運動日期, 運動時間, 運動距離 | 運動時間 is minutes. 運動距離 such as `10米: 10 meter` is `10` meters. |
| 血壓 | 血壓, plus 血壓 Plot | 健康, 量血壓日期, 收縮壓, 舒張壓, 心跳 | Numbers as stored. |
| 睡眠 | 睡眠, plus 睡眠 Plot | 睡眠, 日期, 開始時間, 結束時間, 睡眠時間 | 睡眠時間 is hours (end minus start, including overnight). |
| 清潔 | 喝水量, plus 喝水量 Plot | 喝水, 喝水日期, 喝水時間, 喝水量(杯) | Time is `hh:mm`. Cup amounts such as `1/6` are fractions. |
| 排洩 | 排洩 only | 排洩日期, 排洩時間, 大便量, 小便量, 小便顏色改變, 拉肚子, 肚子痛 | One table. Dates are `yyyy-mm-dd`. Time is `hh:mm`. No plot sheet and no HTML chart. |
| 身體情況 | 身體情況 only | 身體情況日期, 睡眠, 哀叫, 情緒, 走路力量, 說話聲音 | One table. No plot sheet and no HTML chart. |

運動, 血壓, 睡眠, and 喝水量 place each dated file side by side in `Extract.xlsx`. 排洩 and 身體情況 use one header row, with every file’s rows stacked underneath. Titles on those two sheets are a single line (Chinese and Indonesian). A row is written only when a measurement or observation is filled in, so a template row that is only a date, a time, and `?` is skipped.

The same 運動, 血壓, 睡眠, and 喝水量 charts are written to `Extract_plots.html`. 排洩 and 身體情況 are not in that file.

## Sample

`testExample\` has dated workbooks. Processing that folder writes `testExample\output\`.

```bat
processExcel.bat "C:\Users\User\Desktop\app8\testExample"
```

## Tests

```bat
.\.venv\Scripts\python.exe -m unittest test_processExcel.py
```

## Not committed

`.venv\` and `__pycache__\` stay on this PC. Run `processExcel.bat` to create the environment here.
