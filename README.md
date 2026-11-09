# processExcel (app8)

Windows desktop app that reads dated Excel workbooks and writes one summary workbook, Plotly charts, landscape page images, and one PDF. Sheet titles and source labels are Traditional Chinese and Indonesian.

This README describes the **current** app8 tree on `main`. Older releases stay on GitHub as tags. Those tags are not moved, and `main` is not force-pushed, so each old tree remains reachable.

| Tag | Where |
|---|---|
| **`v1`** | First processExcel release (commit `ff2a2d3`). Sheets 運動, 血壓, and 睡眠 only. |
| **`v2`** | 喝水量 from 清潔, plus 排洩 and 身體情況 tables with no plots (commit `6b57b7c`). |
| **`v3`** | Landscape PNG pages, `Extract.pdf`, and **Gmail 傳送 Excel** (commit `be9b652`). |
| **`v4`** | This tree. Diet, sleep, exercise, and water pages described below, plus the 食物 sheet and chart. `main` points here. |

```bat
git fetch origin tag v1
git checkout v1
```

The same commands work for `v2` and `v3`. Return to the new version with `git checkout main`.

| | |
|---|---|
| Program | `processExcel.py` |
| Launch | Double-click `processExcel.bat` |
| Input | Excel files named `YYYYMMDD.xlsx` |
| Output | `output\Extract.xlsx`, `output\Extract_plots.html`, landscape PNG pages, and `output\Extract.pdf` |

## Requirements

- Windows 10/11
- Python 3.12 or newer, with tkinter (`py -3.12` or `python` on PATH)
- Packages in `requirements.txt`: `openpyxl`, `plotly`, and `pillow`

`processExcel.bat` creates `.venv` in this folder and installs those packages the first time it runs. A `.venv` copied from another computer is rebuilt automatically, because that environment still points at the other PC’s Python.

## Run

Double-click **`processExcel.bat`**.

1. Click **Browse** and choose the folder that contains the dated workbooks.
2. Set chart width, height, and font sizes if you want sizes other than the defaults.
3. Click **Process**.
4. Click **Gmail 傳送 Excel** (to the right of Process) to open a Gmail compose window with `Extract.xlsx`, `Extract.pdf`, and `Extract.html` pasted in. Confirm the three files are attached, then send.

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
| 清潔 | 喝水量, plus 喝水量 Plot | 喝水, 喝水日期, 喝水時間, 喝水量(杯) | Time is `hh:mm`. `1/6` is one sixth of a cup. `5 1/2` is 5.5 cups. |
| 飲食 | 食物 | 日期, 早餐食物, 午餐食物, 晚餐食物, 一天食物 | One row per date. 一天食物 is the sum of that date’s portions. The same table is the 食物-日期 chart in `Extract_plots.html`. |
| 排洩 | 排洩 only | 排洩日期, 排洩時間, 大便量, 小便量, 小便顏色改變, 拉肚子, 肚子痛 | One table. Dates are `yyyy-mm-dd`. Time is `hh:mm`. No plot sheet and no HTML chart. |
| 身體情況 | 身體情況 only | 身體情況日期, 睡眠, 哀叫, 情緒, 走路力量, 說話聲音 | One table. No plot sheet and no HTML chart. |

運動, 血壓, 睡眠, and 喝水量 place each dated file side by side in `Extract.xlsx`. 排洩 and 身體情況 use one header row, with every file’s rows stacked underneath. Titles on those two sheets are a single line (Chinese and Indonesian). A row is written only when a measurement or observation is filled in, so a template row that is only a date, a time, and `?` is skipped.

The same 運動, 血壓, 睡眠, 喝水量, and 食物 charts are written to `Extract_plots.html`. 排洩 and 身體情況 are not in that file. The 食物 chart plots 早餐食物, 午餐食物, 晚餐食物, and 一天食物 against 日期.

## Landscape pages and PDF

Process also writes one landscape page series per table. A page is 1600×1132 pixels. When the rows do not fit, the next file is `運動_2.png`, `運動_3.png`, and so on. Every PNG is then combined into `output\Extract.pdf`.

A page keeps rows above the first `?` or `？` in column A, and above the first `?`, `？`, or blank cell in the stop column. Rows at and below that marker are left out, including template rows whose formula result is `00:00:00`.

| Page file | Sheet | What is drawn |
|---|---|---|
| `運動_1.png` | 運動 | 運動日期, 開始時間, 運動距離, and 一天運動距離 on the first row of each date |
| `清潔_1.png` | 清潔 | Columns B, D, F, above the first blank cup |
| `飲食_1.png` | 飲食 | One row per date: 早餐食物, 午餐食物, 晚餐食物, and 一天食物 |
| `喝水量_1.png` | 喝水量 in `Extract.xlsx` | Placed in the PDF immediately after 飲食 |
| `睡眠_1.png` | 睡眠 | 日期, 開始時間, 結束時間, 睡眠時間, and 一天睡眠時間 on the first row of each date |
| `排洩1_1.png` | 排洩 | Columns A, E, G |
| `排洩2_1.png` | 排洩 | Columns A, I, K, M |
| `血壓1_1.png` | 血壓 | Columns A, C, E, F |
| `血壓2_1.png` | 血壓 | Columns A, C, G |
| `身體情況1_1.png` | 身體情況 | Columns A, C, E, G |
| `身體情況2_1.png` | 身體情況 | Columns A, I, K |

一天運動距離 is the sum of that date’s 運動距離. 一天睡眠時間 is the sum of that date’s 睡眠時間. 排洩, 血壓, and 身體情況 each use two page series.

## Gmail

**Gmail 傳送 Excel** opens Gmail in the browser and pastes three files from the selected folder’s `output` directory:

- `Extract.xlsx`
- `Extract.pdf`
- `Extract.html` (copied from `Extract_plots.html` when `Extract.html` is not already there)

Receivers: `ghpiu0661@gmail.com`, `ghjane5168@gmail.com`, `a290772@gmail.com`.

The message stays open so you can check that all three files are attached before you press Send.

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
