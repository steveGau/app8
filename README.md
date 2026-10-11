# processExcel

Windows desktop app that reads dated Excel workbooks and writes one summary workbook plus Plotly charts. Sheet titles and source labels are Traditional Chinese and Indonesian.

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

Only files named like `20261005.xlsx` are used. They are sorted by date. Other `.xlsx` names are skipped. In each workbook the app reads sheets **運動**, **血壓**, and **睡眠**, from row 2 down to the row above `?` or `？`. Empty cells and missing sheets are skipped.

| Sheet | Kept columns | Converted values |
|---|---|---|
| 運動 | 選擇運動, 運動日期, 運動時間, 運動距離 | 運動時間 to minutes. 運動距離 such as `10米: 10 meter` to `10` meters. |
| 血壓 | 健康, 量血壓日期, 收縮壓, 舒張壓, 心跳 | Numbers as stored. |
| 睡眠 | 睡眠, 日期, 開始時間, 結束時間, 睡眠時間 | 睡眠時間 to hours (end minus start, including overnight). |

Each dated file is placed side by side in `Extract.xlsx`. Plot sheets in that workbook are **運動 Plot**, **血壓 Plot**, and **睡眠 Plot**. The same charts are written to `Extract_plots.html`.

## Sample

`testExample\` has `20261005.xlsx`, `20261006.xlsx`, and `20261007.xlsx`. Processing that folder writes:

```bat
processExcel.bat "C:\Users\a2907\Desktop\app8\testExample"
```

Results are saved in `testExample\output\`.

## Tests

```bat
.\.venv\Scripts\python.exe -m unittest test_processExcel.py
```

## Not committed

`.venv\` and `__pycache__\` stay on this PC. Run `processExcel.bat` to create the environment here.
