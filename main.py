import re
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import sqlite3
import pandas as pd
import io

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_FILE = "qms.db"

# Cac sheet tong hop / khong phai nhat ky loi hang ngay -> luon bo qua
SUMMARY_SHEETS = {"产量", "产量 (2)", "汇总", "Sheet1", "Sheet2"}

# Vi tri cot CO DINH trong file "异常处理明细报表" (da doi chieu tren du lieu
# thuc te thang 5 va thang 6/2026 - cac cot nay giu nguyen vi tri du sheet
# co hay khong co dong tieu de van ban '客户'/'重量'/'疋数').
COL_LOT = 2
COL_AREA = 5
COL_LEADER = 6
COL_OPERATOR = 7
COL_SHIFT = 8
COL_REWORK = 9
COL_ERROR_TYPE = 1
COL_CUSTOMER = 15
COL_KG = 24
COL_ROLLS = 25
MIN_COLS_REQUIRED = COL_ROLLS + 1  # 26

# Nhan dien dong du lieu that: cot dau tien phai la ngay dang 2026.05.02,
# 2026-05-02, 2026/05/02 ... (thay vi hardcode '202' / '07/' nhu truoc,
# vi cach cu se sai voi moi thang khac thang 7).
DATE_RE = re.compile(r"^20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}")


def clean(val):
    v = str(val).strip()
    return "" if v.lower() in ["nan", "none", "nat"] else v


def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT, lot TEXT, area TEXT, leader TEXT,
            operator TEXT, machine TEXT, shift TEXT,
            customer TEXT, errorType TEXT, kg REAL, rolls INTEGER, rework TEXT
        )
        """
    )
    conn.commit()
    conn.close()


@app.on_event("startup")
def startup():
    init_db()


@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    conn = None
    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents))

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM records")

        success_count = 0
        sheet_reports = []  # thong ke tung sheet de nguoi dung biet sheet nao bi bo qua

        for sheet_name in xls.sheet_names:
            if sheet_name in SUMMARY_SHEETS:
                continue

            # header=None de khong bi pandas "nuot" mot dong lam ten cot -
            # vi trong file thuc te, dong tieu de van ban nam o vi tri khac
            # nhau tuy sheet (dong 0 hoac dong 1), neu de header mac dinh
            # thi co sheet se mat dong tieu de va co sheet thi khong.
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            if df.shape[1] < MIN_COLS_REQUIRED:
                sheet_reports.append(
                    {"sheet": sheet_name, "rows": 0, "reason": "thieu cot (chi co %d cot)" % df.shape[1]}
                )
                continue

            rows_in_sheet = 0
            for _, row in df.iterrows():
                date_val = clean(row.iloc[0])
                if not DATE_RE.match(date_val):
                    continue  # dong tieu de / dong trong, khong phai du lieu

                try:
                    kg_val = float(row.iloc[COL_KG]) if pd.notna(row.iloc[COL_KG]) else 0.0
                except (ValueError, TypeError):
                    kg_val = 0.0

                try:
                    rolls_val = int(row.iloc[COL_ROLLS]) if pd.notna(row.iloc[COL_ROLLS]) else 0
                except (ValueError, TypeError):
                    rolls_val = 0

                cursor.execute(
                    """
                    INSERT INTO records (date, lot, area, leader, operator, machine, shift,
                                          customer, errorType, kg, rolls, rework)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        date_val,
                        clean(row.iloc[COL_LOT]),
                        clean(row.iloc[COL_AREA]),
                        clean(row.iloc[COL_LEADER]),
                        clean(row.iloc[COL_OPERATOR]),
                        "",  # may: hien khong co cot rieng trong file nguon
                        clean(row.iloc[COL_SHIFT]),
                        clean(row.iloc[COL_CUSTOMER]),
                        clean(row.iloc[COL_ERROR_TYPE]),
                        kg_val,
                        rolls_val,
                        clean(row.iloc[COL_REWORK]),
                    ),
                )
                rows_in_sheet += 1
                success_count += 1

            sheet_reports.append({"sheet": sheet_name, "rows": rows_in_sheet, "reason": None})

        conn.commit()

        empty_sheets = [s["sheet"] for s in sheet_reports if s["rows"] == 0]

        return {
            "status": "success",
            "rows": success_count,
            "message": "Import thanh cong!",
            "sheets_processed": len(sheet_reports),
            "sheets_with_no_data": empty_sheets,
            "sheet_details": sheet_reports,
        }
    except Exception as e:
        if conn is not None:
            conn.rollback()
        return {"status": "error", "message": str(e)}
    finally:
        if conn is not None:
            conn.close()


@app.get("/api/records")
def get_records():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        records = conn.execute("SELECT * FROM records ORDER BY id DESC").fetchall()
        return [dict(r) for r in records]
    finally:
        conn.close()
