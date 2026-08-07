import re
from pathlib import Path
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
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

# Thu muc chua frontend da build (dist/index.html + dist/assets/*), nam
# canh main.py theo dung cau truc trong "cau truc appweb.txt".
BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist"
ASSETS_DIR = DIST_DIR / "assets"

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

# Cot "操作人" (nguoi van hanh) doi khi bi ghep chung ma may o dau, vi du
# "DH14 杜德忠" (ma may DH14 + ten). Chi tach khi khop chac chan dang
# "2 chu cai + 2 chu so, roi den khoang trang, roi den phan con lai" -
# da doi chieu tren du lieu thuc te thang 5-6/2026 (vi du DH17, PC05, DH21...).
# Nhung gia tri khac (vd ma nhan vien rieng le nhu "ec0783", khong co dau
# cach) se KHONG bi dong, giu nguyen trong cot operator nhu truoc, tranh
# cat sai du lieu khong chac chan.
OPERATOR_MACHINE_RE = re.compile(r"^([A-Za-z]{2}\d{2})\s+(\S.*)$")


def split_operator_machine(raw_operator):
    """Tra ve (machine, operator_name). machine = '' neu khong nhan dien duoc."""
    match = OPERATOR_MACHINE_RE.match(raw_operator)
    if match:
        return match.group(1), match.group(2)
    return "", raw_operator

# Nhan dien dong du lieu that: cot dau tien phai la ngay dang 2026.05.02,
# 2026-05-02, 2026/05/02 ... (thay vi hardcode '202' / '07/' nhu truoc,
# vi cach cu se sai voi moi thang khac thang 7). Bat luon nhom nam/thang/ngay
# de chuan hoa ve dinh dang ISO va de xac dinh cac thang co trong file upload.
DATE_RE = re.compile(r"^(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})")


def normalize_date(date_val):
    """Tra ve (ngay_ISO 'YYYY-MM-DD', thang 'YYYY-MM') hoac (None, None)."""
    match = DATE_RE.match(date_val)
    if not match:
        return None, None
    year, month, day = match.groups()
    iso_date = f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    year_month = f"{int(year):04d}-{int(month):02d}"
    return iso_date, year_month


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

        success_count = 0
        sheet_reports = []  # thong ke tung sheet de nguoi dung biet sheet nao bi bo qua
        parsed_rows = []  # gom het du lieu truoc, chi ghi vao DB sau khi da doc xong toan bo file
        months_in_file = set()

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
                iso_date, year_month = normalize_date(date_val)
                if iso_date is None:
                    continue  # dong tieu de / dong trong, khong phai du lieu

                try:
                    kg_val = float(row.iloc[COL_KG]) if pd.notna(row.iloc[COL_KG]) else 0.0
                except (ValueError, TypeError):
                    kg_val = 0.0

                try:
                    rolls_val = int(row.iloc[COL_ROLLS]) if pd.notna(row.iloc[COL_ROLLS]) else 0
                except (ValueError, TypeError):
                    rolls_val = 0

                machine_val, operator_val = split_operator_machine(clean(row.iloc[COL_OPERATOR]))

                parsed_rows.append(
                    (
                        iso_date,
                        clean(row.iloc[COL_LOT]),
                        clean(row.iloc[COL_AREA]),
                        clean(row.iloc[COL_LEADER]),
                        operator_val,
                        machine_val,
                        clean(row.iloc[COL_SHIFT]),
                        clean(row.iloc[COL_CUSTOMER]),
                        clean(row.iloc[COL_ERROR_TYPE]),
                        kg_val,
                        rolls_val,
                        clean(row.iloc[COL_REWORK]),
                    )
                )
                months_in_file.add(year_month)
                rows_in_sheet += 1
                success_count += 1

            sheet_reports.append({"sheet": sheet_name, "rows": rows_in_sheet, "reason": None})

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # Chi xoa du lieu CUA CHINH CAC THANG co trong file dang upload,
        # khong dong cha tat ca (truoc day "DELETE FROM records" xoa sach
        # toan bo bang moi lan upload, nen upload thang 6 se xoa mat du lieu
        # thang 5 da nap truoc do). Nho vay upload nhieu thang lan luot van
        # giu duoc lich su, con upload lai dung 1 thang thi thay the sach
        # du lieu thang do (tranh trung lap).
        for year_month in months_in_file:
            cursor.execute("DELETE FROM records WHERE substr(date, 1, 7) = ?", (year_month,))

        cursor.executemany(
            """
            INSERT INTO records (date, lot, area, leader, operator, machine, shift,
                                  customer, errorType, kg, rolls, rework)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            parsed_rows,
        )

        conn.commit()

        empty_sheets = [s["sheet"] for s in sheet_reports if s["rows"] == 0]

        return {
            "status": "success",
            "rows": success_count,
            "message": "Import thanh cong!",
            "months_replaced": sorted(months_in_file),
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


# --- Phuc vu frontend (dist/) cung mot origin voi API ---
# Truoc day khong co phan nay: file JS build (index-*.js) chi goi duong
# dan tuong doi (vi du fetch("/api/records")) khi KHONG chay tren cong dev
# 5173, nghia la no gia dinh frontend va backend cung dung 1 origin. Nhung
# main.py truoc do khong phuc vu dist/ nen gia dinh do khong bao gio dung -
# mo dist/index.html truc tiep (hoac serve bang mot web server tinh khac)
# se khien moi request /api/... roi vao dung server dang mo index.html
# (khong phai FastAPI tren cong 8000) va that bai.
# Mount tai day, SAU cac route /api de khong bi chan boi StaticFiles.
if ASSETS_DIR.is_dir():
    app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")


@app.get("/")
def serve_index():
    index_file = DIST_DIR / "index.html"
    if index_file.is_file():
        return FileResponse(str(index_file))
    return {"detail": "dist/index.html khong ton tai - chi co API tai /api/*"}
