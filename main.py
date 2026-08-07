"""
QMS Dashboard — Backend FastAPI (bản sửa lỗi hoàn chỉnh)
=========================================================
Đã sửa:
  1. KHÔNG xoá toàn bộ dữ liệu khi upload → chỉ xoá theo tháng có trong file
  2. Xác thực mật khẩu ở SERVER (Bearer token)
  3. Parse KG có dấu phẩy nghìn: "1,172.20" → 1172.2
  4. Tách mã máy khỏi operator: "DH14 杜德忠" → machine="DH14", operator="杜德忠"
  5. Customer: cột 14 (brand) + fallback dynamic search
  6. UNIQUE constraint + INDEX trên bảng records
  7. Validate file type (.xlsx/.xls)
  8. Serve frontend từ dist/ cùng origin với API
  9. Dùng lifespan thay vì @app.on_event (deprecated)
"""

import io
import os
import re
import sqlite3
from contextlib import asynccontextmanager
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# ═══════════════════════════════════════════════════════
# CẤU HÌNH
# ═══════════════════════════════════════════════════════

DB_FILE = "qms.db"

# Mật khẩu upload — nên dùng biến môi trường trong production:
#   export QMS_UPLOAD_PASSWORD="vh@Thang"
UPLOAD_PASSWORD = os.environ.get("QMS_UPLOAD_PASSWORD", "vh@Thang")

BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist"
ASSETS_DIR = DIST_DIR / "assets"

# Sheet tổng hợp / không chứa dữ liệu lỗi hàng ngày → bỏ qua
SUMMARY_SHEETS = {"产量", "产量 (2)", "汇总", "Sheet1", "Sheet2", "合计", ""}

# ─── Vị trí cột (đã đối chiếu dữ liệu thực tế T5 + T6/2026) ───
COL_DATE       = 0   # 日期         2026.06.24
COL_ERROR_TYPE = 1   # 問題類型     布重 / 克重 / 洗后縮率 …
COL_LOT        = 2   # A (缸号)     DP6061703 / DJ6059364 …
COL_DEPT       = 3   # 責任部門     V04.后整部
COL_RESP       = 4   # 責任人       EC1427 / 吴国鹏 …
COL_LEADER     = 5   # 责任车间/组   赵志红 / 莫志乾 …
COL_SUBLEADER  = 6   # 组           李文玉 / 王文楼 …
COL_OPERATOR   = 7   # 操作人       曾文行 / "DH14 杜德忠"
COL_SHIFT      = 8   # 班别         A / B / N
COL_REWORK     = 9   # 是否回修     Y / N
# 10-13: mô tả / kết quả / xác nhận (không lưu vào DB)
COL_CUSTOMER   = 14  # C (brand)    UNIQLO / AEO / TCP …
COL_FACTORY    = 15  # D (factory)  晶苑益力 / REGENT …
COL_KG         = 24  # 产量 (Kg)    1,172.20
COL_ROLLS      = 25  # T (Rolls)    22.00
MIN_COLS       = COL_ROLLS + 1       # tối thiểu 26 cột

# Regex tách mã máy: "DH14 杜德忠" → ("DH14", "杜德忠")
MACHINE_RE = re.compile(r"^([A-Za-z]{2}\d{2})\s+(.+)$")

# Regex parse ngày: 2026.05.02 / 2026-05-02 / 2026/05/02
DATE_RE = re.compile(r"^(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})")

# Từ khoá nhận diện khách hàng (fallback khi cột 14 trống)
KNOWN_CUSTOMER_KEYWORDS = (
    "TORAY", "UNIQLO", "HANSAE", "AEO", "WALMART", "JCP", "CK",
    "FILA", "TCP", "LULULEMON", "A & F", "LACOSTE", "HUGO BOSS",
    "FRUIT", "BABATON", "SONOMA", "CAT & JACK", "LANDS", "MACY",
    "TARGET", "ADORE ME", "DULUTH", "LOLLYTOGS", "M&S", "PTT",
    "ADDIS", "U-KNITS", "SAE-A", "SCAVI", "REGENT", "FASHION",
    "HANSOLL", "COMTEXTI", "SHINSUNG", "SUY", "LEADING", "AMERICAN",
    "MAXHILL", "NAM THUA",
)


# ═══════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════

def clean(val) -> str:
    """Trả về string đã strip; rỗng nếu NaN/None."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "nat", "<na>") else s


def normalize_date(raw):
    """
    Parse ngày → (iso_date 'YYYY-MM-DD', year_month 'YYYY-MM')
    hoặc (None, None) nếu không hợp lệ.
    """
    m = DATE_RE.match(str(raw).strip())
    if not m:
        return None, None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None, None
    return f"{y:04d}-{mo:02d}-{d:02d}", f"{y:04d}-{mo:02d}"


def parse_kg(val) -> float:
    """Parse kg — xử lý dấu phẩy nghìn: '1,172.20' → 1172.2"""
    s = clean(val).replace(",", "").replace("，", "")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def parse_rolls(val) -> int:
    """Parse số cuộn → int."""
    s = clean(val).replace(",", "").replace("，", "")
    try:
        return int(float(s)) if s else 0
    except ValueError:
        return 0


def split_operator_machine(raw: str):
    """'DH14 杜德忠' → ('DH14', '杜德忠');  '曾文行' → ('', '曾文行')"""
    m = MACHINE_RE.match(raw)
    if m:
        return m.group(1), m.group(2)
    return "", raw


def find_customer(row) -> str:
    """
    Ưu tiên cột 14. Nếu trống → tìm trong khoảng idx 13-17
    giá trị chứa từ khoá khách hàng đã biết.
    """
    customer = clean(row.iloc[COL_CUSTOMER]) if len(row) > COL_CUSTOMER else ""
    if customer:
        return customer
    for i in range(13, min(18, len(row))):
        v = clean(row.iloc[i])
        if v and any(kw in v.upper() for kw in KNOWN_CUSTOMER_KEYWORDS):
            return v
    return ""


# ═══════════════════════════════════════════════════════
# DATABASE
# ═══════════════════════════════════════════════════════

def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA encoding = 'UTF-8'")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            date       TEXT    NOT NULL,
            lot        TEXT    DEFAULT '',
            area       TEXT    DEFAULT '',
            leader     TEXT    DEFAULT '',
            operator   TEXT    DEFAULT '',
            machine    TEXT    DEFAULT '',
            shift      TEXT    DEFAULT '',
            customer   TEXT    DEFAULT '',
            error_type TEXT    DEFAULT '',
            kg         REAL    DEFAULT 0,
            rolls      INTEGER DEFAULT 0,
            rework     TEXT    DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(date, lot, error_type, shift)
        )
    """)
    for col in ("date", "customer", "error_type", "leader", "area"):
        conn.execute(
            f"CREATE INDEX IF NOT EXISTS idx_{col} ON records({col})"
        )
    conn.commit()
    conn.close()


# ═══════════════════════════════════════════════════════
# FASTAPI APP
# ═══════════════════════════════════════════════════════

@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield

app = FastAPI(title="QMS Dashboard API", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ═══════════════════════════════════════════════════════
# API: UPLOAD EXCEL
# ═══════════════════════════════════════════════════════

@app.post("/api/upload")
async def upload_excel(
    file: UploadFile = File(...),
    authorization: str = Header(None),
):
    # ── 1. Xác thực (FIX #2) ──────────────────────────
    if authorization != f"Bearer {UPLOAD_PASSWORD}":
        raise HTTPException(
            status_code=401,
            detail="Sai mật khẩu hoặc thiếu Authorization header",
        )

    # ── 2. Validate file type ─────────────────────────
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file Excel (.xlsx/.xls)")

    conn = None
    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents), engine="openpyxl")

        parsed_rows: list[tuple] = []
        months_in_file: set[str] = set()
        sheet_reports: list[dict] = []

        for sheet_name in xls.sheet_names:
            if sheet_name in SUMMARY_SHEETS:
                continue

            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            # Sheet không đủ cột → bỏ qua
            if df.shape[1] < MIN_COLS:
                sheet_reports.append({
                    "sheet": sheet_name,
                    "rows": 0,
                    "reason": f"thiếu cột ({df.shape[1]}/{MIN_COLS})",
                })
                continue

            rows_in_sheet = 0
            for _, row in df.iterrows():
                # Parse ngày — dòng tiêu đề / tổng kết sẽ có date=None → skip
                iso_date, year_month = normalize_date(row.iloc[COL_DATE])
                if iso_date is None:
                    continue

                lot = clean(row.iloc[COL_LOT])
                if not lot:
                    continue  # không có Lot → không phải bản ghi lỗi

                error_type = clean(row.iloc[COL_ERROR_TYPE])
                shift      = clean(row.iloc[COL_SHIFT]).upper()
                rework     = clean(row.iloc[COL_REWORK]).upper()

                # Tách mã máy khỏi operator (FIX: machine luôn trống)
                raw_operator = clean(row.iloc[COL_OPERATOR])
                machine, operator = split_operator_machine(raw_operator)

                # Customer (FIX: dynamic search fallback)
                customer = find_customer(row)

                # KG / Rolls (FIX: dấu phẩy nghìn)
                kg_val    = parse_kg(row.iloc[COL_KG])
                rolls_val = parse_rolls(row.iloc[COL_ROLLS])

                parsed_rows.append((
                    iso_date,                              # date
                    lot,                                   # lot
                    clean(row.iloc[COL_DEPT]),             # area (責任部門)
                    clean(row.iloc[COL_LEADER]),           # leader (责任车间)
                    operator,                              # operator
                    machine,                               # machine
                    shift,                                 # shift
                    customer,                              # customer
                    error_type,                            # error_type
                    kg_val,                                # kg
                    rolls_val,                             # rolls
                    rework,                                # rework
                ))
                months_in_file.add(year_month)
                rows_in_sheet += 1

            sheet_reports.append({
                "sheet": sheet_name,
                "rows": rows_in_sheet,
                "reason": None,
            })

        # ── 3. Ghi vào DB (FIX #1: xoá theo tháng, không xoá hết) ──
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        for ym in months_in_file:
            cursor.execute(
                "DELETE FROM records WHERE substr(date, 1, 7) = ?", (ym,)
            )

        cursor.executemany(
            """INSERT INTO records
               (date, lot, area, leader, operator, machine,
                shift, customer, error_type, kg, rolls, rework)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            parsed_rows,
        )
        conn.commit()

        empty_sheets = [s["sheet"] for s in sheet_reports if s["rows"] == 0]

        return {
            "status": "success",
            "rows": len(parsed_rows),
            "message": f"Import thành công {len(parsed_rows)} dòng",
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


# ═══════════════════════════════════════════════════════
# API: ĐỌC DỮ LIỆU
# ═══════════════════════════════════════════════════════

@app.get("/api/records")
def get_records():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT * FROM records ORDER BY date DESC, id DESC"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


@app.get("/api/summary")
def get_summary():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        total       = conn.execute("SELECT COUNT(*) c FROM records").fetchone()["c"]
        total_kg    = conn.execute("SELECT COALESCE(SUM(kg),0) s FROM records").fetchone()["s"]
        total_rolls = conn.execute("SELECT COALESCE(SUM(rolls),0) s FROM records").fetchone()["s"]
        months      = conn.execute(
            "SELECT DISTINCT substr(date,1,7) m FROM records ORDER BY m"
        ).fetchall()
        return {
            "total_records": total,
            "total_kg": total_kg,
            "total_rolls": total_rolls,
            "months": [r["m"] for r in months],
        }
    finally:
        conn.close()


@app.delete("/api/records")
def delete_all_records(authorization: str = Header(None)):
    """Xoá toàn bộ dữ liệu (cần xác thực)."""
    if authorization != f"Bearer {UPLOAD_PASSWORD}":
        raise HTTPException(status_code=401, detail="Unauthorized")
    conn = sqlite3.connect(DB_FILE)
    try:
        conn.execute("DELETE FROM records")
        conn.commit()
        return {"status": "success", "message": "Đã xoá toàn bộ dữ liệu"}
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════
# SERVE FRONTEND (dist/)
# ═══════════════════════════════════════════════════════

if ASSETS_DIR.is_dir():
    app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")

@app.get("/")
def serve_index():
    index_file = DIST_DIR / "index.html"
    if index_file.is_file():
        return FileResponse(str(index_file))
    return {"detail": "dist/index.html không tồn tại — chỉ có API tại /api/*"}
