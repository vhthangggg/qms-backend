"""
QMS Dashboard — Backend FastAPI (Production-ready cho Render.com)
==================================================================
Các thay đổi cho Render:
  - Port từ PORT env var
  - Database path hỗ trợ Persistent Disk (/data/qms.db)
  - Gunicorn-friendly (stateless)
  - Health check endpoint
  - CORS với domain Render
"""

import io
import os
import re
import sqlite3
from contextlib import asynccontextmanager
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Header, HTTPException, UploadFile, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# ═══════════════════════════════════════════════════════
# CẤU HÌNH TỪ ENVIRONMENT (Render inject)
# ═══════════════════════════════════════════════════════

# Port do Render cung cấp (bắt buộc)
PORT = int(os.environ.get("PORT", 8000))

# Mật khẩu upload — KHÔNG hardcode trên production
UPLOAD_PASSWORD = os.environ.get("QMS_UPLOAD_PASSWORD", "vh@Thang")

# Domain ngoài của Render (ví dụ: https://qms.onrender.com)
RENDER_URL = os.environ.get("RENDER_EXTERNAL_URL", "http://localhost:8000")

# Đường dẫn database — ưu tiên /data nếu có Persistent Disk
BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist"
ASSETS_DIR = DIST_DIR / "assets"

# /data là mount point của Render Persistent Disk
PERSISTENT_DIR = Path("/data")
DB_FILE = (
    PERSISTENT_DIR / "qms.db" if PERSISTENT_DIR.is_dir()
    else BASE_DIR / "qms.db"
)

# ═══════════════════════════════════════════════════════
# CẤU HÌNH NGHIỆP VỤ
# ═══════════════════════════════════════════════════════

SUMMARY_SHEETS = {"产量", "产量 (2)", "汇总", "Sheet1", "Sheet2", "合计", ""}

# Vị trí cột (đã đối chiếu dữ liệu T5 + T6/2026)
COL_DATE       = 0
COL_ERROR_TYPE = 1
COL_LOT        = 2
COL_DEPT       = 3
COL_RESP       = 4
COL_LEADER     = 5
COL_SUBLEADER  = 6
COL_OPERATOR   = 7
COL_SHIFT      = 8
COL_REWORK     = 9
COL_CUSTOMER   = 14
COL_FACTORY    = 15
COL_KG         = 24
COL_ROLLS      = 25
MIN_COLS       = COL_ROLLS + 1

MACHINE_RE = re.compile(r"^([A-Za-z]{2}\d{2})\s+(.+)$")
DATE_RE    = re.compile(r"^(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})")

KNOWN_CUSTOMER_KEYWORDS = (
    "TORAY", "UNIQLO", "HANSAE", "AEO", "WALMART", "JCP", "CK",
    "FILA", "TCP", "LULULEMON", "A & F", "LACOSTE", "HUGO BOSS",
    "FRUIT", "BABATON", "SONOMA", "CAT & JACK", "LANDS", "MACY",
    "TARGET", "ADORE ME", "DULUTH", "LOLLYTOGS", "M&S", "PTT",
    "ADDIS", "U-KNITS", "SAE-A", "SCAVI", "REGENT", "HANSOLL",
)

# ═══════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════

def clean(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "nat", "<na>") else s


def normalize_date(raw):
    m = DATE_RE.match(str(raw).strip())
    if not m:
        return None, None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None, None
    return f"{y:04d}-{mo:02d}-{d:02d}", f"{y:04d}-{mo:02d}"


def parse_kg(val) -> float:
    s = clean(val).replace(",", "").replace("，", "")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def parse_rolls(val) -> int:
    s = clean(val).replace(",", "").replace("，", "")
    try:
        return int(float(s)) if s else 0
    except ValueError:
        return 0


def split_operator_machine(raw: str):
    m = MACHINE_RE.match(raw)
    return (m.group(1), m.group(2)) if m else ("", raw)


def find_customer(row) -> str:
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

def get_db():
    """Tạo connection với timeout để tránh lock trên Render."""
    conn = sqlite3.connect(str(DB_FILE), timeout=30)
    conn.execute("PRAGMA encoding = 'UTF-8'")
    conn.execute("PRAGMA journal_mode = WAL")  # Write-Ahead Logging
    conn.execute("PRAGMA busy_timeout = 5000")
    return conn


def init_db():
    """Khởi tạo DB — tạo thư mục /data nếu chưa có."""
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    conn = get_db()
    try:
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
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════
# FASTAPI APP
# ═══════════════════════════════════════════════════════

@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    print(f"🚀 QMS Dashboard started — DB at {DB_FILE}")
    yield

app = FastAPI(title="QMS Dashboard API", lifespan=lifespan)

# CORS — cho phép domain Render của bạn + localhost (dev)
ALLOWED_ORIGINS = [
    RENDER_URL,
    "http://localhost:5173",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
# Nếu bạn có domain custom, thêm vào đây:
# ALLOWED_ORIGINS.append("https://yourdomain.com")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ═══════════════════════════════════════════════════════
# HEALTH CHECK (Render dùng để monitor)
# ═══════════════════════════════════════════════════════

@app.get("/health")
def health_check():
    """Endpoint để Render health check."""
    try:
        conn = get_db()
        count = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
        conn.close()
        return {
            "status": "healthy",
            "records_count": count,
            "db_path": str(DB_FILE),
            "persistent_disk": PERSISTENT_DIR.is_dir(),
        }
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "unhealthy", "error": str(e)},
        )


# ═══════════════════════════════════════════════════════
# API: UPLOAD EXCEL
# ═══════════════════════════════════════════════════════

@app.post("/api/upload")
async def upload_excel(
    file: UploadFile = File(...),
    authorization: str = Header(None),
):
    if authorization != f"Bearer {UPLOAD_PASSWORD}":
        raise HTTPException(
            status_code=401,
            detail="Sai mật khẩu hoặc thiếu Authorization header",
        )

    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file Excel")

    conn = None
    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents), engine="openpyxl")

        parsed_rows = []
        months_in_file = set()
        sheet_reports = []

        for sheet_name in xls.sheet_names:
            if sheet_name in SUMMARY_SHEETS:
                continue

            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            if df.shape[1] < MIN_COLS:
                sheet_reports.append({
                    "sheet": sheet_name,
                    "rows": 0,
                    "reason": f"thiếu cột ({df.shape[1]}/{MIN_COLS})",
                })
                continue

            rows_in_sheet = 0
            for _, row in df.iterrows():
                iso_date, year_month = normalize_date(row.iloc[COL_DATE])
                if iso_date is None:
                    continue

                lot = clean(row.iloc[COL_LOT])
                if not lot:
                    continue

                error_type = clean(row.iloc[COL_ERROR_TYPE])
                shift      = clean(row.iloc[COL_SHIFT]).upper()
                rework     = clean(row.iloc[COL_REWORK]).upper()

                raw_operator = clean(row.iloc[COL_OPERATOR])
                machine, operator = split_operator_machine(raw_operator)

                customer = find_customer(row)
                kg_val    = parse_kg(row.iloc[COL_KG])
                rolls_val = parse_rolls(row.iloc[COL_ROLLS])

                parsed_rows.append((
                    iso_date, lot,
                    clean(row.iloc[COL_DEPT]),
                    clean(row.iloc[COL_LEADER]),
                    operator, machine, shift, customer, error_type,
                    kg_val, rolls_val, rework,
                ))
                months_in_file.add(year_month)
                rows_in_sheet += 1

            sheet_reports.append({
                "sheet": sheet_name,
                "rows": rows_in_sheet,
                "reason": None,
            })

        # Ghi vào DB
        conn = get_db()
        cursor = conn.cursor()

        for ym in months_in_file:
            cursor.execute(
                "DELETE FROM records WHERE substr(date, 1, 7) = ?", (ym,)
            )

        cursor.executemany(
            """INSERT OR IGNORE INTO records
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
        if conn:
            conn.rollback()
        return {"status": "error", "message": str(e)}
    finally:
        if conn:
            conn.close()


# ═══════════════════════════════════════════════════════
# API: ĐỌC DỮ LIỆU
# ═══════════════════════════════════════════════════════

@app.get("/api/records")
def get_records():
    conn = get_db()
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
    conn = get_db()
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
    if authorization != f"Bearer {UPLOAD_PASSWORD}":
        raise HTTPException(status_code=401, detail="Unauthorized")
    conn = get_db()
    try:
        conn.execute("DELETE FROM records")
        conn.commit()
        return {"status": "success", "message": "Đã xoá toàn bộ dữ liệu"}
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════
# SERVE FRONTEND
# ═══════════════════════════════════════════════════════

if ASSETS_DIR.is_dir():
    app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")

@app.get("/")
def serve_index():
    index_file = DIST_DIR / "index.html"
    if index_file.is_file():
        return FileResponse(str(index_file))
    return {"detail": "dist/index.html không tồn tại"}


# ═══════════════════════════════════════════════════════
# ENTRYPOINT (khi chạy trực tiếp bằng uvicorn)
# ═══════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=PORT,
        log_level="info",
    )
