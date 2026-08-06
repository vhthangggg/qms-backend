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

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT, lot TEXT, area TEXT, leader TEXT, 
            operator TEXT, machine TEXT, shift TEXT, 
            customer TEXT, errorType TEXT, kg REAL, rolls INTEGER, rework TEXT
        )
    ''')
    conn.commit()
    conn.close()

@app.on_event("startup")
def startup():
    init_db()

@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        xls = pd.ExcelFile(io.BytesIO(contents))
        
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Tự động dọn dẹp dữ liệu cũ mỗi khi nạp file mới
        cursor.execute("DELETE FROM records")
        
        success_count = 0
        
        # Quét qua tất cả các sheet trong file Excel
        for sheet_name in xls.sheet_names:
            # Bỏ qua các sheet báo cáo tổng (không phải nhật ký lỗi hàng ngày)
            if sheet_name in ['产量', '汇总', 'Sheet1', 'Sheet2']:
                continue
                
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # AI tìm kiếm cột tự động (dựa vào tên tiếng Trung)
            cus_idx, kg_idx, rolls_idx = None, None, None
            for row_idx in range(min(3, len(df))):
                row_vals = df.iloc[row_idx].values.tolist()
                for i, val in enumerate(row_vals):
                    v = str(val).strip()
                    if v == '客户': cus_idx = i
                    if v == '重量': kg_idx = i
                    if v == '疋数': rolls_idx = i
            
            # Nếu không tìm thấy cột Trọng lượng, bỏ qua sheet đó
            if kg_idx is None:
                continue
                
            def clean(val):
                v = str(val).strip()
                return '' if v.lower() in ['nan', 'none', 'nat'] else v

            # Bóc tách dữ liệu theo đúng cấu trúc thực tế của file
            for i, row in df.iterrows():
                date_val = clean(row.iloc[0])
                # Chỉ xử lý những dòng nào cột đầu tiên là Ngày (vd: 2026.07.01)
                if not (date_val.startswith('202') or date_val.startswith('07/')):
                    continue
                
                try: kg_val = float(row.iloc[kg_idx]) if pd.notna(row.iloc[kg_idx]) else 0.0
                except: kg_val = 0.0
                    
                try: rolls_val = int(row.iloc[rolls_idx]) if pd.notna(row.iloc[rolls_idx]) else 0
                except: rolls_val = 0

                cursor.execute('''
                    INSERT INTO records (date, lot, area, leader, operator, machine, shift, customer, errorType, kg, rolls, rework)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    date_val,             # Ngày
                    clean(row.iloc[2]),   # Lot (缸号)
                    clean(row.iloc[5]),   # Area
                    clean(row.iloc[6]),   # Leader
                    clean(row.iloc[7]),   # Operator
                    '',                   # Máy (bị gộp chung với Op)
                    clean(row.iloc[8]),   # Shift
                    clean(row.iloc[cus_idx]) if cus_idx is not None else '', # Customer
                    clean(row.iloc[1]),   # Error Type
                    kg_val,               # Kg
                    rolls_val,            # Rolls
                    clean(row.iloc[9])    # Rework (Y/N)
                ))
                success_count += 1
                
        conn.commit()
        conn.close()
        return {"status": "success", "rows": success_count, "message": "Import thành công!"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/api/records")
def get_records():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    records = conn.execute("SELECT * FROM records ORDER BY id DESC").fetchall()
    conn.close()
    return [dict(r) for r in records]