from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import sqlite3
import pandas as pd
import io
import os

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
        cursor.execute("DELETE FROM records")
        
        success_count = 0
        for sheet_name in xls.sheet_names:
            if sheet_name in ['产量', '汇总', 'Sheet1', 'Sheet2']:
                continue
                
            df = pd.read_excel(xls, sheet_name=sheet_name)
            cus_idx, kg_idx, rolls_idx = None, None, None
            for row_idx in range(min(3, len(df))):
                row_vals = df.iloc[row_idx].values.tolist()
                for i, val in enumerate(row_vals):
                    v = str(val).strip()
                    if v == '客户': cus_idx = i
                    if v == '重量': kg_idx = i
                    if v == '疋数': rolls_idx = i
            
            if kg_idx is None:
                continue
                
            def clean(val):
                v = str(val).strip()
                return '' if v.lower() in ['nan', 'none', 'nat'] else v

            for i, row in df.iterrows():
                date_val = clean(row.iloc[0])
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
                    date_val, clean(row.iloc[2]), clean(row.iloc[5]), clean(row.iloc[6]), clean(row.iloc[7]), '', clean(row.iloc[8]), 
                    clean(row.iloc[cus_idx]) if cus_idx is not None else '', clean(row.iloc[1]), kg_val, rolls_val, clean(row.iloc[9])
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

# --- ĐOẠN CODE PHỤC VỤ GIAO DIỆN (ĐÃ FIX LỖI NOT FOUND) ---
if os.path.isdir("dist"):
    # Cấp quyền đọc các file thiết kế CSS, JS
    if os.path.isdir("dist/assets"):
        app.mount("/assets", StaticFiles(directory="dist/assets"), name="assets")
    
    # Bắt mọi đường dẫn và hiển thị web chính
    @app.get("/{catchall:path}")
    def serve_react(catchall: str):
        file_path = os.path.join("dist", catchall)
        if os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse("dist/index.html")
