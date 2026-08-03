import sqlite3
import pandas as pd
import os
import warnings
from datetime import datetime

# Membungkam peringatan Pandas agar terminal Anda tetap bersih
warnings.filterwarnings('ignore', category=UserWarning)

DB_NAME = 'pitstop.db'

def get_connection():
    return sqlite3.connect(DB_NAME)

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS driver_km_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nopol TEXT,
            current_km INTEGER,
            report_date DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS custom_intervals (
            nopol TEXT PRIMARY KEY,
            interval_km INTEGER
        )
    ''')
    conn.commit()
    conn.close()
    print("[INFO] Database & Tabel siap digunakan.")

def import_file_to_db(filepath, table_name):
    if not os.path.exists(filepath):
        print(f"[ERROR] File {filepath} tidak ditemukan! Cek kembali nama dan folder.")
        return False
    try:
        try:
            df = pd.read_html(filepath)[0]
        except:
            try:
                df = pd.read_excel(filepath, engine='xlrd')
            except:
                df = pd.read_excel(filepath, engine='openpyxl')
        
        # Format tanggal diseragamkan
        if 'createdate' in df.columns:
            df['createdate'] = pd.to_datetime(df['createdate'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
        if 'Actual Date' in df.columns:
            df['Actual Date'] = pd.to_datetime(df['Actual Date'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
            
        conn = get_connection()
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        print(f"[SUCCESS] Import {len(df)} baris ke '{table_name}'.")
        return True
    except Exception as e:
        print(f"[ERROR] Gagal membaca {filepath}: {e}")
        return False

def cari_kendaraan(keyword):
    conn = get_connection()
    keyword = f"%{keyword}%"
    
    query = """
        WITH RankedHandovers AS (
            SELECT 
                nopol, 
                name, 
                createdate,
                ROW_NUMBER() OVER(PARTITION BY nopol ORDER BY createdate DESC) as rn
            FROM handovers
        )
        SELECT 
            h.nopol, 
            h.name, 
            COALESCE(v."Brand", '') AS brand, 
            COALESCE(v."Series", '') AS series,
            h.createdate AS tgl_handover
        FROM RankedHandovers h
        LEFT JOIN vehicles v ON h.nopol = v."Nomor Polisi"
        WHERE h.rn = 1 AND (h.nopol LIKE ? OR h.name LIKE ?)
        ORDER BY h.createdate DESC
    """
    df = pd.read_sql_query(query, conn, params=(keyword, keyword))
    conn.close()
    return df

def get_km_bengkel_terakhir(nopol):
    conn = get_connection()
    # Logika disamakan dengan export: Mengandalkan Category ERP dan kata OIL/OLI
    query = """
        SELECT Odometer FROM actual_lists 
        WHERE "No. Pol" = ? 
          AND (
              Category LIKE '%Tune Up & Ganti Oli%' 
              OR Description LIKE '%oli%' 
              OR Description LIKE '%oil%'
              OR Description LIKE '%service%'
          )
          AND CAST(Odometer AS INTEGER) > 0
        ORDER BY "Actual Date" DESC, CAST(Odometer AS INTEGER) DESC LIMIT 1
    """
    cursor = conn.cursor()
    cursor.execute(query, (nopol,))
    result = cursor.fetchone()
    conn.close()
    return int(result[0]) if result else 0

def simpan_km_harian(nopol, km_terbaru):
    conn = get_connection()
    cursor = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("INSERT INTO driver_km_reports (nopol, current_km, report_date) VALUES (?, ?, ?)", 
                   (nopol, km_terbaru, now))
    conn.commit()
    conn.close()
    print(f"[SUCCESS] Data KM {km_terbaru} untuk unit {nopol} berhasil disimpan!")

def simpan_interval_khusus(nopol, interval):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("REPLACE INTO custom_intervals (nopol, interval_km) VALUES (?, ?)", (nopol, interval))
    conn.commit()
    conn.close()
    print(f"[SUCCESS] Interval khusus {interval} KM untuk mobil {nopol} berhasil disimpan permanen!")