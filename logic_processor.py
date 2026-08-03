import pandas as pd
from db_manager import get_connection

def calculate_next_km(row):
    odo = row['Odometer_Bengkel']
    
    # 1. Prioritas Utama: Interval Khusus (7000/8000)
    if pd.notna(row.get('interval_km')) and row['interval_km'] > 0:
        return odo + row['interval_km']
        
    transmisi = str(row['Transmition']).upper()
    series = str(row['Series']).upper()
    
    # 2. Aturan Bawaan (EV & Matic)
    if 'EV' in series or 'IONIQ' in series:
        return odo + 15000
    if 'A/T' in transmisi or 'MATIC' in transmisi:
        return odo + 30000
        
    # 3. Default (Manual)
    return odo + 10000

def process_and_export(output_filename="UPDATE_SERVICE_PITSTOP.xlsx"):
    conn = get_connection()
    try:
        # Tarik semua tabel dari SQLite
        df_actual = pd.read_sql_query("SELECT * FROM actual_lists", conn)
        df_vehicle = pd.read_sql_query("SELECT * FROM vehicles", conn)
        df_handover = pd.read_sql_query("SELECT * FROM handovers", conn)
        df_reports = pd.read_sql_query("SELECT * FROM driver_km_reports", conn)
        df_custom = pd.read_sql_query("SELECT * FROM custom_intervals", conn)
        
        # --- LOGIKA FILTER TERBARU (Sesuai Kategori ERP & Typo) ---
        df_actual['Category'] = df_actual['Category'].fillna('')
        df_actual['Description'] = df_actual['Description'].fillna('')
        
        # Ambil SEMUA data yang kategorinya "Tune Up & Ganti Oli"
        mask_kategori_utama = df_actual['Category'].str.contains('Tune Up & Ganti Oli', case=False)
        
        # Ambil data dari kategori "Others" TAPI yang deskripsinya mengandung unsur oli/service
        kata_kunci = 'oli|oil|service|servis|berkala|berskska|tune up|flush|shell|idemitsu|fluid'
        mask_kategori_others = df_actual['Category'].str.contains('Others', case=False) & \
                               df_actual['Description'].str.contains(kata_kunci, case=False, regex=True)
                               
        # Gabungkan kedua kondisi di atas
        df_oli = df_actual[mask_kategori_utama | mask_kategori_others].copy()
        # ----------------------------------------------------------
        
        # Validasi: Buang Odometer 0 (Human error input ERP)
        df_oli['Odometer'] = pd.to_numeric(df_oli['Odometer'], errors='coerce')
        df_oli = df_oli[df_oli['Odometer'] > 0]
        
        # Ambil 1 record bengkel terbaru per mobil
        df_oli['Actual Date'] = pd.to_datetime(df_oli['Actual Date'], errors='coerce')
        df_latest_service = df_oli.sort_values(by=['Actual Date', 'Odometer'], ascending=[False, False]).drop_duplicates(subset=['No. Pol'])
        df_latest_service.rename(columns={'Odometer': 'Odometer_Bengkel'}, inplace=True)
        
        # Ambil laporan WA terbaru (jika ada)
        if not df_reports.empty:
            df_reports['report_date'] = pd.to_datetime(df_reports['report_date'])
            df_latest_report = df_reports.sort_values('report_date', ascending=False).drop_duplicates('nopol')
        else:
            df_latest_report = pd.DataFrame(columns=['nopol', 'current_km'])

        # Ambil driver terbaru dari Handover
        df_handover['createdate'] = pd.to_datetime(df_handover['createdate'], errors='coerce')
        df_latest_driver = df_handover.sort_values('createdate', ascending=False).drop_duplicates('nopol')
        
        # --- PENGGABUNGAN DATA (MERGE) ---
        merged = pd.merge(df_latest_service, df_vehicle, left_on='No. Pol', right_on='Nomor Polisi', how='left')
        merged = pd.merge(merged, df_latest_driver, left_on='No. Pol', right_on='nopol', how='left')
        merged = pd.merge(merged, df_latest_report, left_on='No. Pol', right_on='nopol', how='left')
        merged = pd.merge(merged, df_custom, left_on='No. Pol', right_on='nopol', how='left') 
        
        # --- HITUNG INTERVAL & SISA KM ---
        merged['KM_SELANJUTNYA'] = merged.apply(calculate_next_km, axis=1)
        merged['current_km'] = merged['current_km'].fillna(merged['Odometer_Bengkel'])
        merged['SISA_KM'] = merged['KM_SELANJUTNYA'] - merged['current_km']
        
        # --- SUSUN LAPORAN EXCEL ---
        final_report = pd.DataFrame({
            'TANGGAL SERVICE (BENGKEL)': merged['Actual Date'].dt.strftime('%Y-%m-%d'),
            'NOPOL': merged['No. Pol'],
            'DRIVER': merged['name'],
            'PHONE': merged['phone'],
            'KM SERVICE AWAL': merged['Odometer_Bengkel'],
            'KM SERVICE SELANJUTNYA': merged['KM_SELANJUTNYA'],
            'KM UPDATE WA': merged['current_km'],
            'SISA KM MENUJU SERVICE': merged['SISA_KM'],
            'ATURAN INTERVAL': merged['interval_km'].fillna('Default/Sistem'),
            'KETERANGAN BENGKEL': merged['Description'],
            'GPS': merged['GSM SERVER']
        })
        
        final_report.to_excel(output_filename, index=False)
        print(f"\n[SUCCESS] Laporan diekspor ke: {output_filename}")
        
    except Exception as e:
        print(f"[ERROR] Kegagalan proses: {e}")
    finally:
        conn.close()