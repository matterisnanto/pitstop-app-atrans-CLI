import pandas as pd
from db_manager import get_connection, get_daftar_driver_terbaru

def calculate_next_km(row):
    odo = row.get('Odometer_Bengkel', 0)
    if pd.isna(odo): 
        odo = 0
        
    if pd.notna(row.get('interval_km')) and row.get('interval_km', 0) > 0:
        return odo + row['interval_km']
        
    transmisi = str(row.get('Transmition', '')).upper()
    series_val = str(row.get('series', '')).upper()
    brand_val = str(row.get('brand', '')).upper() 
    teks_kendaraan = series_val + " " + brand_val
    
    if 'EV' in teks_kendaraan or 'IONIQ' in teks_kendaraan or 'BYD' in teks_kendaraan or 'VINFAST' in teks_kendaraan or 'MORRIS GARAGE' in teks_kendaraan or 'MG' in teks_kendaraan:
        return odo + 15000
    if 'A/T' in transmisi or 'MATIC' in transmisi:
        return odo + 30000
        
    return odo + 10000

def process_and_export(output_filename="LAPORAN_LENGKAP_PITSTOP.xlsx"):
    conn = get_connection()
    try:
        # 1. TARIK DATA
        df_actual = pd.read_sql_query("SELECT * FROM actual_lists", conn)
        df_vehicle = pd.read_sql_query("SELECT * FROM vehicles", conn)
        df_reports = pd.read_sql_query("SELECT * FROM driver_km_reports", conn)
        df_custom = pd.read_sql_query("SELECT * FROM custom_intervals", conn)
        
        # 2. AMBIL DRIVER AKTIF (Paling Update, Tanpa Ganda)
        df_driver_aktif = get_daftar_driver_terbaru()
        
        # 3. FILTER SERVICE BENGKEL TERAKHIR
        df_actual['Category'] = df_actual['Category'].fillna('')
        df_actual['Description'] = df_actual['Description'].fillna('')
        
        mask_kategori_utama = df_actual['Category'].str.contains('Tune Up & Ganti Oli', case=False)
        kata_kunci = 'oli|oil|service|servis|berkala|berskska|tune up|flush|shell|idemitsu|fluid'
        mask_kategori_others = df_actual['Category'].str.contains('Others', case=False) & \
                               df_actual['Description'].str.contains(kata_kunci, case=False, regex=True)
                               
        df_oli = df_actual[mask_kategori_utama | mask_kategori_others].copy()
        df_oli['Odometer'] = pd.to_numeric(df_oli['Odometer'], errors='coerce')
        df_oli = df_oli[df_oli['Odometer'] > 0]
        
        df_oli['Actual Date'] = pd.to_datetime(df_oli['Actual Date'], errors='coerce')
        df_latest_service = df_oli.sort_values(by=['Actual Date', 'Odometer'], ascending=[False, False]).drop_duplicates(subset=['No. Pol'])
        df_latest_service.rename(columns={'Odometer': 'Odometer_Bengkel'}, inplace=True)
        
        # 4. FILTER INPUT KM WA TERAKHIR
        if not df_reports.empty:
            df_reports['report_date'] = pd.to_datetime(df_reports['report_date'])
            df_latest_report = df_reports.sort_values('report_date', ascending=False).drop_duplicates('nopol')
        else:
            # Pastikan ada kolom report_date meskipun data masih kosong
            df_latest_report = pd.DataFrame(columns=['nopol', 'current_km', 'report_date'])

        # 5. MERGE KE 1 TABEL LENGKAP
        df_vehicle_lite = df_vehicle[['Nomor Polisi', 'Transmition']] if 'Transmition' in df_vehicle.columns else pd.DataFrame(columns=['Nomor Polisi'])
        
        merged = pd.merge(df_driver_aktif, df_vehicle_lite, left_on='nopol', right_on='Nomor Polisi', how='left')
        merged = pd.merge(merged, df_latest_service, left_on='nopol', right_on='No. Pol', how='left')
        merged = pd.merge(merged, df_latest_report, on='nopol', how='left')
        merged = pd.merge(merged, df_custom, on='nopol', how='left') 
        
        # 6. HITUNG MATEMATIKA
        merged['KM_SELANJUTNYA'] = merged.apply(calculate_next_km, axis=1)
        merged['current_km'] = merged['current_km'].fillna(merged['Odometer_Bengkel'])
        merged['SISA_KM'] = merged['KM_SELANJUTNYA'] - merged['current_km']
        
        # 7. EXPORT JADI 1 SHEET (Dengan Tambahan TANGGAL UPDATE WA)
        final_report = pd.DataFrame({
            'NOPOL': merged['nopol'],
            'DRIVER AKTIF': merged['name'],
            'PHONE': merged['phone'],
            'UNIT KENDARAAN': merged['brand'] + " " + merged['series'],
            'GPS': merged['gps'],
            'TANGGAL HANDOVER TERBARU': pd.to_datetime(merged['tgl_handover']).dt.strftime('%Y-%m-%d').fillna('-'),
            'KM SERVICE AWAL': merged['Odometer_Bengkel'].fillna(0),
            'KM SERVICE SELANJUTNYA': merged['KM_SELANJUTNYA'].fillna(0),
            'KM UPDATE WA': merged['current_km'].fillna(0),
            'TANGGAL UPDATE WA': pd.to_datetime(merged['report_date']).dt.strftime('%Y-%m-%d %H:%M').fillna('-'), # KOLOM BARU
            'SISA KM MENUJU SERVICE': merged['SISA_KM'].fillna(0),
            'ATURAN INTERVAL': merged['interval_km'].fillna('Default/Sistem'),
            'TANGGAL SERVICE (BENGKEL)': merged['Actual Date'].dt.strftime('%Y-%m-%d').fillna('Belum Ada Data'),
            'KETERANGAN BENGKEL': merged['Description'].fillna('-')
        })
        
        final_report.to_excel(output_filename, index=False)
        print(f"\n[SUCCESS] Laporan Terpadu berhasil diekspor ke: {output_filename}")
        
    except Exception as e:
        print(f"[ERROR] Kegagalan proses export: {e}")
    finally:
        conn.close()