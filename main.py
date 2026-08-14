import sys
from db_manager import (
    init_db, import_file_to_db, cari_kendaraan, 
    simpan_km_harian, get_km_bengkel_terakhir, 
    simpan_interval_khusus, get_daftar_driver_terbaru
)
from logic_processor import process_and_export

def print_menu():
    print("\n" + "="*55)
    print(" 📱 SISTEM TRACKING PITSTOP OLI - AKTIF 📱")
    print("="*55)
    print(" 1. Inisialisasi Database (Awal Saja)")
    print(" 2. Sinkronisasi Data ERP (Upload File .xls)")
    print(" 3. Input KM Harian (Dari Foto WA Driver)")
    print(" 4. Seting Interval Khusus (7000/8000 KM)")
    print(" 5. Hitung & Export Laporan LENGKAP (1 Sheet Master)")
    print(" 6. Lihat Daftar Driver Aktif di Layar (Cepat)")
    print(" 7. Keluar")
    print("="*55)

def input_km_harian():
    while True:
        print("\n--- PENCARIAN KENDARAAN ---")
        keyword = input("Masukkan NOPOL atau Nama Driver (atau '0' ke menu): ").strip()
        
        if keyword == '0' or not keyword:
            break
            
        hasil = cari_kendaraan(keyword)
        if hasil.empty:
            print(f"[!] Tidak ada driver atau nopol yang cocok dengan '{keyword}'.")
            continue
            
        print("\nKendaraan Ditemukan:")
        for idx, row in hasil.iterrows():
            print(f"{idx + 1}. NOPOL: {row['nopol']} | Driver: {row['name']} | Unit: {row['brand']} {row['series']}")
            
        try:
            pilih = int(input("\nPilih nomor kendaraan (atau 0 untuk cari lagi): "))
            if pilih == 0 or pilih > len(hasil): 
                continue
                
            nopol_pilihan = hasil.iloc[pilih-1]['nopol']
            km_bengkel_terakhir = get_km_bengkel_terakhir(nopol_pilihan)
            
            print(f"\n=> Anda memilih: {nopol_pilihan}")
            print(f"=> Info: KM Terakhir saat masuk bengkel adalah {km_bengkel_terakhir} KM")
            
            km_baru = int(input("Masukkan angka KM terbaru dari foto WA: ").strip())
            
            if km_baru < km_bengkel_terakhir:
                print(f"\n[DITOLAK] Angka tidak logis! KM tidak mungkin mundur.")
            else:
                simpan_km_harian(nopol_pilihan, km_baru)
        except ValueError:
            print("[!] Input tidak valid. Harus berupa angka.")
            
        lanjut = input("\nLanjut input KM lagi? (y/n): ").strip().lower()
        if lanjut != 'y':
            break

def setting_interval_khusus():
    while True:
        print("\n--- SETING INTERVAL GANTI OLI KHUSUS ---")
        keyword = input("Masukkan NOPOL mobil (atau '0' ke menu): ").strip()
        
        if keyword == '0' or not keyword: 
            break

        hasil = cari_kendaraan(keyword)
        if hasil.empty:
            print(f"[!] Mobil tidak ditemukan.")
            continue
            
        print("\nKendaraan Ditemukan:")
        for idx, row in hasil.iterrows():
            print(f"{idx + 1}. NOPOL: {row['nopol']} | Unit: {row['brand']} {row['series']}")
            
        try:
            pilih = int(input("\nPilih nomor kendaraan (atau 0 untuk cari lagi): "))
            if pilih == 0 or pilih > len(hasil): 
                continue
                
            nopol_pilihan = hasil.iloc[pilih-1]['nopol']
            print(f"\n=> Anda mengatur interval khusus untuk: {nopol_pilihan}")
            interval_baru = int(input("Masukkan angka interval (Contoh: 7000 atau 8000): "))
            
            if interval_baru < 1000:
                print("[DITOLAK] Interval terlalu kecil, tidak logis!")
            else:
                simpan_interval_khusus(nopol_pilihan, interval_baru)
        except ValueError:
            print("[!] Input harus berupa angka.")
            
        lanjut = input("\nLanjut seting interval lagi? (y/n): ").strip().lower()
        if lanjut != 'y':
            break

def lihat_driver_terbaru():
    print("\n[PROSES] Mengambil data driver terbaru (Status: OPEN)...")
    df = get_daftar_driver_terbaru()
    if df.empty:
        print("[!] Belum ada data handover atau tidak ada status open.")
        return
        
    print(f"\nTotal Unit Aktif Terdeteksi: {len(df)} Kendaraan\n")
    print(f"{'No':<4} | {'NOPOL':<10} | {'DRIVER':<22} | {'PHONE':<13} | {'GPS':<12} | {'UNIT'}")
    print("-" * 85)
    
    for idx, row in df.iterrows():
        unit = f"{row['brand']} {row['series']}".strip()
        print(f"{idx+1:<4} | {row['nopol']:<10} | {str(row['name'])[:22]:<22} | {str(row['phone'])[:13]:<13} | {str(row['gps'])[:12]:<12} | {unit}")
    print("-" * 85)
    print("\n[INFO] Untuk Export Excel lengkap, gunakan Menu 5.")

def main():
    while True:
        print_menu()
        pilihan = input("Pilih menu (1-7): ")
        
        if pilihan == '1':
            init_db()
        elif pilihan == '2':
            print("\n[PROSES] Sinkronisasi Data ERP...")
            import_file_to_db('data_excel/VehicleData.xls', 'vehicles')
            import_file_to_db('data_excel/handover.xls', 'handover')
            import_file_to_db('data_excel/ActualList.xls', 'actual_lists')
            print("[SELESAI] Database ter-update.")
        elif pilihan == '3':
            input_km_harian()
        elif pilihan == '4':
            setting_interval_khusus()
        elif pilihan == '5':
            print("\n[PROSES] Menghitung Sisa KM & Export Laporan Master...")
            process_and_export('LAPORAN_LENGKAP_PITSTOP.xlsx')
        elif pilihan == '6':
            lihat_driver_terbaru()
        elif pilihan == '7':
            print("\nMematikan sistem...")
            sys.exit(0)
        else:
            print("\n[WARNING] Pilihan salah.")

if __name__ == "__main__":
    main()