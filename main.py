import sys
from db_manager import init_db, import_file_to_db, cari_kendaraan, simpan_km_harian, get_km_bengkel_terakhir, simpan_interval_khusus
from logic_processor import process_and_export

def print_menu():
    print("\n" + "="*55)
    print(" 📱 SISTEM TRACKING PITSTOP OLI - AKTIF 📱")
    print("="*55)
    print(" 1. Inisialisasi Database (Awal Saja)")
    print(" 2. Sinkronisasi Data ERP (Upload File .xls)")
    print(" 3. Input KM Harian (Dari Foto WA Driver)")
    print(" 4. Seting Interval Khusus (7000/8000 KM)")
    print(" 5. Hitung & Export Laporan Excel")
    print(" 6. Keluar")
    print("="*55)

def input_km_harian():
    print("\n--- PENCARIAN KENDARAAN ---")
    keyword = input("Masukkan NOPOL atau Nama Driver: ").strip()
    if not keyword: return
    
    hasil = cari_kendaraan(keyword)
    if hasil.empty:
        print(f"[!] Tidak ada driver atau nopol yang cocok dengan '{keyword}'.")
        return
        
    print("\nKendaraan Ditemukan:")
    for idx, row in hasil.iterrows():
        print(f"{idx + 1}. NOPOL: {row['nopol']} | Driver: {row['name']} | Unit: {row['brand']} {row['series']}")
        
    try:
        pilih = int(input("\nPilih nomor kendaraan (atau 0 untuk batal): "))
        if pilih == 0 or pilih > len(hasil): return
            
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

def setting_interval_khusus():
    print("\n--- SETING INTERVAL GANTI OLI KHUSUS ---")
    keyword = input("Masukkan NOPOL mobil: ").strip()
    if not keyword: return

    hasil = cari_kendaraan(keyword)
    if hasil.empty:
        print(f"[!] Mobil tidak ditemukan.")
        return
        
    print("\nKendaraan Ditemukan:")
    for idx, row in hasil.iterrows():
        print(f"{idx + 1}. NOPOL: {row['nopol']} | Unit: {row['brand']} {row['series']}")
        
    try:
        pilih = int(input("\nPilih nomor kendaraan (0 batal): "))
        if pilih == 0 or pilih > len(hasil): return
            
        nopol_pilihan = hasil.iloc[pilih-1]['nopol']
        print(f"\n=> Anda mengatur interval khusus untuk: {nopol_pilihan}")
        interval_baru = int(input("Masukkan angka interval (Contoh: 7000 atau 8000): "))
        
        if interval_baru < 1000:
            print("[DITOLAK] Interval terlalu kecil, tidak logis!")
        else:
            simpan_interval_khusus(nopol_pilihan, interval_baru)
    except ValueError:
        print("[!] Input harus berupa angka.")

def main():
    while True:
        print_menu()
        pilihan = input("Pilih menu (1-6): ")
        
        if pilihan == '1':
            init_db()
        elif pilihan == '2':
            print("\n[PROSES] Sinkronisasi Data ERP...")
            import_file_to_db('data_excel/VehicleData.xls', 'vehicles')
            import_file_to_db('data_excel/handover.xls', 'handovers')
            import_file_to_db('data_excel/ActualList.xls', 'actual_lists') # PASTIKAN ADA .xls
            print("[SELESAI] Database ter-update.")
        elif pilihan == '3':
            input_km_harian()
        elif pilihan == '4':
            setting_interval_khusus()
        elif pilihan == '5':
            print("\n[PROSES] Menghitung Sisa KM & Export Laporan...")
            process_and_export('UPDATE_SERVICE_PITSTOP.xlsx')
        elif pilihan == '6':
            print("\nMematikan sistem...")
            sys.exit(0)
        else:
            print("\n[WARNING] Pilihan salah.")

if __name__ == "__main__":
    main()