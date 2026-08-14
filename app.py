import streamlit as st
import pandas as pd
import os
from db_manager import init_db, import_file_to_db, cari_kendaraan, simpan_km_harian, simpan_interval_khusus, get_km_bengkel_terakhir, get_km_wa_terakhir
from logic_processor import process_and_export

# Setup tampilan halaman Web (Lebar dan Judul)
st.set_page_config(page_title="Pitstop Dashboard", page_icon="🚗", layout="wide")

# Membuat folder penampung jika belum ada
os.makedirs("data_excel", exist_ok=True)

# SIDEBAR MENU (Bagian Kiri)
st.sidebar.title("📱 Menu Pitstop")
menu = st.sidebar.radio("Navigasi:", [
    "📊 Dashboard & Laporan", 
    "📝 Input KM Harian", 
    "⚙️ Seting Interval Khusus", 
    "🔄 Sinkronisasi Data ERP"
])
st.sidebar.markdown("---")
st.sidebar.info("Aplikasi terhubung dengan jaringan lokal. Bisa diakses via HP/Laptop lain di WiFi yang sama.")

# HALAMAN 1: DASHBOARD & LAPORAN
if menu == "📊 Dashboard & Laporan":
    st.title("📊 Dashboard Laporan Sisa KM & Driver Aktif")
    st.write("Klik tombol di bawah untuk menghitung ulang data terbaru dari database.")
    
    if st.button("🔄 Generate Laporan Terbaru", type="primary"):
        with st.spinner("Sedang memproses dan menghitung sisa KM..."):
            # Jalankan mesin pemroses yang sudah kita buat
            process_and_export('LAPORAN_LENGKAP_PITSTOP.xlsx')
            
            # Baca hasil Excelnya untuk ditampilkan di Web
            if os.path.exists('LAPORAN_LENGKAP_PITSTOP.xlsx'):
                df_hasil = pd.read_excel('LAPORAN_LENGKAP_PITSTOP.xlsx')
                st.success(f"Berhasil! Menampilkan {len(df_hasil)} unit kendaraan aktif.")
                
                # Tombol Download Excel untuk User
                with open('LAPORAN_LENGKAP_PITSTOP.xlsx', 'rb') as f:
                    st.download_button(
                        label="📥 Download File Excel",
                        data=f,
                        file_name="LAPORAN_LENGKAP_PITSTOP.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Tampilkan tabel interaktif di web
                st.dataframe(df_hasil, use_container_width=True)
            else:
                st.error("Gagal membuat laporan.")

# HALAMAN 2: INPUT KM HARIAN
# HALAMAN 2: INPUT KM HARIAN
elif menu == "📝 Input KM Harian":
    st.title("📝 Input KM Harian (Dari Foto WA)")
    
    keyword = st.text_input("🔍 Cari Kendaraan (Masukkan NOPOL atau Nama Driver):")
    
    if keyword:
        hasil = cari_kendaraan(keyword)
        if hasil.empty:
            st.warning("Kendaraan atau Driver tidak ditemukan.")
        else:
            pilihan = hasil.apply(lambda row: f"{row['nopol']} - {row['name']} ({row['brand']} {row['series']}) | GPS: {row['gps']}", axis=1).tolist()
            terpilih = st.selectbox("Pilih Kendaraan:", pilihan)
            
            if terpilih:
                nopol_pilihan = terpilih.split(" - ")[0]
                
                # Ambil data KM Bengkel dan KM WA (beserta tanggal)
                km_bengkel = get_km_bengkel_terakhir(nopol_pilihan)
                data_wa = get_km_wa_terakhir(nopol_pilihan)
                km_wa = data_wa['km']
                tanggal_wa = data_wa['tanggal']
                
                st.info(f"Odometer terakhir di bengkel : **{km_bengkel} KM**")
                
                if km_wa > 0:
                    st.success(f"Odometer terbaru (Update WA) : **{km_wa} KM** *(Diinput pada: {tanggal_wa})*")
                else:
                    st.warning("Belum ada data update KM dari WA.")
                
                batas_minimum = max(km_bengkel, km_wa)
                
                st.markdown("---")
                km_baru = st.number_input("Masukkan Angka KM Terbaru dari Foto:", min_value=0, step=1)
                
                if st.button("💾 Simpan KM"):
                    if km_baru < batas_minimum:
                        st.error(f"❌ Angka tidak logis! KM tidak boleh lebih kecil dari {batas_minimum} KM.")
                    else:
                        simpan_km_harian(nopol_pilihan, km_baru)
                        st.success(f"✅ KM terbaru untuk {nopol_pilihan} berhasil disimpan!")# HALAMAN 3: SETING INTERVAL

elif menu == "⚙️ Seting Interval Khusus":
    st.title("⚙️ Seting Interval Khusus (7000 / 8000 KM)")
    
    keyword = st.text_input("🔍 Cari Kendaraan (NOPOL/Driver):")
    if keyword:
        hasil = cari_kendaraan(keyword)
        if hasil.empty:
            st.warning("Kendaraan tidak ditemukan.")
        else:
            # Dropdown di sini juga ditambah GPS agar seragam
            pilihan = hasil.apply(lambda row: f"{row['nopol']} - {row['name']} ({row['brand']} {row['series']}) | GPS: {row['gps']}", axis=1).tolist()
            terpilih = st.selectbox("Pilih Kendaraan:", pilihan)
            
            if terpilih:
                nopol_pilihan = terpilih.split(" - ")[0]
                interval_baru = st.number_input("Masukkan Interval Baru (Cth: 7000):", min_value=1000, step=1000)
                
                if st.button("💾 Simpan Interval"):
                    simpan_interval_khusus(nopol_pilihan, interval_baru)
                    st.success(f"✅ Interval {nopol_pilihan} diset menjadi {interval_baru} KM.")
# HALAMAN 4: SINKRONISASI DATA ERP
elif menu == "🔄 Sinkronisasi Data ERP":
    st.title("🔄 Upload & Sinkronisasi Data ERP")
    
    if st.button("🛠️ Inisialisasi Database Pertama Kali"):
        init_db()
        st.success("Database berhasil disiapkan!")
        
    st.markdown("### Upload File Excel dari ERP")
    file_vehicle = st.file_uploader("Upload VehicleData.xls", type=['xls', 'xlsx', 'html'])
    file_handover = st.file_uploader("Upload handover.xls", type=['xls', 'xlsx', 'html'])
    file_actual = st.file_uploader("Upload ActualList.xls", type=['xls', 'xlsx', 'html'])
    
    if st.button("🚀 Proses Sinkronisasi", type="primary"):
        with st.spinner("Menyinkronkan data ke sistem..."):
            # Simpan file yang diupload ke folder data_excel, lalu import
            if file_vehicle:
                path_v = os.path.join("data_excel", file_vehicle.name)
                with open(path_v, "wb") as f: f.write(file_vehicle.getbuffer())
                import_file_to_db(path_v, 'vehicles')
                st.write("✅ Data Kendaraan Terupdate")
                
            if file_handover:
                path_h = os.path.join("data_excel", file_handover.name)
                with open(path_h, "wb") as f: f.write(file_handover.getbuffer())
                import_file_to_db(path_h, 'handover')
                st.write("✅ Data Handover Terupdate")
                
            if file_actual:
                path_a = os.path.join("data_excel", file_actual.name)
                with open(path_a, "wb") as f: f.write(file_actual.getbuffer())
                import_file_to_db(path_a, 'actual_lists')
                st.write("✅ Data Bengkel Terupdate")
            
            if file_vehicle or file_handover or file_actual:
                st.success("Sinkronisasi Selesai!")
            else:
                st.warning("Belum ada file yang diunggah.")