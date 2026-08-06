import streamlit as st
import pandas as pd
import io
import openpyxl
import os
import google.generativeai as genai

# --- KOKURIKULUM & KONFIGURASI AI (AMAN UNTUK GITHUB) ---
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY", "GANTI_DENGAN_KEY_KAMU")
genai.configure(api_key=api_key)
ai_model = genai.GenerativeModel('gemini-3.5-flash')

st.set_page_config(page_title="AuditPilot — AI-Powered Audit Analytics", layout="wide", initial_sidebar_state="expanded")

# --- INITIALIZATION SESSION STATE ---
if "rooms" not in st.session_state:
    st.session_state.rooms = ["Klien Default"]
if "active_room" not in st.session_state:
    st.session_state.active_room = "Klien Default"
if "chat_histories" not in st.session_state:
    st.session_state.chat_histories = {
        "Klien Default": [{"role": "ai", "content": "Halo! Silakan upload file General Ledger di panel kanan dan mulai analisis bersama AI."}]
    }
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data = {}
if "ai_reports" not in st.session_state:
    st.session_state.ai_reports = {}
if "show_panel" not in st.session_state:
    st.session_state.show_panel = True

# --- CSS STYLING MODERN ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }
    .stApp { background-color: #F8F9FA; color: #121212; }
    
    [data-testid="stSidebar"] {
        background-color: #0F1115 !important;
        border-right: 1px solid #2A2D35;
    }
    [data-testid="stSidebar"] div, [data-testid="stSidebar"] p, [data-testid="stSidebar"] h1 { 
        color: #FFFFFF !important; 
    }
    
    .stButton>button { 
        background-color: #556B2F !important; 
        color: #FFFFFF !important; 
        border-radius: 8px; 
        border: none;
        font-weight: 600;
        width: 100%; 
    }
    .stButton>button:hover { background-color: #6B8E23 !important; }
    h1, h2, h3 { color: #556B2F; font-weight: 700; letter-spacing: -0.5px; }
</style>
""", unsafe_allow_html=True)

# ================= 1. SIDEBAR: MANAJEMEN RUANG KERJA (ROOMS) =================
with st.sidebar:
    st.title("📁 Daftar Klien")
    new_room = st.text_input("Buat Ruang Kerja Baru:", placeholder="Misal: PT ABC 2026")
    if st.button("➕ Tambah Room"):
        if new_room and new_room not in st.session_state.rooms:
            st.session_state.rooms.append(new_room)
            st.session_state.chat_histories[new_room] = [{"role": "ai", "content": f"Halo! Ruang kerja {new_room} siap digunakan."}]
            st.session_state.active_room = new_room
            st.rerun()
            
    room_pilihan = st.selectbox("Pilih Ruang Kerja:", st.session_state.rooms, index=st.session_state.rooms.index(st.session_state.active_room))
    if room_pilihan != st.session_state.active_room:
        st.session_state.active_room = room_pilihan
        st.rerun()
        
    st.markdown("---")
    st.markdown("💡 *Gunakan tombol di kanan atas untuk menyembunyikan/menampilkan panel kerja.*")

# ================= 2. HALAMAN UTAMA =================
col_top1, col_top2 = st.columns([8, 2])
with col_top1:
    st.title(f"✈️ AuditPilot ({st.session_state.active_room})")
with col_top2:
    if st.button("🎛️ Panel Data & Filter"):
        st.session_state.show_panel = not st.session_state.show_panel
        st.rerun()

st.markdown("---")

if st.session_state.show_panel:
    chat_col, panel_col = st.columns([1.2, 1])
else:
    chat_col, panel_col = st.columns([1, 0.0001])

# --- KOLOM TENGAH: AI CHAT MANAGER ---
with chat_col:
    st.subheader("Asisten Investigasi AI")
    
    chat_container = st.container(height=500)
    with chat_container:
        for chat in st.session_state.chat_histories[st.session_state.active_room]:
            with st.chat_message(chat["role"]):
                st.write(chat["content"])

    prompt = st.chat_input(f"Tanya soal data {st.session_state.active_room}...")
    if prompt:
        st.session_state.chat_histories[st.session_state.active_room].append({"role": "user", "content": prompt})
        with chat_container:
            with st.chat_message("user"):
                st.write(prompt)
        
        with chat_container:
            with st.chat_message("ai"):
                with st.spinner("AI sedang menganalisis arsip..."):
                    try:
                        df_aktif = st.session_state.uploaded_data.get(st.session_state.active_room, pd.DataFrame())
                        data_ringkas = df_aktif.head(100).to_dict(orient="records") if not df_aktif.empty else "Belum ada data di-upload."
                        
                        prompt_sistem = f"""
                        Kamu adalah Asisten Investigasi Audit Senior di Kantor Akuntan Publik (KAP). 
                        Klien/Ruang Kerja aktif: {st.session_state.active_room}.
                        --- DATA SAMPEL TRANSAKSI ---
                        {data_ringkas}
                        -----------------------------
                        Pertanyaan Auditor: {prompt}
                        
                        Jawab secara spesifik, profesional, dan berdasarkan data di atas.
                        """
                        response = ai_model.generate_content(prompt_sistem)
                        jawaban = response.text
                    except Exception as e:
                        jawaban = f"Terjadi kesalahan pada AI: {e}"
                        
                    st.write(jawaban)
                    
        st.session_state.chat_histories[st.session_state.active_room].append({"role": "ai", "content": jawaban})
        st.rerun()

# --- KOLOM KANAN: PANEL WORKSPACE, FILTER & AI REPORT ---
if st.session_state.show_panel:
    with panel_col:
        st.subheader("📂 Panel Workspace & Audit Rules")
        uploaded_file = st.file_uploader("Upload file General Ledger (Excel)", type=["xlsx", "xls"], key=f"uploader_{st.session_state.active_room}")

        if uploaded_file is not None:
            xls = pd.ExcelFile(uploaded_file)
            sheet_terpilih = st.selectbox("Pilih Sheet:", xls.sheet_names, key=f"sel_sheet_{st.session_state.active_room}")
            df_mentah = pd.read_excel(uploaded_file, sheet_name=sheet_terpilih, header=None)

            with st.spinner("Merapikan struktur kolom..."):
                kata_kunci = ['tanggal', 'date', 'akun', 'coa', 'account', 'keterangan', 'deskripsi', 'uraian', 'debit', 'debet', 'kredit', 'credit', 'mutasi', 'saldo']
                baris_header, skor_tertinggi = 0, 0
                for i in range(min(20, len(df_mentah))):
                    baris = df_mentah.iloc[i].tolist()
                    skor = sum(any(kunci in str(sel).lower() for kunci in kata_kunci) for sel in baris)
                    if skor > skor_tertinggi:
                        skor_tertinggi, baris_header = skor, i
                if skor_tertinggi >= 1:
                    df_mentah.columns = df_mentah.iloc[baris_header]
                    df_mentah = df_mentah.iloc[baris_header + 1:].reset_index(drop=True)
                
                cols = [str(col).strip().lower() for col in df_mentah.columns]
                seen = {}
                new_cols = []
                for c in cols:
                    if c in seen:
                        seen[c] += 1
                        new_cols.append(f"{c}_{seen[c]}")
                    else:
                        seen[c] = 0
                        new_cols.append(c)
                df_mentah.columns = new_cols

                df_mentah = df_mentah.loc[:, ~df_mentah.columns.astype(str).str.contains('unnamed|nan', case=False, na=False)]
                df_mentah = df_mentah.loc[:, df_mentah.columns != '']
                df = df_mentah.dropna(axis=1, how='all').dropna(axis=0, how='all').astype(str)
            
            st.session_state.uploaded_data[st.session_state.active_room] = df

        if st.session_state.active_room in st.session_state.uploaded_data:
            df = st.session_state.uploaded_data[st.session_state.active_room]

            st.markdown("---")
            search_query = st.text_input("🔍 Cari transaksi...", "", key=f"search_{st.session_state.active_room}")
            
            st.write("### ⚡ Parameter & Rules Audit")
            kolom_angka = st.selectbox("Pilih kolom nominal/saldo:", df.columns, key=f"col_angka_{st.session_state.active_room}")
            materialitas = st.number_input("Batas Materialitas (Rp)", min_value=0, value=100000000, step=1000000, key=f"mat_{st.session_state.active_room}")
            chk_material = st.checkbox("🚨 Aktifkan Filter Materialitas", key=f"chk_mat_{st.session_state.active_room}")

            kolom_teks = st.selectbox("Pilih kolom Keterangan/Deskripsi:", df.columns, key=f"col_teks_{st.session_state.active_room}")
            chk_kosong = st.checkbox("👻 Aktifkan Filter Ket. Kosong", key=f"chk_kosong_{st.session_state.active_room}")

            # --- ANOMALY RULES ---
            st.markdown("#### 🤖 Deteksi Anomali Otomatis")
            chk_weekend = st.checkbox("📅 Deteksi Transaksi Weekend (Sabtu / Minggu)", key=f"chk_wknd_{st.session_state.active_room}")
            chk_round = st.checkbox("🎯 Deteksi Angka Bulat Mencurigakan (Kelipatan 1 Juta)", key=f"chk_rnd_{st.session_state.active_room}")
            
            col_duplikat = st.selectbox("Pilih kolom acuan Duplikasi:", df.columns, key=f"col_dup_{st.session_state.active_room}")
            chk_duplicate = st.checkbox("🔁 Deteksi Nilai Kembar Berdasarkan Kolom", key=f"chk_dup_{st.session_state.active_room}")

            # Logika Filter & Rule Gabungan
            df_tampil = df.copy()
            
            if chk_material:
                series_angka = df_tampil[kolom_angka]
                if isinstance(series_angka, pd.DataFrame):
                    series_angka = series_angka.iloc[:, 0]
                angka_bersih = pd.to_numeric(series_angka.astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce')
                df_tampil = df_tampil[angka_bersih > materialitas]
                
            if chk_kosong:
                series_teks = df_tampil[kolom_teks]
                if isinstance(series_teks, pd.DataFrame):
                    series_teks = series_teks.iloc[:, 0]
                mask_kosong = series_teks.astype(str).str.strip().str.lower().isin(['', 'nan', 'none', 'null'])
                df_tampil = df_tampil[mask_kosong]

            if chk_weekend:
                col_date = next((col for col in df_tampil.columns if 'tanggal' in col or 'date' in col), None)
                if col_date:
                    parsed_dates = pd.to_datetime(df_tampil[col_date], errors='coerce')
                    df_tampil = df_tampil[parsed_dates.dt.dayofweek.isin([5, 6])]

            if chk_round:
                angka_seri = pd.to_numeric(df_tampil[kolom_angka].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce').fillna(0)
                df_tampil = df_tampil[(angka_seri > 0) & (angka_seri % 1000000 == 0)]

            if chk_duplicate and col_duplikat:
                series_dup = df_tampil[col_duplikat]
                if isinstance(series_dup, pd.DataFrame):
                    series_dup = series_dup.iloc[:, 0]
                mask_dup = series_dup.duplicated(keep=False)
                df_tampil = df_tampil[mask_dup]

            if search_query:
                mask_search = df_tampil.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)
                df_tampil = df_tampil[mask_search]

            # --- METRIC CARDS ---
            st.markdown("---")
            m1, m2 = st.columns(2)
            m1.metric("Total Data Asli", f"{len(df)} Baris")
            m2.metric("Temuan Saringan", f"{len(df_tampil)} Baris")
                
            st.write(f"📋 **Preview Temuan**")
            st.dataframe(df_tampil, use_container_width=True, height=250)

            # --- AI EXECUTIVE SUMMARY REPORT ---
            st.markdown("---")
            st.write("### 📝 AI Executive Summary Report")
            if st.button("✨ Buat Laporan Naratif Otomatis", key=f"btn_summary_{st.session_state.active_room}"):
                with st.spinner("AI sedang menyusun laporan audit formal..."):
                    try:
                        data_sampel_laporan = df_tampil.head(50).to_dict(orient="records")
                        prompt_summary = f"""
                        Bertindaklah sebagai Manajer Audit Senior di Kantor Akuntan Publik (KAP). 
                        Buatkan Laporan Ringkasan Temuan Audit (Executive Audit Summary Report) yang formal dan profesional berdasarkan data sampel transaksi hasil saringan berikut untuk klien: {st.session_state.active_room}.
                        
                        --- DATA SAMPEL TEMUAN ---
                        {data_sampel_laporan}
                        --------------------------
                        Format laporan wajib menggunakan struktur baku:
                        1. Ringkasan Eksekutif
                        2. Temuan Utama (Key Findings)
                        3. Risiko Audit & Asersi Terkait
                        4. Rekomendasi & Prosedur Lanjutan
                        Gunakan Bahasa Indonesia baku, objektif, dan siap disalin ke Microsoft Word.
                        """
                        res_summary = ai_model.generate_content(prompt_summary)
                        st.session_state.ai_reports[st.session_state.active_room] = res_summary.text
                        st.success("✅ Laporan berhasil disusun!")
                    except Exception as e:
                        st.error(f"Gagal menyusun laporan: {e}")

            if st.session_state.active_room in st.session_state.ai_reports:
                st.text_area(
                    "Hasil Draf Laporan (Salin ke Word):",
                    value=st.session_state.ai_reports[st.session_state.active_room],
                    height=250,
                    key=f"txtarea_report_{st.session_state.active_room}"
                )

            # --- EXPORT KKA FORMAL ---
            st.markdown("---")
            st.write("### 📥 Export Kertas Kerja Audit (KKA)")
            
            df_kka = df_tampil.copy()
            df_kka['Catatan / Temuan Auditor'] = "Perlu konfirmasi & vouching ke dokumen sumber"
            df_kka['Tickmark'] = "✔"
            
            output_export = io.BytesIO()
            with pd.ExcelWriter(output_export, engine='openpyxl') as writer:
                df_kka.to_excel(writer, index=False, sheet_name='KKA_Temuan', startrow=5)
                workbook = writer.book
                worksheet = workbook.active
                
                worksheet['A1'] = "KERTAS KERJA AUDIT (KKA)"
                worksheet['A1'].font = openpyxl.styles.Font(size=14, bold=True, color="556B2F")
                worksheet['A2'] = f"Nama Klien / Ruang Kerja: {st.session_state.active_room}"
                worksheet['A3'] = f"Jumlah Temuan / Sampel Uji: {len(df_kka)} Baris"
                worksheet['A4'] = "Status: Telah disaring otomatis oleh AuditPilot AI Engine"
                
                header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                header_fill = openpyxl.styles.PatternFill(start_color="556B2F", end_color="556B2F", fill_type="solid")
                
                for col_num in range(1, len(df_kka.columns) + 1):
                    cell = worksheet.cell(row=6, column=col_num)
                    cell.font = header_font
                    cell.fill = header_fill

            output_export.seek(0)
            st.download_button(
                label="📥 Download KKA Berformat Resmi (Excel)",
                data=output_export,
                file_name=f"KKA_Formal_{st.session_state.active_room}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_formal_{st.session_state.active_room}",
                use_container_width=True
            )
        else:
            st.info("👆 Silakan upload file Excel terlebih dahulu untuk mulai memfilter data.")