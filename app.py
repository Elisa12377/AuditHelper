import streamlit as st
import pandas as pd
import io
import openpyxl
import os
import google.generativeai as genai

# --- KONFIGURASI AI ---
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY", "GANTI_DENGAN_KEY_KAMU")
genai.configure(api_key=api_key)
ai_model = genai.GenerativeModel('gemini-3.5-flash')

st.set_page_config(page_title="AuditPilot — AI-Powered Audit Analytics", layout="wide", initial_sidebar_state="expanded")

# --- INITIALIZATION SESSION STATE ---
if "rooms" not in st.session_state:
    st.session_state.rooms = ["PT Jaya Abadi 2026", "CV Makmur Sejahtera"]
if "active_room" not in st.session_state:
    st.session_state.active_room = None  # None artinya sedang berada di Dashboard Utama
if "client_vault" not in st.session_state:
    st.session_state.client_vault = {}
if "chat_histories" not in st.session_state:
    st.session_state.chat_histories = {}
if "ai_reports" not in st.session_state:
    st.session_state.ai_reports = {}

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

# ================= 1. SIDEBAR: NAVIGASI UTAMA =================
with st.sidebar:
    st.title("✈️ AuditPilot Menu")
    
    if st.button("🏠 Kembali ke Dashboard Utama"):
        st.session_state.active_room = None
        st.rerun()
        
    st.markdown("---")
    st.subheader("Daftar Klien Aktif")
    for room in st.session_state.rooms:
        if st.button(f"📂 {room}", key=f"side_room_{room}"):
            st.session_state.active_room = room
            st.rerun()

# ================= 2. KONTROL HALAMAN (DASHBOARD VS WORKSPACE KLIEN) =================

if st.session_state.active_room is None:
    # --- TAMPILAN DASHBOARD UTAMA ---
    st.title("📊 Audit Dashboard & Engagement Manager")
    st.markdown("Selamat datang di pusat kendali audit. Pilih ruang klien di bawah ini atau buat penugasan audit baru untuk mulai mengunggah dokumen dan berinteraksi dengan AI Copilot.")
    st.markdown("---")
    
    col_dash1, col_dash2 = st.columns([1.5, 1])
    
    with col_dash1:
        st.subheader("🏢 Daftar Ruang Klien (Rooms)")
        if st.session_state.rooms:
            for room in st.session_state.rooms:
                jumlah_dok = len(st.session_state.client_vault.get(room, []))
                col_r1, col_r2 = st.columns([3, 1])
                with col_r1:
                    st.write(f"**{room}** — *{jumlah_dok} Dokumen di Brankas*")
                with col_r2:
                    if st.button("Buka Workspace", key=f"btn_open_{room}"):
                        st.session_state.active_room = room
                        st.rerun()
                st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)
        else:
            st.info("Belum ada klien. Silakan tambahkan klien baru.")

    with col_dash2:
        st.subheader("➕ Tambah Klien Baru")
        with st.form(key="form_add_room"):
            new_client_name = st.text_input("Nama Perusahaan / Klien & Tahun:", placeholder="Misal: PT Makmur 2026")
            submit_client = st.form_submit_button("Buat Ruang Klien")
            if submit_client and new_client_name:
                if new_client_name not in st.session_state.rooms:
                    st.session_state.rooms.append(new_client_name)
                    st.success(f"Berhasil membuat ruang untuk {new_client_name}!")
                    st.rerun()

else:
    # --- TAMPILAN WORKSPACE KLIEN (SETELAH KLIEN DIPILIH) ---
    klien_ini = st.session_state.active_room
    
    col_top1, col_top2 = st.columns([6, 1])
    with col_top1:
        st.title(f"📁 Workspace Klien: {klien_ini}")
    with col_top2:
        if st.button("🚪 Tutup Ruang"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("---")
    
    # Alur Navigasi Step-by-Step dalam bentuk Tab
    tab_step1, tab_step2, tab_step3 = st.tabs([
        "📂 1. Upload & Document Vault", 
        "📊 2. Analytics & Review", 
        "🤖 3. AI Copilot & Reporting"
    ])

    # --- STEP 1: UPLOAD & DOCUMENT VAULT ---
    with tab_step1:
        st.subheader("Langkah 1: Pengelolaan Dokumen Bertahap")
        st.info("Upload dokumen klien secara fleksibel (GL, TB, Rekening Koran, Stock Opname, dll.). Dokumen akan terakumulasi secara otomatis.")
        
        with st.form(key=f"form_upload_workspace_{klien_ini}", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                 jenis_dok = st.selectbox(
                    "Jenis Dokumen:",
                    ["General Ledger (GL)", "Trial Balance (TB)", "Rekening Koran (Bank)", "Stock Opname", "Cash Opname", "Lainnya"]
                )
            with c2:
                file_upl = st.file_uploader("Pilih File Excel:", type=["xlsx", "xls"])
                
            btn_sub_upl = st.form_submit_button("📥 Simpan ke Brankas Klien")
            if btn_sub_upl and file_upl is not None:
                try:
                    df_v = pd.read_excel(file_upl).astype(str)
                    if klien_ini not in st.session_state.client_vault:
                        st.session_state.client_vault[klien_ini] = []
                    
                    st.session_state.client_vault[klien_ini].append({
                        "nama_file": file_upl.name,
                        "jenis": jenis_dok,
                        "data": df_v,
                        "total_baris": len(df_v)
                    })
                    st.success(f"Berhasil mengunggah {file_upl.name}!")
                except Exception as e:
                    st.error(f"Gagal membaca file: {e}")

        st.markdown("---")
        st.write("### Daftar Dokumen di Brankas Klien Ini:")
        vault_data = st.session_state.client_vault.get(klien_ini, [])
        if vault_data:
            for d in vault_data:
                with st.expander(f"📄 [{d['jenis']}] — {d['nama_file']} ({d['total_baris']} Baris)"):
                    st.dataframe(d['data'].head(15), use_container_width=True)
        else:
            st.warning("Brankas masih kosong. Silakan upload dokumen di atas.")

    # --- STEP 2: ANALYTICS & REVIEW ---
    with tab_step2:
        st.subheader("Langkah 2: Analytics & Pemeriksaan Data")
        vault_data = st.session_state.client_vault.get(klien_ini, [])
        
        if vault_data:
            file_names = [d['nama_file'] for d in vault_data]
            pilihan_file = st.selectbox("Pilih file yang ingin direview:", file_names)
            selected_doc = next(d for d in vault_data if d['nama_file'] == pilihan_file)
            
            st.write(f"Menampilkan preview data untuk: **{selected_doc['nama_file']}**")
            st.dataframe(selected_doc['data'], use_container_width=True, height=350)
        else:
            st.info("Belum ada dokumen yang tersedia untuk dianalisis. Silakan lakukan upload di Langkah 1.")

    # --- STEP 3: AI COPILOT & REPORTING ---
    with tab_step3:
        st.subheader("Langkah 3: Interaksi AI Copilot & Penyusunan Laporan")
        vault_data = st.session_state.client_vault.get(klien_ini, [])
        
        if klien_ini not in st.session_state.chat_histories:
            st.session_state.chat_histories[klien_ini] = [{"role": "ai", "content": f"Halo! Saya AI Copilot siap membantu menganalisis dokumen untuk klien {klien_ini}."}]
            
        chat_box = st.container(height=350)
        with chat_box:
            for chat in st.session_state.chat_histories[klien_ini]:
                with st.chat_message(chat["role"]):
                    st.write(chat["content"])

        user_input = st.chat_input("Tanya AI tentang dokumen klien ini...")
        if user_input:
            st.session_state.chat_histories[klien_ini].append({"role": "user", "content": user_input})
            with chat_box:
                with st.chat_message("user"):
                    st.write(user_input)
            
            with chat_box:
                with st.chat_message("ai"):
                    with st.spinner("AI sedang memeriksa brankas dokumen..."):
                        try:
                            ringkasan_ctx = ""
                            for doc in vault_data:
                                ringkasan_ctx += f"\n- Dokumen {doc['jenis']} ({doc['nama_file']}): {doc['data'].head(15).to_dict(orient='records')}\n"
                            
                            if not ringkasan_ctx:
                                ringkasan_ctx = "Belum ada dokumen di-upload."

                            sys_prompt = f"""
                            Anda adalah Senior Audit AI di KAP untuk klien {klien_ini}.
                            Dokumen tersedia: {ringkasan_ctx}
                            Pertanyaan Auditor: {user_input}
                            Jawab secara profesional berdasarkan data dokumen di atas.
                            """
                            resp = ai_model.generate_content(sys_prompt)
                            ans = resp.text
                        except Exception as e:
                            ans = f"Error AI: {e}"
                        st.write(ans)
                        
            st.session_state.chat_histories[klien_ini].append({"role": "ai", "content": ans})
            st.rerun()

        st.markdown("---")
        if vault_data:
            if st.button("✨ Buat Executive Audit Summary Report Otomatis"):
                with st.spinner("Menyusun laporan..."):
                    try:
                        ctx_rep = ""
                        for doc in vault_data:
                            ctx_rep += f"File {doc['jenis']} ({doc['nama_file']}): {doc['data'].head(10).to_dict(orient='records')}\n"
                            
                        p_sum = f"Buat draf Laporan Audit Eksekutif formal untuk klien {klien_ini} berdasarkan data berikut:\n{ctx_rep}"
                        r_sum = ai_model.generate_content(p_sum)
                        st.session_state.ai_reports[klien_ini] = r_sum.text
                        st.success("Laporan berhasil dibuat!")
                    except Exception as e:
                        st.error(f"Gagal: {e}")

            if klien_ini in st.session_state.ai_reports:
                st.text_area("Draf Laporan Resmi:", value=st.session_state.ai_reports[klien_ini], height=250)
