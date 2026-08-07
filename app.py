import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import os
import io
import google.generativeai as genai

# --- KONFIGURASI AI ---
try:
    api_key = st.secrets.get("GEMINI_API_KEY")
except Exception:
    api_key = os.environ.get("GEMINI_API_KEY", "GANTI_DENGAN_KEY_KAMU")

genai.configure(api_key=api_key)
ai_model = genai.GenerativeModel('gemini-3.5-flash')

st.set_page_config(page_title="AuditPilot — AI-Powered Audit Analytics", layout="wide", initial_sidebar_state="collapsed")

# --- INISIALISASI DATABASE SQLITE (UNTUK AKUN & DATA AMAN) ---
def init_db():
    conn = sqlite3.connect('auditpilot_secure.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS clients (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, client_name TEXT)''')
    conn.commit()
    return conn

db_conn = init_db()

def make_hash(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_user(username, password):
    c = db_conn.cursor()
    c.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, make_hash(password)))
    return c.fetchone()

def add_user(username, password):
    c = db_conn.cursor()
    try:
        c.execute("INSERT INTO users(username, password) VALUES (?, ?)", (username, make_hash(password)))
        db_conn.commit()
        return True
    except:
        return False

# --- FUNGSI PEMBERSIH EXCEL & DETEKSI SHEET/HEADER ---
def get_excel_sheets(uploaded_file):
    try:
        xl = pd.ExcelFile(uploaded_file)
        return xl.sheet_names
    except Exception:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file, engine='openpyxl')
        return xl.sheet_names

def clean_and_detect_header(uploaded_file, sheet_name):
    try:
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str)
    except Exception:
        uploaded_file.seek(0)
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, engine='openpyxl', dtype=str)
    
    header_row_idx = 0
    keywords = ['tanggal', 'date', 'akun', 'account', 'keterangan', 'description', 'debit', 'credit', 'kredit', 'ref', 'penyusutan', 'ppn']
    
    for idx in range(min(15, len(raw_df))):
        row_str = " ".join(raw_df.iloc[idx].dropna().astype(str).values).lower()
        if any(kw in row_str for kw in keywords):
            header_row_idx = idx
            break
            
    uploaded_file.seek(0)
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_idx, dtype=str)
    except Exception:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_idx, engine='openpyxl', dtype=str)
    
    df = df.dropna(how="all")
    df.columns = [str(col).strip() for col in df.columns]
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
    df = df.fillna("").astype(str)
    
    return df, header_row_idx

# --- INITIALIZATION SESSION STATE ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "active_room" not in st.session_state:
    st.session_state.active_room = None
if "client_vault" not in st.session_state:
    st.session_state.client_vault = {}
if "chat_histories" not in st.session_state:
    st.session_state.chat_histories = {}
if "ai_reports" not in st.session_state:
    st.session_state.ai_reports = {}

# --- CSS STYLING ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Space Grotesk', sans-serif; 
    }
    
    header[data-testid="stHeader"] {
        background: transparent !important;
        display: none !important;
    }
    [data-testid="stDecoration"] {
        display: none !important;
    }
    
    .stApp { 
        background: linear-gradient(135deg, #05070B 0%, #0D1124 50%, #150F2D 100%);
        color: #FFFFFF; 
        padding-top: 1rem;
    }
    
    [data-testid="stSidebar"] {
        display: none;
    }
    
    p, span, label, div, .stMarkdown, .stSelectbox label, .stFileUploader label {
        color: #F8FAFC !important;
    }
    
    .stAlert {
        background-color: rgba(99, 102, 241, 0.15) !important;
        color: #E2E8F0 !important;
        border: 1px solid rgba(99, 102, 241, 0.3) !important;
    }
    .stAlert p {
        color: #E2E8F0 !important;
    }
    
    div.stButton > button { 
        background: #F8FAFC !important; 
        color: #0F172A !important; 
        border-radius: 10px !important; 
        border: 1px solid #CBD5E1 !important;
        font-weight: 700 !important;
        padding: 0.5rem 1.2rem !important;
        box-shadow: none !important;
    }
    div.stButton > button:hover { 
        background: #E2E8F0 !important;
        color: #000000 !important;
        border-color: #94A3B8 !important;
    }
    div.stButton > button p, div.stButton > button span {
        color: #0F172A !important;
        font-weight: 700 !important;
    }

    div.stFormSubmitButton > button {
        background: #6366F1 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
    }
    div.stFormSubmitButton > button:hover {
        background: #4F46E5 !important; 
        color: #FFFFFF !important;
    }
    div.stFormSubmitButton > button p, div.stFormSubmitButton > button span {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    .glass-box {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        padding: 28px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
        margin-bottom: 20px;
    }
    
    .stTextInput input, .stChatInput input, textarea, div[data-baseweb="input"] input {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-weight: 600 !important;
    }
    
    textarea[aria-label="Draf Laporan Resmi / KKA (Siap Salin):"] {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-size: 1rem !important;
        border: 1px solid #6366F1 !important;
    }
    
    .stSelectbox>div>div>select {
        background-color: rgba(15, 23, 42, 0.8) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 10px !important;
    }
    
    div[data-baseweb="select"] > div {
        background-color: rgba(15, 23, 42, 0.95) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
    }
    div[data-baseweb="select"] span {
        color: #FFFFFF !important;
    }
    
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"] {
        background-color: #0F172A !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
    }
    
    div[data-baseweb="popover"] div, ul[data-baseweb="menu"] li, span[data-baseweb="tag"], div[role="option"] {
        color: #FFFFFF !important;
        background-color: #0F172A !important;
    }
    
    li[data-baseweb="option"], div[role="option"] {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
    }
    
    li[data-baseweb="option"]:hover, div[role="option"]:hover {
        background-color: #6366F1 !important;
        color: #FFFFFF !important;
    }
    
    [data-testid="stFileUploader"] {
        background-color: rgba(15, 23, 42, 0.6) !important;
        padding: 12px;
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.15);
    }
    [data-testid="stFileUploader"] section {
        background-color: rgba(255, 255, 255, 0.03) !important;
        border: 1px dashed rgba(255, 255, 255, 0.25) !important;
    }
    [data-testid="stFileUploader"] span, [data-testid="stFileUploader"] small, [data-testid="stFileUploader"] div, [data-testid="stFileUploader"] p {
        color: #E2E8F0 !important;
    }
    
    h1, h2, h3 { 
        color: #FFFFFF !important; 
        font-weight: 700; 
        letter-spacing: -0.5px; 
    }
</style>
""", unsafe_allow_html=True)

# ================= SISTEM AUTENTIKASI (LOGIN & REGISTER) =================

if not st.session_state.logged_in:
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center;'>🔐 AuditPilot Secure Portal</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #94a3b8;'>Silakan login atau daftar akun untuk mengamankan data klien Anda.</p>", unsafe_allow_html=True)
        
        tab_log, tab_reg = st.tabs(["🔑 Login", "📝 Daftar Akun Baru"])
        
        with tab_log:
            with st.form("form_login"):
                u_user = st.text_input("Username:")
                u_pass = st.text_input("Password:", type="password")
                btn_login = st.form_submit_button("Masuk ke Sistem")
                
                if btn_login:
                    if check_user(u_user, u_pass):
                        st.session_state.logged_in = True
                        st.session_state.username = u_user
                        st.success("Login berhasil!")
                        st.rerun()
                    else:
                        st.error("Username atau password salah!")
                        
        with tab_reg:
            with st.form("form_register"):
                n_user = st.text_input("Buat Username Baru:")
                n_pass = st.text_input("Buat Password:", type="password")
                btn_reg = st.form_submit_button("Daftar Sekarang")
                
                if btn_reg:
                    if n_user and n_pass:
                        if add_user(n_user, n_pass):
                            st.success("Akun berhasil dibuat! Silakan pindah ke tab Login.")
                        else:
                            st.error("Username sudah terpakai, gunakan yang lain.")
                    else:
                        st.warning("Mohon isi semua kolom.")
        st.markdown("</div>", unsafe_allow_html=True)

else:
    current_user = st.session_state.username
    
    c_db = db_conn.cursor()
    c_db.execute("SELECT client_name FROM clients WHERE username = ?", (current_user,))
    user_rooms = [row[0] for row in c_db.fetchall()]
    if not user_rooms:
        user_rooms = ["Contoh Klien: PT Jaya Abadi 2026"]

    if st.session_state.active_room is None:
        c_nav1, c_nav2, c_nav3 = st.columns([3, 5, 2])
        with c_nav1:
            st.markdown(f"### ✈️ **AUDITPILOT** <br><span style='font-size:0.8rem; color:#818cf8;'>User: {current_user}</span>", unsafe_allow_html=True)
        with c_nav2:
            st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.95rem; margin-top: 10px;'>Secure Workspace &nbsp;|&nbsp; Isolated Data Vault</p>", unsafe_allow_html=True)
        with c_nav3:
            if st.button("🚪 Logout"):
                st.session_state.logged_in = False
                st.session_state.username = ""
                st.session_state.active_room = None
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<h1 style='text-align: center; font-size: 3.5rem; background: linear-gradient(to right, #ffffff, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>Manage Your Audit Smarter</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 1.1rem; max-width: 600px; margin: 0 auto;'>Setiap data klien Anda terenkripsi dan terisolasi privat hanya untuk akun Anda sendiri.</p>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        col_home1, col_home2 = st.columns(2)
        
        with col_home1:
            st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
            st.subheader("🏢 Pilih Klien / Penugasan Anda")
            for room in user_rooms:
                jumlah_dok = len(st.session_state.client_vault.get(room, []))
                cr1, cr2 = st.columns([3, 1])
                with cr1:
                    st.markdown(f"**📂 {room}** <br><span style='color: #94a3b8; font-size: 0.85rem;'>{jumlah_dok} Dokumen di Brankas</span>", unsafe_allow_html=True)
                with cr2:
                    if st.button("Buka", key=f"open_room_{room}"):
                        st.session_state.active_room = room
                        st.rerun()
                st.markdown("<hr style='border-color: rgba(255,255,255,0.05);'>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_home2:
            st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
            st.subheader("➕ Buat Klien / Perusahaan Baru")
            with st.form("form_new_client_acc"):
                new_c_name = st.text_input("Nama Perusahaan & Tahun Buku:", placeholder="Misal: PT Berkah Jaya 2026")
                btn_add_c = st.form_submit_button("Inisialisasi Klien")
                if btn_add_c and new_c_name:
                    if new_c_name not in user_rooms:
                        c_db.execute("INSERT INTO clients (username, client_name) VALUES (?, ?)", (current_user, new_c_name))
                        db_conn.commit()
                    st.session_state.active_room = new_c_name
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    else:
        klien_ini = st.session_state.active_room
        
        col_top_w1, col_top_w2 = st.columns([7, 2])
        with col_top_w1:
            st.markdown(f"## ✈️ Workspace: **{klien_ini}** <span style='font-size:0.8rem; color:#818cf8;'>({current_user})</span>", unsafe_allow_html=True)
        with col_top_w2:
            if st.button("⬅️ Kembali ke Menu Klien"):
                st.session_state.active_room = None
                st.rerun()
                
        st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
        
        tab_w1, tab_w2, tab_w3 = st.tabs([
            "📂 1. Document Vault (Input Data)", 
            "📊 2. Parameter & Rules Audit (Filter)", 
            "🤖 3. AI Copilot & Reporting"
        ])

        with tab_w1:
            st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
            st.subheader("Pusat Unggah Dokumen Klien (Private Vault)")
            st.info("Pilih file Excel Anda. Sistem otomatis mendeteksi sheet, header, dan mendukung kategori dokumen kustom.")
            
            f_dok = st.file_uploader("Pilih File Excel (.xlsx / .xls):", type=["xlsx", "xls"])
            
            if f_dok is not None:
                try:
                    sheet_names = get_excel_sheets(f_dok)
                    selected_sheet = st.selectbox("Pilih Sheet / Lembar Kerja Excel:", sheet_names)
                    
                    j_dok_pilihan = st.selectbox("Kategori Dokumen:", ["General Ledger (GL)", "Trial Balance (TB)", "Rekening Koran (Bank)", "Rincian Penyusutan Aset Tetap", "Ekualisasi Pajak (PPN/PPh)", "Stock Opname", "Lainnya"])
                    
                    if j_dok_pilihan == "Lainnya":
                        custom_doc_name = st.text_input("Ketik Nama / Jenis Dokumen Custom:", placeholder="Misal: Buku Besar Piutang Usaha")
                        j_dok = custom_doc_name if custom_doc_name else "Dokumen Lainnya"
                    else:
                        j_dok = j_dok_pilihan
                    
                    if st.button("🚀 Proses & Unggah ke Brankas Data"):
                        df_v, detected_row = clean_and_detect_header(f_dok, selected_sheet)
                        
                        if klien_ini not in st.session_state.client_vault:
                            st.session_state.client_vault[klien_ini] = []
                        st.session_state.client_vault[klien_ini].append({
                            "nama_file": f"{f_dok.name} [{selected_sheet}]",
                            "jenis": j_dok,
                            "data": df_v,
                            "total_baris": len(df_v)
                        })
                        st.success(f"Berhasil mengunggah dokumen '{j_dok}'! Header terdeteksi pada baris ke-{detected_row + 1}.")
                except Exception as e:
                    st.error(f"Gagal memproses file: {e}")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("### Daftar Arsip Dokumen Masuk:")
            v_data = st.session_state.client_vault.get(klien_ini, [])
            if v_data:
                for item in v_data:
                    with st.expander(f"📄 [{item['jenis']}] — {item['nama_file']} ({item['total_baris']} Baris)"):
                        st.dataframe(item['data'].head(15), use_container_width=True)
            else:
                st.warning("Brankas dokumen masih kosong.")

        with tab_w2:
            st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
            st.subheader("⚡ Parameter & Rules Audit")
            v_data = st.session_state.client_vault.get(klien_ini, [])
            if v_data:
                f_names = [d['nama_file'] for d in v_data]
                p_file = st.selectbox("Pilih file arsip untuk dianalisis:", f_names, key="filter_file_select")
                s_doc = next(d for d in v_data if d['nama_file'] == p_file)
                
                df_original = s_doc['data']
                df_to_show = df_original.copy()
                col_cols = df_to_show.columns.tolist()
                
                global_search = st.text_input("🔍 Cari transaksi...", placeholder="Ketik kata kunci pencarian bebas...")
                if global_search:
                    mask = df_to_show.apply(lambda row: row.astype(str).str.contains(global_search, case=False).any(), axis=1)
                    df_to_show = df_to_show[mask]

                st.markdown("---")
                st.markdown("#### ⚡ Parameter & Rules Audit")
                
                col_mat_col = st.selectbox("Pilih kolom nominal/saldo:", col_cols)
                batas_materialitas = st.number_input("Batas Materialitas (Rp)", value=100000000, step=1000000)
                aktifkan_materialitas = st.checkbox("🚨 Aktifkan Filter Materialitas")
                
                if aktifkan_materialitas:
                    try:
                        numeric_s = pd.to_numeric(df_to_show[col_mat_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce').fillna(0)
                        df_to_show = df_to_show[abs(numeric_s) >= batas_materialitas]
                    except Exception:
                        pass

                col_ket_col = st.selectbox("Pilih kolom Keterangan/Deskripsi:", col_cols, key="sel_ket_col")
                aktifkan_ket_kosong = st.checkbox("🎯 Aktifkan Filter Ket. Kosong")
                
                if aktifkan_ket_kosong:
                    df_to_show = df_to_show[df_to_show[col_ket_col].str.strip() == ""]

                st.markdown("---")
                st.markdown("#### 💀 Deteksi Anomali Otomatis")
                
                deteksi_weekend = st.checkbox("📅 Deteksi Transaksi Weekend (Sabtu / Minggu)")
                deteksi_bulat = st.checkbox("🎯 Deteksi Angka Bulat Mencurigakan (Kelipatan 1 Juta)")
                
                col_duplikasi = st.selectbox("Pilih kolom acuan Duplikasi (Contoh: No Jurnal, Akun, atau Ref):", col_cols)
                deteksi_duplikat = st.checkbox("🔄 Deteksi Nilai Kembar Berdasarkan Kolom Terpilih")
                
                if deteksi_bulat:
                    try:
                        nums = pd.to_numeric(df_to_show[col_mat_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce').fillna(0)
                        df_to_show = df_to_show[(nums != 0) & (nums % 1000000 == 0)]
                    except Exception:
                        pass
                        
                if deteksi_duplikat and col_duplikasi:
                    df_to_show = df_to_show[df_to_show.duplicated(subset=[col_duplikasi], keep=False)]

                st.markdown("---")
                
                c_stat1, c_stat2 = st.columns(2)
                with c_stat1:
                    st.markdown(f"<p style='color: #94a3b8; font-size: 0.9rem;'>Total Data Asli</p><h2 style='color: #ffffff;'>{len(df_original)} Baris</h2>", unsafe_allow_html=True)
                with c_stat2:
                    st.markdown(f"<p style='color: #94a3b8; font-size: 0.9rem;'>Temuan Saringan</p><h2 style='color: #818cf8;'>{len(df_to_show)} Baris</h2>", unsafe_allow_html=True)
                
                st.markdown("#### 📂 Preview Temuan")
                st.dataframe(df_to_show, use_container_width=True, height=350)
            else:
                st.info("Silakan lakukan upload dokumen di Tab 1 terlebih dahulu.")
            st.markdown("</div>", unsafe_allow_html=True)

        with tab_w3:
            st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
            st.subheader("AI Copilot & Pembuatan KKA / Report")
            v_data = st.session_state.client_vault.get(klien_ini, [])
            
            if klien_ini not in st.session_state.chat_histories:
                st.session_state.chat_histories[klien_ini] = [{"role": "ai", "content": f"Halo {current_user}! Saya AI Copilot untuk penugasan {klien_ini}. Silakan tanyakan analisis, pengujian penyusunan ekualisasi, atau minta draf KKA."}]
                
            chat_box = st.container(height=350)
            with chat_box:
                for chat in st.session_state.chat_histories[klien_ini]:
                    with st.chat_message(chat["role"]):
                        st.write(chat["content"])

            u_input = st.chat_input("Ketik pertanyaan audit / minta analisis ekualisasi / buatkan KKA...")
            if u_input:
                st.session_state.chat_histories[klien_ini].append({"role": "user", "content": u_input})
                with chat_box:
                    with st.chat_message("user"):
                        st.write(u_input)
                
                with chat_box:
                    with st.chat_message("ai"):
                        with st.spinner("AI sedang membaca brankas data..."):
                            try:
                                ctx = ""
                                for d in v_data:
                                    ctx += f"\n- {d['jenis']} ({d['nama_file']}): {d['data'].head(15).to_dict(orient='records')}\n"
                                if not ctx: ctx = "Belum ada dokumen."
                                
                                prompt_ai = f"Anda adalah Senior AI Audit untuk {klien_ini}. Data: {ctx}. Pertanyaan/Instruksi Auditor: {u_input}"
                                resp = ai_model.generate_content(prompt_ai)
                                ans = resp.text
                            except Exception as e:
                                ans = f"Error: {e}"
                            st.write(ans)
                st.session_state.chat_histories[klien_ini].append({"role": "ai", "content": ans})
                st.rerun()

            st.markdown("---")
            if v_data:
                if st.button("✨ Buat Draf KKA & Laporan Resmi"):
                    with st.spinner("Menyusun laporan formal..."):
                        try:
                            c_rep = ""
                            for d in v_data:
                                c_rep += f"{d['jenis']} ({d['nama_file']}): {d['data'].head(10).to_dict(orient='records')}\n"
                            p_rep = f"Buat draf Laporan Audit Eksekutif & Kertas Kerja Audit (KKA) formal untuk {klien_ini} berdasarkan:\n{c_rep}"
                            r_rep = ai_model.generate_content(p_rep)
                            st.session_state.ai_reports[klien_ini] = r_rep.text
                            st.success("Draf laporan berhasil disusun!")
                        except Exception as e:
                            st.error(f"Gagal: {e}")

                if klien_ini in st.session_state.ai_reports:
                    st.text_area("Draf Laporan Resmi / KKA (Siap Salin):", value=st.session_state.ai_reports[klien_ini], height=250)
                    
                    report_text = st.session_state.ai_reports[klien_ini]
                    b_bytes = report_text.encode('utf-8')
                    
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label="📥 Unduh KKA (Format Teks/Word .txt)",
                            data=b_bytes,
                            file_name=f"KKA_{klien_ini.replace(' ', '_')}.txt",
                            mime="text/plain"
                        )
                    with col_dl2:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            lines = report_text.split('\n')
                            df_report_export = pd.DataFrame({"Kertas Kerja Audit (KKA) & Laporan Eksekutif": lines})
                            df_report_export.to_excel(writer, index=False, sheet_name='KKA_Summary')
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Unduh KKA (Format Excel .xlsx)",
                            data=output,
                            file_name=f"KKA_{klien_ini.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            st.markdown("</div>", unsafe_allow_html=True)
