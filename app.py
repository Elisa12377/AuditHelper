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

st.set_page_config(page_title="AuditPilot — AI-Powered Audit Analytics", layout="wide", initial_sidebar_state="collapsed")

# --- INITIALIZATION SESSION STATE ---
if "rooms" not in st.session_state:
    st.session_state.rooms = ["PT Jaya Abadi 2026", "CV Makmur Sejahtera"]
if "active_room" not in st.session_state:
    st.session_state.active_room = None
if "client_vault" not in st.session_state:
    st.session_state.client_vault = {}
if "chat_histories" not in st.session_state:
    st.session_state.chat_histories = {}
if "ai_reports" not in st.session_state:
    st.session_state.ai_reports = {}

# --- CSS STYLING: FIX TOTAL WARNA TEKS & KOTAK UPLOAD ---
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
    
    /* Paksa semua teks umum agar terang & jelas */
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
    
    /* TOMBOL: Background putih, teks HITAM tebal, fix tidak transparan */
    .stButton>button { 
        background: #FFFFFF !important; 
        color: #000000 !important; 
        border-radius: 12px !important; 
        border: none !important;
        font-weight: 700 !important;
        padding: 0.6rem 1.4rem !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    .stButton>button p, .stButton>button span, .stButton>button div {
        color: #000000 !important;
        font-weight: 700 !important;
    }
    .stButton>button:hover { 
        background: #E2E8F0 !important;
        color: #000000 !important;
    }
    
    /* KARTU GLASSMORPHISM */
    .glass-box {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        padding: 28px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
        margin-bottom: 20px;
    }
    
    /* INPUT & SELECTBOX */
    .stTextInput>div>div>input, .stSelectbox>div>div>select {
        background-color: rgba(255, 255, 255, 0.08) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 12px !important;
    }
    
    /* FILE UPLOADER CONTAINER: Beri background gelap agar tombol/teks upload di dalamnya kelihatan kontras */
    [data-testid="stFileUploader"] {
        background-color: rgba(255, 255, 255, 0.06);
        padding: 15px;
        border-radius: 14px;
        border: 1px solid rgba(255, 255, 255, 0.15);
    }
    [data-testid="stFileUploader"] section {
        background-color: rgba(255, 255, 255, 0.05) !important;
        border: 1px dashed rgba(255, 255, 255, 0.3) !important;
    }
    [data-testid="stFileUploader"] span, [data-testid="stFileUploader"] small, [data-testid="stFileUploader"] div, [data-testid="stFileUploader"] p, [data-testid="stFileUploader"] button {
        color: #FFFFFF !important;
    }
    
    h1, h2, h3 { 
        color: #FFFFFF !important; 
        font-weight: 700; 
        letter-spacing: -0.5px; 
    }
</style>
""", unsafe_allow_html=True)

# ================= KONTROL ALUR APLIKASI =================

if st.session_state.active_room is None:
    # --- 1. LANDING PAGE / COVER UTAMA ---
    c_nav1, c_nav2, c_nav3 = st.columns([3, 6, 2])
    with c_nav1:
        st.markdown("### ✈️ **AUDITPILOT & CO.**")
    with c_nav2:
        st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.95rem; margin-top: 10px;'>Home &nbsp;&nbsp;|&nbsp;&nbsp; Features &nbsp;&nbsp;|&nbsp;&nbsp; Analytics Engine</p>", unsafe_allow_html=True)
    with c_nav3:
        if st.button("Launch App"):
            st.session_state.active_room = "SETUP_NEW_CLIENT"
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    st.markdown("<h1 style='text-align: center; font-size: 3.8rem; background: linear-gradient(to right, #ffffff, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>Manage Your Audit Smarter</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 1.15rem; max-width: 650px; margin: 0 auto;'>Stay organized, connected, and productive with one powerful AI-driven platform designed to simplify your financial audits and compliance.</p>", unsafe_allow_html=True)
    
    st.markdown("<br><center>", unsafe_allow_html=True)
    if st.button("🚀 Get Started / New Company"):
        st.session_state.active_room = "SETUP_NEW_CLIENT"
        st.rerun()
    st.markdown("</center>", unsafe_allow_html=True)
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    
    with col_stat1:
        st.markdown("""
        <div class="glass-box">
            <h3 style="color: #818cf8; margin-bottom: 5px;">1.5 M+</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">Trusted financial records analyzed globally.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_stat2:
        st.markdown("""
        <div class="glass-box" style="text-align: center;">
            <h3 style="color: #34d399; margin-bottom: 5px;">93% Accuracy</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">AI Compliance & Anomaly Risk Score.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_stat3:
        st.markdown("""
        <div class="glass-box">
            <h3 style="color: #f43f5e; margin-bottom: 5px;">120 K+</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">Automated Working Papers Generated.</p>
        </div>
        """, unsafe_allow_html=True)

elif st.session_state.active_room == "SETUP_NEW_CLIENT":
    # --- 2. SETUP / NEW COMPANY FORM ---
    col_top_back, _ = st.columns([1, 6])
    with col_top_back:
        if st.button("⬅️ Back to Home"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_form1, col_form2 = st.columns(2)
    with col_form1:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("🏢 Buat Perusahaan Baru (New Company)")
        st.markdown("Daftarkan entitas klien baru untuk membuka ruang kerja investigasi.")
        
        with st.form(key="form_new_comp"):
            nama_perusahaan = st.text_input("Nama Perusahaan Klien:", placeholder="Misal: PT Nusantara Jaya Tbk")
            tahun_audit = st.text_input("Periode Tahun Buku:", placeholder="2026")
            sub_comp = st.form_submit_button("Inisialisasi Workspace")
            
            if sub_comp and nama_perusahaan:
                full_room_name = f"{nama_perusahaan} ({tahun_audit})"
                if full_room_name not in st.session_state.rooms:
                    st.session_state.rooms.append(full_room_name)
                st.session_state.active_room = full_room_name
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with col_form2:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("📂 Pilih Klien Terdaftar")
        for room in st.session_state.rooms:
            c_r1, c_r2 = st.columns([3, 1])
            with c_r1:
                st.write(f"**{room}**")
            with c_r2:
                if st.button("Buka", key=f"select_exist_{room}"):
                    st.session_state.active_room = room
                    st.rerun()
            st.markdown("<hr style='border-color: rgba(255,255,255,0.05);'>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

else:
    # --- 3. WORKSPACE UTAMA (VAULT, ANALYTICS, AI COPILOT) ---
    klien_ini = st.session_state.active_room
    
    col_top_w1, col_top_w2 = st.columns([7, 2])
    with col_top_w1:
        st.markdown(f"## ✈️ Workspace: **{klien_ini}**")
    with col_top_w2:
        if st.button("🚪 Keluar ke Home"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
    
    tab_w1, tab_w2, tab_w3 = st.tabs([
        "📂 1. Document Vault (Input Data)", 
        "📊 2. Analytics & Visuals", 
        "🤖 3. AI Copilot & Reporting"
    ])

    with tab_w1:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("Pusat Unggah Dokumen Klien (Rolling Basis)")
        st.info("Unggah dokumen secara bertahap (GL, TB, Rekening Koran, Stock Opname, dll.). Data akan terakumulasi otomatis.")
        
        with st.form(key=f"ws_upload_{klien_ini}", clear_on_submit=True):
            uc1, uc2 = st.columns(2)
            with uc1:
                j_dok = st.selectbox("Kategori Dokumen:", ["General Ledger (GL)", "Trial Balance (TB)", "Rekening Koran (Bank)", "Stock Opname", "Cash Opname", "Lainnya"])
            with uc2:
                f_dok = st.file_uploader("Pilih File Excel:", type=["xlsx", "xls"])
                
            btn_up = st.form_submit_button("Unggah ke Brankas Data")
            if btn_up and f_dok is not None:
                try:
                    df_v = pd.read_excel(f_dok).astype(str)
                    if klien_ini not in st.session_state.client_vault:
                        st.session_state.client_vault[klien_ini] = []
                    st.session_state.client_vault[klien_ini].append({
                        "nama_file": f_dok.name,
                        "jenis": j_dok,
                        "data": df_v,
                        "total_baris": len(df_v)
                    })
                    st.success(f"Berhasil mengunggah {f_dok.name}!")
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
        st.subheader("Analytics Engine & Data Visuals")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        if v_data:
            f_names = [d['nama_file'] for d in v_data]
            p_file = st.selectbox("Pilih file arsip untuk dianalisis:", f_names)
            s_doc = next(d for d in v_data if d['nama_file'] == p_file)
            
            st.markdown("#### 📈 Preview & Grafik Analitik")
            st.dataframe(s_doc['data'], use_container_width=True, height=350)
        else:
            st.info("Silakan lakukan upload dokumen di Tab 1 terlebih dahulu.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_w3:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("AI Copilot & Pembuatan KKA / Report")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        
        if klien_ini not in st.session_state.chat_histories:
            st.session_state.chat_histories[klien_ini] = [{"role": "ai", "content": f"Halo! Saya AI Copilot untuk penugasan {klien_ini}. Silakan tanyakan analisis atau minta penyusunan draf KKA."}]
            
        chat_box = st.container(height=350)
        with chat_box:
            for chat in st.session_state.chat_histories[klien_ini]:
                with st.chat_message(chat["role"]):
                    st.write(chat["content"])

        u_input = st.chat_input("Ketik pertanyaan audit / minta buatkan KKA...")
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
                            
                            prompt_ai = f"Anda adalah Senior AI Audit untuk {klien_ini}. Data: {ctx}. Pertanyaan: {u_input}"
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
        st.markdown("</div>", unsafe_allow_html=True) 
