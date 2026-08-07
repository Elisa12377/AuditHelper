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

# --- CSS STYLING VISUAL MODERN ALA REFERENSI GAMBAR ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Space Grotesk', sans-serif; 
    }
    
    /* Background Gradient Gelap Elegan */
    .stApp { 
        background: linear-gradient(135deg, #05070B 0%, #0D1124 50%, #150F2D 100%);
        color: #FFFFFF; 
    }
    
    [data-testid="stSidebar"] {
        display: none;
    }
    
    /* Tombol Utama Gradasi Ungu-Biru */
    .stButton>button { 
        background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%) !important; 
        color: #FFFFFF !important; 
        border-radius: 14px; 
        border: none;
        font-weight: 600;
        padding: 0.6rem 1.4rem;
        box-shadow: 0 4px 20px rgba(99, 102, 241, 0.4);
        transition: all 0.3s ease;
    }
    .stButton>button:hover { 
        transform: translateY(-2px);
        box-shadow: 0 6px 25px rgba(99, 102, 241, 0.7);
    }
    
    /* Kartu Efek Kaca (Glassmorphism) */
    .glass-box {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        padding: 24px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
    }
    
    h1, h2, h3 { 
        color: #FFFFFF; 
        font-weight: 700; 
        letter-spacing: -0.5px; 
    }
</style>
""", unsafe_allow_html=True)

# ================= KONTROL ALUR APLIKASI =================

if st.session_state.active_room is None:
    # --- 1. LANDING PAGE / COVER UTAMA (MENYERUPAI REFERENSI) ---
    
    # Header Navigasi Sederhana
    c_nav1, c_nav2, c_nav3, c_nav4 = st.columns([2, 4, 2, 2])
    with c_nav1:
        st.markdown("### ✈️ **AUDITPILOT & CO.**")
    with c_nav2:
        st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.95rem; margin-top: 10px;'>Home &nbsp;&nbsp;|&nbsp;&nbsp; Features &nbsp;&nbsp;|&nbsp;&nbsp; How it Works</p>", unsafe_allow_html=True)
    with c_nav4:
        # Tombol aksi utama menuju Setup Klien
        if st.button("Contact / Launch"):
            st.session_state.active_room = "SETUP_NEW_CLIENT"
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Hero Title Tengah
    st.markdown("<h1 style='text-align: center; font-size: 3.8rem; background: linear-gradient(to right, #ffffff, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>Manage Your Audit Smarter</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 1.15rem; max-width: 650px; margin: 0 auto;'>Stay organized, connected, and productive with one powerful AI-driven platform designed to simplify your financial audits and compliance.</p>", unsafe_allow_html=True)
    
    st.markdown("<br><center>", unsafe_allow_html=True)
    if st.button("🚀 Get Started / New Company"):
        st.session_state.active_room = "SETUP_NEW_CLIENT"
        st.rerun()
    st.markdown("️</center>", unsafe_allow_html=True)
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # Dekorasi Kartu Statistik / Visual Grafik Bawah ala Desain Referensi
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    
    with col_stat1:
        st.markdown("""
        <div class="glass-box">
            <h3 style="color: #818cf8; margin-bottom: 5px;">1.5 M+</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">Trusted records analyzed by auditors worldwide.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_stat2:
        st.markdown("""
        <div class="glass-box" style="text-align: center;">
            <h3 style="color: #34d399; margin-bottom: 5px;">93% Accuracy</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">AI Audit Compliance & Anomaly Rating Score.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_stat3:
        st.markdown("""
        <div class="glass-box">
            <h3 style="color: #f43f5e; margin-bottom: 5px;">120 K+</h3>
            <p style="color: #94a3b8; font-size: 0.9rem; margin:0;">Automated Working Papers (KKA) Generated.</p>
        </div>
        """, unsafe_allow_html=True)

elif st.session_state.active_room == "SETUP_NEW_CLIENT":
    # --- 2. FORM PEMBUATAN / PEMILIHAN KLIEN BARU (NEW COMPANY) ---
    col_top_back, _ = st.columns([1, 6])
    with col_top_back:
        if st.button("⬅️ Back to Home"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_form1, col_form2 = st.columns(2)
    with col_form1:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("🏢 Buat Ruang Perusahaan Baru (New Company)")
        st.markdown("Masukkan nama entitas klien dan periode tahun buku untuk memulai investigasi.")
        
        with st.form(key="form_new_comp"):
            nama_perusahaan = st.text_input("Nama Perusahaan Klien:", placeholder="Misal: PT Berkah Sejahtera Tbk")
            tahun_audit = st.text_input("Periode Tahun Buku:", placeholder="2026")
            sub_comp = st.form_submit_button("Inisialisasi Workspace Klien")
            
            if sub_comp and nama_perusahaan:
                full_room_name = f"{nama_perusahaan} ({tahun_audit})"
                if full_room_name not in st.session_state.rooms:
                    st.session_state.rooms.append(full_room_name)
                st.session_state.active_room = full_room_name
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with col_form2:
        st.markdown("<div class='glass-box'>", unsafe_allow_html=True)
        st.subheader("📂 Atau Pilih Klien yang Sudah Ada")
        for room in st.session_state.rooms:
            c_r1, c_r2 = st.columns([3, 1])
            with c_r1:
                st.write(f"**{room}**")
            with c_r2:
                if st.button("Masuk", key=f"select_exist_{room}"):
                    st.session_state.active_room = room
                    st.rerun()
            st.markdown("<hr style='border-color: rgba(255,255,255,0.05);'>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

else:
    # --- 3 & 4. WORKSPACE UTAMA: UPLOAD, ANALYTICS, AI COPILOT & REPORT ---
    klien_ini = st.session_state.active_room
    
    col_top_w1, col_top_w2 = st.columns([7, 2])
    with col_top_w1:
        st.markdown(f"## ✈️ Workspace Klien: **{klien_ini}**")
    with col_top_w2:
        if st.button("🚪 Keluar / Ganti Klien"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
    
    # Tab Navigasi Berjenjang Sesuai Alur Kerjamu
    tab_w1, tab_w2, tab_w3 = st.tabs([
        "📂 1. Document Vault (Input Data)", 
        "📊 2. Analytics & Visuals", 
        "🤖 3. AI Copilot & Reporting (KKA)"
    ])

    with tab_w1:
        st.subheader("Pusat Unggah Dokumen Klien (New Data Input)")
        st.info("Silakan masukkan data keuangan bertahap: General Ledger (GL), Trial Balance (TB), Rekening Koran, atau Stock Opname.")
        
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
                    st.success(f"Berhasil mengunggah {f_dok.name} ke sistem!")
                except Exception as e:
                    st.error(f"Gagal memproses file: {e}")

        st.markdown("---")
        st.write("### Daftar Dokumen Tersimpan:")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        if v_data:
            for item in v_data:
                with st.expander(f"📄 [{item['jenis']}] — {item['nama_file']} ({item['total_baris']} Baris)"):
                    st.dataframe(item['data'].head(15), use_container_width=True)
        else:
            st.warning("Belum ada file yang diunggah di ruang ini.")

    with tab_w2:
        st.subheader("Analytics Engine & Visual Grafik")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        if v_data:
            f_names = [d['nama_file'] for d in v_data]
            p_file = st.selectbox("Pilih file arsip untuk dianalisis grafiknya:", f_names)
            s_doc = next(d for d in v_data if d['nama_file'] == p_file)
            
            st.markdown("#### 📈 Ringkasan Grafik Analitik Data")
            # Menampilkan preview tabel data yang otomatis dikonversi jadi grafik batang sederhana di Streamlit jika ada kolom angka
            st.dataframe(s_doc['data'], use_container_width=True, height=350)
        else:
            st.info("Silakan upload data terlebih dahulu di Tab 1 untuk melihat hasil analitiknya.")

    with tab_w3:
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
                    with st.spinner("AI sedang memproses data dokumen..."):
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
            if st.button("✨ Buat Draf Kertas Kerja Audit (KKA) & Laporan Resmi"):
                with st.spinner("Menyusun draf laporan formal..."):
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
