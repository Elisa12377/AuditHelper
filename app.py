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

# --- CSS STYLING MODERN (DARK BLUE / PURPLE GRADIENT) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Space Grotesk', sans-serif; 
    }
    
    .stApp { 
        background: linear-gradient(135deg, #090B10 0%, #111425 50%, #1A1333 100%);
        color: #FFFFFF; 
    }
    
    [data-testid="stSidebar"] {
        display: none;
    }
    
    .stButton>button { 
        background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%) !important; 
        color: #FFFFFF !important; 
        border-radius: 12px; 
        border: none;
        font-weight: 600;
        padding: 0.6rem 1.2rem;
        box-shadow: 0 4px 14px rgba(99, 102, 241, 0.4);
        transition: all 0.3s ease;
    }
    .stButton>button:hover { 
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(99, 102, 241, 0.6);
    }
    
    h1, h2, h3 { 
        color: #FFFFFF; 
        font-weight: 700; 
        letter-spacing: -0.5px; 
    }
</style>
""", unsafe_allow_html=True)

# ================= KONTROL HALAMAN =================

if st.session_state.active_room is None:
    # --- LANDING PAGE / DASHBOARD UTAMA ---
    col_nav1, col_nav2, col_nav3 = st.columns([3, 6, 2])
    with col_nav1:
        st.markdown("### ✈️ **AUDITPILOT & CO.**")
    with col_nav3:
        if st.button("Mulai Audit Sekarang"):
            st.session_state.active_room = st.session_state.rooms[0] if st.session_state.rooms else "Klien Utama"
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    st.markdown("<h1 style='text-align: center; font-size: 3.5rem; background: linear-gradient(to right, #ffffff, #a5b4fc); -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>Manage Your Audit Smarter</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 1.2rem; max-width: 600px; margin: 0 auto;'>Stay organized, connected, and productive with one powerful AI-driven platform designed to simplify your financial audits.</p>", unsafe_allow_html=True)
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col_home1, col_home2 = st.columns([1.2, 1])
    
    with col_home1:
        st.markdown("### 🏢 Pilih Penugasan Klien Aktif")
        st.markdown("Pilih salah satu ruang kerja atau buat penugasan baru di bawah ini:")
        
        for room in st.session_state.rooms:
            jumlah_dok = len(st.session_state.client_vault.get(room, []))
            c_r1, c_r2 = st.columns([3, 1])
            with c_r1:
                st.markdown(f"**📂 {room}** <br><span style='color: #94a3b8; font-size: 0.9rem;'>{jumlah_dok} Dokumen tersimpan di brankas</span>", unsafe_allow_html=True)
            with c_r2:
                if st.button("Buka", key=f"home_open_{room}"):
                    st.session_state.active_room = room
                    st.rerun()
            st.markdown("<hr style='border-color: rgba(255,255,255,0.05);'>", unsafe_allow_html=True)

    with col_home2:
        st.markdown("### ➕ Buat Klien Baru")
        with st.form(key="form_home_client"):
            new_c = st.text_input("Nama Perusahaan / Klien & Tahun:", placeholder="Misal: PT Teknologi Maju 2026")
            sub_c = st.form_submit_button("Inisialisasi Workspace")
            if sub_c and new_c:
                if new_c not in st.session_state.rooms:
                    st.session_state.rooms.append(new_c)
                    st.session_state.active_room = new_c
                    st.rerun()

else:
    # --- WORKSPACE KLIEN ---
    klien_ini = st.session_state.active_room
    
    col_top1, col_top2 = st.columns([7, 2])
    with col_top1:
        st.markdown(f"## ✈️ Workspace: **{klien_ini}**")
    with col_top2:
        if st.button("⬅️ Kembali ke Home"):
            st.session_state.active_room = None
            st.rerun()
            
    st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
    
    tab_w1, tab_w2, tab_w3 = st.tabs([
        "📂 1. Document Vault (Brankas)", 
        "📊 2. Analytics Engine", 
        "🤖 3. AI Copilot & Report"
    ])

    with tab_w1:
        st.subheader("Pusat Unggah Dokumen Klien")
        st.info("Unggah dokumen secara bertahap (GL, TB, Rekening Koran, Stock Opname, dll.).")
        
        with st.form(key=f"ws_upload_{klien_ini}", clear_on_submit=True):
            uc1, uc2 = st.columns(2)
            with uc1:
                j_dok = st.selectbox("Jenis Dokumen:", ["General Ledger (GL)", "Trial Balance (TB)", "Rekening Koran (Bank)", "Stock Opname", "Cash Opname", "Lainnya"])
            with uc2:
                f_dok = st.file_uploader("Pilih File (Excel):", type=["xlsx", "xls"])
                
            btn_up = st.form_submit_button("Unggah ke Brankas")
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

        st.markdown("---")
        st.write("### Arsip Dokumen Masuk:")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        if v_data:
            for item in v_data:
                with st.expander(f"📄 [{item['jenis']}] — {item['nama_file']} ({item['total_baris']} Baris)"):
                    st.dataframe(item['data'].head(15), use_container_width=True)
        else:
            st.warning("Brankas dokumen masih kosong.")

    with tab_w2:
        st.subheader("Analytics & Review Data")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        if v_data:
            f_names = [d['nama_file'] for d in v_data]
            p_file = st.selectbox("Pilih file untuk diperiksa:", f_names)
            s_doc = next(d for d in v_data if d['nama_file'] == p_file)
            st.dataframe(s_doc['data'], use_container_width=True, height=400)
        else:
            st.info("Belum ada data untuk dianalisis.")

    with tab_w3:
        st.subheader("AI Audit Copilot & Reporting")
        v_data = st.session_state.client_vault.get(klien_ini, [])
        
        if klien_ini not in st.session_state.chat_histories:
            st.session_state.chat_histories[klien_ini] = [{"role": "ai", "content": f"Halo! Saya AI Copilot untuk penugasan {klien_ini}. Silakan tanyakan apa saja mengenai dokumen yang sudah diunggah."}]
            
        chat_box = st.container(height=350)
        with chat_box:
            for chat in st.session_state.chat_histories[klien_ini]:
                with st.chat_message(chat["role"]):
                    st.write(chat["content"])

        u_input = st.chat_input("Ketik pertanyaan analisis...")
        if u_input:
            st.session_state.chat_histories[klien_ini].append({"role": "user", "content": u_input})
            with chat_box:
                with st.chat_message("user"):
                    st.write(u_input)
            
            with chat_box:
                with st.chat_message("ai"):
                    with st.spinner("AI sedang membaca brankas..."):
                        try:
                            ctx = ""
                            for d in v_data:
                                ctx += f"\n- {d['jenis']} ({d['nama_file']}): {d['data'].head(15).to_dict(orient='records')}\n"
                            if not ctx: ctx = "Belum ada data."
                            
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
            if st.button("✨ Buat Laporan Ringkasan Eksekutif Otomatis"):
                with st.spinner("Menyusun laporan audit..."):
                    try:
                        c_rep = ""
                        for d in v_data:
                            c_rep += f"{d['jenis']} ({d['nama_file']}): {d['data'].head(10).to_dict(orient='records')}\n"
                        p_rep = f"Buat draf Laporan Audit Eksekutif formal untuk {klien_ini} berdasarkan:\n{c_rep}"
                        r_rep = ai_model.generate_content(p_rep)
                        st.session_state.ai_reports[klien_ini] = r_rep.text
                        st.success("Laporan selesai disusun!")
                    except Exception as e:
                        st.error(f"Gagal: {e}")

            if klien_ini in st.session_state.ai_reports:
                st.text_area("Draf Laporan Resmi:", value=st.session_state.ai_reports[klien_ini], height=250)
