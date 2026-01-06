import streamlit as st
import pandas as pd
from datetime import date, timedelta
import datetime
import altair as alt 
from fpdf import FPDF
import io
import time

# ======================================
# 1. GLOBAL CONFIG & SECURITY
# ======================================
st.set_page_config(page_title="Audit Dashboard Pro", layout="wide")

USERNAME_TARGET = "useraudit"
PASSWORD_TARGET = "user123"

st.markdown("""
    <style>
    .block-container { padding-left: 1rem; padding-right: 1rem; padding-top: 2rem; }
    .stFileUploader { border: 2px dashed #1F497D; padding: 15px; border-radius: 10px; background-color: #f7fbff; }
    [data-testid="stMetricValue"] { font-size: 28px; color: #1F497D; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ======================================
# 2. HELPER FUNCTIONS
# ======================================

def normalize(col):
    return col.astype(str).str.upper().str.strip().str.replace(r"\s+", " ", regex=True)

def to_time(col):
    return pd.to_datetime(col, format='%H:%M:%S', errors="coerce").dt.time

def highlight_violation(s):
    return ['background-color: #ffcccc' if s.name == 'CreateTim' else '' for _ in s] 

def highlight_top_rank(s, top_n=3):
    df_style = pd.DataFrame('', index=s.index, columns=s.columns)
    if 'Total Pelanggaran' in s.columns:
        df_style.loc[s.index[:top_n], 'Total Pelanggaran'] = 'background-color: #ffd70030'
    return df_style

def calculate_minutes_diff(row):
    """Menghitung selisih menit antara waktu aktivitas dengan batas absen."""
    trx_dt = datetime.datetime.combine(row['TRXDATE'], row['CreateTim'])
    start_dt = row['Start_S']
    end_dt = row['End_S']
    
    if trx_dt < start_dt:
        diff = start_dt - trx_dt
        return f"-{int(diff.total_seconds() / 60)} Min"
    elif trx_dt > end_dt:
        diff = trx_dt - end_dt
        return f"+{int(diff.total_seconds() / 60)} Min"
    return "0 Min"

# FUNGSI GENERATE PDF (VERSI FORMAL UI/UX PREMIUM UNTUK ATASAN)
def create_pdf(df_detail, s_date, e_date):
    # Setup Dokumen Landscape A4
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # --- KONFIGURASI WARNA & STYLE ---
    PRIMARY_COLOR = (31, 73, 125)   # Midnight Blue
    SECONDARY_COLOR = (242, 242, 242) # Zebra Gray
    ACCENT_COLOR = (44, 62, 80)     # Dark Gray
    TEXT_COLOR = (40, 40, 40)
    WHITE = (255, 255, 255)
    
    # --- 1. KOP SURAT / HEADER FORMAL ---
    pdf.set_fill_color(*PRIMARY_COLOR)
    pdf.rect(0, 0, 297, 45, 'F') 
    
    pdf.set_text_color(*WHITE)
    pdf.set_font("Arial", 'B', 24)
    pdf.cell(0, 15, "OFFICIAL AUDIT REPORT", 0, 1, 'C')
    
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 5, "LAPORAN ANALISIS AKTIVITAS DI LUAR JAM KERJA", 0, 1, 'C')
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 10, f"Periode Pemeriksaan: {s_date} s/d {e_date}", 0, 1, 'C')
    
    pdf.ln(22) # Jeda dari header
    
    # --- 2. EXECUTIVE SUMMARY BOXES ---
    pdf.set_text_color(*PRIMARY_COLOR)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "I. RINGKASAN EKSEKUTIF", 0, 1, 'L')
    pdf.line(10, pdf.get_y(), 60, pdf.get_y())
    pdf.ln(5)
    
    total_kasus = len(df_detail)
    total_staff = df_detail['AuthName'].nunique()
    
    # Render 2 Kotak Statistik
    pdf.set_fill_color(*SECONDARY_COLOR)
    pdf.set_draw_color(*PRIMARY_COLOR)
    pdf.set_text_color(*TEXT_COLOR)
    
    # Kotak 1
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(80, 20, f"  Total Pelanggaran Aktivitas : {total_kasus} Kasus", 1, 0, 'L', True)
    pdf.cell(10, 20, "", 0, 0) # Spacer
    # Kotak 2
    pdf.cell(80, 20, f"  Total Personel Terlibat    : {total_staff} Orang", 1, 1, 'L', True)
    
    pdf.ln(10)

    # --- 3. TOP OFFENDERS TABLE ---
    pdf.set_font("Arial", 'B', 14)
    pdf.set_text_color(*PRIMARY_COLOR)
    pdf.cell(0, 10, "II. DAFTAR PERINGKAT PELANGGAR TERTINGGI", 0, 1, 'L')
    pdf.ln(2)
    
    df_rank = df_detail.groupby("AuthName").size().reset_index(name='Total').sort_values('Total', ascending=False).head(5)
    
    # Header Tabel Ranking
    pdf.set_fill_color(*ACCENT_COLOR)
    pdf.set_text_color(*WHITE)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(20, 10, "Rank", 1, 0, 'C', True)
    pdf.cell(150, 10, "Nama Karyawan", 1, 0, 'C', True)
    pdf.cell(40, 10, "Frekuensi Pelanggaran", 1, 1, 'C', True)
    
    # Body Tabel Ranking
    pdf.set_text_color(*TEXT_COLOR)
    pdf.set_font("Arial", '', 10)
    for i, (idx, row) in enumerate(df_rank.iterrows(), 1):
        pdf.cell(20, 9, str(i), 1, 0, 'C')
        pdf.cell(150, 9, f"  {row['AuthName']}", 1, 0, 'L')
        pdf.cell(40, 9, f"{row['Total']} Kali", 1, 1, 'C')

    # --- 4. DETAIL ACTIVITY TABLE ---
    pdf.add_page() # Pindah halaman khusus untuk detail
    pdf.set_font("Arial", 'B', 14)
    pdf.set_text_color(*PRIMARY_COLOR)
    pdf.cell(0, 10, "III. DATA DETAIL AKTIVITAS PEMERIKSAAN", 0, 1, 'L')
    pdf.ln(2)
    
    # Header Tabel Detail
    pdf.set_fill_color(*PRIMARY_COLOR)
    pdf.set_text_color(*WHITE)
    pdf.set_font("Arial", 'B', 8)
    
    # No(10), Tanggal(20), Nama(45), IN(20), OUT(20), Aktv Jam(22), Selisih(25), Tipe(115)
    cols = [10, 20, 50, 22, 22, 22, 22, 109]
    headers = ["No", "Tanggal", "Nama Karyawan", "Jam IN", "Jam OUT", "Jam Aktv", "Selisih", "Tipe Aktivitas"]
    
    for i in range(len(headers)):
        pdf.cell(cols[i], 10, headers[i], 1, 0, 'C', True)
    pdf.ln()

    # Body Tabel Detail
    pdf.set_text_color(*TEXT_COLOR)
    pdf.set_font("Arial", '', 7.5)
    
    for i, (idx, row) in enumerate(df_detail.iterrows(), 1):
        # Zebra coloring logic
        fill = True if i % 2 == 0 else False
        pdf.set_fill_color(250, 250, 250)
        
        # New page check
        if pdf.get_y() > 180:
            pdf.add_page()
            pdf.set_fill_color(*PRIMARY_COLOR); pdf.set_text_color(*WHITE); pdf.set_font("Arial", 'B', 8)
            for j in range(len(headers)): pdf.cell(cols[j], 10, headers[j], 1, 0, 'C', True)
            pdf.ln(); pdf.set_text_color(*TEXT_COLOR); pdf.set_font("Arial", '', 7.5)

        pdf.cell(cols[0], 8, str(i), 1, 0, 'C', fill)
        pdf.cell(cols[1], 8, str(row['TRXDATE']), 1, 0, 'C', fill)
        pdf.cell(cols[2], 8, f"  {str(row['AuthName'])[:30]}", 1, 0, 'L', fill)
        pdf.cell(cols[3], 8, str(row['IN']), 1, 0, 'C', fill)
        pdf.cell(cols[4], 8, str(row['OUT']), 1, 0, 'C', fill)
        pdf.cell(cols[5], 8, str(row['CreateTim']), 1, 0, 'C', fill)
        pdf.cell(cols[6], 8, str(row['Selisih_Waktu']), 1, 0, 'C', fill)
        pdf.cell(cols[7], 8, f"  {str(row['Sumber Aktivitas'])}", 1, 1, 'L', fill)
        
    # --- 5. FOOTER ---
    pdf.ln(5)
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    current_ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    pdf.cell(0, 10, f"Dokumen ini bersifat rahasia. Dibuat secara otomatis pada: {current_ts}", 0, 1, 'R')
    
    return pdf.output(dest='S').encode('latin-1')

# ======================================
# 4. LOGIN SYSTEM WITH ANIMATION
# ======================================

def login_form():
    _, col_center, _ = st.columns([1, 1.2, 1]) 
    with col_center:
        st.markdown("<h2 style='text-align: center; color: #1F497D;'>🔒 Sistem Audit Pro</h2>", unsafe_allow_html=True)
        with st.container(border=True):
            with st.form("login_form"):
                u = st.text_input("Username")
                p = st.text_input("Password", type="password")
                if st.form_submit_button("Masuk 🚀", use_container_width=True):
                    if u == USERNAME_TARGET and p == PASSWORD_TARGET:
                        with st.status("Mengautentikasi...", expanded=False) as status:
                            st.write("🔍 Memverifikasi kredensial...")
                            time.sleep(0.7)
                            status.update(label="Akses Diberikan!", state="complete")
                        st.session_state['logged_in'] = True
                        st.rerun() 
                    else:
                        st.error("❌ Username atau Password salah.")

# ======================================
# 5. MAIN APPLICATION
# ======================================

def show_main_app():
    if st.sidebar.button("Keluar (Logout) 🚪"):
        st.session_state['logged_in'] = False
        st.rerun()

    st.title("📊 Dashboard Audit Aktivitas Luar Jam Kerja")
    st.markdown("---")

    uploaded_file = st.file_uploader("📂 Upload File Report Excel", type=["xlsx"])

    if uploaded_file:
        with st.spinner('🔄 Menyiapkan data...'):
            time.sleep(1)
            xls = pd.ExcelFile(uploaded_file, engine='openpyxl') 
            if "DATA ABSEN" not in xls.sheet_names:
                st.error("❌ Sheet 'DATA ABSEN' tidak ditemukan."); st.stop()
                
            df_absen = pd.read_excel(xls, sheet_name="DATA ABSEN")
            df_absen["AuthName"] = normalize(df_absen["Nama"])
            df_absen["Tanggal"] = pd.to_datetime(df_absen["Tanggal"], errors="coerce").dt.date
            df_absen["IN"] = to_time(df_absen["IN"])
            df_absen["OUT"] = to_time(df_absen["OUT"])
            df_absen = df_absen.dropna(subset=["AuthName", "Tanggal", "IN", "OUT"])

        dates = st.date_input("📅 Pilih Periode Audit", value=(date.today() - timedelta(days=30), date.today()))
        
        if isinstance(dates, tuple) and len(dates) == 2:
            start_date, end_date = dates
            
            with st.status("🕵️ Menganalisis aktivitas...", expanded=False) as status:
                report_sheets = ["Report Item Correct", "Report Void - Bill Cancellation", "Report Return", "Report Print Duplicate"]
                hasil = []

                for sheet in report_sheets:
                    if sheet in xls.sheet_names:
                        st.write(f"Memeriksa {sheet}...")
                        df = pd.read_excel(xls, sheet_name=sheet)
                        df["AuthName"] = normalize(df["AuthName"])
                        df["TRXDATE"] = pd.to_datetime(df["TRXDATE"], errors="coerce").dt.date
                        df["CreateTim"] = to_time(df["CreateTim"]) 

                        df = df[(df["TRXDATE"] >= start_date) & (df["TRXDATE"] <= end_date)]
                        df_merge = df.merge(df_absen, left_on=["AuthName", "TRXDATE"], right_on=["AuthName", "Tanggal"], how="inner")
                        
                        if not df_merge.empty:
                            df_merge['Start_S'] = df_merge.apply(lambda r: datetime.datetime.combine(r['TRXDATE'], r['IN']), axis=1)
                            df_merge['End_S'] = df_merge.apply(lambda r: datetime.datetime.combine(r['TRXDATE'], r['OUT']) + (timedelta(days=1) if r['IN'] > r['OUT'] else timedelta(0)), axis=1)
                            
                            df_v = df_merge[(df_merge.apply(lambda r: datetime.datetime.combine(r['TRXDATE'], r['CreateTim']), axis=1) < df_merge['Start_S']) | 
                                            (df_merge.apply(lambda r: datetime.datetime.combine(r['TRXDATE'], r['CreateTim']), axis=1) > df_merge['End_S'])].copy()
                            
                            if not df_v.empty:
                                df_v["Sumber Aktivitas"] = sheet
                                df_v["Selisih_Waktu"] = df_v.apply(calculate_minutes_diff, axis=1)
                                hasil.append(df_v)
                
                status.update(label="Analisis Selesai!", state="complete")

            if hasil:
                final_df = pd.concat(hasil, ignore_index=True).sort_values(by=['TRXDATE', 'CreateTim'])
                
                st.markdown("### 📈 Ringkasan Eksekutif")
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Pelanggaran", f"{len(final_df)} Kasus")
                m2.metric("Personel Terlibat", f"{final_df['AuthName'].nunique()} Orang")
                m3.metric("Tipe Terbanyak", final_df['Sumber Aktivitas'].mode()[0])

                st.markdown("---")
                st.subheader("🌡️ Heatmap Pola Waktu Pelanggaran")
                heatmap_data = final_df.copy()
                heatmap_data['Jam'] = heatmap_data['CreateTim'].apply(lambda x: x.hour)
                heatmap_data['Hari'] = pd.to_datetime(heatmap_data['TRXDATE']).dt.day_name()
                
                heatmap = alt.Chart(heatmap_data).mark_rect().encode(
                    x=alt.X('Jam:O', title='Jam Kejadian'),
                    y=alt.Y('Hari:O', title='Hari', sort=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']),
                    color=alt.Color('count():Q', scale=alt.Scale(scheme='reds')),
                    tooltip=['Hari', 'Jam', 'count()']
                ).properties(height=300).interactive()
                st.altair_chart(heatmap, use_container_width=True)

                st.markdown("---")
                col_t, col_c = st.columns([1, 1])
                with col_t:
                    st.subheader("🏆 Peringkat Pelanggar")
                    df_ranking = final_df.groupby("AuthName").size().reset_index(name='Total Pelanggaran').sort_values(by='Total Pelanggaran', ascending=False)
                    st.dataframe(df_ranking.style.apply(highlight_top_rank, axis=None), hide_index=True, use_container_width=True)
                with col_c:
                    st.subheader("📊 Visualisasi Top 10")
                    chart = alt.Chart(df_ranking.head(10)).mark_bar(color='#F7C351').encode(
                        y=alt.Y('AuthName', sort='-x'), x='Total Pelanggaran'
                    ).properties(height=300)
                    st.altair_chart(chart, use_container_width=True)

                st.markdown("---")
                st.subheader("🚨 Detail Pemeriksaan & Selisih Waktu")
                st.dataframe(final_df[["TRXDATE", "AuthName", "IN", "OUT", "CreateTim", "Sumber Aktivitas"]].style.apply(highlight_violation, axis=0), use_container_width=True, hide_index=True)
                
                if st.button("📄 Siapkan Laporan PDF Formal"):
                    with st.spinner('Menyusun PDF...'):
                        pdf_bytes = create_pdf(final_df, start_date, end_date)
                        st.download_button(label="⬇️ Download PDF Sekarang", data=pdf_bytes, file_name=f"Official_Report_{start_date}.pdf", mime="application/pdf")
            else:
                st.success("✅ Tidak ditemukan aktivitas mencurigakan.")
    else:
        st.info("⬆️ Silakan unggah file Excel.")

# EXECUTION
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if st.session_state['logged_in']: show_main_app()
else: login_form()