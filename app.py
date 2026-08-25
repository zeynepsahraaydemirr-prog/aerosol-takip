import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime

# --- SİSTEM GİRİŞ ŞİFRESİ ---
DOGRU_SIFRE = "1927"

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Aerosol TPM & Arıza Takip",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS: SİYAH-BEYAZ YÜKSEK KONTRAST ---
st.markdown("""
<style>
    .stApp {
        background-color: #FFFFFF !important;
    }
    
    section[data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 2px solid #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #000000 !important;
    }

    h1, h2, h3, h4, h5, h6, label, p, span {
        color: #000000 !important;
        font-weight: 700 !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="select"] div,
    div[data-baseweb="select"] input {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 2px solid #000000 !important;
        border-radius: 6px !important;
    }
    div[data-baseweb="select"] span {
        color: #000000 !important;
        font-weight: 800 !important;
    }
    div[data-baseweb="select"] svg {
        fill: #000000 !important;
    }

    ul[data-baseweb="menu"], ul[data-baseweb="menu"] * {
        background-color: #FFFFFF !important;
        color: #000000 !important;
    }

    .stTextInput input {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 2px solid #000000 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    
    ::placeholder {
        color: #64748B !important;
        opacity: 1 !important;
    }

    .stButton>button, .stDownloadButton>button {
        background-color: #000000 !important;
        color: #FFFFFF !important;
        font-weight: 800 !important;
        border-radius: 6px !important;
        border: 2px solid #000000 !important;
        padding: 10px 18px !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #334155 !important;
        color: #FFFFFF !important;
    }
    .stButton>button *, .stDownloadButton>button * {
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

# --- ŞİFRE KONTROL MEKANİZMASI ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if st.session_state.password_input == DOGRU_SIFRE:
        st.session_state.authenticated = True
    else:
        st.session_state.authenticated = False
        st.error("Hatalı PIN Kodu! Lütfen tekrar deneyiniz.")

if not st.session_state.authenticated:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_lock1, col_lock2, col_lock3 = st.columns([1, 1.5, 1])
    with col_lock2:
        st.title("🔒 Giriş Doğrulama")
        st.caption("Aerosol TPM & Arıza Portalı Güvenli Erişim")
        st.text_input(
            "Erişim PIN Kodunu Giriniz:",
            type="password",
            key="password_input",
            placeholder="PIN Kodunu yazıp Enter'a basınız",
            on_change=check_password
        )
        st.button("Giriş Yap", on_click=check_password, use_container_width=True)
    st.stop()

# --- VERİTABANI BAĞLANTISI ---
conn = sqlite3.connect("aerosol_db.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS maintenance_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    line_name TEXT,
    station TEXT,
    error_title TEXT,
    intervened_by TEXT,
    intervention_type TEXT,
    solution_applied TEXT,
    duration_min INTEGER,
    created_at TIMESTAMP
)
""")
conn.commit()

# --- YAN MENÜ: EXCEL VE SİSTEM YÖNETİMİ ---
with st.sidebar:
    st.markdown("### 📁 Excel ve Sistem Yönetimi")
    
    # 1. Excel İndirme Butonu
    df_export = pd.read_sql_query("SELECT id as 'Kayıt No', line_name as 'Hat', station as 'İstasyon', error_title as 'Arıza Tanımı', intervened_by as 'Müdahale Eden', intervention_type as 'Yöntem', solution_applied as 'Nasıl Çözüldü', duration_min as 'Duruş Süresi (Dk)', created_at as 'Tarih' FROM maintenance_logs", conn)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Arıza Kayıtları')
    
    st.download_button(
        label="📥 Güncel Verileri Excel'e Dök (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Aerosol_Ariza_Kayitlari_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.divider()
    
    # 2. Excel Yükleme
    st.markdown("#### 📤 Excel Dosyası Yükle")
    uploaded_file = st.file_uploader("Dosya Seçiniz (.xlsx)", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            excel_df = pd.read_excel(uploaded_file)
            if st.button("Excel Verilerini Sisteme Aktar", use_container_width=True):
                for _, row in excel_df.iterrows():
                    cursor.execute("""
                    INSERT INTO maintenance_logs (line_name, station, error_title, intervened_by, intervention_type, solution_applied, duration_min, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Hat', row.get('line_name', 'A1'))),
                        str(row.get('İstasyon', row.get('station', 'Diğer'))),
                        str(row.get('Arıza Tanımı', row.get('error_title', 'Arıza'))),
                        str(row.get('Müdahale Eden', row.get('intervened_by', 'Bakım'))),
                        str(row.get('Yöntem', row.get('intervention_type', 'Fiziksel Müdahale'))),
                        str(row.get('Nasıl Çözüldü', row.get('solution_applied', 'Çözüm uygulandı'))),
                        int(row.get('Duruş Süresi (Dk)', row.get('duration_min', 10))),
                        str(row.get('Tarih', datetime.now().strftime("%Y-%m-%d %H:%M")))
                    ))
                conn.commit()
                st.success("Excel başarıyla içeri aktarıldı!")
                st.rerun()
        except Exception as e:
            st.error(f"Hata: {e}")

    st.divider()
    
    # 3. Sıfırlama
    st.markdown("#### 🗑️ Veritabanı Temizle")
    if st.button("Tüm Veritabanını Sıfırla", use_container_width=True):
        cursor.execute("DELETE FROM maintenance_logs")
        cursor.execute("DELETE FROM sqlite_sequence WHERE name='maintenance_logs'")
        conn.commit()
        if os.path.exists("ariza_kayitlari.xlsx"):
            try:
                os.remove("ariza_kayitlari.xlsx")
            except:
                pass
        st.success("Veritabanı sıfırlandı!")
        st.rerun()
        
    st.divider()
    if st.button("🚪 Çıkış Yap", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# --- ÜST BAŞLIK ALANI ---
st.title("🏭 Aerosol Üretim Hattı • Arıza & Bilgi Bankası")
st.caption("TPM Tekrarlayan Arıza Takip, Otonom Bakım ve Duruş Çözüm Portalı")
st.divider()

# --- VERİLERİ ÇEK & KPI HESAPLA ---
df_total = pd.read_sql_query("SELECT * FROM maintenance_logs", conn)

total_records = len(df_total)
total_loss = int(df_total["duration_min"].sum()) if total_records > 0 else 0
remote_count = len(df_total[df_total["intervention_type"] == "Uzaktan Bağlantı"]) if total_records > 0 else 0
top_station = str(df_total["station"].mode()[0]) if total_records > 0 else "-"

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
col_m1.metric("Toplam Arıza", f"{total_records} Adet")
col_m2.metric("Toplam Duruş", f"{total_loss} Dk")
col_m3.metric("Uzaktan Müdahale", f"{remote_count} Kayıt")
col_m4.metric("Kritik İstasyon", top_station)

st.divider()

# --- SEKMELER ---
tab1, tab2 = st.tabs(["⚡ Hızlı Arıza Kaydı & Çözüm Bul", "📊 Duruş Analizi & Kronik Hatalar"])

# --- 1. SEKME: ARIZA GİRİŞİ ---
with tab1:
    col_left, col_right = st.columns(2, gap="large")
    
    with col_left:
        st.subheader("1. Arıza ve İstasyon")
        
        # Hat Seçimi
        line_selection = st.selectbox("Üretim Hattı", ["A1", "A2", "A3", "A4", "Diğer (Manuel Giriş)"])
        if line_selection == "Diğer (Manuel Giriş)":
            line = st.text_input("Hat Kodunu Yazınız", placeholder="Örn: A5, Mini Dolum...")
        else:
            line = line_selection

        # İstasyon Seçimi
        station_selection = st.selectbox("İstasyon / Makine", [
            "QR / Kamera Kontrol",
            "Dolum Ünitesi",
            "Valf Çakma",
            "Gazlama (Gasser)",
            "Su Banyosu (Leak Test)",
            "Aktüatör / Kapak Takma",
            "Paketleme & Koli",
            "Diğer (Manuel Giriş)"
        ])
        if station_selection == "Diğer (Manuel Giriş)":
            station = st.text_input("İstasyon / Makine Adını Yazınız", placeholder="Örn: Etiketleme...")
        else:
            station = station_selection

        error_title = st.text_input("Arıza Tanımı / Belirti", placeholder="Örn: QR kod silik çıktı")
        
        # Çözüm Geçmişi Arama
        st.markdown(f"##### 💡 **{station}** İçin Çözüm Geçmişi")
        if error_title.strip() != "":
            query = """
            SELECT error_title as 'Arıza', solution_applied as 'Nasıl Çözüldü', intervened_by as 'Müdahale Eden', created_at as 'Tarih'
            FROM maintenance_logs 
            WHERE station = ? AND (LOWER(error_title) LIKE ? OR LOWER(solution_applied) LIKE ?)
            ORDER BY id DESC LIMIT 3
            """
            search_param = f"%{error_title.strip().lower()}%"
            past_data = pd.read_sql_query(query, conn, params=(station, search_param, search_param))
        else:
            query = """
            SELECT error_title as 'Arıza', solution_applied as 'Nasıl Çözüldü', intervened_by as 'Müdahale Eden', created_at as 'Tarih'
            FROM maintenance_logs 
            WHERE station = ?
            ORDER BY id DESC LIMIT 3
            """
            past_data = pd.read_sql_query(query, conn, params=(station,))

        if not past_data.empty:
            st.dataframe(past_data, use_container_width=True, hide_index=True)
        else:
            st.caption("Bu istasyon için henüz kayıtlı bir çözüm bulunamadı.")

    with col_right:
        st.subheader("2. Müdahale & Çözüm")
        intervened_by = st.text_input("Müdahale Eden Kişi / Ekip", placeholder="Örn: Oğuz, Fatih...")
        
        # Müdahale Yöntemi
        method_selection = st.radio(
            "Müdahale Yöntemi",
            ["Fiziksel Müdahale", "Uzaktan Bağlantı", "Operatör Ayarı / Temizlik", "Diğer (Manuel Giriş)"],
            horizontal=True
        )
        if method_selection == "Diğer (Manuel Giriş)":
            intervention_type = st.text_input("Müdahale Yöntemini Yazınız", placeholder="Örn: Kalibrasyon, Sensör Ayarı...")
        else:
            intervention_type = method_selection

        duration = st.number_input("Duruş Süresi (Dakika)", min_value=1, max_value=480, value=10, step=5)
        solution = st.text_input("Problem Nasıl Çözüldü? (Kısa Açıklama)", placeholder="Örn: Kamera lensi temizlendi, fotosel açısı düzeltildi")
        
        if st.button("Kaydı Tamamla ve Sisteme İşle", use_container_width=True):
            if error_title.strip() == "" or solution.strip() == "" or intervened_by.strip() == "" or line.strip() == "" or station.strip() == "" or intervention_type.strip() == "":
                st.error("Lütfen tüm alanları doldurunuz.")
            else:
                cursor.execute("""
                INSERT INTO maintenance_logs (line_name, station, error_title, intervened_by, intervention_type, solution_applied, duration_min, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (line.strip(), station.strip(), error_title.strip(), intervened_by.strip(), intervention_type.strip(), solution.strip(), duration, datetime.now().strftime("%Y-%m-%d %H:%M")))
                conn.commit()
                
                # Arka Planda Excel Güncelle
                df_to_save = pd.read_sql_query("SELECT id as 'Kayıt No', line_name as 'Hat', station as 'İstasyon', error_title as 'Arıza Tanımı', intervened_by as 'Müdahale Eden', intervention_type as 'Yöntem', solution_applied as 'Nasıl Çözüldü', duration_min as 'Duruş Süresi (Dk)', created_at as 'Tarih' FROM maintenance_logs", conn)
                df_to_save.to_excel("ariza_kayitlari.xlsx", index=False)
                
                st.success("Kayıt başarıyla eklendi!")
                st.rerun()

# --- 2. SEKME: ANALİZ & GRAFİKLER ---
with tab2:
    df = pd.read_sql_query("SELECT * FROM maintenance_logs", conn)
    
    if not df.empty:
        g1, g2 = st.columns(2, gap="medium")
        with g1:
            st.subheader("İstasyon Bazında Toplam Duruş (Dk)")
            station_loss = df.groupby("station")["duration_min"].sum().reset_index()
            st.bar_chart(station_loss.set_index("station"), color="#000000")
            
        with g2:
            st.subheader("Hat Bazında Toplam Duruş (Dk)")
            line_loss = df.groupby("line_name")["duration_min"].sum().reset_index()
            st.bar_chart(line_loss.set_index("line_name"), color="#475569")
            
        st.divider()
        st.subheader("🚨 Tekrarlayan Hatalar (Kronik Duruşlar)")
        repeat_df = df.groupby(["line_name", "station", "error_title"]).size().reset_index(name='Tekrar Sayısı')
        chronic = repeat_df[repeat_df['Tekrar Sayısı'] >= 2]
        
        if not chronic.empty:
            st.warning("Bu hatalar birden fazla kez tekrarlanmıştır:")
            st.dataframe(chronic, use_container_width=True, hide_index=True)
        else:
            st.info("Henüz 2 ve üzeri tekrarlayan kronik bir arıza kaydı yok.")
            
        st.divider()
        st.subheader("📋 Tüm Arıza Kayıtları")
        st.dataframe(df.tail(20).iloc[::-1], use_container_width=True, hide_index=True)
    else:
        st.info("Sistemde henüz kayıtlı bir duruş verisi bulunmamaktadır.")