import streamlit as st
import pandas as pd
import io
# Logic fonksiyonlarını diğer dosyadan import ediyoruz
from analysis_logic import run_1_year_analysis, run_3_year_no_publication_analysis, REQUIRED_COMMON_COLUMNS

# --- Sayfa Ayarları ve Stil ---
st.set_page_config(
    page_title="Akademik Performans Analiz Portalı",
    page_icon="🔬",
    layout="wide"
)

# Daha şık bir görünüm için basit CSS dokunuşları
st.markdown("""
<style>
    /* Ana başlık stili */
    .st-emotion-cache-10trblm {
        color: #2c3e50; /* Koyu mavi-gri */
        font-weight: bold;
    }
    /* Buton stili */
    .stButton>button {
        color: #ffffff;
        background-color: #3498db; /* Canlı mavi */
        border-radius: 8px;
        border: none;
        padding: 10px 20px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2980b9; /* Hover rengi */
    }
    /* Sidebar stili */
    .st-emotion-cache-16txtl3 {
        background-color: #f8f9fa; /* Açık gri arka plan */
    }
</style>
""", unsafe_allow_html=True)


# --- Sidebar: Kontrol Paneli ---
with st.sidebar:
    st.image("https://www.yildiz.edu.tr/media/files/yildiz_logo_dikdortgen.png", width=200) # Örnek bir logo
    st.title("⚙️ Kontrol Paneli")
    st.markdown("---")

    analysis_choice = st.radio(
        "**1. Analiz Türünü Seçin:**",
        ("1 Yıllık Detaylı Analiz", "3 Yıllık Odaklı Analiz"),
        help="1 yıllık analiz tüm metrikleri içerir. 3 yıllık analiz ise özellikle yayını olmayan akademisyenlere odaklanır."
    )
    st.markdown("---")
    
    uploaded_file = st.file_uploader(
        "**2. Excel Dosyasını Yükleyin:**",
        type=["xlsx"]
    )
    
    with st.expander("❓ Gerekli Sütunlar Nelerdir?"):
        st.info("Yükleyeceğiniz Excel dosyası aşağıdaki sütunları içermelidir:")
        # Gerekli sütunları logic dosyasından dinamik olarak alalım
        st.code('\n'.join(REQUIRED_COMMON_COLUMNS), language='text')

# --- Ana Sayfa İçeriği ---
st.title("🔬 Akademik Yayın Performans Analizi")
st.markdown("##### Yıldız Teknik Üniversitesi Bilgi İşlem Daire Başkanlığı - Rektörlük Raporlama Birimi")
st.markdown("---")

if not uploaded_file:
    st.info("Başlamak için lütfen sol panelden bir Excel dosyası yükleyin.")
    st.image("https://i.imgur.com/gYvW5eF.png", caption="Veri Analizi İllüstrasyonu") # Hoş bir karşılama resmi

else:
    st.success(f"✅ **{uploaded_file.name}** dosyası başarıyla yüklendi.")
    st.markdown("Analizi başlatmak için aşağıdaki butona tıklayın.")
    
    if st.button("🚀 Analizi Başlat!"):
        try:
            df_original = pd.read_excel(uploaded_file)
            
            # Bellekte bir Excel dosyası oluşturmak için stream
            output = io.BytesIO()

            with st.spinner("🧠 Analiz yapılıyor, grafikler oluşturuluyor... Lütfen bekleyin."):
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    analysis_successful = False
                    error_message = None

                    if analysis_choice == "1 Yıllık Detaylı Analiz":
                        analysis_successful, error_message = run_1_year_analysis(df_original.copy(), writer)
                        output_filename = "akademik_analiz_1_yillik_rapor.xlsx"
                    else: # 3 Yıllık Odaklı Analiz
                        analysis_successful, error_message = run_3_year_no_publication_analysis(df_original.copy(), writer)
                        output_filename = "akademik_analiz_3_yillik_rapor.xlsx"

            if analysis_successful:
                st.balloons()
                st.success("🎉 Analiz başarıyla tamamlandı!")
                
                # İndirme butonunu göstermek için veriyi hazırla
                output.seek(0)
                st.download_button(
                    label="📥 Raporu İndir",
                    data=output,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Oluşturulan Excel raporunu bilgisayarınıza indirin."
                )
            else:
                # Logic dosyasından dönen hata mesajını göster
                st.error(f"❌ Analiz sırasında bir hata oluştu: \n\n {error_message}")

        except Exception as e:
            st.error(f"Beklenmedik bir hata oluştu: {e}")
            st.error("Lütfen dosya formatının doğru olduğundan ve tüm gerekli sütunların bulunduğundan emin olun.")