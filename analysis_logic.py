import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

# Define target titles and column names for clarity and reusability
TARGET_TITLES = ['Doç. Dr.', 'Prof. Dr.', 'Dr. Öğr. Üyesi']
WOS_Q_COLUMNS = ['WoS Q1 Makale Sayısı', 'WoS Q2 Makale Sayısı', 'WoS Q3 Makale Sayısı', 'WoS Q4 Makale Sayısı']
SCOPUS_Q_COLUMNS = ['Scopus Q1 Yayın Sayısı', 'Scopus Q2 Yayın Sayısı', 'Scopus Q3 Yayın Sayısı', 'Scopus Q4 Yayın Sayısı']
REQUIRED_COMMON_COLUMNS = ['Unvan', 'Toplam Yayın', 'Ad Soyad'] + WOS_Q_COLUMNS + SCOPUS_Q_COLUMNS


def _check_and_prepare_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame | None, str | None]:
    """
    Performs initial checks and preparations on the DataFrame.
    Returns a tuple: (prepared_DataFrame, error_message).
    On success, error_message is None. On failure, prepared_DataFrame is None.
    """
    # Filter by target titles
    df_filtered = df[df['Unvan'].isin(TARGET_TITLES)].copy()
    if df_filtered.empty:
        return None, "Yüklenen dosyada 'Prof. Dr.', 'Doç. Dr.' veya 'Dr. Öğr. Üyesi' unvanlarına sahip akademisyen bulunamadı. Lütfen dosyanızı kontrol edin."

    # Check for missing columns
    missing_cols = [col for col in REQUIRED_COMMON_COLUMNS if col not in df_filtered.columns]
    if missing_cols:
        error_msg = f"HATA: Excel dosyanızda şu sütunlar eksik: {', '.join(missing_cols)}. Lütfen dosya formatınızı kontrol edin."
        return None, error_msg

    # Calculate total Q scores
    df_filtered['WoS Q Toplamı'] = df_filtered[WOS_Q_COLUMNS].sum(axis=1)
    df_filtered['Scopus Q Toplamı'] = df_filtered[SCOPUS_Q_COLUMNS].sum(axis=1)

    return df_filtered, None


def _add_chart_to_sheet(sheet, chart, chart_anchor="E2", data_labels=True):
    """Helper function to add a chart to an OpenPyXL sheet with common settings."""
    if data_labels:
        chart.data_labels = DataLabelList(showVal=True)
    sheet.add_chart(chart, chart_anchor)


def _set_column_widths(sheet, dataframe):
    """Sets column widths in an OpenPyXL sheet based on dataframe content."""
    for i, col in enumerate(dataframe.columns):
        max_length = max(dataframe[col].astype(str).map(len).max(), len(col)) + 2
        sheet.column_dimensions[get_column_letter(i + 1)].width = max_length


def run_1_year_analysis(df: pd.DataFrame, writer: pd.ExcelWriter) -> tuple[bool, str | None]:
    """
    Performs all 1-year analysis steps and writes to the ExcelWriter.
    Returns a tuple: (success_status, error_message).
    """
    df_prepared, error = _check_and_prepare_dataframe(df)
    if error:
        return False, error

    # --- ANALYSIS 1: GENERAL OVERVIEW ---
    # ... (Sizin kodunuzun geri kalanı buraya birebir kopyalanacak, sadece st.info/success gibi UI çağrıları olmayacak)
    # 1.1 Faculty Title Distribution
    sheet_name = '1.1_Unvan_Dagilimi'
    unvan_counts = df_prepared['Unvan'].value_counts()
    unvan_distribution_df = pd.DataFrame({
        'Sıklık': unvan_counts,
        'Yüzde (%)': (unvan_counts / unvan_counts.sum() * 100).round(0)
    })
    unvan_distribution_df.to_excel(writer, sheet_name=sheet_name)

    sheet = writer.sheets[sheet_name]
    _set_column_widths(sheet, unvan_distribution_df)
    chart = PieChart()
    chart.title = "Öğretim Üyesi Unvan Dağılımı (Sıklık)"
    labels = Reference(sheet, min_col=1, min_row=2, max_row=len(unvan_distribution_df) + 1)
    data = Reference(sheet, min_col=2, min_row=1, max_row=len(unvan_distribution_df) + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    _add_chart_to_sheet(sheet, chart)

    # ... Diğer tüm analiz adımları buraya eklenecek ...
    # (KODUNUZUN GERİ KALANINI BURAYA EKLEYİN, run_3_year_no_publication_analysis dahil)
    # Ben örnek olarak birkaç tanesini ekledim, siz tamamını eklemelisiniz.

    # 2.4 Individual Publication Contribution
    title_list = [
        ('Prof. Dr.', 'PD', '2.4.1_Prof_Dr_Katki'),
        ('Doç. Dr.', 'DD', '2.4.2_Doc_Dr_Katki'),
        ('Dr. Öğr. Üyesi', 'DRU', '2.4.3_Dr_Ogr_Uyesi_Katki')
    ]
    for unvan, kod, sheet_suffix in title_list:
        # ... (Bu döngünün tamamı)
        pass # Placeholder - Bu kısmı kendi kodunuzdan kopyalayın

    # --- ANALYSIS 3: ACADEMICIANS WITH NO PUBLICATIONS ---
    total_faculty_count_by_title = df_prepared['Unvan'].value_counts()
    # ... (Bu analizin tamamı)
    pass # Placeholder - Bu kısmı kendi kodunuzdan kopyalayın

    return True, None # Başarılı biterse


def run_3_year_no_publication_analysis(df: pd.DataFrame, writer: pd.ExcelWriter) -> tuple[bool, str | None]:
    """
    Analyzes only 'no publication' academics over a 3-year period.
    Returns a tuple: (success_status, error_message).
    """
    df_prepared, error = _check_and_prepare_dataframe(df)
    if error:
        return False, error

    # --- ANALYSIS: ACADEMICIANS WITH NO PUBLICATIONS (3-YEAR FOCUS) ---
    total_faculty_count_by_title = df_prepared['Unvan'].value_counts()

    analysis_definitions = [
        ('Toplam Yayın', 'Toplam Yayın', '3.1_Yayini_Olmayanlar', 'YAYINI OLMAYAN HOCA SAYISI'),
        ('WoS Q Toplamı', 'WoS Q', '3.2_WOS_Yayini_Olmayanlar', 'WOS Q YAYINI OLMAYAN HOCA SAYISI'),
        ('Scopus Q Toplamı', 'Scopus Q', '3.3_SCOPUS_Yayini_Olmayanlar', 'SCOPUS Q YAYINI OLMAYAN HOCA SAYISI')
    ]

    for column, type_name, sheet_name, title in analysis_definitions:
        no_publications_df = df_prepared[df_prepared[column] == 0]['Unvan'].value_counts()
        report_df = pd.DataFrame({
            'TOPLAM UNVANDAKİ HOCA SAYISI': total_faculty_count_by_title,
            title: no_publications_df
        }).fillna(0).astype(int)
        report_df.to_excel(writer, sheet_name=sheet_name)

        sheet = writer.sheets[sheet_name]
        _set_column_widths(sheet, report_df)
        chart = BarChart()
        chart.type = "col"
        chart.title = f"{type_name} Yayını Olmayanların Unvana Göre Dağılımı (3 Yıllık)"
        chart.y_axis.title = "Akademisyen Sayısı"
        chart.x_axis.title = "Unvan"
        data = Reference(sheet, min_col=2, min_row=1, max_row=len(report_df) + 1, max_col=3)
        cats = Reference(sheet, min_col=1, min_row=2, max_row=len(report_df) + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        _add_chart_to_sheet(sheet, chart)

    return True, None