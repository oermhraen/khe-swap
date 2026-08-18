import streamlit as st
import PyPDF2
import re
import io
import openpyxl

# Kod haritasi
KOD_HARITASI = {
    "503": "KHE-P01", "504": "KHE-P02", "508": "KHE-P04", "509": "KHE-P05",
    "510": "KHE-P06", "513": "KHE-P07", "514": "KHE-P08", "517": "KHE-P09",
    "520": "KHE-P10", "521": "KHE-P11", "522": "KHE-P12", "535": "KHE-P14",
    "547": "KHE-P15", "550": "KHE-P16", "562": "KHE-P17", "707": "KHE-S03",
    "708": "KHE-S04"
}

def pdf_verilerini_cek(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
        
    veriler = {}
    
    # Genel ve Esanjor Ozellikleri
    tarih_match = re.search(r"Tarih\s+([\d.]+)", text)
    veriler["Tarih"] = tarih_match.group(1) if tarih_match else ""
    
    model_match = re.search(r"Model\s+(?:MIT\s+)?(\d+)", text)
    if model_match:
        mit_kodu = model_match.group(1)
        veriler["Model_Kodlu"] = KOD_HARITASI.get(mit_kodu, f"Bulunamadi ({mit_kodu})")
        veriler["Model_Raw"] = f"MIT {mit_kodu}"
    else:
        veriler["Model_Kodlu"] = ""
        veriler["Model_Raw"] = ""

    kapasite_match = re.search(r"Kapasite\s+([\d,]+)", text)
    veriler["Kapasite"] = kapasite_match.group(1) if kapasite_match else ""
    
    plaka_match = re.search(r"Toplam Plaka Sayısı\s+(\d+)", text)
    veriler["Plaka_Sayisi"] = plaka_match.group(1) if plaka_match else ""
    
    dizilim_match = re.search(r"Plaka Dizilimi\s+([A-Za-z0-9\s+]+)", text)
    veriler["Plaka_Dizilimi"] = dizilim_match.group(1).strip() if dizilim_match else ""
    
    alan_match = re.search(r"Isı Tran[s]?fer Alanı\s+([\d,]+)", text)
    veriler["Isi_Transfer_Alani"] = alan_match.group(1) if alan_match else ""
    
    marjin_match = re.search(r"Eşanjör marjini\s+([\d,]+)", text)
    veriler["Esanjor_Marjini"] = marjin_match.group(1) if marjin_match else ""
    
    k_degeri_match = re.search(r"Görev k değeri\s+([\d\s/]+)\s*W", text)
    veriler["K_Degeri"] = k_degeri_match.group(1).strip() if k_degeri_match else ""
    
    lmtd_match = re.search(r"LMTD\s+([\d,]+)", text)
    veriler["LMTD"] = lmtd_match.group(1) if lmtd_match else ""

    # Iki sutunlu veriler (Primer ve Sekonder)
    def ikili_cek(kalip):
        match = re.search(kalip, text)
        if match:
            return match.group(1).strip(), match.group(2).strip()
        return "", ""

    veriler["Primer_Akiskan"], veriler["Sekonder_Akiskan"] = ikili_cek(r"Akışkan Cinsi\s+(\S+)\s+(\S+)")
    veriler["Primer_Gecis"], veriler["Sekonder_Gecis"] = ikili_cek(r"Geçiş Sayısı\s+(\d+)\s+(\d+)")
    veriler["Primer_Debi"], veriler["Sekonder_Debi"] = ikili_cek(r"Akışkan Debisi\s+([\d,]+)\s*m³/h\s+([\d,]+)")
    veriler["Primer_Giris_Sicakligi"], veriler["Sekonder_Giris_Sicakligi"] = ikili_cek(r"Giriş Sıcaklığı\s+([\d,]+)\s*°C\s+([\d,]+)")
    veriler["Primer_Cikis_Sicakligi"], veriler["Sekonder_Cikis_Sicakligi"] = ikili_cek(r"Çıkış Sıcaklığı\s+([\d,]+)\s*°C\s+([\d,]+)")
    veriler["Primer_Basinc_Kaybi"], veriler["Sekonder_Basinc_Kaybi"] = ikili_cek(r"Basınç Kaybı\s+([\d,]+)\s*kPa\s+([\d,]+)")
    veriler["Primer_Plaka_Basinc"], veriler["Sekonder_Plaka_Basinc"] = ikili_cek(r"Plakalardaki basınç kaybı\s+([\d,]+)\s*kPa\s+([\d,]+)")
    veriler["Primer_Baglanti_Basinc"], veriler["Sekonder_Baglanti_Basinc"] = ikili_cek(r"Bağlantılardaki basınç kaybı\s+([\d,]+)\s*kPa\s+([\d,]+)")
    veriler["Primer_Hiz"], veriler["Sekonder_Hiz"] = ikili_cek(r"Kanal Akışkan Hızı\s+([\d,]+)\s*m/s\s+([\d,]+)")
    veriler["Primer_Kirlenme"], veriler["Sekonder_Kirlenme"] = ikili_cek(r"Kirlenme faktörü\s+([\d,]+)\s*\(m² K\)/W\s+([\d,]+)")
    
    veriler["Primer_Yogunluk"], veriler["Sekonder_Yogunluk"] = ikili_cek(r"Yoğunluk\s+([\d,]+)\s*kg/m³\s+([\d,]+)")
    veriler["Primer_Ozgul_Isi"], veriler["Sekonder_Ozgul_Isi"] = ikili_cek(r"Özgül Isı\s+(\d+)\s*J/\(kg K\)\s+(\d+)")
    veriler["Primer_Iletkenlik"], veriler["Sekonder_Iletkenlik"] = ikili_cek(r"Termal İletkenlik\s+([\d,]+)\s*W/\(m K\)\s+([\d,]+)")
    veriler["Primer_Viskozite"], veriler["Sekonder_Viskozite"] = ikili_cek(r"Viskozite\s+([\d,]+)\s*cP\s+([\d,]+)")

    # Malzeme Listesi
    plaka_malzeme_match = re.search(r"Plaka Malzemesi\s+(.+)", text)
    veriler["Plaka_Malzemesi"] = plaka_malzeme_match.group(1).strip() if plaka_malzeme_match else ""
    
    conta_malzeme_match = re.search(r"Conta Malzemesi\s+(.+)", text)
    veriler["Conta_Malzemesi"] = conta_malzeme_match.group(1).strip() if conta_malzeme_match else ""
    
    govde_malzeme_match = re.search(r"Gövde Malzemesi\s+(.+)", text)
    veriler["Govde_Malzemesi"] = govde_malzeme_match.group(1).strip() if govde_malzeme_match else ""
    
    # Baglantilar (M1 => M2 ve dis yapisi)
    baglanti_p1_match = re.search(r"Primer Devre\s+(M\d+\s*=>\s*M\d+)", text)
    veriler["Baglanti_Primer_1"] = baglanti_p1_match.group(1) if baglanti_p1_match else ""
    
    baglanti_tipi_match = re.search(r"(2\"\s*Dıştan Dişli\s*CS)", text) 
    veriler["Baglanti_Tipi"] = baglanti_tipi_match.group(1) if baglanti_tipi_match else ""
    
    baglanti_s1_match = re.search(r"Sekonder Devre\s+(M\d+\s*=>\s*M\d+)", text)
    veriler["Baglanti_Sekonder_1"] = baglanti_s1_match.group(1) if baglanti_s1_match else ""

    # Agirlik ve Basinc
    agirlik_match = re.search(r"Ağırlık Boş / Dolu\s+([\d,\s/]+)\s*kg", text)
    veriler["Agirlik"] = agirlik_match.group(1).strip() if agirlik_match else ""
    
    hacim_match = re.search(r"İç Hacim Primer / Sekonder\s+([\d,\s/]+)\s*dm³", text)
    veriler["Hacim"] = hacim_match.group(1).strip() if hacim_match else ""
    
    dizayn_basinc_match = re.search(r"Dizayn / Test Basıncı\s+([\d,\s/]+)\s*bar", text)
    veriler["Dizayn_Basinci"] = dizayn_basinc_match.group(1).strip() if dizayn_basinc_match else ""
    
    calisma_sicakligi_match = re.search(r"Min/Max Çalışma Sıcaklığı\s+([-\d,\s/]+)\s*°C", text)
    veriler["Calisma_Sicakligi"] = calisma_sicakligi_match.group(1).strip() if calisma_sicakligi_match else ""
    
    fark_basinc_match = re.search(r"Maksimum Diferansiyel Basınç Farkı\s+(\d+)\s*bar", text)
    veriler["Max_Fark_Basinc"] = fark_basinc_match.group(1).strip() if fark_basinc_match else ""

    return veriler

def excele_yaz(excel_file_path, v):
    # Github uzerindeki sabit dosyayi okuyoruz
    wb = openpyxl.load_workbook(excel_file_path)
    sheet = wb.active 
    
    # Genel ve Esanjor Ozellikleri
    sheet["E3"] = v.get("Tarih", "")
    sheet["A3"] = v.get("Model_Kodlu", "")
    sheet["B6"] = v.get("Kapasite", "")
    sheet["B7"] = v.get("Model_Raw", "")
    sheet["B8"] = v.get("Plaka_Sayisi", "")
    sheet["B9"] = v.get("Plaka_Dizilimi", "")
    sheet["B10"] = v.get("Isi_Transfer_Alani", "")
    sheet["B11"] = v.get("Esanjor_Marjini", "")
    sheet["B12"] = v.get("K_Degeri", "")
    sheet["B13"] = v.get("LMTD", "")
    
    # Primer Devre
    sheet["B15"] = v.get("Primer_Akiskan", "")
    sheet["B16"] = v.get("Primer_Gecis", "")
    sheet["B17"] = v.get("Primer_Debi", "")
    sheet["B18"] = v.get("Primer_Giris_Sicakligi", "")
    sheet["B19"] = v.get("Primer_Cikis_Sicakligi", "")
    sheet["B20"] = v.get("Primer_Basinc_Kaybi", "")
    sheet["B21"] = v.get("Primer_Plaka_Basinc", "")
    sheet["B22"] = v.get("Primer_Baglanti_Basinc", "")
    sheet["B23"] = v.get("Primer_Hiz", "")
    sheet["B24"] = v.get("Primer_Kirlenme", "")
    
    # Sekonder Devre
    sheet["D15"] = v.get("Sekonder_Akiskan", "")
    sheet["D16"] = v.get("Sekonder_Gecis", "")
    sheet["D17"] = v.get("Sekonder_Debi", "")
    sheet["D18"] = v.get("Sekonder_Giris_Sicakligi", "")
    sheet["D19"] = v.get("Sekonder_Cikis_Sicakligi", "")
    sheet["D20"] = v.get("Sekonder_Basinc_Kaybi", "")
    sheet["D21"] = v.get("Sekonder_Plaka_Basinc", "")
    sheet["D22"] = v.get("Sekonder_Baglanti_Basinc", "")
    sheet["D23"] = v.get("Sekonder_Hiz", "")
    sheet["D24"] = v.get("Sekonder_Kirlenme", "")
    
    # Akiskan Ozellikleri
    sheet["B27"] = v.get("Primer_Yogunluk", "")
    sheet["B28"] = v.get("Primer_Ozgul_Isi", "")
    sheet["B29"] = v.get("Primer_Iletkenlik", "")
    sheet["B30"] = v.get("Primer_Viskozite", "")
    
    sheet["D27"] = v.get("Sekonder_Yogunluk", "")
    sheet["D28"] = v.get("Sekonder_Ozgul_Isi", "")
    sheet["D29"] = v.get("Sekonder_Iletkenlik", "")
    sheet["D30"] = v.get("Sekonder_Viskozite", "")
    
    # Malzeme ve Yapisal Detaylar
    sheet["B32"] = v.get("Plaka_Malzemesi", "")
    sheet["B33"] = v.get("Conta_Malzemesi", "")
    sheet["B34"] = v.get("Govde_Malzemesi", "")
    
    # Baglantilar ve Basinc/Agirlik
    sheet["B36"] = v.get("Baglanti_Primer_1", "")
    sheet["B37"] = v.get("Baglanti_Tipi", "")
    sheet["B38"] = v.get("Baglanti_Sekonder_1", "")
    sheet["B39"] = v.get("Baglanti_Tipi", "")
    sheet["B40"] = v.get("Agirlik", "")
    sheet["B41"] = v.get("Hacim", "")
    sheet["B42"] = v.get("Dizayn_Basinci", "")
    sheet["B43"] = v.get("Calisma_Sicakligi", "")
    sheet["B44"] = v.get("Max_Fark_Basinc", "")
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- STREAMLIT ARAYUZU ---
st.set_page_config(page_title="PDF'ten Excel'e Veri Aktarimi", layout="centered")
st.title("Teknik Belge Excel Olusturucu")

# Sadece PDF yukleme alani
uploaded_pdf = st.file_uploader("Teknik Belgeyi (PDF) Yukle", type="pdf")

# GitHub'a yukleyecegin sablon dosyasinin tam adi
sablon_excel_yolu = "teknik.xlsx" 

if uploaded_pdf:
    if st.button("Uygula ve Excel'i Hazirla"):
        with st.spinner("Okunuyor..."):
            cekilen_veriler = pdf_verilerini_cek(uploaded_pdf)
            
            try:
                hazir_excel = excele_yaz(sablon_excel_yolu, cekilen_veriler)
                st.success("Aktarim Basarili!")
                st.download_button(
                    label="Hazirlanan Excel Dosyasini Indir",
                    data=hazir_excel,
                    file_name="teknik_cikti_doldurulmus.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except FileNotFoundError:
                st.error(f"Hata: '{sablon_excel_yolu}' dosyasi bulunamadi. Lutfen dosyanin GitHub reposunda app.py ile ayni dizinde oldugundan emin ol.")
