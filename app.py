import streamlit as st
import pandas as pd
import os
import time
import tempfile
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Gecikme Zammı Hesaplama",
    page_icon="📊",
    layout="centered"
)

# ==============================
# STYLE
# ==============================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    .main-header h1 { font-size: 1.8rem; margin: 0; font-weight: 700; }
    .main-header p  { font-size: 0.9rem; opacity: 0.8; margin: 0.5rem 0 0; }
    
    .info-box {
        background: #f0f7ff;
        border-left: 4px solid #0f3460;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 1.5rem;
        font-size: 0.88rem;
        color: #1a1a2e;
    }
    .stDataFrame { border-radius: 8px; overflow: hidden; }
    .stButton > button {
        background: linear-gradient(135deg, #0f3460, #1a1a2e);
        color: white;
        border: none;
        padding: 0.6rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.95rem;
        width: 100%;
        cursor: pointer;
        transition: opacity 0.2s;
    }
    .stButton > button:hover { opacity: 0.85; }
    .step-badge {
        display: inline-block;
        background: #0f3460;
        color: white;
        border-radius: 50%;
        width: 24px; height: 24px;
        text-align: center;
        line-height: 24px;
        font-size: 0.8rem;
        font-weight: 700;
        margin-right: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER
# ==============================
st.markdown("""
<div class="main-header">
    <h1>📊 Gecikme Zammı Hesaplama</h1>
    <p>GİB Dijital Hizmetler — Otomatik Form Doldurma & İndirme</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
    <b>Nasıl Çalışır?</b><br>
    Excel dosyanızı yükleyin → Veriler otomatik olarak GİB sitesine girilir → PDF ve Excel çıktıları indirilir.<br><br>
    <b>Excel Formatı:</b> A sütunu: Ödenecek Miktar | B sütunu: Vade Tarihi | C sütunu: Ödeme Tarihi
</div>
""", unsafe_allow_html=True)

# ==============================
# ŞABLON EXCEL İNDİRME
# ==============================
st.markdown("### <span class='step-badge'>1</span> Şablon İndir (İsteğe Bağlı)", unsafe_allow_html=True)

def sablon_olustur():
    df = pd.DataFrame({
        "Ödenecek Miktar (A)": [1000.00, 2500.50, 750.00],
        "Vade Tarihi (B)":      ["01.01.2023", "15.03.2023", "10.06.2023"],
        "Ödeme Tarihi (C)":     ["15.06.2024", "20.08.2024", "01.01.2025"],
    })
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

st.download_button(
    label="📥 Örnek Excel Şablonunu İndir",
    data=sablon_olustur(),
    file_name="gecikme_zammi_sablon.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# ==============================
# DOSYA YÜKLEME
# ==============================
st.markdown("### <span class='step-badge'>2</span> Excel Dosyanızı Yükleyin", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Excel dosyasını seçin (.xlsx)",
    type=["xlsx"],
    help="A: Miktar, B: Vade Tarihi, C: Ödeme Tarihi formatında olmalıdır."
)

valid_rows = []

def tarihe_cevir(deger):
    """datetime objesi, string veya sayısal Excel tarihi → GG.AA.YYYY string"""
    if deger is None or (isinstance(deger, float) and pd.isna(deger)):
        return None
    # Zaten datetime/date objesi ise (Excel'in otomatik parse ettiği)
    if hasattr(deger, "strftime"):
        return deger.strftime("%d.%m.%Y")
    s = str(deger).strip()
    # Saat kısmını at: "2024-01-01 00:00:00" → "2024-01-01"
    s = s.split(" ")[0]
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return None

if uploaded_file:
    try:
        # dtype=str KULLANMA — datetime'ları doğal haliyle oku
        df_raw = pd.read_excel(uploaded_file, header=None)
        df_raw.columns = ["Ödenecek Miktar", "Vade Tarihi", "Ödeme Tarihi"] + \
                         [f"Sütun_{i}" for i in range(3, len(df_raw.columns))]
        df_raw = df_raw[["Ödenecek Miktar", "Vade Tarihi", "Ödeme Tarihi"]].dropna(how="all")

        for _, row in df_raw.iterrows():
            m = row["Ödenecek Miktar"]
            v = tarihe_cevir(row["Vade Tarihi"])
            o = tarihe_cevir(row["Ödeme Tarihi"])
            if pd.notna(m) and v and o:
                valid_rows.append({"miktar": str(m).strip(), "vade": v, "odeme": o})

        if len(valid_rows) == 0:
            st.error("❌ Geçerli satır bulunamadı. Tarih formatını kontrol edin.")
        else:
            st.success(f"✅ Dosya okundu — **{len(valid_rows)} geçerli satır** işlenmeye hazır.")

    except Exception as e:
        st.error(f"❌ Dosya okunamadı: {e}")

st.divider()

# ==============================
# HESAPLA
# ==============================
st.markdown("### <span class='step-badge'>3</span> Hesaplamayı Başlat", unsafe_allow_html=True)

if not valid_rows:
    st.info("Lütfen önce geçerli bir Excel dosyası yükleyin.")
else:
    st.markdown(f"**{len(valid_rows)} satır** işlenmeye hazır.")

    if st.button("🚀 GİB Sitesinde Hesapla ve İndir", disabled=(len(valid_rows) == 0)):

        # Playwright kurulu mu kontrol et
        try:
            from playwright.sync_api import sync_playwright
            playwright_available = True
        except ImportError:
            playwright_available = False

        if not playwright_available:
            st.error("""
❌ **Playwright yüklü değil.**

`requirements.txt` dosyanıza şunları ekleyin:
```
playwright
openpyxl
pandas
```
Ardından terminalde çalıştırın:
```
playwright install chromium
```
""")
            st.stop()

        # ==============================
        # PLAYWRIGHT AKIŞI
        # ==============================
        progress = st.progress(0, text="Tarayıcı başlatılıyor...")
        log_area  = st.empty()
        logs      = []

        def log(msg):
            logs.append(msg)
            log_area.code("\n".join(logs), language=None)

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(accept_downloads=True)
                page    = context.new_page()

                log("🌐 GİB sitesine bağlanılıyor...")
                page.goto("https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama", timeout=30000)
                page.evaluate("document.body.style.zoom='50%'")
                progress.progress(10, text="Site yüklendi")

                def satir_sayisi():
                    return len(page.query_selector_all("input[id^='odenecekMiktar']"))

                def yeni_satir_ekle_ve_bekle():
                    once = satir_sayisi()
                    btn = page.query_selector("button[aria-label='add']")
                    btn.scroll_into_view_if_needed()
                    btn.click()
                    deadline = time.time() + 10
                    while time.time() < deadline:
                        if satir_sayisi() > once:
                            return True
                        time.sleep(0.2)
                    return False

                toplam = len(valid_rows)
                for i, row in enumerate(valid_rows):
                    form_index = satir_sayisi()

                    page.click(f"#gecikmeTipi{form_index}")
                    page.click("li:text('Gecikme Zammı')")

                    for field_id, value in [
                        (f"odenecekMiktar{form_index}", row["miktar"]),
                        (f"vadeTarihi{form_index}",     row["vade"]),
                        (f"odemeTarihi{form_index}",    row["odeme"]),
                    ]:
                        el = page.query_selector(f"#{field_id}")
                        el.fill(value)

                    log(f"✅ Satır {i+1}/{toplam}: {row['miktar']} TL | Vade: {row['vade']} | Ödeme: {row['odeme']}")
                    pct = int(10 + (i + 1) / toplam * 50)
                    progress.progress(pct, text=f"Satır dolduruluyor: {i+1}/{toplam}")

                    if i < toplam - 1:
                        if not yeni_satir_ekle_ve_bekle():
                            log("❌ Yeni satır eklenemedi, duruluyor.")
                            break

                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(0.3)

                # Hesapla
                log("⏳ Hesaplama butonu bekleniyor...")
                progress.progress(65, text="Hesaplanıyor...")
                page.wait_for_selector("#submit:enabled", timeout=15000)
                page.click("#submit")
                log("✅ HESAPLA tıklandı")

                # PDF İndir
                progress.progress(75, text="PDF indiriliyor...")
                page.wait_for_selector("#exportPdfButton:enabled", timeout=20000)

                with tempfile.TemporaryDirectory() as tmp:
                    # PDF
                    with page.expect_download() as dl_info:
                        page.click("#exportPdfButton")
                    dl = dl_info.value
                    pdf_path = os.path.join(tmp, dl.suggested_filename or "gecikme_zammi.pdf")
                    dl.save_as(pdf_path)
                    log(f"📄 PDF indirildi: {dl.suggested_filename}")
                    progress.progress(85, text="Excel indiriliyor...")

                    # Excel
                    time.sleep(2)
                    with page.expect_download() as xl_dl_info:
                        page.get_by_text("Excel'e Aktar").click()
                    xl_dl = xl_dl_info.value
                    xl_path = os.path.join(tmp, xl_dl.suggested_filename or "gecikme_zammi.xlsx")
                    xl_dl.save_as(xl_path)
                    log(f"📊 Excel indirildi: {xl_dl.suggested_filename}")

                    progress.progress(100, text="Tamamlandı!")
                    log("🎉 İşlem başarıyla tamamlandı!")

                    # Dosyaları oku ve indirme butonları sun
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
                    with open(xl_path, "rb") as f:
                        xl_bytes = f.read()

                browser.close()

            st.success("✅ Hesaplama tamamlandı! Aşağıdan dosyaları indirin.")
            st.divider()
            st.markdown("### <span class='step-badge'>4</span> Çıktıları İndir", unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📄 PDF İndir",
                    data=pdf_bytes,
                    file_name="gecikme_zammi_sonuc.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col2:
                st.download_button(
                    label="📊 Excel İndir",
                    data=xl_bytes,
                    file_name="gecikme_zammi_sonuc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            progress.empty()
            st.error(f"❌ Hata oluştu: {e}")
            log(f"❌ HATA: {e}")

# ==============================
# FOOTER
# ==============================
st.divider()
st.markdown(
    "<p style='text-align:center; color:#999; font-size:0.8rem;'>"
    "GİB Gecikme Zammı Hesaplama Aracı • Otomatik form doldurma servisi"
    "</p>",
    unsafe_allow_html=True
)
