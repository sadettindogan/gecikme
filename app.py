import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time
import tempfile
import math
import zipfile
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# --- TARİH FORMATI ---
def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- ANA UYGULAMA ---
st.title("Gecikme Zammı Rapor Portalı")
st.info("""
**Kurallar:**
* Başlık olmadan A Sütunu: Tutar, B Sütunu: Vade Tarihi, C Sütunu: Ödeme Tarihi olmalıdır.
* **Tutar Kontrolü:** Tutarlar en fazla 2 ondalık basamak içerebilir (Örn: 1000,01 kabul edilir, 1000,001 reddedilir).
""")

if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None

yuklenen_dosya = st.file_uploader("Excel dosyanızı yükleyin (.xlsx)", type=["xlsx"])

if yuklenen_dosya:
    if st.button("🚀 Başlat"):
        st.session_state.zip_bytes = None
        tmp_dir = tempfile.mkdtemp()

        try:
            yuklenen_dosya.seek(0)
            wb_orijinal = load_workbook(yuklenen_dosya, data_only=True)
            sheet_orijinal = wb_orijinal.active

            satirlar = []
            hatali_veriler = []

            for satir in range(1, sheet_orijinal.max_row + 1):
                a = sheet_orijinal[f"A{satir}"].value
                b = sheet_orijinal[f"B{satir}"].value
                c = sheet_orijinal[f"C{satir}"].value

                if a is not None and b and c:
                    tutar_metin = str(a).replace(',', '.')
                    if '.' in tutar_metin:
                        ondalik_kisim = tutar_metin.split('.')[1]
                        if len(ondalik_kisim) > 2:
                            hatali_veriler.append(f"Satır {satir}: {a}")
                    satirlar.append((a, b, c))

            if hatali_veriler:
                st.error("⚠️ Aşağıdaki satırlarda 2'den fazla ondalık basamak var:")
                for hata in hatali_veriler:
                    st.warning(hata)
                st.stop()

            if not satirlar:
                st.error("Geçerli veri bulunamadı.")
                st.stop()

            MAX_GRUP = 25
            grup_sayisi = math.ceil(len(satirlar) / MAX_GRUP)
            progress = st.progress(0)
            log = st.empty()
            sonuclar = {}

            for grup_no in range(grup_sayisi):
                baslangic = grup_no * MAX_GRUP
                bitis     = min(baslangic + MAX_GRUP, len(satirlar))
                grup      = satirlar[baslangic:bitis]
                etiket    = f"{baslangic + 1}-{bitis}"

                with sync_playwright() as p:
                    browser = p.chromium.launch(
                        headless=True,
                        executable_path="/usr/bin/chromium",
                        args=["--no-sandbox", "--disable-dev-shm-usage"]
                    )
                    context = browser.new_context(accept_downloads=True)
                    page    = context.new_page()

                    page.goto("https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama", timeout=60000)
                    page.wait_for_load_state("networkidle")
                    time.sleep(2)

                    def satir_sayisi():
                        return len(page.query_selector_all("input[id^='odenecekMiktar']"))

                    def dropdown_sec(form_index):
                        dropdown_id = f"gecikmeTipi{form_index}"
                        page.wait_for_selector(f"#{dropdown_id}", timeout=10000).click()
                        time.sleep(0.5)
                        gecikme_li = page.query_selector("li[data-value='Gecikme Zammı']") or page.query_selector("li:has-text('Gecikme Zammı')")
                        if gecikme_li:
                            gecikme_li.click()
                        else:
                            page.keyboard.press("Escape")

                    for idx, (miktar, vade, odeme) in enumerate(grup):
                        form_idx = idx
                        if idx > 0: # İlk satır zaten var, diğerleri için ekle
                            page.click("button[aria-label='add']")
                            time.sleep(0.5)

                        dropdown_sec(form_idx)
                        page.fill(f"#odenecekMiktar{form_idx}", str(miktar))
                        page.fill(f"#vadeTarihi{form_idx}", tarih_str(vade))
                        page.keyboard.press("Escape")
                        page.fill(f"#odemeTarihi{form_idx}", tarih_str(odeme))
                        page.keyboard.press("Escape")

                    # --- HESAPLA VE BEKLE ---
                    log.info(f"🔄 Grup {grup_no+1}: Hesaplanıyor...")
                    page.click("#submit")
                    
                    # Zaman aşımı hatasını önlemek için süreyi 30 saniyeye çıkardık
                    # Ayrıca sayfanın aşağı kaydırılması gerekebilir
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")

                    # PDF butonunu bekle (Süre artırıldı)
                    try:
                        page.wait_for_selector("#exportPdfButton:not([disabled])", timeout=30000)
                    except:
                        log.error("❌ PDF butonu aktifleşmedi. Sayfa yavaş olabilir.")
                        raise

                    # PDF İndir
                    log.info(f"📥 PDF indiriliyor ({etiket})...")
                    pdf_yolu = os.path.join(tmp_dir, f"rapor_{etiket}.pdf")
                    with page.expect_download() as dl_info:
                        page.click("#exportPdfButton")
                    dl_info.value.save_as(pdf_yolu)

                    # Excel İndir
                    log.info(f"📥 Excel indiriliyor ({etiket})...")
                    excel_yolu = os.path.join(tmp_dir, f"rapor_{etiket}.xlsx")
                    with page.expect_download() as xl_info:
                        page.get_by_text("Excel'e Aktar").click()
                    xl_info.value.save_as(excel_yolu)

                    browser.close()

                # GİB Excel verilerini oku
                wb_gib = load_workbook(excel_yolu, data_only=True)
                sheet_gib = wb_gib.active
                for i in range(len(grup)):
                    g_val = sheet_gib[f"G{3 + i}"].value
                    sheet_orijinal.cell(row=baslangic + 1 + i, column=4, value=g_val)

                with open(pdf_yolu, "rb") as f: sonuclar[f"rapor_{etiket}.pdf"] = f.read()
                with open(excel_yolu, "rb") as f: sonuclar[f"rapor_{etiket}.xlsx"] = f.read()
                
                progress.progress((grup_no + 1) / grup_sayisi)

            sonuc_buffer = io.BytesIO()
            wb_orijinal.save(sonuc_buffer)
            sonuclar["toplu_sonuc.xlsx"] = sonuc_buffer.getvalue()

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for d_ad, icerik in sonuclar.items():
                    zf.writestr(d_ad, icerik)
            
            st.session_state.zip_bytes = zip_buffer.getvalue()
            log.empty()
            st.success("✅ İşlem tamamlandı!")

        except Exception as e:
            st.error(f"❌ Bir hata oluştu: {str(e)}")

if st.session_state.zip_bytes:
    st.download_button("📦 Sonuçları İndir (ZIP)", st.session_state.zip_bytes, "sonuclar.zip", "application/zip")
