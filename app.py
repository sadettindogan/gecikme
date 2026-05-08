import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time
import tempfile
import json
from datetime import date

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Raporu (25 Satır)", page_icon="📄")

# --- ZİYARETÇİ SAYACI ---
SAYAC_DOSYA = "ziyaretci_sayac.json"
ADMIN_SIFRE = "gecikme2024"

def sayac_yukle():
    if os.path.exists(SAYAC_DOSYA):
        with open(SAYAC_DOSYA, "r") as f:
            return json.load(f)
    return {"toplam": 0, "bugun": 0, "bugun_tarih": str(date.today())}

def sayac_kaydet(veri):
    with open(SAYAC_DOSYA, "w") as f:
        json.dump(veri, f)

def ziyareti_kaydet():
    if "ziyaret_sayildi" not in st.session_state:
        st.session_state.ziyaret_sayildi = True
        veri = sayac_yukle()
        bugun = str(date.today())
        if veri.get("bugun_tarih") != bugun:
            veri["bugun"] = 0
            veri["bugun_tarih"] = bugun
        veri["toplam"] = veri.get("toplam", 0) + 1
        veri["bugun"] = veri.get("bugun", 0) + 1
        sayac_kaydet(veri)

ziyareti_kaydet()

# Admin Paneli
if st.query_params.get("admin") == ADMIN_SIFRE:
    veri = sayac_yukle()
    with st.sidebar:
        st.markdown("### 👁️ İstatistikler")
        st.metric("Toplam Ziyaret", veri.get("toplam", 0))
        st.metric("Bugün", veri.get("bugun", 0))

def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- ANA UYGULAMA ---
st.title("📄 Gecikme Zammı Rapor Portalı")
st.info("Bu portal GİB sistemi üzerinden otomatik olarak **25 satıra kadar** hesaplama yapabilir.")

yuklenen_dosya = st.file_uploader("Excel Dosyanızı Yükleyin (.xlsx)", type=["xlsx"])

if yuklenen_dosya:
    if st.button("🚀 25 Satırlık Sorgulamayı Başlat"):
        tmp_dir = tempfile.mkdtemp()
        
        try:
            wb = load_workbook(yuklenen_dosya, data_only=True)
            sheet = wb.active
            
            # Veri Toplama
            satirlar = []
            for satir in range(1, sheet.max_row + 1):
                a = sheet[f"A{satir}"].value
                b = sheet[f"B{satir}"].value
                c = sheet[f"C{satir}"].value
                if a is not None and b is not None and c is not None:
                    satirlar.append((a, b, c))

            if not satirlar:
                st.error("Excel'de geçerli veri bulunamadı.")
                st.stop()

            if len(satirlar) > 25:
                st.error(f"❌ Excel'de {len(satirlar)} satır var. Maksimum 25 satır yükleyebilirsiniz.")
                st.stop()

            st.success(f"✅ {len(satirlar)} satır doğrulandı. İşlem sırasına alındı.")
            progress = st.progress(0)
            log = st.empty()

            with sync_playwright() as p:
                # Browser Ayarları (Sunucu uyumlu)
                browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()

                log.info("🌐 GİB sistemine bağlanılıyor...")
                page.goto("https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama")
                page.wait_for_load_state("networkidle")
                time.sleep(2)

                def satir_sayisi():
                    return len(page.query_selector_all("input[id^='odenecekMiktar']"))

                def yeni_satir_ekle():
                    once = satir_sayisi()
                    btn = page.query_selector("button[aria-label='add']")
                    btn.scroll_into_view_if_needed()
                    btn.click()
                    deadline = time.time() + 10
                    while time.time() < deadline:
                        if satir_sayisi() > once: return True
                        time.sleep(0.2)
                    return False

                def gecikme_turu_sec(f_idx):
                    try:
                        dropdown = page.wait_for_selector(f"#gecikmeTipi{f_idx}", timeout=5000)
                        dropdown.scroll_into_view_if_needed()
                        dropdown.click()
                        time.sleep(0.4)
                        # Gecikme Zammı seçeneğini bul ve tıkla
                        secenek = page.query_selector("li[data-value='Gecikme Zammı']") or \
                                  page.query_selector("li:has-text('Gecikme Zammı')")
                        if secenek:
                            secenek.click()
                        else:
                            page.keyboard.press("Escape")
                    except: pass

                # --- VERİ GİRİŞ DÖNGÜSÜ ---
                for idx, (miktar, vade, odeme) in enumerate(satirlar):
                    f_idx = satir_sayisi()
                    log.info(f"⏳ Satır {idx+1}/{len(satirlar)} dolduruluyor...")
                    
                    gecikme_turu_sec(f_idx)
                    
                    page.fill(f"#odenecekMiktar{f_idx}", str(miktar))
                    
                    v_inp = page.wait_for_selector(f"#vadeTarihi{f_idx}")
                    v_inp.fill(tarih_str(vade))
                    page.keyboard.press("Escape")
                    
                    o_inp = page.wait_for_selector(f"#odemeTarihi{f_idx}")
                    o_inp.fill(tarih_str(odeme))
                    page.keyboard.press("Escape")

                    if idx < len(satirlar) - 1:
                        if not yeni_satir_ekle():
                            st.warning("Yeni satır açılırken bir sorun oluştu.")
                            break
                    
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    progress.progress((idx + 1) / len(satirlar))

                log.info("🔄 Hesaplama yapılıyor...")
                page.click("#submit")
                time.sleep(5)

                # PDF ve EXCEL İNDİRME
                pdf_yolu = os.path.join(tmp_dir, "Rapor.pdf")
                with page.expect_download() as d1:
                    page.click("#exportPdfButton")
                d1.value.save_as(pdf_yolu)

                excel_yolu = os.path.join(tmp_dir, "Rapor.xlsx")
                with page.expect_download() as d2:
                    page.get_by_text("Excel'e Aktar").click()
                d2.value.save_as(excel_yolu)

                browser.close()

            # İndirme Butonları
            col1, col2 = st.columns(2)
            with col1:
                with open(pdf_yolu, "rb") as f:
                    st.download_button("📥 PDF Raporu", f.read(), "Gecikme_Raporu.pdf")
            with col2:
                with open(excel_yolu, "rb") as f:
                    st.download_button("📊 Excel Raporu", f.read(), "Gecikme_Raporu.xlsx")
            
            st.success("✅ İşlem başarıyla tamamlandı!")

        except Exception as e:
            st.error(f"❌ Hata oluştu: {str(e)}")
