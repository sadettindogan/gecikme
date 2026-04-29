import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time
import tempfile

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# --- TARİH FORMATI ---
def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- ANA UYGULAMA ---
st.title("📄 Gecikme Zammı Rapor Portalı")
st.write("Excel dosyanızı yükleyin (A: Tutar, B: Vade Tarihi, C: Ödeme Tarihi).")
st.warning("⚠️ GİB sistemi en fazla **10 satır** veri kabul etmektedir.")

yuklenen_dosya = st.file_uploader("Dosya Seçin (.xlsx)", type=["xlsx"])

if yuklenen_dosya:
    if st.button("🚀 Hesaplamayı Başlat (GİB Otomasyonu)"):
        tmp_dir = tempfile.mkdtemp()

        try:
            wb = load_workbook(yuklenen_dosya, data_only=True)
            sheet = wb.active

            satirlar = []
            for satir in range(1, sheet.max_row + 1):
                a = sheet[f"A{satir}"].value
                b = sheet[f"B{satir}"].value
                c = sheet[f"C{satir}"].value
                if a and b and c:
                    satirlar.append((a, b, c))

            if not satirlar:
                st.error("Excel dosyasında geçerli satır bulunamadı.")
                st.stop()

            if len(satirlar) > 10:
                st.error("❌ En fazla 10 satır girilebilir.")
                st.stop()

            st.write(f"📊 **{len(satirlar)} satır** işlenecek.")
            progress = st.progress(0)
            log = st.empty()

            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    executable_path="/usr/bin/chromium",
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()

                log.info("🌐 GİB sitesi açılıyor...")
                page.goto("https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama")
                page.wait_for_load_state("networkidle")
                time.sleep(3)

                def satir_sayisi():
                    return len(page.query_selector_all("input[id^='odenecekMiktar']"))

                def yeni_satir_ekle():
                    once = satir_sayisi()
                    btn = page.query_selector("button[aria-label='add']")
                    btn.scroll_into_view_if_needed()
                    btn.click()
                    deadline = time.time() + 10
                    while time.time() < deadline:
                        if satir_sayisi() > once:
                            time.sleep(0.5)
                            return True
                        time.sleep(0.2)
                    return False

                def dropdown_sec(form_index):
                    dropdown_id = f"gecikmeTipi{form_index}"
                    mevcut = page.query_selector(f"input[name='{dropdown_id}']")
                    if mevcut:
                        mevcut_deger = mevcut.get_attribute("value") or ""
                        if "Gecikme Zammı" in mevcut_deger or "gecikmeZammi" in mevcut_deger.lower():
                            return
                    dropdown_div = page.wait_for_selector(f"#{dropdown_id}", timeout=5000)
                    dropdown_div.scroll_into_view_if_needed()
                    dropdown_div.click()
                    time.sleep(0.5)
                    page.wait_for_selector("ul[role='listbox']", timeout=5000)
                    time.sleep(0.3)
                    gecikme_li = page.query_selector("li[data-value='Gecikme Zammı']") or \
                                 page.query_selector("li:has-text('Gecikme Zammı')")
                    if gecikme_li:
                        gecikme_li.click()
                        time.sleep(0.3)
                    else:
                        page.keyboard.press("Escape")

                def satir_doldur(miktar, vade, odeme, son_mu):
                    form_index = satir_sayisi()
                    vade_s  = tarih_str(vade)
                    odeme_s = tarih_str(odeme)

                    dropdown_sec(form_index)

                    inp_miktar = page.wait_for_selector(f"#odenecekMiktar{form_index}", timeout=5000)
                    inp_miktar.click()
                    inp_miktar.fill(str(miktar))

                    inp_vade = page.wait_for_selector(f"#vadeTarihi{form_index}", timeout=5000)
                    inp_vade.click()
                    inp_vade.fill(vade_s)
                    page.keyboard.press("Escape")
                    time.sleep(0.2)

                    inp_odeme = page.wait_for_selector(f"#odemeTarihi{form_index}", timeout=5000)
                    inp_odeme.click()
                    inp_odeme.fill(odeme_s)
                    page.keyboard.press("Escape")
                    time.sleep(0.2)

                    if not son_mu:
                        if not yeni_satir_ekle():
                            st.warning("Yeni satır eklenemedi, işlem durdu.")
                            return False
                    return True

                for idx, (miktar, vade, odeme) in enumerate(satirlar):
                    son_mu = (idx == len(satirlar) - 1)
                    log.info(f"⏳ Satır {idx+1} işleniyor...")
                    ok = satir_doldur(miktar, vade, odeme, son_mu)
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(0.3)
                    progress.progress((idx + 1) / len(satirlar))
                    if not ok:
                        break

                log.info("🔄 Hesaplama yapılıyor...")
                page.wait_for_selector("#submit:enabled", timeout=15000)
                page.click("#submit")
                time.sleep(4)

                log.info("📥 PDF indiriliyor...")
                page.wait_for_selector("#exportPdfButton:enabled", timeout=15000)
                pdf_yolu = os.path.join(tmp_dir, "Gecikme_Zammi_Raporu.pdf")
                with page.expect_download() as dl_info:
                    page.click("#exportPdfButton")
                dl_info.value.save_as(pdf_yolu)

                log.info("📥 Excel indiriliyor...")
                time.sleep(2)
                excel_yolu = os.path.join(tmp_dir, "Gecikme_Zammi_Raporu.xlsx")
                with page.expect_download() as xl_info:
                    page.get_by_text("Excel'e Aktar").click()
                xl_info.value.save_as(excel_yolu)

                browser.close()

            if os.path.exists(pdf_yolu):
                with open(pdf_yolu, "rb") as f:
                    st.download_button(
                        label="📥 PDF Raporunu İndir",
                        data=f.read(),
                        file_name="Gecikme_Zammi_Raporu.pdf",
                        mime="application/pdf"
                    )

            if os.path.exists(excel_yolu):
                with open(excel_yolu, "rb") as f:
                    st.download_button(
                        label="📊 Excel Raporunu İndir",
                        data=f.read(),
                        file_name="Gecikme_Zammi_Raporu.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            st.success("✅ İşlem başarıyla tamamlandı!")

        except Exception as e:
            st.error(f"❌ Bir hata oluştu: {str(e)}")
