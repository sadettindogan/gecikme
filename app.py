import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time
import tempfile
import math
import zipfile
import io
from decimal import Decimal, ROUND_DOWN

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# --- TARİH FORMATI ---
def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- ONDALIK KONTROL ---
def ondalik_kontrol(deger):
    """Değerin en fazla 2 ondalık basamaklı olup olmadığını kontrol eder."""
    try:
        d = Decimal(str(round(float(deger), 10)))
        # Normalize edip ondalık kısmını al
        isaretli = d.normalize()
        kok = d.as_tuple()
        ondalik_basamak = max(0, -kok.exponent)
        return ondalik_basamak <= 2
    except Exception:
        return True  # Sayı değilse geçir, başka kontrol yakalar

# --- ANA UYGULAMA ---
st.title("📄 Gecikme Zammı Rapor Portalı")
st.write("Satır sayısı sınırlaması yoktur. Başlık olmadan A Sütunu: Tutar, B Sütunu: Vade Tarihi, C Sütunu: Ödeme Tarihi olan Excel dosyanızı yükleyin veya aşağıya sürükleyip bırakın. Başlata tıkladıktan sonra Tamamlandı görene kadar bekleyin.")

if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None

yuklenen_dosya = st.file_uploader("Dosya Seçin (.xlsx)", type=["xlsx"])

if yuklenen_dosya:
    if st.button("🚀 Başlat"):
        st.session_state.zip_bytes = None
        tmp_dir = tempfile.mkdtemp()

        try:
            yuklenen_dosya.seek(0)
            wb_orijinal = load_workbook(yuklenen_dosya, data_only=True)
            sheet_orijinal = wb_orijinal.active

            satirlar = []
            hatali_satirlar = []
            for satir in range(1, sheet_orijinal.max_row + 1):
                a = sheet_orijinal[f"A{satir}"].value
                b = sheet_orijinal[f"B{satir}"].value
                c = sheet_orijinal[f"C{satir}"].value
                if a and b and c:
                    if not ondalik_kontrol(a):
                        hatali_satirlar.append(satir)
                    else:
                        satirlar.append((a, b, c))

            if hatali_satirlar:
                st.error(f"❌ A sütununda 3 veya daha fazla ondalık basamak içeren satırlar var: {hatali_satirlar}. Lütfen düzeltin.")
                st.stop()

            if not satirlar:
                st.error("Excel dosyasında geçerli satır bulunamadı.")
                st.stop()

            MAX_GRUP = 25
            grup_sayisi = math.ceil(len(satirlar) / MAX_GRUP)
            st.write(f"📊 **{len(satirlar)} satır** — **{grup_sayisi} grup** halinde işlenecek.")

            progress = st.progress(0)
            log = st.empty()

            sonuclar = {}

            for grup_no in range(grup_sayisi):
                baslangic = grup_no * MAX_GRUP
                bitis     = min(baslangic + MAX_GRUP, len(satirlar))
                grup      = satirlar[baslangic:bitis]
                etiket    = f"{baslangic + 1}-{bitis}"

                log.info(f"🚀 Grup {grup_no + 1}/{grup_sayisi} işleniyor: Satır {etiket}")

                with sync_playwright() as p:
                    browser = p.chromium.launch(
                        headless=True,
                        executable_path="/usr/bin/chromium",
                        args=["--no-sandbox", "--disable-dev-shm-usage"]
                    )
                    context = browser.new_context(accept_downloads=True)
                    page    = context.new_page()

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

                    for idx, (miktar, vade, odeme) in enumerate(grup):
                        son_mu = (idx == len(grup) - 1)
                        log.info(f"⏳ Grup {grup_no+1} — Satır {idx+1}/{len(grup)} işleniyor...")
                        ok = satir_doldur(miktar, vade, odeme, son_mu)
                        page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(0.3)
                        if not ok:
                            break

                    log.info("🔄 Hesaplama yapılıyor...")
                    page.wait_for_selector("#submit:enabled", timeout=15000)
                    page.click("#submit")
                    time.sleep(8)

                    log.info(f"📥 PDF indiriliyor ({etiket})...")
                    page.wait_for_selector("#exportPdfButton:enabled", timeout=30000)
                    pdf_yolu = os.path.join(tmp_dir, f"xvb_{etiket}.pdf")
                    with page.expect_download() as dl_info:
                        page.click("#exportPdfButton")
                    dl_info.value.save_as(pdf_yolu)

                    log.info(f"📥 Excel indiriliyor ({etiket})...")
                    time.sleep(2)
                    excel_yolu = os.path.join(tmp_dir, f"xvb_{etiket}.xlsx")
                    with page.expect_download() as xl_info:
                        page.get_by_text("Excel'e Aktar").click()
                    xl_info.value.save_as(excel_yolu)

                    browser.close()

                # GİB Excel'inden G sütununu oku (3. satırdan başlıyor)
                wb_gib = load_workbook(excel_yolu, data_only=True)
                sheet_gib = wb_gib.active
                for i in range(len(grup)):
                    g_degeri = sheet_gib[f"G{3 + i}"].value
                    if g_degeri is not None:
                        try:
                            g_degeri = round(float(g_degeri), 2)
                        except (ValueError, TypeError):
                            pass
                    cell = sheet_orijinal.cell(row=baslangic + 1 + i, column=4, value=g_degeri)
                    cell.number_format = "#,##0.00"

                with open(pdf_yolu, "rb") as f:
                    sonuclar[f"xvb_{etiket}.pdf"] = f.read()
                with open(excel_yolu, "rb") as f:
                    sonuclar[f"xvb_{etiket}.xlsx"] = f.read()

                progress.progress((grup_no + 1) / grup_sayisi)

            # Orijinal Excel'i D sütunuyla birlikte kaydet
            sonuc_buffer = io.BytesIO()
            wb_orijinal.save(sonuc_buffer)
            sonuclar["sonuc_dosyasi.xlsx"] = sonuc_buffer.getvalue()

            # Tüm dosyaları tek ZIP'e koy
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for dosya_adi, icerik in sonuclar.items():
                    zf.writestr(dosya_adi, icerik)
            st.session_state.zip_bytes = zip_buffer.getvalue()

            log.empty()
            st.success("✅ Tamamlandı!")

        except Exception as e:
            st.error(f"❌ Bir hata oluştu: {str(e)}")

# Tek indirme butonu
if st.session_state.zip_bytes:
    st.download_button(
        label="📦 İndir",
        data=st.session_state.zip_bytes,
        file_name="xvb_raporlar.zip",
        mime="application/zip"
    )
