import streamlit as st
import pandas as pd
import time
import os
import tempfile
from io import BytesIO
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook

st.set_page_config(
    page_title="GİB Gecikme Zammı Hesaplayıcı",
    page_icon="🧾",
    layout="centered"
)

MAX_FORM_SATIRI = 25

st.title("🧾 GİB Gecikme Zammı Hesaplayıcı")
st.markdown(
    "GİB'in [Gecikme Zammı ve Faiz Hesaplama](https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama) "
    "aracına otomatik veri girerek PDF ve Excel sonuçlarını indirir."
)

st.divider()

# ─── Excel Şablonu İndir ───────────────────────────────────────────────────────
with st.expander("📄 Excel şablonunu indir"):
    st.markdown(
        "Başlık **olmadan** doldurun:\n\n"
        "- **A sütunu:** Miktar (TL)\n"
        "- **B sütunu:** Beyanname Tarihi (GG.AA.YYYY)\n"
        "- **C sütunu:** Ödenecek ya da Ödenen Tarih (GG.AA.YYYY)\n\n"
        "En fazla **25 satır** girilebilir. Excel dosyanızı sürükleyip bırakabilirsiniz."
    )
    ornek = pd.DataFrame({
        "A": [1000.00, 2500.50],
        "B": ["01.01.2023", "15.03.2023"],
        "C": ["10.06.2024", "20.07.2024"],
    })
    buf = BytesIO()
    ornek.to_excel(buf, index=False, header=False)
    st.download_button(
        "⬇️ Şablon İndir",
        data=buf.getvalue(),
        file_name="gib_sablon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()


# ─── Playwright Hesaplama ─────────────────────────────────────────────────────
def run_hesaplama(satirlar):
    progress = st.progress(0, text="Tarayıcı başlatılıyor…")
    log_area = st.empty()
    loglar = []

    def log(msg):
        loglar.append(msg)
        log_area.code("\n".join(loglar), language=None)

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_hedef   = os.path.join(tmpdir, "gecikme_zam.pdf")
        excel_hedef = os.path.join(tmpdir, "gecikme_zam.xlsx")

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(accept_downloads=True)
                page    = context.new_page()

                log("🌐 GİB sayfası açılıyor…")
                page.goto(
                    "https://dijital.gib.gov.tr/hesaplamalar/GecikmeZamVeFaizHesaplama",
                    timeout=30000,
                )
                page.evaluate("document.body.style.zoom='50%'")
                progress.progress(10, text="Sayfa açıldı.")

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
                            return True
                        time.sleep(0.2)
                    return False

                def gecikme_turu_sec(idx):
                    dd = page.query_selector(f"#gecikmeTipi{idx}")
                    dd.scroll_into_view_if_needed()
                    dd.click()
                    time.sleep(0.3)
                    try:
                        page.wait_for_selector("li[data-value='Gecikme Zammı']", state="visible", timeout=5000)
                        li = page.query_selector("li[data-value='Gecikme Zammı']")
                        if li:
                            li.click()
                            time.sleep(0.2)
                            return
                    except Exception:
                        pass
                    page.keyboard.press("Escape")
                    time.sleep(0.2)

                toplam = len(satirlar)
                for i, satir in enumerate(satirlar):
                    form_idx = satir_sayisi()
                    gecikme_turu_sec(form_idx)
                    for field_id, value in [
                        (f"odenecekMiktar{form_idx}", satir["Miktar (TL)"]),
                        (f"vadeTarihi{form_idx}",     satir["Beyanname Tarihi"]),
                        (f"odemeTarihi{form_idx}",    satir["Ödeme Tarihi"]),
                    ]:
                        el = page.query_selector(f"#{field_id}")
                        el.fill(str(value))

                    log(f"✅ Satır {i+1}/{toplam} forma girildi.")
                    pct = int(10 + (i + 1) / toplam * 50)
                    progress.progress(pct, text=f"Satır {i+1}/{toplam} işlendi")

                    if i < toplam - 1:
                        if not yeni_satir_ekle():
                            st.error("❌ Yeni satır eklenemedi. İşlem durdu.")
                            browser.close()
                            return

                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(0.3)

                log("⏳ Hesaplanıyor…")
                progress.progress(65, text="Hesaplama başlatılıyor…")
                page.wait_for_selector("#submit:enabled", timeout=15000)
                page.click("#submit")
                time.sleep(5)

                # PDF
                log("📥 PDF indiriliyor…")
                progress.progress(75, text="PDF indiriliyor…")
                page.wait_for_selector("#exportPdfButton:enabled", timeout=15000)
                with page.expect_download() as dl:
                    page.click("#exportPdfButton")
                dl.value.save_as(pdf_hedef)
                log("✅ PDF indirildi.")

                # Excel
                log("📥 Excel indiriliyor…")
                progress.progress(88, text="Excel indiriliyor…")
                time.sleep(2)
                with page.expect_download() as xl:
                    page.get_by_text("Excel'e Aktar").click()
                xl.value.save_as(excel_hedef)
                log("✅ Excel indirildi.")

                browser.close()

            progress.progress(100, text="✅ Tamamlandı!")
            log("🎉 Tüm işlemler tamamlandı!")

            st.divider()
            st.subheader("📥 Sonuçları İndir")
            col1, col2 = st.columns(2)

            if os.path.exists(pdf_hedef):
                with open(pdf_hedef, "rb") as f:
                    col1.download_button(
                        "📄 PDF İndir",
                        data=f.read(),
                        file_name="gecikme_zam.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            if os.path.exists(excel_hedef):
                with open(excel_hedef, "rb") as f:
                    col2.download_button(
                        "📊 Excel İndir",
                        data=f.read(),
                        file_name="gecikme_zam.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

        except Exception as e:
            st.error(f"İşlem sırasında hata: {e}")
            progress.empty()


# ─── Dosya Yükle ──────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📂 Excel dosyanızı yükleyin (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        wb = load_workbook(BytesIO(uploaded.read()), data_only=True)
        sheet = wb.active

        satirlar = []
        for i in range(1, sheet.max_row + 1):
            miktar = sheet[f"A{i}"].value
            vade   = sheet[f"B{i}"].value
            odeme  = sheet[f"C{i}"].value
            if None not in (miktar, vade, odeme):
                satirlar.append({
                    "Miktar (TL)":       miktar,
                    "Beyanname Tarihi":  vade  if isinstance(vade,  str) else vade.strftime("%d.%m.%Y"),
                    "Ödeme Tarihi":      odeme if isinstance(odeme, str) else odeme.strftime("%d.%m.%Y"),
                })

        if not satirlar:
            st.error("Excel dosyasında geçerli veri bulunamadı. A, B, C sütunlarını kontrol edin.")
            st.stop()

        if len(satirlar) > MAX_FORM_SATIRI:
            st.error(f"❌ {len(satirlar)} satır bulundu. En fazla {MAX_FORM_SATIRI} satır işlenebilir.")
            st.stop()

        st.success(f"✅ {len(satirlar)} satır veri okundu.")
        st.dataframe(pd.DataFrame(satirlar), use_container_width=True)

        if st.button("🚀 Hesapla ve İndir", type="primary", use_container_width=True):
            run_hesaplama(satirlar)

    except Exception as e:
        st.error(f"Dosya okunurken hata oluştu: {e}")
