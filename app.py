import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook, Workbook
import os
import time
import tempfile
import math
import zipfile
import io
import re

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# --- TARİH FORMATI ---
def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- TUTAR STR ---
def miktar_ham_str(deger):
    if isinstance(deger, float):
        return str(deger).replace(".", ",")
    if isinstance(deger, int):
        return str(deger)
    return str(deger).strip()

# --- VALİDASYON ---
def miktar_dogrula(satirlar):
    hatalar = []
    for idx, (a, b, c) in enumerate(satirlar):
        satir_no = idx + 1
        ham   = str(a).strip()
        deger = miktar_ham_str(a)
        if "." in deger:
            hatalar.append((satir_no, ham, "Nokta (.) içeriyor — ondalık ayracı virgül (,) olmalıdır."))
            continue
        if "," in deger:
            parcalar = deger.split(",")
            if len(parcalar) > 2:
                hatalar.append((satir_no, ham, "Birden fazla virgül içeriyor."))
            elif len(parcalar[1]) > 2:
                hatalar.append((satir_no, ham, f"Ondalık kısmı {len(parcalar[1])} hane — en fazla 2 hane olabilir."))
                continue
        temiz = deger.replace(",", "")
        if not temiz.isdigit():
            hatalar.append((satir_no, ham, "Geçersiz karakter (yalnızca rakam ve virgül kabul edilir)."))
    return hatalar

# --- PARSE ---
TARIH_PATTERN = re.compile(r'\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b')

def tarih_normalize(s):
    m = TARIH_PATTERN.match(s.strip())
    if m:
        gun, ay, yil = m.group(1), m.group(2), m.group(3)
        return f"{gun.zfill(2)}.{ay.zfill(2)}.{yil}"
    return s.strip()

def parse_yapistirilmis_metin(metin):
    satirlar = []
    hatali_satirlar = []
    lines = metin.strip().splitlines()
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        hucreler = line.split("\t")
        if len(hucreler) >= 3:
            tutar = hucreler[0].strip()
            vade  = tarih_normalize(hucreler[1].strip())
            odeme = tarih_normalize(hucreler[2].strip())
            if tutar and vade and odeme:
                satirlar.append((tutar, vade, odeme))
                continue
        tarihler = TARIH_PATTERN.findall(line)
        if len(tarihler) >= 2:
            ilk_eslesme = TARIH_PATTERN.search(line)
            tutar = line[:ilk_eslesme.start()].strip().rstrip("\t ,")
            gun1, ay1, yil1 = tarihler[0]
            gun2, ay2, yil2 = tarihler[1]
            vade  = f"{gun1.zfill(2)}.{ay1.zfill(2)}.{yil1}"
            odeme = f"{gun2.zfill(2)}.{ay2.zfill(2)}.{yil2}"
            if tutar and vade and odeme:
                satirlar.append((tutar, vade, odeme))
                continue
            else:
                hatali_satirlar.append((i + 1, line, "Tutar ayrıştırılamadı."))
                continue
        hatali_satirlar.append((i + 1, line, "3 sütun bulunamadı. Sekme ayracı eksik olabilir."))
    return satirlar, hatali_satirlar

def wb_olustur(satirlar):
    wb = Workbook()
    sheet = wb.active
    for idx, (a, b, c) in enumerate(satirlar):
        sheet[f"A{idx+1}"] = a
        sheet[f"B{idx+1}"] = b
        sheet[f"C{idx+1}"] = c
    return wb

# ============================================================
# SESSION STATE
# ============================================================
for key, default in [
    ("zip_bytes", None),
    ("sonuc_satirlar", []),
    ("onizleme_satirlar", []),
    ("onizleme_hatalari", []),
    ("onizleme_aktif", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# BAŞLIK
# ============================================================
st.title("📄 Gecikme Zammı Rapor Portalı")
st.caption("A: Tutar · B: Vade Tarihi · C: Ödeme Tarihi — başlık satırı olmadan")

# ============================================================
# GİRİŞ
# ============================================================
st.markdown("### Veri Girişi")

yapistir_metni = st.text_area(
    "Excel'den kopyalayıp buraya yapıştırın (Ctrl+V)",
    height=160,
    placeholder="1500\t01.01.2023\t15.06.2023\n2300\t15.03.2023\t20.09.2023",
    key="ta_yapistir",
)

col_enter, col_bos = st.columns([1, 5])
with col_enter:
    enter_tiklandi = st.button("↵ Enter", use_container_width=True)

with st.expander("📂 Dosya yükle (.xlsx)"):
    yuklenen_dosya = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

# --- ÖNİZLEMEYİ TETİKLE ---
if enter_tiklandi and yapistir_metni.strip():
    onizleme, hatalar = parse_yapistirilmis_metin(yapistir_metni)
    st.session_state.onizleme_satirlar = onizleme
    st.session_state.onizleme_hatalari = hatalar
    st.session_state.onizleme_aktif    = True
    st.session_state.sonuc_satirlar    = []
    st.session_state.zip_bytes         = None

if yuklenen_dosya is not None:
    yuklenen_dosya.seek(0)
    wb_tmp    = load_workbook(yuklenen_dosya, data_only=True)
    sheet_tmp = wb_tmp.active
    satirlar_tmp = []
    for satir in range(1, sheet_tmp.max_row + 1):
        a = sheet_tmp[f"A{satir}"].value
        b = sheet_tmp[f"B{satir}"].value
        c = sheet_tmp[f"C{satir}"].value
        if a and b and c:
            satirlar_tmp.append((miktar_ham_str(a), tarih_str(b), tarih_str(c)))
    st.session_state.onizleme_satirlar = satirlar_tmp
    st.session_state.onizleme_hatalari = []
    st.session_state.onizleme_aktif    = True
    st.session_state.sonuc_satirlar    = []
    st.session_state.zip_bytes         = None

# ============================================================
# ÖNİZLEME
# ============================================================
if st.session_state.onizleme_aktif:
    onizleme = st.session_state.onizleme_satirlar
    hatalar  = st.session_state.onizleme_hatalari

    if hatalar:
        st.warning(f"⚠️ {len(hatalar)} satırda format sorunu:")
        st.table([{"Satır": s, "İçerik": ic, "Hata": h} for s, ic, h in hatalar])

    if onizleme:
        st.success(f"✅ {len(onizleme)} satır algılandı")
        st.dataframe(
            [{"Tutar": a, "Vade Tarihi": b, "Ödeme Tarihi": c} for a, b, c in onizleme],
            use_container_width=True,
            hide_index=True,
        )

        # ============================================================
        # BAŞLAT
        # ============================================================
        if st.button("🚀 Başlat", type="primary"):
            st.session_state.sonuc_satirlar = []
            st.session_state.zip_bytes      = None
            tmp_dir  = tempfile.mkdtemp()
            satirlar = onizleme

            if yuklenen_dosya is not None:
                yuklenen_dosya.seek(0)
                wb_orijinal = load_workbook(yuklenen_dosya, data_only=True)
            else:
                wb_orijinal = wb_olustur(satirlar)
            sheet_orijinal = wb_orijinal.active

            hatalar_val = miktar_dogrula(satirlar)
            if hatalar_val:
                st.error(f"❌ {len(hatalar_val)} satırda hata — işlem başlatılmadı.")
                st.table([{"Satır": s, "Değer": d, "Hata": h} for s, d, h in hatalar_val])
                st.stop()

            st.success(f"✅ {len(satirlar)} satır geçerli. İşlem başlıyor...")

            MAX_GRUP    = 25
            grup_sayisi = math.ceil(len(satirlar) / MAX_GRUP)
            st.write(f"📊 {len(satirlar)} satır — {grup_sayisi} grup")

            progress = st.progress(0)
            log      = st.empty()
            sonuclar = {}

            try:
                for grup_no in range(grup_sayisi):
                    baslangic = grup_no * MAX_GRUP
                    bitis     = min(baslangic + MAX_GRUP, len(satirlar))
                    grup      = satirlar[baslangic:bitis]
                    etiket    = f"{baslangic + 1}-{bitis}"

                    log.info(f"🚀 Grup {grup_no + 1}/{grup_sayisi} — Satır {etiket}")

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
                            btn  = page.query_selector("button[aria-label='add']")
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
                            vade_s     = tarih_str(vade)
                            odeme_s    = tarih_str(odeme)
                            dropdown_sec(form_index)
                            inp_miktar = page.wait_for_selector(f"#odenecekMiktar{form_index}", timeout=5000)
                            inp_miktar.click()
                            inp_miktar.fill(miktar_ham_str(miktar))
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
                            log.info(f"⏳ Grup {grup_no+1} — {idx+1}/{len(grup)}")
                            ok = satir_doldur(miktar, vade, odeme, son_mu)
                            page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                            time.sleep(0.3)
                            if not ok:
                                break

                        log.info("🔄 Hesaplama yapılıyor...")
                        page.wait_for_selector("#submit:enabled", timeout=15000)
                        page.click("#submit")
                        time.sleep(4)

                        log.info(f"📥 PDF indiriliyor ({etiket})...")
                        page.wait_for_selector("#exportPdfButton:enabled", timeout=15000)
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

                    # GİB Excel G sütunu
                    wb_gib    = load_workbook(excel_yolu, data_only=True)
                    sheet_gib = wb_gib.active
                    for i in range(len(grup)):
                        g_degeri = sheet_gib[f"G{3 + i}"].value
                        if g_degeri is not None:
                            try:
                                g_str = f"{float(str(g_degeri).replace(',', '.')):.2f}".replace(".", ",")
                            except (ValueError, TypeError):
                                g_str = str(g_degeri)
                        else:
                            g_str = None
                        hucre = sheet_orijinal.cell(row=baslangic + 1 + i, column=4, value=g_str)
                        hucre.number_format = "@"

                        a_val, b_val, c_val = grup[i]
                        st.session_state.sonuc_satirlar.append((
                            miktar_ham_str(a_val),
                            tarih_str(b_val),
                            tarih_str(c_val),
                            g_str or "",
                        ))

                    with open(pdf_yolu, "rb") as f:
                        sonuclar[f"xvb_{etiket}.pdf"] = f.read()
                    with open(excel_yolu, "rb") as f:
                        sonuclar[f"xvb_{etiket}.xlsx"] = f.read()

                    progress.progress((grup_no + 1) / grup_sayisi)

                sonuc_buffer = io.BytesIO()
                wb_orijinal.save(sonuc_buffer)
                sonuclar["sonuc_dosyasi.xlsx"] = sonuc_buffer.getvalue()

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for dosya_adi, icerik in sonuclar.items():
                        zf.writestr(dosya_adi, icerik)
                st.session_state.zip_bytes = zip_buffer.getvalue()

                log.empty()
                st.success("✅ Tamamlandı!")

            except Exception as e:
                st.error(f"❌ Bir hata oluştu: {str(e)}")

# ============================================================
# SONUÇ — sadece kopyalanabilir kod bloğu
# ============================================================
if st.session_state.sonuc_satirlar:
    st.markdown("---")
    tsv = "\n".join("\t".join(row) for row in st.session_state.sonuc_satirlar)
    st.code(tsv, language=None)

if st.session_state.zip_bytes:
    st.download_button(
        label="📦 İndir (ZIP)",
        data=st.session_state.zip_bytes,
        file_name="xvb_raporlar.zip",
        mime="application/zip"
    )
