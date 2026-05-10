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
from datetime import datetime

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# --- TARİH FORMATI ---
def tarih_str(t):
    return t.strftime("%d.%m.%Y") if hasattr(t, "strftime") else str(t)

# --- A SÜTUNU: HÜCRE DEĞERİNİ GÜVENLİ STR'YE ÇEVİR ---
def miktar_ham_str(deger):
    if isinstance(deger, float):
        return str(deger).replace(".", ",")
    if isinstance(deger, int):
        return str(deger)
    return str(deger).strip()

# --- A SÜTUNU VALİDASYONU ---
def miktar_dogrula(satirlar):
    hatalar = []
    for idx, (a, b, c) in enumerate(satirlar):
        satir_no = idx + 1
        ham   = str(a).strip()
        deger = miktar_ham_str(a)

        if "." in deger:
            hatalar.append((satir_no, ham, "Nokta (.) içeriyor — ondalık ayracı olarak virgül (,) kullanılmalıdır."))
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
            hatalar.append((satir_no, ham, "Geçersiz karakter içeriyor (yalnızca rakam ve virgül kabul edilir)."))

    return hatalar

# --- KOPYALA YAPIŞTIR METNİNİ PARSE ET ---
# Tarih pattern: 1.01.2024 veya 01.01.2024 veya 1/01/2024 vb.
TARIH_PATTERN = re.compile(r'\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b')

def tarih_normalize(s):
    """
    '1.01.2024' → '01.01.2024' (GIB her zaman 2 haneli gün/ay bekler)
    Hem nokta hem slash ayracını kabul eder.
    """
    m = TARIH_PATTERN.match(s.strip())
    if m:
        gun, ay, yil = m.group(1), m.group(2), m.group(3)
        return f"{gun.zfill(2)}.{ay.zfill(2)}.{yil}"
    return s.strip()

def parse_yapistirilmis_metin(metin):
    """
    Excel'den kopyalanmış metni satır listesine çevirir.
    Strateji:
      1. Önce sekme (\\t) ile ayır — Excel'in standart formatı.
      2. 3 parça çıkmazsa satırdan tarihleri regex ile ayıkla,
         geriye kalan kısmı tutar olarak al.
    Her satır: (tutar_str, vade_str, odeme_str)
    """
    satirlar = []
    hatali_satirlar = []

    lines = metin.strip().splitlines()
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # --- Yöntem 1: Sekme ile ayır ---
        hucreler = line.split("\t")
        if len(hucreler) >= 3:
            tutar = hucreler[0].strip()
            vade  = tarih_normalize(hucreler[1].strip())
            odeme = tarih_normalize(hucreler[2].strip())
            if tutar and vade and odeme:
                satirlar.append((tutar, vade, odeme))
                continue

        # --- Yöntem 2: Satırdan tarihleri regex ile çıkar ---
        tarihler = TARIH_PATTERN.findall(line)
        if len(tarihler) >= 2:
            # İlk tarihin başlangıç pozisyonunu bul, öncesi tutar
            ilk_eslesme = TARIH_PATTERN.search(line)
            tutar = line[:ilk_eslesme.start()].strip()
            # Sekme/boşluk varsa temizle
            tutar = tutar.rstrip("\t ,")

            gun1, ay1, yil1 = tarihler[0]
            gun2, ay2, yil2 = tarihler[1]
            vade  = f"{gun1.zfill(2)}.{ay1.zfill(2)}.{yil1}"
            odeme = f"{gun2.zfill(2)}.{ay2.zfill(2)}.{yil2}"

            if tutar and vade and odeme:
                satirlar.append((tutar, vade, odeme))
                continue
            else:
                hatali_satirlar.append((i + 1, line, "Tutar ayrıştırılamadı — lütfen sekmeyle ayrılmış şekilde yapıştırın."))
                continue

        # --- Hiçbiri çalışmadı ---
        hatali_satirlar.append((i + 1, line, "3 sütun bulunamadı (Tutar | Vade Tarihi | Ödeme Tarihi). Sekme ayracı eksik olabilir."))

    return satirlar, hatali_satirlar

# --- VERİ KAYNAĞI SEÇİMİ ---
def veri_al(yuklenen_dosya, yapistir_metni):
    """
    Dosya yüklendi ise dosyadan, yoksa yapıştırılan metinden satırları döner.
    Dönüş: (satirlar, parse_hatalari)
    """
    if yuklenen_dosya is not None:
        yuklenen_dosya.seek(0)
        wb = load_workbook(yuklenen_dosya, data_only=True)
        sheet = wb.active
        satirlar = []
        for satir in range(1, sheet.max_row + 1):
            a = sheet[f"A{satir}"].value
            b = sheet[f"B{satir}"].value
            c = sheet[f"C{satir}"].value
            if a and b and c:
                satirlar.append((a, b, c))
        return satirlar, [], wb
    else:
        satirlar, hatalar = parse_yapistirilmis_metin(yapistir_metni)
        # Yapıştırma modunda orijinal wb'yi bellekte oluştur (D sütunu için)
        wb = Workbook()
        sheet = wb.active
        for idx, (a, b, c) in enumerate(satirlar):
            sheet[f"A{idx+1}"] = a
            sheet[f"B{idx+1}"] = b
            sheet[f"C{idx+1}"] = c
        return satirlar, hatalar, wb

# ============================================================
# ANA UYGULAMA
# ============================================================
st.title("📄 Gecikme Zammı Rapor Portalı")
st.write("Satır sayısı sınırlaması yoktur.")
st.write(
    "**A Sütunu: Tutar**, **B Sütunu: Vade Tarihi**, **C Sütunu: Ödeme Tarihi** — "
    "başlık satırı olmadan Excel dosyası yükleyin **veya** aşağıya kopyalayıp yapıştırın."
)
st.write("Başlata tıkladıktan sonra Tamamlandı görene kadar bekleyin.")

if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None
if "sonuc_satirlar" not in st.session_state:
    st.session_state.sonuc_satirlar = []  # [(tutar, vade, odeme, gecikme_zammi), ...]

# --- GİRİŞ YÖNTEMİ ---
tab_yukle, tab_yapistir = st.tabs(["📂 Dosya Yükle", "📋 Kopyala & Yapıştır"])

yuklenen_dosya   = None
yapistir_metni   = ""

with tab_yukle:
    yuklenen_dosya = st.file_uploader("Dosya Seçin (.xlsx)", type=["xlsx"])

with tab_yapistir:
    st.markdown(
        "Excel'de **A, B, C sütunlarını** seçip `Ctrl+C` ile kopyalayın, "
        "ardından aşağıya `Ctrl+V` ile yapıştırın."
    )
    yapistir_metni = st.text_area(
        "Verilerinizi buraya yapıştırın",
        height=200,
        placeholder="Tutar\tVade Tarihi\tÖdeme Tarihi\n1500\t01.01.2023\t15.06.2023\n...",
    )

    # Önizleme
    if yapistir_metni.strip():
        onizleme, parse_hatalari_on = parse_yapistirilmis_metin(yapistir_metni)
        if parse_hatalari_on:
            st.warning(f"⚠️ {len(parse_hatalari_on)} satırda format sorunu var (aşağıda gösterilir).")
        if onizleme:
            st.success(f"✅ {len(onizleme)} satır algılandı — önizleme:")
            st.dataframe(
                [{"Tutar": a, "Vade Tarihi": b, "Ödeme Tarihi": c} for a, b, c in onizleme[:10]],
                use_container_width=True,
            )
            if len(onizleme) > 10:
                st.caption(f"... ve {len(onizleme) - 10} satır daha")

# Başlat butonu — her iki modda da ortak
veri_var = yuklenen_dosya is not None or yapistir_metni.strip()

if veri_var:
    if st.button("🚀 Başlat"):
        st.session_state.zip_bytes = None
        st.session_state.sonuc_satirlar = []
        tmp_dir = tempfile.mkdtemp()

        try:
            satirlar, parse_hatalari, wb_orijinal = veri_al(yuklenen_dosya, yapistir_metni)
            sheet_orijinal = wb_orijinal.active

            # Parse hataları (kopyala-yapıştır moduna özgü)
            if parse_hatalari:
                st.error(f"❌ {len(parse_hatalari)} satırda format hatası var:")
                st.table([
                    {"Satır No": s, "İçerik": i, "Hata": h}
                    for s, i, h in parse_hatalari
                ])
                st.stop()

            if not satirlar:
                st.error("Geçerli satır bulunamadı.")
                st.stop()

            # --- VALİDASYON ---
            hatalar = miktar_dogrula(satirlar)
            if hatalar:
                st.error(
                    f"❌ **{len(hatalar)} satırda hata bulundu.** "
                    "Lütfen aşağıdaki satırları düzeltin ve tekrar deneyin. "
                    "İşlem başlatılmadı."
                )
                st.table([
                    {"Satır No": s, "Girilen Değer": d, "Hata": h}
                    for s, d, h in hatalar
                ])
                st.stop()

            st.success(f"✅ Validasyon başarılı — {len(satirlar)} satır geçerli. İşlem başlıyor...")

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
                        log.info(f"⏳ Grup {grup_no+1} — Satır {idx+1}/{len(grup)} işleniyor...")
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

                # GİB Excel'inden G sütununu oku
                wb_gib = load_workbook(excel_yolu, data_only=True)
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

                    # Sonuç satırını session_state'e kaydet
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

            # Sonuç Excel'i kaydet
            sonuc_buffer = io.BytesIO()
            wb_orijinal.save(sonuc_buffer)
            sonuclar["sonuc_dosyasi.xlsx"] = sonuc_buffer.getvalue()

            # ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for dosya_adi, icerik in sonuclar.items():
                    zf.writestr(dosya_adi, icerik)
            st.session_state.zip_bytes = zip_buffer.getvalue()

            log.empty()
            st.success("✅ Tamamlandı!")

        except Exception as e:
            st.error(f"❌ Bir hata oluştu: {str(e)}")

# --- SONUÇ TABLOSU + KOPYALA BUTONU ---
if st.session_state.sonuc_satirlar:
    st.markdown("---")
    st.subheader("📋 Sonuç — Excel'e Yapıştırmaya Hazır")

    satirlar = st.session_state.sonuc_satirlar
    df_goster = [
        {"Tutar": a, "Vade Tarihi": b, "Ödeme Tarihi": c, "Gecikme Zammı": d}
        for a, b, c, d in satirlar
    ]
    st.dataframe(df_goster, use_container_width=True)

    # Sekmeyle ayrılmış metin oluştur (Excel'e yapıştırılabilir)
    tsv_satirlar = ["\t".join([a, b, c, d]) for a, b, c, d in satirlar]
    tsv_metin = "\n".join(tsv_satirlar)

    st.markdown(
        "Aşağıdaki metni kopyalayıp doğrudan Excel'e yapıştırabilirsiniz "
        "(A, B, C, D sütunları olarak yapışır):"
    )
    st.code(tsv_metin, language=None)

    st.info(
        "💡 **Nasıl yapıştırılır:** Yukarıdaki metni seçip kopyalayın "
        "(veya sağ üstteki kopyala ikonuna tıklayın), "
        "Excel'de hedef hücreye tıklayın ve Ctrl+V yapın."
    )

# --- İNDİRME BUTONU ---
if st.session_state.zip_bytes:
    st.download_button(
        label="📦 İndir (ZIP)",
        data=st.session_state.zip_bytes,
        file_name="xvb_raporlar.zip",
        mime="application/zip"
    )
