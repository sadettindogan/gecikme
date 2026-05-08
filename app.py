import streamlit as st
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time
import tempfile
import json
from datetime import date

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Gecikme Zammı Otomasyonu", page_icon="📄")

# ... (Sayacı ve Admin Fonksiyonları Aynı Kalacak) ...

# --- ANA UYGULAMA ---
st.title("📄 Gecikme Zammı Rapor Portalı")
st.write("Excel dosyanızı yükleyin (A: Tutar, B: Vade Tarihi, C: Ödeme Tarihi).")

# --- DEĞİŞİKLİK 1: Kullanıcı Bilgilendirmesi ---
st.warning("⚠️ GİB sistemi için otomatik olarak **25 satıra** kadar destek sağlanmaktadır.")

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

            # --- DEĞİŞİKLİK 2: Satır Kontrol Sınırı ---
            if len(satirlar) > 25:
                st.error("❌ En fazla 25 satır işlenebilir. Lütfen dosyanızı küçültün.")
                st.stop()

            st.write(f"📊 **{len(satirlar)} satır** işlenecek.")
            
            # ... (Playwright ve GİB Otomasyon Kısmı Aynı Kalacak) ...
            # Not: Kodun geri kalanı dinamik olduğu için (satir_sayisi() fonksiyonu gibi) 
            # 25 satıra kadar otomatik olarak 'Ekle' butonuna basmaya devam edecektir.
