# ExcellOrder: Excel Tabanlı Sipariş Yönetim Sistemi

ExcellOrder is an easy-to-use, fast, and professional web-based tool that transforms complex Excel product catalogs into clean, customer-ready order forms. Designed for sellers who need a quick and robust solution for local order management.

## 🚀 Key Features

- **Dynamic Column Mapping (Manual A-Z):** No matter how your Excel is structured, you can map columns (Sku, Name, Price, Images, etc.) using a simple A-B-C selection interface.
- **Smart Image Extraction:** Automatically extracts embedded images from Excel cells, supporting multiple images per row.
- **Category Support:** Automatically converts Excel tabs (sheets) into UI categories/tabs.
- **Currency Support (USD/EUR/TRY):** Choose your preferred currency during import; the system handles UI display and Excel export formatting accordingly.
- **HTMX Powered UI:** Fast, reactive experience for searching, filtering, and cart management without page reloads.
- **Professional Export:** Generates a locked, print-ready Excel order sheet with automated sum formulas for your customers.
- **Local & Private:** Your data stays on your machine (SQLite), making it safe and private.

## 🚀 Özellikler (Turkish)

- **Dinamik Sütun Eşleştirme (A-Z):** Excel yapınız nasıl olursa olsun, sütunları (Katalog No, İsim, Fiyat, Resim) A-B-C seçimiyle kolayca eşleştirebilirsiniz.
- **Akıllı Görsel Ayıklama:** Hücrelere gömülü resimleri otomatik yakalar (Aynı satırda çoklu resim desteği dahil).
- **Para Birimi Desteği (USD/EUR/TL):** İçe aktarma sırasında para biriminizi seçin; sistem hem arayüzde hem de çıktı Excel'inde tüm formatları ona göre ayarlar.
- **HTMX ile Akıcı Arayüz:** Sayfa yenilenmeden hızlıca ürün ekleme, silme ve arama.
- **Profesyonel Çıktı:** Müşteriye özel, formüllü ve korumalı Excel sipariş formu oluşturur.

## 🛠️ Teknoloji Yığını

- **Backend:** Python, Flask
- **Veritabanı:** SQLite (SQLAlchemy)
- **Frontend:** HTML5, CSS3 (Glassmorphism), HTMX
- **Veri İşleme:** Pandas, Openpyxl, Pillow

## 📦 Kurulum

1. Depoyu indirin:
   ```bash
   git clone https://github.com/kullaniciadi/excellorder.git
   cd excellorder
   ```

2. Bağımlılıkları yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

3. Uygulamayı başlatın:
   ```bash
   python app.py
   ```

4. Tarayıcınızda açın:
   `http://localhost:5001`

## 📖 Kullanım Senaryosu

1. Ürünlerinizin olduğu bir Excel dosyasını sisteme yükleyin.
2. Açılan ekranda Ürün Kodu, Ad, Fiyat ve Resimlerin hangi sütunlarda (A, B, C...) olduğunu seçin.
3. Katalogdan ürünlerinizi seçip miktar girerek sipariş listenizi oluşturun.
4. "Müşteri Excel'i Oluştur" butonuyla temiz, sadece seçili ürünlerin olduğu ve toplamları otomatik hesaplayan Excel dosyanızı indirin.

## 📄 Lisans

Bu proje MIT lisansı ile lisanslanmıştır.
