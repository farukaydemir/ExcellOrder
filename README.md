# ExcellOrder: Excel Tabanlı Sipariş Yönetim Sistemi

ExcellOrder, karmaşık Excel ürün kataloglarını profesiyonel sipariş formlarına dönüştüren, kurulumu kolay ve hızlı bir web uygulamasıdır. Özellikle kendi kullanımı için hızlı çözüm arayan satıcılar için tasarlanmıştır.

## 🚀 Özellikler

- **Esnek Excel Okuma:** Sütun yapısı ne olursa olsun, kullanıcı arayüzü üzerinden "Hangi veri hangi kolonda?" eşleştirmesi yapabilirsiniz.
- **Akıllı Görsel Çıkarımı:** Excel hücrelerine gömülü resimleri otomatik olarak ayıklar ve ürünlerle eşleştirir (Hücre başına çoklu resim desteği dahil).
- **Kategori Desteği:** Excel sekmelerini (sheets) otomatik olarak kategorilere dönüştürür.
- **HTMX ile Hızlı Arayüz:** Sayfa yenilenmeden arama, filtreleme ve sepete ekleme işlemleri.
- **Profesyonel Çıktı:** Müşteriye gönderilmeye hazır, formüllü (toplam tutar hesaplayan) ve korumalı Excel dosyası oluşturur.
- **Yerel ve Güvenli:** Verileriniz kendi bilgisayarınızda (SQLite) tutulur, dışarıya kapalıdır.

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
