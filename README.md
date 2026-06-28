# 🍽️ Yemek Tarif Platformu

Tarif yönetimi, malzeme takibi, kullanıcı değerlendirmeleri ve kişisel notları kapsayan lüks tasarımlı bir tarif yönetim sistemi. PyQt5 ile geliştirilmiş, SQLite veritabanı destekli, koyu/açık tema desteği ve QPainter tabanlı istatistik grafikleri içermektedir.

> ⚠️ Bu uygulama giriş (login) ekranı içermemektedir. Uygulama doğrudan ana ekranla başlar.

---

## 📸 Ekran Görüntüleri

> **[SCREENSHOT: Ana ekran — header, dashboard kartları ve menü, koyu lüks tema]**

---

## 📋 Özellikler

---

### 📊 Dashboard (Ana Ekran)

- Toplam Tarif, Kategoriler, Kullanıcılar, Değerlendirmeler KPI kartları
- Kişisel Notlar ve Versiyon sayısı ek KPI kartları
- Tarih ve istatistik gösterimi header'da
- Koyu / Açık tema toggle butonu

> **[SCREENSHOT: Dashboard — 6 KPI kart ve tema butonu]**

---

### 🍽️ Tarifler

- Tarif ekleme, düzenleme, silme
- Tarif adı, kategori, hazırlama süresi, pişirme süresi, zorluk, porsiyon, açıklama
- Anlık arama (ad, kategori, zorluk)
- Kategori ve zorluk filtresi
- Tarif Wizard — adım adım tarif oluşturma diyaloğu (malzeme → talimat → değerlendirme sekmeli)
- Tarif detay görünümü (malzemeler + talimatlar + değerlendirmeler + notlar + versiyonlar)
- CSV dışa aktarma
- Excel export (openpyxl)

> **[SCREENSHOT: Tarifler sayfası — liste, arama ve filtre]**

> **[SCREENSHOT: Tarif Wizard — adım adım oluşturma diyaloğu]**

---

### 🥕 Malzemeler (Tarif İçi)

- Tarife bağlı malzeme ekleme, düzenleme, silme
- Malzeme adı, miktar, birim
- Tarif detay sekmesinde listelenir

> **[SCREENSHOT: Tarif detayı — Malzemeler sekmesi]**

---

### 👨‍🍳 Talimatlar (Tarif İçi)

- Tarife bağlı adım adım talimat ekleme
- Adım numarası ve açıklama
- Tarif detay sekmesinde listelenir

> **[SCREENSHOT: Tarif detayı — Talimat sekmesi]**

---

### ⭐ Değerlendirmeler

- Tarife kullanıcı bazlı puan (1–5 yıldız) ve yorum ekleme
- Ortalama puan hesaplama ve gösterimi
- Değerlendirme listesi (kullanıcı + puan + yorum + tarih)
- Kullanıcı seçimi ile değerlendirme ekleme diyaloğu

> **[SCREENSHOT: Değerlendirmeler sayfası — puan listesi]**

---

### 👥 Kullanıcılar

- Kullanıcı ekleme (ad, soyad, email)
- Kullanıcı listesi ve email benzersizlik kontrolü
- Kullanıcı bazlı değerlendirme geçmişi

> **[SCREENSHOT: Kullanıcılar sayfası]**

---

### ✅ Yapılan Tarifler

- Tarifi "yapıldı" olarak işaretleme
- Yapılma tarihi ve not ekleme
- Yapılan tarif listesi ve istatistikleri

> **[SCREENSHOT: Yapılan Tarifler sayfası]**

---

### 📝 Kişisel Notlar *(Tier 4)*

- Serbest metin kişisel not ekleme
- Not başlığı, içerik ve tarih
- Notlar listesi ve düzenleme
- Not silme

> **[SCREENSHOT: Kişisel Notlar sayfası]**

---

### 📊 İstatistikler

- QPainter bar grafik: kategori bazlı tarif sayıları
- QPainter pasta grafik: zorluk dağılımı
- Genel özet istatistikleri

> **[SCREENSHOT: İstatistikler sayfası — bar ve pasta grafikler]**

---

### 📜 Versiyonlar *(Tier 4 — Tarif İçi)*

- Tarif bazlı versiyon geçmişi
- Versiyon notları ve tarihleri
- Versiyon karşılaştırma

> **[SCREENSHOT: Tarif detayı — Versiyonlar sekmesi]**

---

### 🔍 Gelişmiş Arama

- Tüm alanlarda (ad, malzeme, kategori, zorluk, puan) çapraz arama
- Arama sonuç listesi ve doğrudan tarife gitme

> **[SCREENSHOT: Gelişmiş Arama diyaloğu]**

---

## ❓ Final Soruları

### Sistemde hangi kullanıcılar veya nesneler vardır?

**Kullanıcı Türleri:**
- Bu uygulamada oturum açma sistemi yoktur. Tek bir genel kullanıcı arayüzü mevcuttur.
- **Kayıtlı Kullanıcılar** — Tarif değerlendirmek için sisteme ad/soyad/email ile eklenen kişiler

**Ana Nesneler / Varlıklar:**
- **Tarif** — tarif_id, ad, kategori, hazırlama süresi, pişirme süresi, zorluk, porsiyon, açıklama, oluşturma tarihi
- **Malzeme** — malzeme_id, tarif, ad, miktar, birim
- **Talimat** — talimat_id, tarif, adım no, açıklama
- **Kategori** — kategori_id, ad
- **Kullanıcı** — kullanici_id, ad, soyad, email
- **Değerlendirme** — degerlendirme_id, tarif, kullanıcı, puan (1–5), yorum, tarih
- **Yapılan Tarif** — yapilan_id, tarif, yapılma tarihi, not
- **Kişisel Not** — not_id, başlık, içerik, tarih
- **Tarif Versiyonu** — versiyon_id, tarif, versiyon notu, tarih

---

### Kullanıcı sistemde hangi işlemleri gerçekleştirebilir?

**Uygulama Kullanıcısı (giriş gerektirmez):**
- Tarif ekleyebilir, düzenleyebilir, silebilir
- Tarif Wizard ile adım adım yeni tarif oluşturabilir
- Tarife malzeme ve talimat ekleyebilir, düzenleyebilir, silebilir
- Tarifi "yapıldı" olarak işaretleyebilir ve not ekleyebilir
- Tarif versiyonu oluşturabilir ve geçmişi görüntüleyebilir
- Kullanıcı ekleyebilir ve yönetebilir
- Tarife puan ve yorum ile değerlendirme ekleyebilir
- Gelişmiş arama yapabilir
- Kişisel not ekleyebilir, düzenleyebilir, silebilir
- İstatistik ve grafikleri görüntüleyebilir
- CSV ve Excel olarak tarif listesi export edebilir
- Koyu / Açık tema arasında geçiş yapabilir

---

## 🖥️ Teknolojiler

| Teknoloji | Kullanım Alanı |
|-----------|----------------|
| Python 3.9+ | Ana programlama dili |
| PyQt5 | GUI Framework |
| SQLite3 | Veritabanı yönetimi (`@contextmanager`) |
| QPainter | İstatistik grafikleri (Bar, Pie) |
| openpyxl | Excel export |
| csv | CSV export |
