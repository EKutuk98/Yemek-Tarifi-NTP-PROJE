# 🍽️ Yemek Tarif Platformu

Tarif yönetimi, malzeme takibi, kullanıcı değerlendirmeleri ve kişisel notları kapsayan lüks tasarımlı bir tarif yönetim sistemi. PyQt5 ile geliştirilmiş, SQLite veritabanı destekli, koyu/açık tema desteği ve QPainter tabanlı istatistik grafikleri içermektedir.

> ⚠️ Bu uygulama giriş (login) ekranı içermemektedir. Uygulama doğrudan ana ekranla başlar.

---

## 📸 Ekran Görüntüleri

<img width="954" height="459" alt="image" src="https://github.com/user-attachments/assets/7dad4494-a918-48ef-aa92-f86ce77e9a74" />


---

## 📋 Özellikler

---

### 📊 Dashboard (Ana Ekran)

- Toplam Tarif, Kategoriler, Kullanıcılar, Değerlendirmeler KPI kartları
- Kişisel Notlar ve Versiyon sayısı ek KPI kartları
- Tarih ve istatistik gösterimi header'da
- Koyu / Açık tema toggle butonu

<img width="1279" height="745" alt="image" src="https://github.com/user-attachments/assets/085606d4-55d7-4e39-be84-329f795eb33a" />


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


<img width="279" height="283" alt="image" src="https://github.com/user-attachments/assets/ddf491fc-533b-4eec-921b-33af75a12196" />
<img width="954" height="482" alt="image" src="https://github.com/user-attachments/assets/95d636e0-74ae-496a-846e-affcf3997d4e" />




<img width="1279" height="704" alt="image" src="https://github.com/user-attachments/assets/3e2c0810-c7e1-4097-86df-2b65012ff614" />
<img width="1279" height="749" alt="image" src="https://github.com/user-attachments/assets/b04d9111-f533-42ef-b25c-8269a325e4b2" />
<img width="1277" height="743" alt="image" src="https://github.com/user-attachments/assets/b40aa091-d000-423b-8105-01cb84c37c9d" />
<img width="1279" height="749" alt="image" src="https://github.com/user-attachments/assets/cd51fb92-66e8-4243-9f97-f51f60c85365" />




---

### 🥕 Malzemeler (Tarif İçi)

- Tarife bağlı malzeme ekleme, düzenleme, silme
- Malzeme adı, miktar, birim
- Tarif detay sekmesinde listelenir

<img width="301" height="415" alt="image" src="https://github.com/user-attachments/assets/134e9a39-612c-4372-b9c1-f88c49eda287" />
<img width="301" height="33" alt="image" src="https://github.com/user-attachments/assets/ec822ced-7829-4e0b-a33b-e9a4c07e4e6c" />
<img width="929" height="419" alt="image" src="https://github.com/user-attachments/assets/e8845896-f4b7-4d5a-9be9-73ffff2e24c5" />
<img width="522" height="450" alt="image" src="https://github.com/user-attachments/assets/ebdddf92-bc96-4a08-b9d7-379a0d30f09b" />

(GÖRSEL ÇALIŞIYOR İSTEDİĞİNİZ GİBİ BİR GÖRSEL EKLEYEBİLİRSİNİZ EĞER FOTOĞRAFI OLMASINI İSTERSENİZ) 


---

### 👨‍🍳 Talimatlar (Tarif İçi)

- Tarife bağlı adım adım talimat ekleme
- Adım numarası ve açıklama
- Tarif detay sekmesinde listelenir
<img width="536" height="454" alt="image" src="https://github.com/user-attachments/assets/f070c4dc-bf3c-4c5d-a03a-5fe8d97beb68" />
<img width="526" height="451" alt="image" src="https://github.com/user-attachments/assets/97bc7a70-8bf5-4e88-b5f9-eeb64ba10ba9" />
<img width="526" height="453" alt="image" src="https://github.com/user-attachments/assets/a208b248-a912-4004-a875-65ebb9907634" />



---

### ⭐ Değerlendirmeler
<img width="1279" height="761" alt="image" src="https://github.com/user-attachments/assets/cdf5d661-bccf-4379-b074-2f862ce04c0b" />

- Tarife kullanıcı bazlı puan (1–5 yıldız) ve yorum ekleme
- Ortalama puan hesaplama ve gösterimi
- Değerlendirme listesi (kullanıcı + puan + yorum + tarih)
- Kullanıcı seçimi ile değerlendirme ekleme diyaloğu



---

### 👥 Kullanıcılar

- Kullanıcı ekleme (ad, soyad, email)
- Kullanıcı listesi ve email benzersizlik kontrolü
- Kullanıcı bazlı değerlendirme geçmişi

<img width="1279" height="761" alt="image" src="https://github.com/user-attachments/assets/2c46677b-d535-4df7-a05a-f3f544bf894c" />
<img width="1279" height="761" alt="image" src="https://github.com/user-attachments/assets/0d00ee26-a70a-431a-a3fa-1d4354e99aa1" />
<img width="1279" height="758" alt="image" src="https://github.com/user-attachments/assets/1d31c336-017d-4612-b446-33b693cf3439" />



---

### ✅ Yapılan Tarifler

- Tarifi "yapıldı" olarak işaretleme
- Yapılma tarihi ve not ekleme
- Yapılan tarif listesi ve istatistikleri

<img width="1279" height="761" alt="image" src="https://github.com/user-attachments/assets/4a0bafa2-d300-4c12-8a14-25dfc0bc0f36" />


---

### 📝 Kişisel Notlar *(Tier 4)*

- Serbest metin kişisel not ekleme
- Not başlığı, içerik ve tarih
- Notlar listesi ve düzenleme
- Not silme

<img width="1279" height="763" alt="image" src="https://github.com/user-attachments/assets/bf31fd76-c434-451b-a1cf-66f99acbf46f" />


---

### 📊 İstatistikler

- QPainter bar grafik: kategori bazlı tarif sayıları
- QPainter pasta grafik: zorluk dağılımı
- Genel özet istatistikleri

<img width="1279" height="761" alt="image" src="https://github.com/user-attachments/assets/8f2770e8-7001-4ff3-8d30-62441f6ed31b" />


---

### 📜 Versiyonlar *(Tier 4 — Tarif İçi)*

- Tarif bazlı versiyon geçmişi
- Versiyon notları ve tarihleri
- Versiyon karşılaştırma

<img width="1279" height="766" alt="image" src="https://github.com/user-attachments/assets/0254f597-6f8c-4051-994d-efb6ed5976c5" />


---

### 🔍 Gelişmiş Arama

- Tüm alanlarda (ad, malzeme, kategori, zorluk, puan) çapraz arama
- Arama sonuç listesi ve doğrudan tarife gitme

<img width="275" height="281" alt="image" src="https://github.com/user-attachments/assets/fc9b4f9e-d0c0-478e-b5b3-eff6f1c870d2" />


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
