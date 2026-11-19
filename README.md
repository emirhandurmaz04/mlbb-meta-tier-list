# Mobile Legends Meta Analiz Aracı

Mobile Legends: Bang Bang'in resmi rank istatistik sayfasından kahraman verilerini **otomatik olarak çeken**, her kahramana bir **güç puanı (Hero Power)** veren ve güncel meta'yı **tier list** olarak gösteren bir masaüstü uygulaması.

Python · PyQt5 · Selenium · openpyxl

---

## Özellikler

- **Otomatik veri çekme:** Selenium ile [mobilelegends.com/rank](https://www.mobilelegends.com/rank) sayfası açılır, seçilen zaman aralığı (son 1–30 gün) ve rank filtresi (tümü veya Epic – Mythical Glory+) uygulanır. Sayfadaki liste kaydırıldıkça yüklendiği için, satır sayısı artmayı bırakana kadar otomatik kaydırılır. Ardından her kahramanın seçilme, kazanma ve yasaklanma oranları ile portresi alınır.
- **Hero Power ve tier hesaplama:** Her kahraman aşağıdaki formülle puanlanır ve sıralamadaki yerine göre SSS'den F'ye kadar bir tier'a yerleştirilir.
- **Meta Tier List tablosu:** Rol, lane ve tier'a göre birlikte çalışan filtreler, isimle arama, sütuna tıklayarak sıralama. Bir kahramanın rolü ve lane'i tabloda çoklu seçimle düzenlenebilir.
- **Excel'e aktarma:** Tablo, tier renkleriyle birlikte `.xlsx` dosyası olarak kaydedilir.
- **Lane × Tier tablosu:** Satırlarda tier'lar, sütunlarda lane'ler, hücrelerde kahraman portreleri. Portre boyutu, 7 tier'ın tamamı tek ekrana sığacak şekilde otomatik ayarlanır.
- **Yeni kahraman tespiti:** Listede olmayan bir kahraman çıktığında program rol ve lane bilgisini soran bir pencere açar ve seçimi kaydeder.
- **Kalıcı veri:** Son sorgu ve tier list kaydedilir; program açılınca son sonuçlar otomatik yüklenir.

---

## Hero Power Formülü

```
Hero Power = (1.4 × seçilme_oranı × kazanma_oranı) + (3 × yasaklanma_oranı)
```

- **Seçilme × kazanma:** Hem çok seçilen hem de çok kazanan kahramanları öne çıkarır.
- **Yasaklanma:** Sık banlanan bir kahraman maçlarda az görünse de oyuncular tarafından güçlü kabul edilir. Bu yüzden ayrıca ağırlıklandırılır.

Kahramanlar Hero Power'a göre sıralanır ve yüzdelik dilimlere göre tier'lara ayrılır:

| Tier | Sıralamadaki dilim |
|---|---|
| SSS | İlk %3 |
| SS | %3 – %13 |
| S | %13 – %33 |
| A | %33 – %53 |
| B | %53 – %73 |
| C | %73 – %90 |
| F | Son %10 |

---

## Kurulum ve Çalıştırma

Bilgisayarda **Google Chrome** kurulu olmalıdır. Uygun ChromeDriver, `webdriver-manager` tarafından otomatik indirilir.

```bash
pip install -r requirements.txt
python mobile_legends_tier_list.py
```

1. **Veri Çek** sekmesinde zaman aralığını ve rank'ı seç, **Veri Çek** butonuna bas.
2. İşlem bitince **Meta Tier List** ve **Lane × Tier Tablosu** sekmelerinde sonuçları incele.

## Proje Yapısı

```
├── mobile_legends_tier_list.py   # Uygulamanın tamamı
├── requirements.txt              # Gerekli kütüphaneler
├── output/                       # Program çalışınca otomatik oluşur
│   ├── mobile_legends_heroes.txt # Çekilen ham veriler
│   ├── heroes_tier_list.txt      # Son hesaplanan tier list
│   ├── hero_roles.json           # Kahraman rol / lane atamaları
│   └── last_query.json           # Son seçilen rank ve zaman aralığı
└── IMAGES/                       # Program çalışınca otomatik oluşur
    ├── HEROES/                   # İndirilen kahraman portreleri
    └── RANKS/                    # Rank logoları (isteğe bağlı, elle eklenir)
```

## Notlar

- Veriler sitedeki HTML yapısından okunduğu için, site tasarımı değişirse koddaki XPath seçicilerinin güncellenmesi gerekebilir.
- Bu proje eğitim amaçlı kişisel bir projedir; Moonton veya Mobile Legends: Bang Bang ile bir bağlantısı yoktur. Oyuna ait görseller depoya eklenmemiştir.

---

**Emirhan Durmaz** · [LinkedIn](https://www.linkedin.com/in/emirhandurmaz04) · [GitHub](https://github.com/emirhandurmaz04)
