# Su Ayak İzi Sektör Kataloğu

Sektörlere ve süreçlere göre tipik su yoğunluğu, baskın su ayak izi
türü (mavi/yeşil/gri su) ve başlıca su risklerinin; ISO 14046, AWS
Standard, CDP Su Güvenliği ve GRI 303 gibi başlıca çerçevelere göre
düzenlenmiş, aranabilir ve filtrelenebilir kataloğu.

## Dosya yapısı

```
suayakizi.py                     Uygulama (route'lar, veri doğrulama, sunucu başlatma)
su_ayak_izi_sektorleri.json      Sektör/süreç veri seti (koddan ayrı, bağımsız yönetilir)
static/icons/*.png               PWA ikonları
requirements.txt                 Çalışma zamanı bağımlılığı (Flask, openpyxl)
requirements-dev.txt             Test bağımlılığı (pytest)
Dockerfile                       Prod imajı tanımı
tests/                           pytest testleri (veri bütünlüğü + route'lar)
```

## Hızlı başlangıç (geliştirme / Termux dahil)

```bash
pip install -r requirements.txt
python suayakizi.py
```

Tarayıcıdan aç: http://127.0.0.1:8081

Farklı port/host için ortam değişkenleri kullanılabilir:

```bash
SUAYAK_PORT=9000 SUAYAK_HOST=0.0.0.0 python suayakizi.py
```

## Prod ortamı

```bash
pip install -r requirements.txt gunicorn
gunicorn -w 2 -b 0.0.0.0:8081 suayakizi:app
```

## Docker

```bash
docker build -t su-ayak-izi-katalogu .
docker run -p 8081:8081 su-ayak-izi-katalogu
```

## Testler

```bash
pip install -r requirements-dev.txt
pytest
```

`tests/test_data.py` veri setinin bütünlüğünü kontrol eder.
`tests/test_routes.py` Flask uç noktalarını (`/`, `/manifest.json`,
`/sw.js`, `/icons/...`, `/api/sectors`, `/export/xlsx`, 404 sayfası)
ve arayüzün varsayılan olarak açık temada açıldığını test eder.

## API

`/api/sectors` uç noktası, katalogdaki verileri JSON olarak döner ve
şu opsiyonel sorgu parametrelerini destekler:

- `sector`, sektöre göre filtreler (örn. `1. TEKSTİL VE HAZIR GİYİM`)
- `type`, su ayak izi türüne göre filtreler (`Mavi Su`, `Yeşil Su`, `Gri Su`, `Karma`)
- `q`, süreç / standart / risk / azaltım / açıklama üzerinde serbest metin arama

Örnek: `/api/sectors?type=Gri+Su&q=boyama`

`/export/xlsx` aynı `sector` / `type` / `q` parametrelerini kabul eder
ve anlık filtreye göre biçimlendirilmiş bir `.xlsx` dosyası indirir.

## Arayüz özellikleri

- **Diğer kardeş projelerden farklı olarak varsayılan tema açık, su
  temalı bir mavidir** (koyu "derin deniz" teması isteğe bağlı olarak
  seçilebilir); bu, kullanıcının doğrudan talebi üzerine bilinçli bir
  tasarım tercihidir
- Sektör ve su ayak izi türüne göre filtreleme, serbest metin arama
  (istemci tarafında, sunucuya gidip gelmeden)
- Her kartta süreç açıklaması, su yoğunluğu rozeti, ilgili standart ve
  azaltım uygulaması bilgisi doğrudan görünür
- Filtrelenmiş sonuçları kurumsal biçimli, renk hiyerarşili bir Excel
  (.xlsx) raporu olarak indirme
- TR / EN arayüz dili geçişi; yalnızca arayüz metinleri çevrilir,
  süreç adları, açıklamalar ve standart isimleri kaynak dilinde sabit
  kalır
- Gerçek çevrimdışı kullanım: servis çalışanı uygulama kabuğunu
  önbelleğe alır
- Filtre durumu URL'ye yansır (bağlantı paylaşılabilir)

## Veri güncelleme

Sektör/süreç verisini güncellemek için yalnızca
`su_ayak_izi_sektorleri.json` dosyasını düzenlemek yeterlidir; kodda
değişiklik gerekmez. Uygulama başlarken bu dosyayı okur ve şunları
doğrular:

- Her kaydın zorunlu alanları dolu mu
- `Su Yoğunluğu` değeri yalnızca `Düşük`, `Orta`, `Yüksek` veya
  `Çok Yüksek` olabilir
- `Su Ayak İzi Türü` değeri yalnızca `Mavi Su`, `Yeşil Su`, `Gri Su`
  veya `Karma` olabilir

## Sorumluluk reddi

Bu katalog genel bilgilendirme amaçlıdır. Su yoğunluğu düzeyleri
göreli/nitel kategorilerdir, kesin ölçüm değeri (ör. m3/ton gibi somut
rakamlar) yerine geçmez; gerçek su tüketimi konum, teknoloji ve
ölçeğe göre önemli ölçüde değişir. Bilinçli olarak somut sayısal
değer içermez, çünkü bu tür rakamlar hızla güncelliğini yitirebilir
ve yanlış yorumlanma riski taşır. Bu içerik yatırım, hukuki veya
teknik danışmanlık teşkil etmez; saha bazlı değerlendirme için yetkin
bir danışmana veya ilgili standart kuruluşunun (ISO, AWS, CDP, GRI)
resmi dokümanlarına başvurulmalıdır.

## Kapsam dışı bırakılanlar

- Kesin sayısal su tüketim değerleri (m3/ton, litre/birim gibi),
  çünkü bu değerler teknoloji ve coğrafyaya göre çok değişken olup
  yanıltıcı kesinlik izlenimi verebilir
- Konum bazlı su stresi haritalama (ör. WRI Aqueduct entegrasyonu),
  çünkü bu katalog sektör/süreç düzeyinde genel bir referans olmayı
  hedefler, saha veya tesis bazlı canlı risk haritalaması değil
- Veritabanı (SQLite/Postgres), çünkü veri seti küçük ve nadiren
  değiştiğinden JSON dosyası yeterli görüldü
