# Operasyon Muhasebesi: Fişten Bilançoya — Lab

Kurs boyunca kullandığımız örnek işletme **Kahve Durağı Ltd. Şti.** tamamen **kurgusaldır**: kahve
demleme ekipmanı satan küçük bir dükkân. Bu klasörde onun Ocak 2026 defteri gerçek bir elektronik tablo
dosyası (`.xlsx`) olarak üretilir, formülleri denetlenir ve ders ekranları bu dosyadan çekilir.

> Unvan, vergi no (0000000000), adres, telefon (0222 000 00 00), e-posta (defter@example.com),
> müşteri "Örnek Kafe Ltd." ve banka "Örnek Bank" uydurmadır.
> **Bu lab mali müşavirlik hizmeti değildir. Kendi durumunuz için mali müşavirinize danışın.**

## 1. Tek komut

```bash
./kontrol.sh            # ortamı kurar, kitabı üretir, denetler (K1–K7), negatif test, görüntüleyici → GEÇTİ/KALDI
./kontrol.sh --ekran    # + bütün ders ekran görüntülerini yeniden çeker (ekran/<ders>/)
```

Gerekenler: `uv` (Python 3.12'yi ve kilitli paketleri kendisi kurar: openpyxl 3.0.10, pycel 1.0b30,
playwright 1.63.0 — `araclar/pyproject.toml` + `araclar/uv.lock`). Ekran görüntüsü için Playwright
Chromium'u (`~/.cache/ms-playwright`). LibreOffice ya da Excel gerekmez.

## 2. Öğrenci ne kullanır?

Öğrenci Python çalıştırmaz. Derslerde iki dosya indirilir ve **kendi elektronik tablo programında**
(Excel, LibreOffice Calc, Google E-Tablolar …) açılır:

| Dosya | Ne işe yarar |
|---|---|
| `defter/cikti/kahve-duragi-2026-01.xlsx` | Kahve Durağı'nın Ocak 2026 defteri: 12 sayfa, bütün toplamlar formül |
| `defter/cikti/bos-sablon.xlsx` | Boş yevmiye + formüllü büyük defter/mizan/gelir tablosu/bilanço — öğrenci kendi kayıtlarını girer |

Kitabın içinde yeşil **kontrol hücreleri** vardır (DENK / EŞİT / TUTUYOR). Bir tutarı yanlış yazınca
kırmızıya döner; öğrenci denetimi programa gerek kalmadan dosyanın içinde görür.

## 3. Komutlar (üretim ekibi)

Hepsi `lab/` klasöründen:

```bash
uv run --project araclar python defter/uret.py              # defter/cikti/kahve-duragi-2026-01.xlsx
uv run --project araclar python defter/uret.py --sablon     # defter/cikti/bos-sablon.xlsx
uv run --project araclar python defter/kontrol.py           # K1–K7 (çıpa dahil)
uv run --project araclar python defter/kontrol.py dosya.xlsx   # başka bir kitap (öğrenci teslimi) — çıpasız
uv run --project araclar python defter/kontrol.py --negatif # bozuk kopyalar yakalanıyor mu
uv run --project araclar python defter/goruntule.py --sun   # görüntüleyici: http://localhost:8768/
uv run --project araclar python araclar/ekranlari_uret.py 2.2   # yalnız bir dersin ekranları
```

### Denetim (`defter/kontrol.py`)

| Kod | Ne kontrol eder |
|---|---|
| K1 | her yevmiye maddesinde borç = alacak (hücre değerlerinden, formülden bağımsız) |
| K2 | yevmiye borç toplamı = alacak toplamı |
| K3 | mizan: hareket ve kalan toplamları denk; her hesap kodu hesap planında |
| K4 | bilanço: aktif = pasif |
| K5 | kitabın içindeki bütün kontrol hücreleri olumlu |
| K6 | hiçbir formül hata vermiyor (formüller pycel ile gerçekten hesaplanır) |
| K7 | **çıpa**: 22 sonuç, koddan bağımsız elle hesaplanmış `defter/veri/beklenen.yml` ile aynı |

**Neden çıpa?** Borç ve alacak yer değiştirirse (madde 10 ters kaydedilirse) K1–K6'nın hepsi yine GEÇER —
denklik doğruluk değildir. Bunu yalnız dışarıdan konmuş beklenen değerler yakalar. `--negatif` bu durumu
her çalıştırmada yeniden kanıtlar (Ders 4.3'ün ekranı).

### Elle sağlama (çıpanın kaynağı)

Bankalar (102): 300.000 − 20.000 − 108.000 + 72.000 − 30.000 − 4.800 + 48.000 − 100.000 = **157.200**
· Kasa (100): 20.000 − 9.600 = **10.400** · İndirilecek KDV (191): 18.000 + 1.600 + 30.000 + 5.000 + 800 − 20.000 = **35.400**
· SMM: 0 + 150.000 − 100.000 = **50.000** · Giderler (632): 8.000 + 25.000 + 4.000 + 1.500 = **38.500**
· Dönem kârı: 100.000 − 50.000 − 38.500 = **11.500** · Aktif: 10.400 + 157.200 + 100.000 + 35.400 + 90.000 − 1.500 = **391.500**
· Pasif: 80.000 + 300.000 + 11.500 = **391.500**.

## 4. Kaynaklı değerler ve sadeleştirmeler

Oran, had ve tarihler `defter/veri/ayarlar.yml` içinde kaynağı ve okunduğu tarihle (2026-09-25) durur;
kitabın **Bilgi** sayfasında da yazılıdır. Ayrıntı: `../planlama/KAYNAKLAR.yml`.

| Değer | Kullanım | Kaynak |
|---|---|---|
| KDV genel oranı %20 | bütün alış/satışlarda **varsayım olarak** | 2007/13033 BKK md. 1/(a) + 7346 sayılı CK (RG 7.7.2023) |
| Doğrudan gider haddi 12.000 TL (2026) | barkod okuyucu 8.000 TL → gider | VUK md. 313 + 588 SN VUK GT (RG 30.12.2025) |
| Faydalı ömür 5 yıl | tezgâh ve raf sistemi | **ÖRNEK** — amortisman listesindeki karşılığı doğrulanmadı |
| Vergi takvimi tarihleri | Takvim sayfası | GİB 2026 vergi takvimi (okundu 2026-09-25) |

Bilinçli sadeleştirmeler (derste söylenir):

- Dönem **Ocak 2026 ara dönemi**; aylık amortisman payı (yıllık / 12) yalnız yönetim tablosu içindir — vergide amortisman yıllık ayrılır (VUK md. 320).
- Giderler ticaret işletmesi için 7/B seçeneğiyle (79 Gider Çeşitleri) izlenir ve ay sonunda 632'ye aktarılır (MSUGT).
- Devreden KDV 191'de bırakıldı; bazı hesap planlarında ayrı hesapta izlenir (açık soru).
- Personel, bordro, stopaj, kurumlar vergisi karşılığı kapsam dışı; gelir tablosu **vergi öncesi**.
- Görüntüleyici bir elektronik tablo programını taklit etmez; dosyanın içeriğini gösterir. Muhasebe yazılımlarının (Paraşüt, Logo, Mikro …) ekranı lab'da yoktur ve taklit edilmez.

## 5. Klasörler

```
lab/
  kontrol.sh               tek komutluk denetim
  defter/
    veri/                  hesap-plani.csv · yevmiye.csv · stok-sayim.csv · ayarlar.yml · beklenen.yml (çıpa)
    uret.py                .xlsx üretir (formüllü)
    kontrol.py             K1–K7 + --negatif
    goruntule.py           gerçek .xlsx → goruntuleyici/*.html (değerler pycel ile hesaplanır)
    cikti/                 kahve-duragi-2026-01.xlsx · bos-sablon.xlsx
  goruntuleyici/           üretilmiş görüntüleyici sayfaları (SITE-HARITASI.md)
  araclar/                 ekran.py (1280×800) · terminal.py (740×462) · ekranlari_uret.py · pyproject + uv.lock
  ekran/<ders>/            ders ekran görüntüleri (+ terminal ham .txt)
  bitirme/                 SENARYO-SUBAT-2026.md (kurgusal bitirme senaryosu)
```
