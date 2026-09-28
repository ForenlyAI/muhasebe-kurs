# Lab sözleşmesi — dosyalar, sayfalar, hücre adresleri (Muhasebe 101)

Bu dosya ders metinleriyle lab arasındaki sözleşmedir: bir ders "Mizan sayfası, E23" diyorsa o hücre
burada yazdığı gibi vardır. Adres değişirse önce burası, sonra `defter/veri/beklenen.yml` ve
`araclar/ekranlari_uret.py` güncellenir. Son doğrulama: `./kontrol.sh --ekran` GEÇTİ (2026-09-25).

## Dosyalar

| Yol | Tür | Kim kullanır |
|---|---|---|
| `defter/cikti/kahve-duragi-2026-01.xlsx` | gerçek .xlsx, 12 sayfa, formüllü | öğrenci (indirir), ekranlar |
| `defter/cikti/bos-sablon.xlsx` | gerçek .xlsx: Bilgi, Hesap Planı, boş Yevmiye (5–200. satır), Mizan, Gelir Tablosu, Bilanço | öğrenci (1.2, 2.4, bitirme) |
| `defter/veri/*.csv`, `ayarlar.yml` | kitabın girdisi | üretim |
| `defter/veri/beklenen.yml` | çıpa (elle hesap) — 22 değer | `kontrol.py` K7 |
| `bitirme/SENARYO-SUBAT-2026.md` | kurgusal bitirme senaryosu | öğrenci (4.4) |

## Kitabın sayfaları ve görüntüleyici adresleri

Görüntüleyici: `uv run --project araclar python defter/goruntule.py --sun` → `http://localhost:8768/<kimlik>.html`
Parametreler: `?sec=E23` (hücreyi seçer, formülü üstte gösterir) · `&vurgu=A23:F23` (aralığı boyar, kaydırır).

| Sayfa | Kimlik | Önemli hücreler |
|---|---|---|
| Bilgi | `bilgi` | B15 KDV genel oranı (%20, kaynaklı) · B16 doğrudan gider haddi (12.000) · B17 örnek faydalı ömür (5) · C15–C17 kaynak metinleri |
| Hesap Planı | `hesap-plani` | A5:E22 — 18 hesap (kod, ad, grup, tür, normal bakiye) |
| Yevmiye | `yevmiye` | A5:I41 — 15 madde (madde, tarih, belge, açıklama, hesap, hesap adı, borç, alacak, kontrol) · I her maddenin ilk satırı: DENK · G201/H201 toplam (1.030.400) · I201 DENK |
| Büyük Defter | `buyuk-defter` | hesap blokları: 100 Kasa A4 · 102 Bankalar A10 (bakiye F19 = 157.200) · 120 A22 · 153 A28 · 191 A34 · 255 A44 · 257 A49 · 320 A54 · 391 A60 · 500 A67 · 600 A72 · 621 A78 · 632 A83 · 793 A88 · 794 A94 · 796 A99 · 798 A104 |
| Mizan | `mizan` | satır 5–22 hesaplar · C23/D23 hareket toplamı · E23/F23 kalan toplamı (520.000) · E26–E29 kontroller |
| Banka Ekstresi | `banka-ekstresi` | A5:E13 ekstre (kurgusal banka) · E16 ekstre bakiyesi 157.150 · E17 defter 157.200 · E18 fark (50) |
| KDV | `kdv` | A5:H11 belge bazında KDV · D14 hesaplanan 20.000 · D15 indirilecek 55.400 · D17 ödenecek 0 · D18 devreden 35.400 · D22:D28 ÖRNEK ay (D27 ödenecek 600) · D31:D33 KDV dahil 1.200 → matrah 1.000 |
| Stok | `stok` | A5:E8 sayım · E9 dönem sonu stok 100.000 · E12:E15 SMM hesabı (E15 = 50.000) · E16–E17 kontroller |
| Amortisman | `amortisman` | B5 maliyet 90.000 · B7 ömür (ÖRNEK) · B8 %20 · B9 %40 · A12:G16 normal ve azalan bakiyeler tablosu · B20 Ocak payı 1.500 · A26:C27 doğrudan gider testi |
| Gelir Tablosu | `gelir-tablosu` | B4 brüt satış 100.000 · B7 SMM 50.000 · B8 brüt kâr · B9 genel yönetim 38.500 · B11 dönem kârı 11.500 · B15 kontrol |
| Bilanço | `bilanco` | A5:B15 aktif · D5:E13 pasif · B17 = E17 = 391.500 · B19 DENK |
| Takvim | `takvim` | A5:D12 — GİB 2026 takviminden sekiz son gün (kaynaklı) |

## Ekranlar (lab/ekran/<ders>/) — `araclar/ekranlari_uret.py`

| Ders | Dosya | Ne gösterir |
|---|---|---|
| 1.1 | `bilanco-denklik.png` | Bilanço, aktif toplamı = pasif toplamı, B19 DENK |
| 1.2 | `yevmiye-ilk-maddeler.png` | Yevmiye madde 1–3 vurgulu, I5 formülü üstte |
| 1.2 | `kontrol.png` + `.txt` | `defter/kontrol.py` gerçek çıktısı (K1–K7 GEÇTİ) |
| 1.3 | `hesap-plani.png` | Hesap Planı sayfası |
| 1.3 | `buyuk-defter-kasa.png` | Büyük Defter, 100 Kasa bloğu |
| 1.4 | `belge-sutunu.png` | Yevmiye belge sütunu (C) vurgulu |
| 2.1 | `satis-belgeleri.png` | madde 6 (e-Arşiv) ve 7 (e-Fatura) satırları |
| 2.2 | `kdv-ozet.png` | KDV dönem özeti, devreden 35.400 |
| 2.2 | `kdv-ornek-ay.png` | ÖRNEK ay: devreden KDV'nin kullanılması |
| 2.3 | `buyuk-defter-banka.png` | 102 Bankalar bloğu, bakiye 157.200 |
| 2.3 | `banka-mutabakat.png` | ekstre ↔ defter mutabakatı, fark (50) |
| 2.4 | `mizan.png` | Mizan toplamları ve kontroller |
| 3.1 | `stok-smm.png` | sayım → SMM 50.000 |
| 3.2 | `amortisman-tablosu.png` | normal ve azalan bakiyeler, 5 yıl |
| 3.2 | `dogrudan-gider-haddi.png` | 12.000 TL haddi testi |
| 3.3 | `gelir-tablosu.png` | dönem kârı 11.500 |
| 3.4 | `bilanco.png` | bilanço, dönem kârı pasifte |
| 4.1 | `vergi-takvimi.png` | Takvim sayfası (GİB 2026) |
| 4.3 | `negatif-test.png` + `.txt` | `kontrol.py --negatif`: ters kayıt yalnız çıpayla yakalanır |
| 4.4 | `bilgi-sayfasi.png` | parametreler ve kaynakları |

Henüz ekranı olmayan dersler: 4.2 (mali müşavirle çalışmak — teslim listesi dersin kendi metni; ŞEMA).
Üçüncü taraf ürün ekranı (GİB portalı, muhasebe yazılımları) çekilmez ve taklit edilmez; gerekiyorsa resmi
sayfanın bağlantısı verilir.

## Portlar

8768 (görüntüleyici). 8765 bu makinede başka bir süreçte dolu olduğu için kullanılmadı.
