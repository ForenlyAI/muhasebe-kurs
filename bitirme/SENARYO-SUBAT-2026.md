# Bitirme senaryosu — Kahve Durağı, Şubat 2026 (KURGUSAL)

Kendi işletmenin verisini kullanmak istemeyenler için hazır senaryo. İşletme, kişiler, banka ve bütün
tutarlar **uydurmadır**. Gerçek kişi adı, TC kimlik no, IBAN, gerçek firma unvanı kullanma.

> Bu çalışma mali müşavirlik hizmeti değildir. Kendi durumunuz için mali müşavirinize danışın.

**Başlangıç:** `defter/cikti/kahve-duragi-2026-01.xlsx` dosyasının bir **kopyası**. Aynı hesap dönemi
(2026) sürdüğü için açılış maddesi yazılmaz: Şubat maddeleri Yevmiye'de 16'dan başlayarak Ocak'ın altına
eklenir. Mizan, Gelir Tablosu ve Bilanço böylece **01.01–28.02.2026 birikimli** olur (başlıkları güncelle).
KDV sayfasına Şubat bölümünü, Stok sayfasına Şubat sayımını, Amortisman sayfasına vitrin dolabını ekle.
Şubat'ın dönem başı stoku Ocak sonu sayımıdır (100.000 TL).

*Kendi işletmenle yapacaksan:* `defter/cikti/bos-sablon.xlsx` ile yeni kurulmuş (kurgusal) bir işletme gibi
sermaye maddesinden başla; gerçek tutar, gerçek müşteri ya da banka bilgisi yazma.

**Oran:** KDV için çalışma kitabındaki genel oran (Bilgi!B15, kaynağı yazılı). Bu senaryoda bütün
alış/satışlara genel oran uygulandığı **varsayılır**.

| # | Tarih | Olay (belge) |
|---|---|---|
| 1 | 03.02 | Örnek Kafe Ltd.'ye vadeli satış — matrah 30.000 TL + KDV (satış e-Faturası ÖRN-S-042) |
| 2 | 05.02 | Bireysel müşterilere satışlar — matrah 45.000 TL + KDV, kartla tahsil, bankaya geçti (e-Arşiv faturalar ÖRN-S-043…070) |
| 3 | 07.02 | Satıcıya kalan borcun tamamı (80.000 TL) bankadan ödendi (dekont ÖRN-D-005) |
| 4 | 10.02 | Satılmak üzere mal alındı — matrah 60.000 TL + KDV, vadeli (alış e-Faturası ÖRN-A-006) |
| 5 | 12.02 | Dükkân için elektrikli su ısıtıcısı — 11.000 TL + KDV, kasadan (alış faturası ÖRN-A-007). Doğrudan gider mi, amortisman mı? Haddi Bilgi!B16'dan kontrol et |
| 6 | 14.02 | Vitrin dolabı — 40.000 TL + KDV, bankadan (alış e-Faturası ÖRN-A-008). Amortisman tablosuna ekle (örnek ömür 5 yıl) |
| 7 | 20.02 | Örnek Kafe Ltd.'den 36.000 TL tahsilat (dekont ÖRN-D-006) |
| 8 | 25.02 | Şubat dükkân kirası — 25.000 TL + KDV, bankadan (alış e-Faturası ÖRN-A-009) |
| 9 | 27.02 | Elektrik ve internet — 4.500 TL + KDV, bankadan (alış e-Faturası ÖRN-A-010) |
| 10 | 28.02 | Banka ekstresinde hesap işletim ücreti 50 TL (Ocak'ta deftere işlenmeyen 50 TL de var — Banka Ekstresi sayfası) |
| 11 | 28.02 | Ay sonu sayımı: kalan malların maliyeti toplamı 85.000 TL (sayım listesi ÖRN-SAY-02) |
| 12 | 28.02 | Şubat amortisman payları: tezgâh ve raf sistemi + vitrin dolabı |
| 13 | 28.02 | Gider çeşitlerinin gelir tablosu hesabına aktarımı (7/B) |
| 14 | 28.02 | Şubat KDV mahsubu — Ocak'tan devreden KDV'yi hesaba kat |

**Takvim:** Şubat 2026 KDV beyanının son günü GİB 2026 takvimine göre 30.03.2026'dır (Takvim sayfası).

Teslim ve geçme ölçütleri: `planlama/course.yml` → `capstone` (Ç1–Ç8).
