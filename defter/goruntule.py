"""Gerçek .xlsx dosyasını tarayıcıda gösteren yerel görüntüleyici üretir (ders ekranları için).

Kullanım (lab klasöründen):
    uv run --project araclar python defter/goruntule.py                 # goruntuleyici/ klasörünü üretir
    uv run --project araclar python defter/goruntule.py --sun           # üretir ve http://localhost:8768 adresinde sunar
    uv run --project araclar python defter/goruntule.py dosya.xlsx      # başka bir kitap

Görüntüleyici bir elektronik tablo programını taklit ETMEZ; kitabın içeriğini olduğu gibi gösterir:
hücre değerleri formüllerin pycel ile hesaplanmış sonucudur, hücreye tıklayınca (ya da ?sec=C5) formül
çubuğunda hücredeki gerçek formül görünür. Sayfa üstünde dosyanın adı ve md5 özeti yazar.

Adresler: /index.html (sayfa listesi) · /<sayfa-kimligi>.html?sec=HÜCRE&vurgu=A5:I9
Sayfa kimlikleri SITE-HARITASI.md'de.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import html
import json
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from pycel import ExcelCompiler

DEFTER = Path(__file__).resolve().parent
LAB = DEFTER.parent
KITAP = DEFTER / "cikti" / "kahve-duragi-2026-01.xlsx"
HEDEF = LAB / "goruntuleyici"
PORT = 8768

KIMLIK = {
    "Bilgi": "bilgi", "Hesap Planı": "hesap-plani", "Yevmiye": "yevmiye", "Büyük Defter": "buyuk-defter",
    "Mizan": "mizan", "KDV": "kdv", "Stok": "stok", "Amortisman": "amortisman",
    "Gelir Tablosu": "gelir-tablosu", "Bilanço": "bilanco",
    "Banka Ekstresi": "banka-ekstresi", "Takvim": "takvim",
}
OLUMLU = {"DENK", "EŞİT", "TUTUYOR", "AKTARILDI", "doğrudan gider yazılabilir"}
OLUMSUZ_ON = ("DENK DEĞİL", "FARK VAR", "EŞİT DEĞİL", "AKTARILMADI", "HESAP PLANINDA YOK")


def tl(x: float, ondalik: int = 2) -> str:
    s = f"{abs(x):,.{ondalik}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return s


def bicimle(v, fmt: str) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "DOĞRU" if v else "YANLIŞ"
    if "YYYY" in fmt.upper():
        if isinstance(v, (int, float)):
            v = from_excel(v)
        if isinstance(v, (dt.datetime, dt.date)):
            return v.strftime("%d.%m.%Y")
    if isinstance(v, (int, float)):
        if "%" in fmt:
            ondalik = 1 if "0.0%" in fmt else 0
            return "%" + tl(v * 100, ondalik)   # Türkçe yazım: %20
        if "#,##0.00" in fmt:
            if round(v, 2) == 0 and '"-"' in fmt:
                return "-"
            return f"({tl(v)})" if v < 0 else tl(v)
        if "#,##0" in fmt:
            return tl(v, 0)
        return str(int(v)) if float(v).is_integer() else tl(v)
    return str(v)


def renk(hucre) -> str | None:
    f = hucre.fill
    if f is not None and f.fill_type == "solid":
        rgb = f.fgColor.rgb
        if isinstance(rgb, str) and len(rgb) >= 6 and rgb not in ("00000000",):
            return "#" + rgb[-6:]
    return None


SABLON = """<!DOCTYPE html>
<html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{baslik} — Kahve Durağı defter görüntüleyici</title>
<style>
:root {{ --zemin:#f6f3ee; --kagit:#ffffff; --cizgi:#d9d2c5; --baslik:#2b2622; --soluk:#6f675d; --vurgu:#8a5a2b;
        --secili:#1f6feb; --ok:#1d7a3a; --ok-zemin:#e3f1e6; --hata:#b42318; --hata-zemin:#fde7e4; }}
* {{ box-sizing: border-box; }}
html, body {{ margin:0; height:100%; background:var(--zemin); color:var(--baslik);
  font: 13px/1.35 "Arimo","DejaVu Sans",sans-serif; }}
header {{ display:flex; align-items:center; gap:14px; padding:10px 16px; background:#2b2622; color:#f6f3ee; }}
header .ad {{ font-weight:700; font-size:15px; letter-spacing:.2px; }}
header .dosya {{ font: 12px "DejaVu Sans Mono", monospace; color:#d8cfc2; }}
header .rozet {{ margin-left:auto; background:#c9772b; color:#fff; font-weight:700; font-size:11px; padding:3px 8px; border-radius:3px; }}
nav {{ display:flex; gap:2px; padding:6px 12px 0; background:#e9e3d8; border-bottom:1px solid var(--cizgi); overflow-x:auto; }}
nav a {{ padding:6px 12px; color:var(--baslik); text-decoration:none; border:1px solid transparent; border-bottom:none;
  border-radius:4px 4px 0 0; white-space:nowrap; }}
nav a.aktif {{ background:var(--kagit); border-color:var(--cizgi); font-weight:700; color:var(--vurgu); }}
.formul {{ display:flex; align-items:stretch; border-bottom:1px solid var(--cizgi); background:var(--kagit); }}
.formul .ref {{ width:92px; padding:6px 10px; border-right:1px solid var(--cizgi); font: 12px "DejaVu Sans Mono", monospace; color:var(--soluk); }}
.formul .fx {{ padding:6px 8px; color:var(--soluk); font-style:italic; border-right:1px solid var(--cizgi); }}
.formul .icerik {{ flex:1; padding:6px 10px; font: 12.5px "DejaVu Sans Mono", monospace; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
main {{ height: calc(100% - 136px); overflow:auto; background:var(--kagit); }}
table {{ border-collapse:collapse; table-layout:fixed; }}
th {{ position:sticky; top:0; z-index:2; background:#efeae1; color:var(--soluk); font-weight:400; font-size:11px;
  border:1px solid var(--cizgi); padding:2px 0; }}
th.no, td.no {{ position:sticky; left:0; z-index:1; width:40px; background:#efeae1; color:var(--soluk); text-align:center; font-size:11px; border:1px solid var(--cizgi); }}
th.no {{ z-index:3; }}
td {{ border:1px solid #ece6db; padding:3px 6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; height:22px; cursor:cell; }}
td.s {{ text-align:right; font-variant-numeric: tabular-nums; }}
td.b {{ font-weight:700; }}
td.i {{ font-style:italic; color:var(--soluk); font-size:11.5px; }}
td.buyuk {{ font-size:17px; font-weight:700; }}
td.wrap {{ white-space:normal; }}
td.tasar {{ overflow:visible; position:relative; z-index:1; }}
td.ok {{ color:var(--ok); font-weight:700; background:var(--ok-zemin) !important; }}
td.hata {{ color:var(--hata); font-weight:700; background:var(--hata-zemin) !important; }}
td.f::after {{ content:""; }}
td.sec {{ outline:2px solid var(--secili); outline-offset:-2px; }}
td.vurgu {{ box-shadow: inset 0 0 0 9999px rgba(201,119,43,.14); }}
footer {{ position:fixed; bottom:0; left:0; right:0; padding:4px 16px; background:#e9e3d8; color:var(--soluk); font-size:11px; border-top:1px solid var(--cizgi); }}
</style></head>
<body>
<header><span class="ad">Kahve Durağı · defter görüntüleyici</span>
<span class="dosya">{dosya} · md5 {md5}</span><span class="rozet">KURGUSAL · EĞİTİM AMAÇLI</span></header>
<nav>{sekmeler}</nav>
<div class="formul"><div class="ref" id="ref">A1</div><div class="fx">fx</div><div class="icerik" id="icerik"></div></div>
<main id="alan"><table style="width:{tablo_genislik}px">{kolonlar}<thead><tr><th class="no"></th>{basliklar}</tr></thead><tbody>{govde}</tbody></table></main>
<footer>Değerler dosyadaki formüllerin hesaplanmış sonucudur (pycel). Hücreye tıklayın: formül yukarıda görünür. Mali müşavirlik hizmeti değildir.</footer>
<script>
const F = {formuller};
const ref = document.getElementById('ref'), icerik = document.getElementById('icerik');
function sec(td) {{
  document.querySelectorAll('td.sec').forEach(x => x.classList.remove('sec'));
  td.classList.add('sec'); ref.textContent = td.dataset.h;
  icerik.textContent = F[td.dataset.h] !== undefined ? F[td.dataset.h] : td.textContent;
}}
document.querySelectorAll('td[data-h]').forEach(td => td.addEventListener('click', () => sec(td)));
function kol(s) {{ let n = 0; for (const c of s) n = n * 26 + c.charCodeAt(0) - 64; return n; }}
const p = new URLSearchParams(location.search);
if (p.get('vurgu')) {{
  const [a, b] = p.get('vurgu').toUpperCase().split(':'); const m1 = a.match(/([A-Z]+)(\\d+)/), m2 = (b || a).match(/([A-Z]+)(\\d+)/);
  const c1 = kol(m1[1]), r1 = +m1[2], c2 = kol(m2[1]), r2 = +m2[2];
  document.querySelectorAll('td[data-h]').forEach(td => {{ const m = td.dataset.h.match(/([A-Z]+)(\\d+)/);
    const c = kol(m[1]), r = +m[2]; if (c >= c1 && c <= c2 && r >= r1 && r <= r2) td.classList.add('vurgu'); }});
}}
const hedef = document.querySelector(`td[data-h="${{(p.get('sec') || 'A1').toUpperCase()}}"]`);
if (hedef) {{ sec(hedef); }}
const ilkVurgu = document.querySelector('td.vurgu');
const kaydir = ilkVurgu || (p.get('sec') ? hedef : null);
if (kaydir && +kaydir.dataset.h.match(/\\d+/)[0] > 18) {{
  document.getElementById('alan').scrollTop = Math.max(0, kaydir.offsetTop - 72);
}}
</script>
</body></html>
"""


def uret(yol: Path) -> list[str]:
    wb = load_workbook(yol)
    x = ExcelCompiler(filename=str(yol))
    md5 = hashlib.md5(yol.read_bytes()).hexdigest()[:10]
    HEDEF.mkdir(exist_ok=True)
    adlar = [ws.title for ws in wb.worksheets]
    uretilen = []
    for ws in wb.worksheets:
        kimlik = KIMLIK.get(ws.title, ws.title.lower().replace(" ", "-"))
        sekmeler = "".join(
            f'<a href="{KIMLIK.get(a, a)}.html"{" class=\"aktif\"" if a == ws.title else ""}>{html.escape(a)}</a>' for a in adlar)
        # son dolu satır: değeri boş olmayan son satır (+ toplam satırı)
        son_satir, son_sutun, degerler, formuller = 1, 1, {}, {}
        for satir in ws.iter_rows():
            for h in satir:
                if h.value is None:
                    continue
                v = h.value
                if isinstance(v, str) and v.startswith("="):
                    formuller[h.coordinate] = v
                    v = x.evaluate(f"{ws.title}!{h.coordinate}")
                degerler[h.coordinate] = v
                if v not in (None, ""):
                    son_satir = max(son_satir, h.row)
                    son_sutun = max(son_sutun, h.column)
        genislikler = []
        for c in range(1, son_sutun + 1):
            harf = get_column_letter(c)
            w = ws.column_dimensions[harf].width if harf in ws.column_dimensions else None
            genislikler.append(int((w or 10) * 6.5 + 10))
        kolonlar = '<colgroup><col style="width:40px">' + "".join(f'<col style="width:{w}px">' for w in genislikler) + "</colgroup>"
        basliklar = "".join(f"<th>{get_column_letter(c)}</th>" for c in range(1, son_sutun + 1))
        govde = []
        for r in range(1, son_satir + 1):
            hucreler = [f'<td class="no">{r}</td>']
            for c in range(1, son_sutun + 1):
                h = ws.cell(row=r, column=c)
                v = degerler.get(h.coordinate)
                metin = bicimle(v, h.number_format or "General")
                sinif = []
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    sinif.append("s")
                if h.font is not None and h.font.b:
                    sinif.append("b")
                if h.font is not None and h.font.i:
                    sinif.append("i")
                if h.font is not None and h.font.sz and h.font.sz >= 14:
                    sinif.append("buyuk")
                if h.alignment is not None and h.alignment.wrap_text:
                    sinif.append("wrap")
                if isinstance(v, str) and metin and "s" not in sinif and not (h.alignment is not None and h.alignment.wrap_text):
                    sag = degerler.get(ws.cell(row=r, column=c + 1).coordinate) if c < son_sutun else None
                    if sag in (None, "") and h.fill.fill_type != "solid":
                        sinif.append("tasar")
                if isinstance(v, str) and v in OLUMLU:
                    sinif.append("ok")
                elif isinstance(v, str) and (v.startswith(OLUMSUZ_ON) or v.startswith("#")):
                    sinif.append("hata")
                stil = f' style="background:{renk(h)}"' if renk(h) and "ok" not in sinif and "hata" not in sinif else ""
                if h.alignment is not None and h.alignment.horizontal == "center":
                    stil = stil[:-1] + ';text-align:center"' if stil else ' style="text-align:center"'
                if h.alignment is not None and h.alignment.wrap_text:
                    stil = stil[:-1] + ';height:auto"' if stil else ' style="height:auto"'
                hucreler.append(f'<td data-h="{h.coordinate}" class="{" ".join(sinif)}"{stil}>{html.escape(metin)}</td>')
            govde.append("<tr>" + "".join(hucreler) + "</tr>")
        sayfa = SABLON.format(
            baslik=html.escape(ws.title), dosya=html.escape(yol.name), md5=md5, sekmeler=sekmeler,
            kolonlar=kolonlar, tablo_genislik=40 + sum(genislikler), basliklar=basliklar, govde="".join(govde),
            formuller=json.dumps(formuller, ensure_ascii=False).replace("</", "<\\/"))
        (HEDEF / f"{kimlik}.html").write_text(sayfa, encoding="utf-8")
        uretilen.append(f"{kimlik}.html")
    liste = "".join(f'<li><a href="{KIMLIK.get(a, a)}.html">{html.escape(a)}</a></li>' for a in adlar)
    (HEDEF / "index.html").write_text(
        f'<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8"><title>Kahve Durağı — defter görüntüleyici</title>'
        f'<meta http-equiv="refresh" content="0; url=bilgi.html"></head><body><h1>{html.escape(yol.name)}</h1>'
        f'<ul>{liste}</ul></body></html>', encoding="utf-8")
    return uretilen


def sun() -> None:
    import functools
    import http.server
    isleyici = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(HEDEF))
    with http.server.ThreadingHTTPServer(("127.0.0.1", PORT), isleyici) as s:
        print(f"Görüntüleyici: http://localhost:{PORT}/  (durdurmak için Ctrl+C)")
        try:
            s.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    arg = [a for a in sys.argv[1:] if not a.startswith("--")]
    yol = Path(arg[0]).resolve() if arg else KITAP
    sayfalar = uret(yol)
    print(f"{len(sayfalar)} sayfa → {HEDEF.relative_to(LAB)}/ ({', '.join(sayfalar)})")
    if "--sun" in sys.argv:
        sun()
