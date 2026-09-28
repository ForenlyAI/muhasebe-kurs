"""Görüntüleyicinin (gerçek .xlsx'in içeriği) gerçek ekran görüntüsü — Playwright + Chromium.

Kullanım (lab klasöründen; önce `defter/goruntule.py` ile sayfalar üretilmiş olmalı):
    uv run --project araclar python araclar/ekran.py <ders> <ad> <sayfa> [--sec C5] [--vurgu A5:I9] [--tam]

Örnek:
    uv run --project araclar python araclar/ekran.py 2.4 mizan mizan --sec E23 --vurgu A23:F23

Python içinden:
    from ekran import sunucu, tarayici_ac, yeni_sayfa, kaydet
    with sunucu() as adres, tarayici_ac() as t:
        s = yeni_sayfa(t); s.goto(adres + "/mizan.html?sec=E23"); kaydet(s, "2.4", "mizan")

- Masaüstü: 1280×800 görüş alanı. Yazı tipleri yüklenmeden çekilmez; `--tam` tüm sayfa boyunu çeker.
- Sunucu: Python'un kendi http.server'ı, goruntuleyici/ klasörü, http://localhost:8768 (yalnız bu bilgisayar).
"""

from __future__ import annotations

import argparse
import contextlib
import functools
import http.server
import socket
import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ortak import LAB, MASAUSTU, MOBIL, ekran_yolu  # noqa: E402

PORT = 8768


def port_acik(port: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", port)) == 0


class _SessizIsleyici(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):  # noqa: D401
        pass


@contextlib.contextmanager
def sunucu(port: int = PORT):
    """goruntuleyici/ klasörünü yerelde sunar, adresini verir, çıkışta durdurur."""
    if port_acik(port):
        raise RuntimeError(f"{port} portu dolu — önce oradaki sunucuyu durdurun.")
    isleyici = functools.partial(_SessizIsleyici, directory=str(LAB / "goruntuleyici"))
    s = http.server.ThreadingHTTPServer(("127.0.0.1", port), isleyici)
    t = threading.Thread(target=s.serve_forever, daemon=True)
    t.start()
    try:
        yield f"http://localhost:{port}"
    finally:
        s.shutdown()
        s.server_close()


@contextlib.contextmanager
def tarayici_ac():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        t = p.chromium.launch()
        try:
            yield t
        finally:
            t.close()


def yeni_sayfa(tarayici, mobil: bool = False):
    if mobil:
        return tarayici.new_page(viewport=MOBIL, device_scale_factor=1, is_mobile=True, has_touch=True,
                                 locale="tr-TR", reduced_motion="reduce")
    return tarayici.new_page(viewport=MASAUSTU, locale="tr-TR", reduced_motion="reduce")


def kaydet(sayfa, ders: str, ad: str, tam: bool = False) -> Path:
    sayfa.evaluate("document.fonts.ready.then(() => true)")
    sayfa.wait_for_timeout(150)
    yol = ekran_yolu(ders, ad)
    sayfa.screenshot(path=str(yol), full_page=tam)
    print(yol.relative_to(LAB))
    return yol


def adres(taban: str, sayfa: str, sec: str | None = None, vurgu: str | None = None) -> str:
    q = "&".join(x for x in (f"sec={sec}" if sec else "", f"vurgu={vurgu}" if vurgu else "") if x)
    return f"{taban}/{sayfa}.html" + (f"?{q}" if q else "")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("ders")
    ap.add_argument("ad")
    ap.add_argument("sayfa")
    ap.add_argument("--sec")
    ap.add_argument("--vurgu")
    ap.add_argument("--tam", action="store_true")
    a = ap.parse_args()
    with sunucu() as taban, tarayici_ac() as t:
        s = yeni_sayfa(t)
        s.goto(adres(taban, a.sayfa, a.sec, a.vurgu), wait_until="networkidle")
        kaydet(s, a.ders, a.ad, a.tam)


if __name__ == "__main__":
    main()
