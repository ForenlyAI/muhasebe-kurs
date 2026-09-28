#!/usr/bin/env bash
# Muhasebe 101 — lab kontrolü (Kahve Durağı, KURGUSAL).
#
#   ./kontrol.sh             ortam (uv sync) → kitabı üret → denetle (K1–K7) → negatif test → görüntüleyiciyi kur → GEÇTİ/KALDI
#   ./kontrol.sh --ekran     + bütün ders ekran görüntülerini yeniden çek (ekran/<ders>/)
#
# Hiçbir yere yayın yapmaz, dış servise bağlanmaz; yalnız bu bilgisayarda çalışır.
set -u
cd "$(dirname "$0")"
EKRAN=0
for a in "$@"; do
  case "$a" in
    --ekran) EKRAN=1 ;;
    *) echo "bilinmeyen seçenek: $a"; exit 2 ;;
  esac
done
PY=(uv run --project araclar python)
HATA=0
export LC_ALL=C.UTF-8
sonuc() { local n=$((34 - ${#1})); [ $n -lt 1 ] && n=1; printf '%s%*s%s\n' "$1" "$n" '' "$2"; }
adim() {  # adim "ad" komut...
  local ad="$1"; shift
  local cikti; cikti=$("$@" 2>&1); local kod=$?
  if [ $kod -eq 0 ]; then sonuc "$ad" "GEÇTİ"; else sonuc "$ad" "KALDI"; echo "$cikti" | tail -15; HATA=1; fi
}

echo "uv $(uv --version | cut -d' ' -f2) · python $(cd araclar && uv run python -c 'import platform;print(platform.python_version())' 2>/dev/null)"
echo
adim "ortam (uv sync --locked)"         bash -c "cd araclar && uv sync --locked -q"
adim "kitap üret (uret.py)"             "${PY[@]}" defter/uret.py
adim "boş şablon üret (--sablon)"       "${PY[@]}" defter/uret.py --sablon
echo
"${PY[@]}" defter/kontrol.py || HATA=1
echo
"${PY[@]}" defter/kontrol.py --negatif || HATA=1
echo
adim "boş şablon formül hatası yok"     bash -c "${PY[*]} defter/kontrol.py defter/cikti/bos-sablon.xlsx | grep -q 'GEÇTİ  K6'"
adim "görüntüleyici (goruntule.py)"     "${PY[@]}" defter/goruntule.py
for s in bilgi hesap-plani yevmiye buyuk-defter mizan kdv stok amortisman gelir-tablosu bilanco banka-ekstresi takvim; do
  [ -s "goruntuleyici/$s.html" ] || { sonuc "görüntüleyici $s.html" "KALDI (yok)"; HATA=1; }
done
if [ "$EKRAN" = 1 ]; then
  echo
  "${PY[@]}" araclar/ekranlari_uret.py || HATA=1
fi
echo
if [ "$HATA" = 0 ]; then echo "SONUÇ: GEÇTİ"; else echo "SONUÇ: KALDI"; fi
exit $HATA
