# -*- coding: utf-8 -*-
"""
LİSTECİ v3 — motor (arayüzden bağımsız iş mantığı)

  • matris_oku()      : sınav görev matrisinden dersleri ve salon/gözetmen görevlerini çıkarır
  • liste_oku()       : şube/öğrenci listesini (xlsx / xls / csv) okur, başlığı kendisi bulur
  • eslestir()        : dosya adını matristeki derslere eşleştirir (hafıza → hazırlık → benzerlik)
  • plan_yap()        : öğrencileri salonlara dağıtır, sığmayanları ayrıca bildirir
  • ders_excel()      : bir dersin şablona yazılmış Excel listesini üretir
  • paket_olustur()   : tüm dersler + RAPOR.xlsx → ZIP
"""
import difflib
import io
import os
import re
import zipfile
from collections import OrderedDict, defaultdict

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SURUM = "3.0"

# ─────────────────────────────────────────────────────────────
# Şablon yerleşimi (sablon.xlsx)
# ─────────────────────────────────────────────────────────────
ILK_SATIR = 7          # öğrenci listesinin başladığı satır
NORMAL_IMZA = 45       # Sayfa1 (normal salon) imza satırı
BUYUK_IMZA = 75        # Sayfa2 (büyük salon) imza satırı
NORMAL_KAP_MAX = NORMAL_IMZA - ILK_SATIR   # 38 — bundan fazlası imza alanına taşar
BUYUK_KAP_MAX = BUYUK_IMZA - ILK_SATIR     # 68

VARSAYILAN_AYAR = {
    "normal_kap": 34,
    "buyuk_salonlar": ["205", "305"],
    "buyuk_kap": 68,
}

HAZIRLIK_SORUMLU = "HAZIRLIK KOORDİNATÖRLÜĞÜ"

# ─────────────────────────────────────────────────────────────
# Metin yardımcıları
# ─────────────────────────────────────────────────────────────
TR_MAP = str.maketrans("çğıöşüîâûÇĞİÖŞÜÎÂÛ", "CGIOSUIAUCGIOSUIAU")
TR_ALFABE = "ABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZ"


def bos_mu(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == "" or str(v).strip().lower() in ("nan", "none", "nat")


def ascii_buyuk(metin):
    """Türkçe harfleri sadeleştirip büyük harfe çevirir (karşılaştırma için)."""
    return str(metin).translate(TR_MAP).upper()


def sade(metin):
    """Sadece A-Z ve 0-9 kalır."""
    return re.sub(r"[^A-Z0-9]", "", ascii_buyuk(metin))


def tr_buyuk(metin):
    """Türkçe kurallarıyla büyük harf: i→İ, ı→I (Python'un upper()'ı 'Ali'yi 'ALI' yapar)."""
    s = str(metin).replace("i", "İ").replace("ı", "I")
    return s.upper()


def tr_sira(metin):
    """Türkçe alfabe sırasına göre sıralama anahtarı (Ç, Ğ, İ, Ö, Ş, Ü doğru yerde)."""
    s = tr_buyuk(str(metin).strip()).translate(str.maketrans("ÂÎÛ", "AİU"))
    return "".join(chr(300 + TR_ALFABE.index(c)) if c in TR_ALFABE else c for c in s)


def dosya_govdesi(ad):
    govde, uz = os.path.splitext(str(ad))
    return govde if uz.lower() in (".xlsx", ".xls", ".csv", ".xlsm") else str(ad)


# ─────────────────────────────────────────────────────────────
# Ad eşleştirme
# ─────────────────────────────────────────────────────────────
ROMA = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}

DOLGU = {
    "VE", "ILE", "DERS", "DERSI", "DERSLERI", "LISTE", "LISTESI", "LISTELERI", "OGRENCI", "OGRENCILER",
    "OGRENCILERI", "OGR", "SUBE", "SUBESI", "SINIF", "SINIFI", "SINAV", "SINAVI", "FINAL", "VIZE",
    "BUTUNLEME", "BUT", "GRUP", "GR", "XLSX", "XLS", "CSV", "KOPYA", "KOPYASI", "YENI", "SON", "GUNCEL",
    "DONEM", "GUZ", "BAHAR", "YARIYIL", "OGRETIM", "NORMAL", "IKINCI", "IO", "NO",
}


def tokenler(metin):
    """Ad → (anlamlı kelimeler, sayılar). Roma rakamları sayıya çevrilir: 'II' = '2'."""
    t = ascii_buyuk(dosya_govdesi(metin))
    t = re.sub(r"[’‘'`´ʼʻ′‛\u2019\u2018\u02bc\u02bb\u00b4]", "", t)   # Kur'an / Kur’an / Kur‘an → KURAN
    t = re.sub(r"([A-Z])-I(?=\s+[A-Z]{2})", r"\1", t)   # Kur'an-ı Kerim, Siyer-i Nebi: izafet, rakam değil
    t = re.sub(r"(\d+)\s*\.", r"\1 ", t)
    kelimeler, sayilar = [], set()
    for tok in re.findall(r"[A-Z]+|\d+", t):
        if tok.isdigit():
            if len(tok) <= 2:
                sayilar.add(int(tok))
            continue                               # yıl, ders kodu vb. uzun sayılar atlanır
        if tok in ROMA:
            sayilar.add(ROMA[tok])
            continue
        if tok in DOLGU or len(tok) == 1:
            continue
        kelimeler.append(tok)
    return kelimeler, sayilar


def _kelime_benzer(a, b):
    if a == b:
        return 1.0
    if min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a)):
        return 0.9                                 # kısaltma: TEF ~ TEFSIR, MUSIKI ~ MUSIKISI
    r = difflib.SequenceMatcher(None, a, b).ratio()
    return r if r >= 0.8 else 0.0                  # yazım hatası toleransı


def ad_skoru(dosya_adi, ders_parcasi):
    """0–1 arası benzerlik. Rakamlar farklıysa (I ↔ II) doğrudan 0."""
    fw, fn = tokenler(dosya_adi)
    cw, cn = tokenler(ders_parcasi)
    if not fw or not cw:
        return 0.0
    if fn and cn and fn != cn:
        return 0.0
    recall = sum(max(_kelime_benzer(c, f) for f in fw) for c in cw) / len(cw)
    prec = sum(max(_kelime_benzer(f, c) for c in cw) for f in fw) / len(fw)
    if recall == 0 or prec == 0:
        return 0.0
    s = 2 * recall * prec / (recall + prec)
    if recall >= 0.95:                             # dersin bütün kelimeleri dosya adında var
        s = max(s, 0.85 + 0.15 * prec)
    if bool(fn) != bool(cn):                       # birinde 'I' var, ötekinde yok
        s *= 0.9
    return round(min(s, 1.0), 3)


def ders_parcalari(ders_adi):
    """Birleşik sütunlar: 'Tefsir I - Hadis I' gibi. Tam ad da parça sayılır."""
    parcalar = [ders_adi]
    for p in re.split(r"\s[-–/]\s|\n|/", str(ders_adi)):
        p = p.strip()
        if p and p != ders_adi and tokenler(p)[0]:
            parcalar.append(p)
    return parcalar


# ─────────────────────────────────────────────────────────────
# Hazırlık
# ─────────────────────────────────────────────────────────────
def hazirlik_mi(metin):
    """'Hazırlık A', 'HZ-B', 'HZC' → evet.   'Hz. Peygamber'in Hayatı' → hayır."""
    toks = re.findall(r"[A-Z0-9]+", ascii_buyuk(dosya_govdesi(metin)))
    for i, tok in enumerate(toks):
        if tok.startswith("HAZIRLIK"):
            return True
        if re.fullmatch(r"HZ[A-F]?", tok):
            sonraki = toks[i + 1] if i + 1 < len(toks) else ""
            if tok == "HZ" and len(sonraki) > 1 and not sonraki.isdigit() and sonraki not in DOLGU:
                continue                            # 'Hz. Muhammed' — hazırlık değil
            return True
    return False


# Matriste adında "hazırlık" geçmeyen ama Arapça hazırlık sınıfına ait dersler.
# Arayüzün kenar çubuğundan değiştirilebilir.
VARSAYILAN_HZ_DERSLERI = [
    "Sarf-Nahiv", "Sarf", "Nahiv", "Okuma Anlama", "Okuma-Anlama", "Sözlü Anlatım",
    "Yazılı Anlatım", "Dinleme Anlama", "Dinleme-Anlama", "Konuşma", "Arapça Hazırlık",
]


def hazirlik_dersi_mi(ders_adi, hz_dersleri=None):
    """Matristeki ders bir hazırlık sınavı mı? ('Hazırlık' geçiyorsa ya da hazırlık ders adlarından biriyse)"""
    if hazirlik_mi(ders_adi):
        return True
    d = sade(ders_adi)
    if "KURAN" in d:
        return False
    for k in (VARSAYILAN_HZ_DERSLERI if hz_dersleri is None else hz_dersleri):
        k = sade(k)
        if not k:
            continue
        if d == k:
            return True
        if len(k) >= 8 and k in d:      # 'SARFNAHIV A', 'OKUMAANLAMA (HZ)' gibi
            return True
    return False


def hazirlik_sube_harfi(metin):
    t = ascii_buyuk(dosya_govdesi(metin))
    m = re.search(r"(?:HAZIRLIK|HZ)[\s\-_\.]*([A-F])(?![A-Z])", t)
    if m:
        return m.group(1)
    m = re.search(r"(?<![A-Z])([A-F])[\s\-_\.]*SUBE", t)
    if m:
        return m.group(1)
    harfler = re.findall(r"(?<![A-Z])([A-F])(?![A-Z])", t)
    return harfler[-1] if harfler else None


def hazirlik_eslesir(ders_adi, dosya_adi):
    h1, h2 = hazirlik_sube_harfi(ders_adi), hazirlik_sube_harfi(dosya_adi)
    return (h1 == h2) if (h1 and h2) else True


# ─────────────────────────────────────────────────────────────
# Hoca adları
# ─────────────────────────────────────────────────────────────
def hoca_anahtari(isim):
    return sade(isim)


def benzersiz_hocalar(seri):
    gorulen, sonuc = set(), []
    for h in seri:
        if bos_mu(h) or str(h).strip() == "Belirtilmedi":
            continue
        h_str = str(h).strip()
        a = hoca_anahtari(h_str)
        if a and a not in gorulen:
            gorulen.add(a)
            sonuc.append(h_str)
    return sonuc


# ─────────────────────────────────────────────────────────────
# Eşleştirme: 1) hafıza  2) ders kodu  3) hazırlık  4) ders adı (birebir)
# ─────────────────────────────────────────────────────────────
KOD_RE = re.compile(r"(?<![A-Z0-9])([A-Z]{2,5})[\s\-_.]?(\d{3,4})(?![0-9])")


def kodlar(metin):
    """'İLH101 Tefsir I', 'ILH 101', 'ilh-101' → {'ILH101'}. Yıl gibi sayılar (2025) kod sayılmaz."""
    t = ascii_buyuk(dosya_govdesi(metin))
    sonuc = set()
    for harf, sayi in KOD_RE.findall(t):
        if len(sayi) == 4 and sayi[:2] in ("19", "20"):
            continue
        if harf in DOLGU or harf in ROMA:
            continue
        sonuc.add(harf + sayi)
    return sonuc


OGRETIM_RE = re.compile(r"(?<![A-Z])(I|N)\s*\.?\s*O(?![A-Z])\.?")


def ogretim(metin):
    """'(İ.Ö)', 'İÖ', 'İkinci Öğretim' → 'IO';  '(N.Ö)', 'NÖ', 'Normal Öğretim' → 'NO';  yoksa None."""
    t = ascii_buyuk(dosya_govdesi(metin))
    if "IKINCI OGRETIM" in t:
        return "IO"
    if "NORMAL OGRETIM" in t:
        return "NO"
    m_ = OGRETIM_RE.search(t)
    return (m_.group(1) + "O") if m_ else None


def kodsuz(metin):
    """Ad karşılaştırması için temizlenmiş metin: kod, İ.Ö/N.Ö etiketi ve parantezden sonrası atılır.
    'İİF106 İSLAM İBADET ESASLARI (N.Ö) 2' → 'ISLAM IBADET ESASLARI'"""
    t = ascii_buyuk(dosya_govdesi(metin))
    t = KOD_RE.sub(" ", t)
    t = re.sub(r"IKINCI OGRETIM|NORMAL OGRETIM", " ", t)
    t = OGRETIM_RE.sub(" ", t)
    t = re.sub(r"\(.*$", " ", t, flags=re.S)      # parantez ve sonrası: (İ.Ö), (A Şubesi), liste no
    t = re.sub(r"[\s_\-]+\d{1,2}\s*$", " ", t) if "(" in ascii_buyuk(dosya_govdesi(metin)) else t
    return t


def ad_anahtari(metin):
    """Karşılaştırma anahtarı: kod, uzantı, dolgu kelimeler, Türkçe harf farkı atılır; I/II = 1/2.
    'İLH101 - Tefsir I', 'tefsir 1.xlsx', 'Tefsir I Öğrenci Listesi' → aynı anahtar."""
    kel, say = tokenler(kodsuz(metin))
    if not kel:
        return None
    return " ".join(sorted(kel)) + "#" + ",".join(str(x) for x in sorted(say))


def ad_parcala(metin):
    """Ad → (kelimeler, sayılar). Kelime = (metin, kısaltma_mı).
    'İ.ARAŞTIRMALARINDA USÜL' → [('I', kısaltma), ('ARASTIRMALARINDA',), ('USUL',)]
    Noktalı harf(ler)den sonra kelime geliyorsa kısaltmadır ('İ.', 'TEF.'), Roma rakamı değil."""
    t = kodsuz(metin)
    t = re.sub(r"[’‘'`´ʼʻ′‛]", "", t)
    t = re.sub(r"([A-Z])-I(?=\s+[A-Z]{2})", r"\1", t)          # Kur'an-ı Kerim, Ehl-i Kitap
    kelimeler, sayilar = [], set()
    for m_ in re.finditer(r"([A-Z]+)(\s*\.)?(?=(\s*)([A-Z]?))|(\d+)", t):
        if m_.group(5):
            if len(m_.group(5)) <= 2:
                sayilar.add(int(m_.group(5)))
            continue
        tok, nokta, sonraki = m_.group(1), m_.group(2), m_.group(4)
        if nokta and sonraki:                      # 'İ.ARAŞTIRMA', 'TEF. METİNLERİ'
            kelimeler.append((tok, True))
            continue
        if tok in ROMA:
            sayilar.add(ROMA[tok])
            continue
        if tok in DOLGU or len(tok) == 1:
            continue
        kelimeler.append((tok, False))
    return kelimeler, sayilar


def ad_ayni_mi(dosya_adi, ders_parcasi):
    """Kelime kelime aynı mı? Kısaltmalı kelime, uzun hâlinin başıyla tutarsa aynı sayılır."""
    fw, fn = ad_parcala(dosya_adi)
    cw, cn = ad_parcala(ders_parcasi)
    if not fw or not cw or len(fw) != len(cw):
        return False
    if fn != cn and not (fn and not cn):           # dosyadaki tek sayı liste numarası olabilir ('Siyer 2')
        return False
    for (f, fk), (c, ck) in zip(fw, cw):
        if f == c:
            continue
        if ck and f.startswith(c):
            continue
        if fk and c.startswith(f):
            continue
        return False
    return True


def ders_ad_anahtarlari(ders_adi):
    """Bir matris hücresindeki bütün ad anahtarları (birleşik sınavlar: 'Tefsir I - Hadis I', alt alta yazılanlar)."""
    anahtarlar = set()
    for p in ders_parcalari(ders_adi):
        a = ad_anahtari(p)
        if a:
            anahtarlar.add(a)
    return anahtarlar


def eslestir(dosya_adi, dersler, hafiza=None, hz_dersleri=None):
    """
    Dönüş: {"idler": [...], "durum": "otomatik"|"hafiza"|"kontrol"|"yok", "aciklama": str, "skor": float}
    Tahmin yürütmez: kod ya da ad birebir tutmuyorsa boş bırakır, en yakın dersi yalnızca açıklamada söyler.
    """
    hafiza = hafiza or {}
    hz = lambda ad: hazirlik_dersi_mi(ad, hz_dersleri)
    anahtar = sade(dosya_govdesi(dosya_adi))

    # 1) Daha önce elle yapılmış eşleştirme
    if anahtar in hafiza:
        hedef = set(hafiza[anahtar])
        idler = [d["id"] for d in dersler if sade(d["ders_adi"]) in hedef]
        if idler:
            return {"idler": idler, "durum": "hafiza", "aciklama": "Önceki elle eşleştirmeden hatırlandı", "skor": 1.0}

    # 2) Ders kodu
    f_kod = kodlar(dosya_adi)
    if f_kod:
        idler = [d["id"] for d in dersler if kodlar(d["ders_adi"]) & f_kod]
        if idler:
            return {"idler": idler, "durum": "otomatik",
                    "aciklama": f"Ders kodu {', '.join(sorted(f_kod))}" + (f" → {len(idler)} sınav" if len(idler) > 1 else ""),
                    "skor": 1.0}

    # 3) Hazırlık: şube harfi tutan bütün hazırlık sınavları
    if hazirlik_mi(dosya_adi):
        hz_dersler = [d for d in dersler if hz(d["ders_adi"])]
        idler = [d["id"] for d in hz_dersler if hazirlik_eslesir(d["ders_adi"], dosya_adi)]
        if idler:
            harf = hazirlik_sube_harfi(dosya_adi)
            harfli_ders_var = any(hazirlik_sube_harfi(d["ders_adi"]) for d in hz_dersler)
            ack = f"Hazırlık {harf} şubesi → {len(idler)} hazırlık sınavı" if harf else f"Hazırlık listesi → {len(idler)} hazırlık sınavı"
            belirsiz = (not harf) and harfli_ders_var   # dosyada şube yok ama matriste şubeye göre ayrılmış sınav var
            return {"idler": idler, "durum": "kontrol" if belirsiz else "otomatik", "aciklama": ack, "skor": 1.0}
        return {"idler": [], "durum": "yok", "aciklama": "Hazırlık listesi ama matriste uygun hazırlık sınavı yok", "skor": 0.0}

    # 4) Ders adı (kod ve parantezden sonrası atılmış hâliyle, kelime kelime; kısaltmalar tolere edilir)
    idler = [d["id"] for d in dersler if any(ad_ayni_mi(dosya_adi, p) for p in ders_parcalari(d["ders_adi"]))]
    if idler:
        return {"idler": idler, "durum": "otomatik",
                "aciklama": "Ders adı aynı" + (f" → {len(idler)} sınav" if len(idler) > 1 else ""),
                "skor": 1.0}

    # Eşleşmedi: en yakın dersi yalnızca bilgi olarak göster (seçmez)
    en_iyi, d1 = 0.0, None
    for d in dersler:
        if hz(d["ders_adi"]):
            continue
        s = max(ad_skoru(kodsuz(dosya_adi), kodsuz(p)) for p in ders_parcalari(d["ders_adi"]))
        if s > en_iyi:
            en_iyi, d1 = s, d
    ack = "Kod ya da ad tutmadı"
    if d1 is not None and en_iyi >= 0.6:
        ack += f" — en yakın: {d1['ders_adi']}"
    return {"idler": [], "durum": "yok", "aciklama": ack, "skor": en_iyi}


# ─────────────────────────────────────────────────────────────
# Matris
# ─────────────────────────────────────────────────────────────
def _hucre_metni(v):
    if bos_mu(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _kirmizi_mi(hucre):
    try:
        f = hucre.fill
        if f is None or f.fill_type is None:
            return False
        renk = f"{f.start_color.index} {f.fgColor.rgb}".upper()
        return "FF0000" in renk
    except Exception:
        return False


def matris_oku(veri):
    """Dönüş: (dersler, uyarılar)"""
    wb = load_workbook(io.BytesIO(veri), data_only=True)
    ws = wb.active
    uyarilar = []

    birlesik = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= 4:
            for r in range(rng.min_row, min(rng.max_row, 4) + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    birlesik[(r, c)] = (rng.min_row, rng.min_col)

    def deger(r, c):
        v = ws.cell(row=r, column=c).value
        if v is None and (r, c) in birlesik:
            v = ws.cell(*birlesik[(r, c)]).value
        return v

    sinir = ws.max_column
    for c in range(1, ws.max_column + 1):
        if any("GOZETMEN" in sade(ws.cell(row=r, column=c).value or "") for r in (1, 2, 3, 4)):
            sinir = c - 1
            break

    # Gözetmen satırları (D sütunu), kırmızı satıra / 'Sınava girecek öğrenci' satırına kadar
    gozetmen_satirlari = []
    for r in range(6, ws.max_row + 1):
        hedef = ws.cell(row=r, column=4)
        if _kirmizi_mi(hedef) or "SINAVAGIRECEKOGRENCI" in sade(hedef.value or ""):
            break
        gozetmen_satirlari.append((r, _hucre_metni(hedef.value)))

    dersler = []
    for c in range(5, sinir + 1):
        ders_val = deger(4, c)
        if bos_mu(ders_val):
            continue
        tarih_val, saat_val = deger(1, c), deger(3, c)

        if hasattr(tarih_val, "strftime"):
            tarih = tarih_val.strftime("%d.%m.%Y")
        else:
            tarih = str(tarih_val).split(" ")[0] if not bos_mu(tarih_val) else "Tarih_Yok"
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", tarih):
                y, a, g = tarih.split("-")
                tarih = f"{g}.{a}.{y}"

        if hasattr(saat_val, "strftime"):
            saat = saat_val.strftime("%H:%M")
        else:
            saat = _hucre_metni(saat_val)
            saat = saat[:5] if ":" in saat else saat

        salonlar = OrderedDict()
        for r, goz in gozetmen_satirlari:
            s = _hucre_metni(ws.cell(row=r, column=c).value)
            if s and len(s) < 10:
                salonlar.setdefault(s, [])
                if goz and goz not in salonlar[s]:
                    salonlar[s].append(goz)

        ders = {
            "id": len(dersler),
            "tarih": tarih,
            "saat": saat,
            "ders_adi": str(ders_val).strip(),
            "hucre": ws.cell(row=4, column=c).coordinate,
            "salonlar": [{"salon": s, "gozetmenler": g} for s, g in salonlar.items()],
        }
        if not ders["salonlar"]:
            uyarilar.append(f"{ders['ders_adi']} ({tarih} {saat}): matriste salon/gözetmen görevi bulunamadı.")
        dersler.append(ders)
    return dersler, uyarilar


# ─────────────────────────────────────────────────────────────
# Öğrenci listeleri
# ─────────────────────────────────────────────────────────────
def _baslik_turu(v):
    h = sade(v)
    if not h:
        return None
    if h in ("SIRA", "SIRANO", "SN", "SNO", "SIRANUMARASI", "SIRANUMARA"):
        return "sira"
    if h in ("SOYAD", "SOYADI", "SOYISIM", "SOYISMI"):
        return "soyad"
    if "SOYAD" in h or "SOYISIM" in h:
        return "adsoyad"
    if "NUMARA" in h or h in ("NO", "OGRENCINO", "OGRNO", "OKULNO", "OGRENCINO"):
        return "no"
    if h in ("AD", "ADI", "ISIM", "ISMI", "OGRENCIADI", "OGRENCIAD", "ADLARI"):
        return "ad"
    if any(k in h for k in ("SORUMLU", "OGRETIMELEMAN", "OGRETIMUYE", "HOCA", "DERSVEREN", "OGRETMEN", "AKADEMISYEN")):
        return "sorumlu"
    return None


def no_temizle(v):
    if bos_mu(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = re.sub(r"\s+", "", str(v))
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def _no_gibi(v):
    return bool(re.fullmatch(r"[A-Za-z]{0,3}\d{5,}", no_temizle(v)))


def _tablodan_ogrenciler(ham):
    """Ham tablodan (header=None) öğrenci satırlarını çıkarır. Dönüş: (df, notlar) ya da (None, neden)."""
    notlar = []
    kolon = {}
    baslik_satiri = None
    for r in range(min(25, len(ham))):
        turler = {}
        for c, v in enumerate(ham.iloc[r].tolist()):
            t = _baslik_turu(v)
            if t and t not in turler:
                turler[t] = c
        if "no" in turler and ("ad" in turler or "adsoyad" in turler):
            kolon, baslik_satiri = turler, r
            break

    if baslik_satiri is None:
        # Başlık yok: öğrenci numarasına benzeyen ilk hücreden konumla tahmin et
        for r in range(min(40, len(ham))):
            satir = ham.iloc[r].tolist()
            c = next((i for i, v in enumerate(satir) if _no_gibi(v)), None)
            if c is not None:
                kolon = {"no": c, "ad": c + 1, "soyad": c + 2}
                if c + 3 < ham.shape[1]:
                    kolon["sorumlu"] = c + 3
                baslik_satiri = r - 1
                notlar.append("Başlık satırı bulunamadı; sütunlar konumdan tahmin edildi (No, Ad, Soyad, Sorumlu).")
                break
    if baslik_satiri is None:
        return None, "Öğrenci numarası içeren satır bulunamadı."

    veri = ham.iloc[baslik_satiri + 1:].reset_index(drop=True)
    ncol = veri.shape[1]

    def sutun(t):
        c = kolon.get(t)
        return veri.iloc[:, c] if c is not None and c < ncol else pd.Series([None] * len(veri))

    no = sutun("no").map(no_temizle)
    if "adsoyad" in kolon and "soyad" not in kolon:
        tam = sutun("adsoyad").map(lambda v: "" if bos_mu(v) else str(v).strip())
        ad = tam.map(lambda s: " ".join(s.split()[:-1]) if len(s.split()) > 1 else s)
        soyad = tam.map(lambda s: s.split()[-1] if len(s.split()) > 1 else "")
    else:
        ad = sutun("ad" if "ad" in kolon else "adsoyad").map(lambda v: "" if bos_mu(v) else str(v).strip())
        soyad = sutun("soyad").map(lambda v: "" if bos_mu(v) else str(v).strip())

    if "sorumlu" in kolon:
        sor = (sutun("sorumlu").map(lambda v: None if bos_mu(v) else " ".join(str(v).split()))
               .ffill().bfill().fillna("Belirtilmedi"))
    else:
        sor = pd.Series(["Belirtilmedi"] * len(veri))
        notlar.append("Ders sorumlusu sütunu yok; imza alanına 'Belirtilmedi' yazılacak.")

    df = pd.DataFrame({"No": no, "Ad": ad, "Soyad": soyad, "Sorumlu": sor})
    gecerli = df["No"].map(lambda s: bool(re.fullmatch(r"[A-Za-z]{0,3}\d{4,}", s))) & ((df["Ad"] != "") | (df["Soyad"] != ""))
    atlanan = int(((df["No"] != "") & ~gecerli).sum())
    df = df[gecerli].reset_index(drop=True)
    if atlanan:
        notlar.append(f"{atlanan} satır geçerli öğrenci numarası olmadığı için atlandı.")
    if df.empty:
        return None, "Tabloda geçerli öğrenci satırı yok."
    return df, notlar


def liste_oku(veri, ad):
    """Dönüş: {"df": DataFrame|None, "notlar": [...], "hata": str|None, "hazirlik": bool}"""
    sonuc = {"df": None, "notlar": [], "hata": None, "hazirlik": hazirlik_mi(ad)}
    uz = os.path.splitext(ad)[1].lower()
    try:
        if uz == ".csv":
            ham = None
            for enc in ("utf-8-sig", "cp1254", "latin-1"):
                try:
                    ham = pd.read_csv(io.BytesIO(veri), header=None, sep=None, engine="python", encoding=enc, dtype=object)
                    break
                except UnicodeDecodeError:
                    continue
            sayfalar = {"csv": ham}
        else:
            sayfalar = pd.read_excel(io.BytesIO(veri), header=None, dtype=object, sheet_name=None)
    except ImportError:
        sonuc["hata"] = ".xls okunamadı: sunucuda 'xlrd' yok. requirements.txt'ye xlrd ekleyin ya da dosyayı .xlsx kaydedin."
        return sonuc
    except Exception as e:
        neden = str(e)
        if "format cannot be determined" in neden or "zip file" in neden.lower() or "Unsupported format" in neden:
            neden = "Dosya Excel olarak açılamadı (bozuk ya da farklı bir biçim). Excel'de açıp .xlsx olarak yeniden kaydet."
        sonuc["hata"] = neden
        return sonuc

    son_neden = "Boş dosya."
    for _, ham in (sayfalar or {}).items():
        if ham is None or ham.empty:
            continue
        df, notlar = _tablodan_ogrenciler(ham)
        if df is not None:
            if sonuc["hazirlik"]:
                df["Sorumlu"] = HAZIRLIK_SORUMLU
                notlar = [n for n in notlar if "sorumlusu" not in n]
            df["Kaynak"] = ad
            sonuc["df"], sonuc["notlar"] = df, notlar
            return sonuc
        son_neden = notlar
    sonuc["hata"] = son_neden
    return sonuc


# ─────────────────────────────────────────────────────────────
# Ders bazında birleştirme, sıralama, salon planı
# ─────────────────────────────────────────────────────────────
def ders_ogrencileri(dosya_dfleri):
    """Birden çok listeyi birleştirir; aynı öğrenci iki listede varsa bir kez alınır."""
    dfs = [d for d in dosya_dfleri if d is not None and not d.empty]
    if not dfs:
        return pd.DataFrame(columns=["No", "Ad", "Soyad", "Sorumlu", "Kaynak"]), 0
    df = pd.concat(dfs, ignore_index=True)
    once = len(df)
    df = df.drop_duplicates(subset=["No"], keep="first").reset_index(drop=True)
    return df, once - len(df)


def sirala(df, kriter="No"):
    d = df.copy()
    # Hoca grupları: kalabalık grup önce, küçük grup sona; 'Belirtilmedi' en sonda.
    # Aynı hocanın farklı yazımları (Öğr.Gör. / Öğr. Gör.) tek grup sayılır.
    anahtar = d["Sorumlu"].map(hoca_anahtari)
    sayim = anahtar.value_counts()
    d["_h"] = anahtar.map(lambda a: ("1" if a == "BELIRTILMEDI" else "0")
                          + str(100000 - int(sayim[a])).zfill(6) + a)
    d["_no"] = d["No"].map(lambda x: re.sub(r"\D", "", x).zfill(20) + x)
    d["_ad"] = d["Ad"].map(tr_sira)
    d["_soyad"] = d["Soyad"].map(tr_sira)
    sira = {"No": ["_h", "_no"], "Ad": ["_h", "_ad", "_soyad", "_no"], "Soyad": ["_h", "_soyad", "_ad", "_no"]}
    d = d.sort_values(by=sira.get(kriter, sira["No"]), kind="stable")
    return d.drop(columns=["_h", "_no", "_ad", "_soyad"]).reset_index(drop=True)


def ayar_duzelt(ayar):
    a = dict(VARSAYILAN_AYAR)
    a.update(ayar or {})
    a["normal_kap"] = max(1, min(int(a["normal_kap"]), NORMAL_KAP_MAX))
    a["buyuk_kap"] = max(1, min(int(a["buyuk_kap"]), BUYUK_KAP_MAX))
    a["buyuk_salonlar"] = [str(s).strip() for s in a["buyuk_salonlar"] if str(s).strip()]
    return a


def salon_buyuk_mu(salon, ayar):
    return str(salon).strip() in set(ayar["buyuk_salonlar"])


def kapasite(salonlar, ayar):
    return sum(ayar["buyuk_kap"] if salon_buyuk_mu(s["salon"], ayar) else ayar["normal_kap"] for s in salonlar)


def plan_yap(n, salonlar, ayar):
    """Önce normal salonlar doldurulur, kalan büyük salonlara eşit bölünür. Dönüş: ([(salon, adet, büyük_mü)], sığmayan)"""
    plan, kalan = [], n
    kucuk = [s for s in salonlar if not salon_buyuk_mu(s["salon"], ayar)]
    buyuk = [s for s in salonlar if salon_buyuk_mu(s["salon"], ayar)]
    for s in kucuk:
        if kalan <= 0:
            break
        k = min(ayar["normal_kap"], kalan)
        plan.append((s, k, False))
        kalan -= k
    if buyuk and kalan > 0:
        p, a = divmod(kalan, len(buyuk))
        for s in buyuk:
            k = min(ayar["buyuk_kap"], p + (1 if a > 0 else 0))
            a = max(0, a - 1)
            if k > 0:
                plan.append((s, k, True))
                kalan -= k
    return plan, kalan


def cakismalar(dersler, ders_dfleri):
    """Aynı tarih+saatte iki farklı sınava yazılmış öğrenciler."""
    grup = defaultdict(list)
    for d in dersler:
        grup[(d["tarih"], d["saat"])].append(d)
    sonuc = []
    for (tarih, saat), ds in grup.items():
        if len(ds) < 2:
            continue
        nerede = defaultdict(list)
        for d in ds:
            df = ders_dfleri.get(d["id"])
            if df is None or df.empty:
                continue
            for _, r in df.iterrows():
                nerede[r["No"]].append((d["ders_adi"], f"{r['Ad']} {r['Soyad']}"))
        for no, l in nerede.items():
            if len(l) > 1:
                sonuc.append({"Tarih": tarih, "Saat": saat, "No": no, "Öğrenci": l[0][1],
                              "Sınavlar": " | ".join(x[0] for x in l)})
    return sonuc


# ─────────────────────────────────────────────────────────────
# Excel üretimi
# ─────────────────────────────────────────────────────────────
def _sayfa_adi(ad, mevcut):
    ad = re.sub(r"[\\/*?:\[\]]", "-", ad)[:31]
    temel, i = ad, 2
    while ad in mevcut:
        ek = f"_{i}"
        ad = temel[: 31 - len(ek)] + ek
        i += 1
    return ad


def ders_excel(ders, df, sablon_veri, ayar, kriter="No"):
    """Dönüş: (xlsx_bytes | None, bilgi)"""
    ayar = ayar_duzelt(ayar)
    df = sirala(df, kriter)
    plan, sigmayan = plan_yap(len(df), ders["salonlar"], ayar)
    bilgi = {"plan": [], "sigmayan": df.iloc[len(df) - sigmayan:] if sigmayan else df.iloc[0:0], "uyarilar": []}
    if not plan:
        bilgi["uyarilar"].append("Salon görevi olmadığı için liste üretilemedi.")
        return None, bilgi

    wb = load_workbook(io.BytesIO(sablon_veri))
    sablon_sayfalari = list(wb.sheetnames)
    idx = 0
    for s, k, buyuk in plan:
        kaynak = "Sayfa2" if buyuk else "Sayfa1"
        if kaynak not in wb.sheetnames:
            kaynak = sablon_sayfalari[0]
        sh = wb.copy_worksheet(wb[kaynak])
        sh.title = _sayfa_adi(f"Sinif_{s['salon']}", wb.sheetnames)
        dilim = df.iloc[idx: idx + k]

        sh["B3"], sh["G5"], sh["C4"], sh["G4"] = ders["ders_adi"], ders["saat"], s["salon"], ders["tarih"]

        r_alt = BUYUK_IMZA if buyuk else NORMAL_IMZA
        goz = s["gozetmenler"]
        sh[f"E{r_alt}"] = "\n".join(goz)
        if len(goz) > 1:
            sh[f"E{r_alt}"].alignment = Alignment(wrapText=True, vertical="center")

        hocalar = benzersiz_hocalar(dilim["Sorumlu"])
        t_h = sh[f"A{r_alt}"]
        t_h.value = "\n".join(hocalar)
        t_h.alignment = Alignment(wrapText=True, vertical="center")
        t_h.font = Font(bold=True, size=10 if len(hocalar) < 3 else 8)

        for i, (_, row) in enumerate(dilim.iterrows()):
            r = ILK_SATIR + i
            sh[f"C{r}"] = int(row["No"]) if row["No"].isdigit() and len(row["No"]) < 16 else row["No"]
            sh[f"D{r}"] = tr_buyuk(row["Ad"])
            sh[f"E{r}"] = tr_buyuk(row["Soyad"])
            sh[f"F{r}"].value = tr_buyuk(row["Sorumlu"])
            sh[f"F{r}"].font = Font(size=10)
            sh[f"F{r}"].alignment = Alignment(shrink_to_fit=True)
        idx += k
        bilgi["plan"].append({"Salon": s["salon"], "Gözetmen": ", ".join(goz), "Öğrenci": k,
                              "Hocalar": ", ".join(hocalar)})

    for ad in sablon_sayfalari:
        if ad in wb.sheetnames and len(wb.sheetnames) > 1:
            wb.remove(wb[ad])

    if sigmayan:
        sh = wb.create_sheet("YERLESEMEYEN")
        sh.append([f"Salon kapasitesi yetersiz: {sigmayan} öğrenci yerleştirilemedi"])
        sh["A1"].font = Font(bold=True, color="C00000")
        sh.append(["No", "Ad", "Soyad", "Ders Sorumlusu", "Liste"])
        for _, r in bilgi["sigmayan"].iterrows():
            sh.append([r["No"], tr_buyuk(r["Ad"]), tr_buyuk(r["Soyad"]), r["Sorumlu"], r["Kaynak"]])
        bilgi["uyarilar"].append(f"{sigmayan} öğrenci salonlara sığmadı (YERLESEMEYEN sayfasına yazıldı).")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), bilgi


def _tablo_sayfasi(wb, ad, satirlar, basliklar):
    sh = wb.create_sheet(ad)
    sh.append(basliklar)
    for c in range(1, len(basliklar) + 1):
        sh.cell(row=1, column=c).font = Font(bold=True, color="FFFFFF")
        sh.cell(row=1, column=c).fill = PatternFill("solid", start_color="2C3E50")
    for s in satirlar:
        sh.append([s.get(b, "") for b in basliklar])
    for c, b in enumerate(basliklar, 1):
        en = max([len(str(b))] + [len(str(s.get(b, ""))) for s in satirlar]) if satirlar else len(b)
        sh.column_dimensions[get_column_letter(c)].width = min(60, en + 2)
    sh.freeze_panes = "A2"
    return sh


def dosya_adi_yap(ders):
    temiz = re.sub(r"[^\w]", "_", ders["ders_adi"][:30], flags=re.UNICODE)
    return f"{ders['tarih'].replace('.', '_')}_{ders['saat'].replace(':', '.')}_{temiz}.xlsx"


def paket_olustur(dersler, ders_dfleri, ders_dosyalari, sablon_veri, ayar, siralamalar,
                  eslesmeyenler=(), okuma_notlari=(), cakisma=()):
    """Dönüş: (zip_bytes, özet_satırları, üretilen_sayısı)"""
    ayar = ayar_duzelt(ayar)
    zbuf = io.BytesIO()
    ozet, sigmayanlar, uretilen = [], [], 0
    kullanilan = set()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        for d in dersler:
            df = ders_dfleri.get(d["id"])
            satir = {"Tarih": d["tarih"], "Saat": d["saat"], "Ders": d["ders_adi"],
                     "Listeler": ", ".join(ders_dosyalari.get(d["id"], [])),
                     "Öğrenci": 0 if df is None else len(df),
                     "Salonlar": ", ".join(s["salon"] for s in d["salonlar"]),
                     "Kapasite": kapasite(d["salonlar"], ayar), "Durum": ""}
            if df is None or df.empty:
                satir["Durum"] = "Liste yok — dosya üretilmedi"
                ozet.append(satir)
                continue
            veri, bilgi = ders_excel(d, df, sablon_veri, ayar, siralamalar.get(d["id"], "No"))
            if veri is None:
                satir["Durum"] = "; ".join(bilgi["uyarilar"])
                ozet.append(satir)
                continue
            ad = dosya_adi_yap(d)
            yol = f"{d['tarih'].replace('.', '_')}/{ad}"
            if yol in kullanilan:
                yol = yol[:-5] + f"_{d['id']}.xlsx"
            kullanilan.add(yol)
            z.writestr(yol, veri)
            uretilen += 1
            satir["Durum"] = "; ".join(bilgi["uyarilar"]) or "Tamam"
            ozet.append(satir)
            for _, r in bilgi["sigmayan"].iterrows():
                sigmayanlar.append({"Ders": d["ders_adi"], "Tarih": d["tarih"], "Saat": d["saat"],
                                    "No": r["No"], "Ad": r["Ad"], "Soyad": r["Soyad"], "Liste": r["Kaynak"]})

        rwb = Workbook()
        rwb.remove(rwb.active)
        _tablo_sayfasi(rwb, "Dersler", ozet, ["Tarih", "Saat", "Ders", "Öğrenci", "Kapasite", "Salonlar", "Listeler", "Durum"])
        _tablo_sayfasi(rwb, "Yerlesemeyen", sigmayanlar, ["Ders", "Tarih", "Saat", "No", "Ad", "Soyad", "Liste"])
        _tablo_sayfasi(rwb, "Cakismalar", list(cakisma), ["Tarih", "Saat", "No", "Öğrenci", "Sınavlar"])
        _tablo_sayfasi(rwb, "Eslesmeyen_Dosyalar", [{"Dosya": e} for e in eslesmeyenler], ["Dosya"])
        _tablo_sayfasi(rwb, "Okuma_Notlari", list(okuma_notlari), ["Dosya", "Not"])
        rb = io.BytesIO()
        rwb.save(rb)
        z.writestr("00_RAPOR.xlsx", rb.getvalue())
    return zbuf.getvalue(), ozet, uretilen
