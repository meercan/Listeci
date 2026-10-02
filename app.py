# -*- coding: utf-8 -*-
"""LİSTECİ v3 — Streamlit arayüzü. İş mantığı motor.py içinde."""
import hashlib
import json
import os

import pandas as pd
import streamlit as st

import motor as m

st.set_page_config(page_title="LİSTECİ v3", page_icon="📋", layout="wide")

st.markdown("""
<style>
  .main .block-container { padding-top: 1.5rem; max-width: 1300px; }
  h1 { color: #2c3e50; text-align: center; font-weight: 800; margin-bottom: .2rem; }
  .alt-baslik { text-align:center; color:#7f8c8d; margin-bottom: 1.5rem; }
  .adim { background:#2c3e50; color:white; padding:.45rem .9rem; border-radius:6px;
          font-weight:700; margin: 1.4rem 0 .8rem 0; }
  .dosya-satir { font-size: 14px; line-height: 1.35; }
  .kucuk { color:#7f8c8d; font-size: 12.5px; }
</style>
""", unsafe_allow_html=True)

HAFIZA_DOSYASI = "eslestirme_hafizasi.json"
DURUM_ROZET = {
    "otomatik": "✅ Otomatik",
    "hafiza": "🧠 Hafızadan",
    "manuel": "✍️ Elle",
    "kontrol": "⚠️ Kontrol et",
    "yok": "❌ Eşleşmedi",
    "hata": "📛 Okunamadı",
}


# ─────────────────────────────────────────────────────────────
# Giriş
# ─────────────────────────────────────────────────────────────
def gecerli_sifre():
    try:
        return str(st.secrets.get("SIFRE", "asu123"))
    except Exception:
        return "asu123"


def sifre_kontrol():
    if st.session_state.get("giris"):
        return True
    st.title("📋 LİSTECİ")
    st.subheader("🔒 Yetkili girişi")
    with st.form("giris_formu"):
        s = st.text_input("Erişim şifresi", type="password")
        if st.form_submit_button("Giriş yap", type="primary"):
            if s == gecerli_sifre():
                st.session_state["giris"] = True
                st.rerun()
            else:
                st.error("Hatalı şifre.")
    return False


# ─────────────────────────────────────────────────────────────
# Durum
# ─────────────────────────────────────────────────────────────
def hafiza_yukle():
    try:
        with open(HAFIZA_DOSYASI, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def hafiza_kaydet():
    try:
        with open(HAFIZA_DOSYASI, "w", encoding="utf-8") as f:
            json.dump(st.session_state["hafiza"], f, ensure_ascii=False, indent=1)
    except Exception:
        pass   # bulutta yazılamazsa sorun değil; kenar çubuğundan indirilebilir


def durum_hazirla():
    v = st.session_state
    v.setdefault("matris", None)          # {"ad","hash","dersler","uyarilar"}
    v.setdefault("dosyalar", {})          # ad → {"veri","hash","okuma","sira","eslesme","durum","aciklama"}
    v.setdefault("sayac", 0)
    v.setdefault("son_yuklenenler", set())
    v.setdefault("hafiza", hafiza_yukle())
    v.setdefault("cikti", None)


def sec_anahtari(ad):
    return f"sec_{st.session_state['dosyalar'][ad]['sira']}"


def dosyayi_eslestir(ad):
    v = st.session_state
    d = v["dosyalar"][ad]
    dersler = v["matris"]["dersler"] if v["matris"] else []
    if d["okuma"]["df"] is None:
        e = {"idler": [], "durum": "hata", "aciklama": str(d["okuma"]["hata"])}
        if dersler:  # okunamasa bile hangi derse ait olduğunu göstermek faydalı
            e2 = m.eslestir(ad, dersler, v["hafiza"], v["ayar_hz"])
            e["idler"] = e2["idler"]
    else:
        e = m.eslestir(ad, dersler, v["hafiza"], v["ayar_hz"])
    d["durum"], d["aciklama"] = e["durum"], e["aciklama"]
    v[sec_anahtari(ad)] = list(e["idler"])


def hepsini_eslestir():
    for ad in st.session_state["dosyalar"]:
        dosyayi_eslestir(ad)
    st.session_state["cikti"] = None


def elle_degisti(ad):
    """Kullanıcı bir dosyanın derslerini değiştirdi → durumu 'manuel' yap ve hafızaya yaz."""
    v = st.session_state
    d = v["dosyalar"][ad]
    secim = v.get(sec_anahtari(ad), [])
    if d["durum"] != "hata":
        d["durum"] = "manuel" if secim else "yok"
        d["aciklama"] = "Elle seçildi" if secim else "Elle boşaltıldı"
    dersler = {x["id"]: x for x in v["matris"]["dersler"]}
    anahtar = m.sade(m.dosya_govdesi(ad))
    if secim:
        v["hafiza"][anahtar] = [m.sade(dersler[i]["ders_adi"]) for i in secim if i in dersler]
    else:
        v["hafiza"].pop(anahtar, None)
    hafiza_kaydet()
    v["cikti"] = None


def ders_etiketi(d):
    return f"{d['tarih']} {d['saat']} · {d['ders_adi']}"


# ─────────────────────────────────────────────────────────────
# Kenar çubuğu
# ─────────────────────────────────────────────────────────────
def kenar_cubugu():
    v = st.session_state
    with st.sidebar:
        st.header("⚙️ Ayarlar")
        nk = st.number_input("Normal salon kapasitesi", 1, m.NORMAL_KAP_MAX, 34,
                             help=f"Şablonun Sayfa1'inde en fazla {m.NORMAL_KAP_MAX} satır var (7–44).")
        bs = st.text_input("Büyük salonlar (virgülle)", "205, 305",
                           help="Bu salonlar şablonun Sayfa2'sine (uzun liste) yazılır.")
        bk = st.number_input("Büyük salon kapasitesi", 1, m.BUYUK_KAP_MAX, 68,
                             help=f"Şablonun Sayfa2'sinde en fazla {m.BUYUK_KAP_MAX} satır var (7–74).")
        v["ayar"] = m.ayar_duzelt({"normal_kap": nk, "buyuk_kap": bk,
                                   "buyuk_salonlar": [x for x in bs.replace(";", ",").split(",")]})
        v["varsayilan_sira"] = st.radio("Varsayılan sıralama", ["No", "Ad", "Soyad"], horizontal=True)

        with st.expander("🎓 Hazırlık sınavı adları"):
            st.caption("Matriste bu adlarla geçen sınavlara bütün Arapça hazırlık listeleri bağlanır. Her satıra bir ad.")
            hz_metin = st.text_area("Hazırlık dersleri", "\n".join(m.VARSAYILAN_HZ_DERSLERI),
                                    height=200, label_visibility="collapsed")
            hz = [x.strip() for x in hz_metin.splitlines() if x.strip()]
            if hz != v.get("ayar_hz"):
                ilk = "ayar_hz" not in v
                v["ayar_hz"] = hz
                if not ilk and v["dosyalar"]:
                    hepsini_eslestir()

        st.divider()
        st.subheader("🧠 Eşleştirme hafızası")
        st.caption(f"{len(v['hafiza'])} dosya adı hatırlanıyor. Elle yaptığın her eşleştirme buraya eklenir; "
                   "bir sonraki sınav döneminde aynı adlı dosyalar kendiliğinden bağlanır.")
        st.download_button("⬇️ Hafızayı indir", json.dumps(v["hafiza"], ensure_ascii=False, indent=1).encode("utf-8"),
                           "eslestirme_hafizasi.json", "application/json", width="stretch")
        yuk = st.file_uploader("Hafıza yükle (.json)", type=["json"], key="hafiza_up")
        if yuk is not None and v.get("hafiza_up_hash") != hashlib.md5(yuk.getvalue()).hexdigest():
            try:
                v["hafiza"].update(json.loads(yuk.getvalue().decode("utf-8")))
                v["hafiza_up_hash"] = hashlib.md5(yuk.getvalue()).hexdigest()
                hafiza_kaydet()
                hepsini_eslestir()
                st.success("Hafıza yüklendi.")
            except Exception as e:
                st.error(f"Hafıza okunamadı: {e}")

        with st.expander("🛠️ Şablon"):
            sb = st.file_uploader("Farklı şablon kullan (isteğe bağlı)", type=["xlsx"], key="sablon_up")
            v["sablon_veri"] = sb.getvalue() if sb is not None else None

        st.divider()
        if st.button("🔄 Yeni oturum (her şeyi temizle)", width="stretch"):
            hafiza = v["hafiza"]
            for k in list(v.keys()):
                if k not in ("giris",):
                    del v[k]
            v["hafiza"] = hafiza
            st.rerun()
        if st.button("🚪 Çıkış", width="stretch"):
            v["giris"] = False
            st.rerun()
        st.caption(f"LİSTECİ v{m.SURUM}")


# ─────────────────────────────────────────────────────────────
# Adımlar
# ─────────────────────────────────────────────────────────────
def adim_matris():
    v = st.session_state
    st.markdown('<div class="adim">1 · Sınav görev matrisi</div>', unsafe_allow_html=True)
    f = st.file_uploader("Sınav görev matrisini seç (.xlsx)", type=["xlsx"], key="matris_up")
    if f is not None:
        veri = f.getvalue()
        h = hashlib.md5(veri).hexdigest()
        if not v["matris"] or v["matris"]["hash"] != h:
            try:
                dersler, uyarilar = m.matris_oku(veri)
                v["matris"] = {"ad": f.name, "hash": h, "dersler": dersler, "uyarilar": uyarilar}
                hepsini_eslestir()
            except Exception as e:
                st.error(f"Matris okunamadı: {e}")
                return False

    if not v["matris"]:
        st.info("Başlamak için matris dosyasını yükle.")
        return False

    mt = v["matris"]
    dersler = mt["dersler"]
    hz_sayi = sum(1 for d in dersler if m.hazirlik_dersi_mi(d["ders_adi"], v["ayar_hz"]))
    c1, c2, c3 = st.columns(3)
    c1.metric("Ders", len(dersler))
    c2.metric("Hazırlık sınavı", hz_sayi)
    c3.metric("Toplam salon kapasitesi", sum(m.kapasite(d["salonlar"], v["ayar"]) for d in dersler))
    for u in mt["uyarilar"]:
        st.warning(u)
    with st.expander(f"📅 Matristen okunan dersler ({mt['ad']})"):
        st.dataframe(pd.DataFrame([{
            "Tarih": d["tarih"], "Saat": d["saat"], "Ders": d["ders_adi"],
            "Tür": "Hazırlık" if m.hazirlik_dersi_mi(d["ders_adi"], v["ayar_hz"]) else "",
            "Salonlar": ", ".join(s["salon"] for s in d["salonlar"]),
            "Kapasite": m.kapasite(d["salonlar"], v["ayar"]), "Hücre": d["hucre"],
        } for d in dersler]), hide_index=True, width="stretch")
    return True


def adim_listeler():
    v = st.session_state
    st.markdown('<div class="adim">2 · Öğrenci listeleri ve eşleştirme</div>', unsafe_allow_html=True)
    yuklenen = st.file_uploader("Tüm şube/öğrenci listelerini buraya topluca bırak",
                                type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="ogr_up")
    yuklenen = yuklenen or []

    simdiki = set()
    for f in yuklenen:
        veri = f.getvalue()
        h = hashlib.md5(veri).hexdigest()
        simdiki.add(f.name)
        mevcut = v["dosyalar"].get(f.name)
        if mevcut and mevcut["hash"] == h:
            continue
        v["sayac"] += 1
        v["dosyalar"][f.name] = {"veri": veri, "hash": h, "okuma": m.liste_oku(veri, f.name),
                                 "sira": v["sayac"], "durum": "", "aciklama": ""}
        dosyayi_eslestir(f.name)
        v["cikti"] = None
    for ad in v["son_yuklenenler"] - simdiki:      # yükleyiciden çıkarılan dosyalar
        if ad in v["dosyalar"]:
            v.pop(sec_anahtari(ad), None)
            del v["dosyalar"][ad]
            v["cikti"] = None
    v["son_yuklenenler"] = simdiki

    dosyalar = v["dosyalar"]
    if not dosyalar:
        st.info("Listeleri yükleyince sistem dosya adlarına bakıp her birini ilgili sınava bağlayacak.")
        return False

    dersler = v["matris"]["dersler"]
    ders_map = {d["id"]: d for d in dersler}
    say = {k: 0 for k in DURUM_ROZET}
    for d in dosyalar.values():
        say[d["durum"]] = say.get(d["durum"], 0) + 1
    c = st.columns(5)
    c[0].metric("Dosya", len(dosyalar))
    c[1].metric("✅ Eşleşti", say["otomatik"] + say["hafiza"] + say["manuel"])
    c[2].metric("⚠️ Kontrol et", say["kontrol"])
    c[3].metric("❌ Eşleşmedi", say["yok"])
    c[4].metric("📛 Okunamadı", say["hata"])

    def satir(ad):
        d = dosyalar[ad]
        ok = d["okuma"]
        sol, sag = st.columns([2, 3])
        with sol:
            n = 0 if ok["df"] is None else len(ok["df"])
            ek = " · 🎓 hazırlık" if ok["hazirlik"] else ""
            st.markdown(f'<div class="dosya-satir"><b>📄 {ad}</b><br>'
                        f'<span class="kucuk">{DURUM_ROZET.get(d["durum"], "")} — {d["aciklama"]} · {n} öğrenci{ek}</span></div>',
                        unsafe_allow_html=True)
            for n_ in ok["notlar"]:
                st.caption(f"ℹ️ {n_}")
        with sag:
            st.multiselect("Sınav(lar)", options=list(ders_map.keys()),
                           format_func=lambda i: ders_etiketi(ders_map[i]),
                           key=sec_anahtari(ad), on_change=elle_degisti, args=(ad,),
                           label_visibility="collapsed", placeholder="Bu liste hangi sınava ait? Seç…")

    sorunlu = [a for a, d in dosyalar.items() if d["durum"] in ("kontrol", "yok", "hata", "manuel")]
    sorunlu.sort(key=lambda a: ["hata", "yok", "kontrol", "manuel"].index(dosyalar[a]["durum"]))
    tamam = sorted(a for a in dosyalar if a not in sorunlu)

    if sorunlu:
        bekleyen = sum(1 for a in sorunlu if dosyalar[a]["durum"] != "manuel")
        if bekleyen:
            st.warning(f"**{bekleyen} dosyaya bakman gerekiyor.** Sağdaki kutudan doğru sınavı seç; "
                       "seçimin hafızaya alınır, bir dahaki sefere aynı ad kendiliğinden eşleşir.")
        with st.container(border=True):
            for a in sorunlu:
                satir(a)
    else:
        st.success("Bütün listeler bir sınava bağlandı.")

    with st.expander(f"✅ Otomatik eşleşenler ({len(tamam)}) — gözden geçirmek için aç"):
        for a in tamam:
            satir(a)

    for ad, d in dosyalar.items():
        if d["okuma"]["df"] is not None and d["okuma"]["df"].empty:
            st.warning(f"{ad}: liste boş.")
    return True


def ders_verileri():
    """Her ders için birleşik öğrenci tablosu, dosya adları ve mükerrer sayısı."""
    v = st.session_state
    dersler = v["matris"]["dersler"]
    ders_dosya = {d["id"]: [] for d in dersler}
    for ad, d in v["dosyalar"].items():
        if d["okuma"]["df"] is None:
            continue
        for i in v.get(sec_anahtari(ad), []):
            if i in ders_dosya:
                ders_dosya[i].append(ad)
    ders_df, mukerrer = {}, {}
    for d in dersler:
        df, tekrar = m.ders_ogrencileri([v["dosyalar"][a]["okuma"]["df"] for a in ders_dosya[d["id"]]])
        ders_df[d["id"]] = df
        mukerrer[d["id"]] = tekrar
    return ders_df, ders_dosya, mukerrer


def adim_onizleme():
    v = st.session_state
    st.markdown('<div class="adim">3 · Önizleme ve kontrol</div>', unsafe_allow_html=True)
    dersler = v["matris"]["dersler"]
    ders_df, ders_dosya, mukerrer = ders_verileri()
    v["siralama"] = v.get("siralama", {})

    satirlar = []
    for d in dersler:
        n = len(ders_df[d["id"]])
        kap = m.kapasite(d["salonlar"], v["ayar"])
        _, sigmayan = m.plan_yap(n, d["salonlar"], v["ayar"])
        if n == 0:
            durum = "⚪ Liste yok"
        elif not d["salonlar"]:
            durum = "🔴 Salon yok"
        elif sigmayan:
            durum = f"🔴 {sigmayan} öğrenci sığmıyor"
        elif mukerrer[d["id"]]:
            durum = f"🟡 {mukerrer[d['id']]} mükerrer öğrenci ayıklandı"
        else:
            durum = "🟢 Hazır"
        satirlar.append({"id": d["id"], "Tarih": d["tarih"], "Saat": d["saat"], "Ders": d["ders_adi"],
                         "Liste": len(ders_dosya[d["id"]]), "Öğrenci": n, "Kapasite": kap,
                         "Durum": durum, "Sıralama": v["siralama"].get(d["id"], v["varsayilan_sira"])})
    tablo = pd.DataFrame(satirlar).set_index("id")

    kirmizi = sum(1 for s in satirlar if s["Durum"].startswith("🔴"))
    bos = sum(1 for s in satirlar if s["Durum"].startswith("⚪"))
    c = st.columns(4)
    c[0].metric("🟢 Hazır", sum(1 for s in satirlar if s["Durum"][0] in "🟢🟡"))
    c[1].metric("🔴 Sorunlu", kirmizi)
    c[2].metric("⚪ Listesi olmayan", bos)
    c[3].metric("Toplam öğrenci", int(tablo["Öğrenci"].sum()) if len(tablo) else 0)

    duzen = st.data_editor(
        tablo, width="stretch", hide_index=True, key="onizleme_tablo",
        disabled=["Tarih", "Saat", "Ders", "Liste", "Öğrenci", "Kapasite", "Durum"],
        column_config={
            "Sıralama": st.column_config.SelectboxColumn("Sıralama", options=["No", "Ad", "Soyad"], required=True),
            "Ders": st.column_config.TextColumn(width="large"),
            "Durum": st.column_config.TextColumn(width="medium"),
        },
    )
    for i, s in duzen["Sıralama"].items():
        v["siralama"][i] = s

    if kirmizi:
        st.error("🔴 işaretli derslerde salon kapasitesi yetmiyor ya da matriste salon yok. "
                 "Liste yine üretilir; sığmayan öğrenciler dosyada **YERLESEMEYEN** sayfasına ve rapora yazılır.")

    cak = m.cakismalar(dersler, ders_df)
    if cak:
        with st.expander(f"⚠️ Aynı saatte iki sınavı olan öğrenciler ({len(cak)})", expanded=True):
            st.dataframe(pd.DataFrame(cak), hide_index=True, width="stretch")

    with st.expander("🔍 Ders detayı (salon dağılımı ve öğrenciler)"):
        dolu = [d for d in dersler if len(ders_df[d["id"]])]
        if dolu:
            sec = st.selectbox("Ders", dolu, format_func=ders_etiketi, key="detay_ders")
            df = m.sirala(ders_df[sec["id"]], v["siralama"].get(sec["id"], v["varsayilan_sira"]))
            plan, sigmayan = m.plan_yap(len(df), sec["salonlar"], v["ayar"])
            st.markdown("**Listeler:** " + ", ".join(ders_dosya[sec["id"]]))
            idx, ps = 0, []
            for s, k, b in plan:
                dil = df.iloc[idx: idx + k]
                ps.append({"Salon": s["salon"] + (" (büyük)" if b else ""), "Gözetmen": ", ".join(s["gozetmenler"]),
                           "Öğrenci": k, "İmza alanındaki hocalar": ", ".join(m.benzersiz_hocalar(dil["Sorumlu"]))})
                idx += k
            if sigmayan:
                ps.append({"Salon": "❗ SIĞMAYAN", "Gözetmen": "", "Öğrenci": sigmayan, "İmza alanındaki hocalar": ""})
            st.dataframe(pd.DataFrame(ps), hide_index=True, width="stretch")
            st.dataframe(df[["No", "Ad", "Soyad", "Sorumlu", "Kaynak"]], hide_index=True, width="stretch", height=300)
        else:
            st.caption("Henüz öğrenci listesi bağlanmış ders yok.")
    return ders_df, ders_dosya, cak


def adim_uret(ders_df, ders_dosya, cak):
    v = st.session_state
    st.markdown('<div class="adim">4 · Listeleri üret</div>', unsafe_allow_html=True)
    sablon = v.get("sablon_veri")
    if sablon is None:
        if not os.path.exists("sablon.xlsx"):
            st.error("⚠️ Depoda 'sablon.xlsx' yok. GitHub'a yükle ya da kenar çubuğundan şablon seç.")
            return
        with open("sablon.xlsx", "rb") as f:
            sablon = f.read()

    bekleyen = [a for a, d in v["dosyalar"].items() if d["durum"] in ("kontrol", "yok")]
    if bekleyen:
        st.warning(f"{len(bekleyen)} dosya hâlâ kontrol bekliyor ya da eşleşmedi; bunlar listelere girmeyecek.")

    if st.button("🔥 TÜM LİSTELERİ HAZIRLA", type="primary", width="stretch"):
        with st.spinner("Listeler hazırlanıyor…"):
            notlar = [{"Dosya": a, "Not": n} for a, d in v["dosyalar"].items() for n in d["okuma"]["notlar"]]
            notlar += [{"Dosya": a, "Not": "OKUNAMADI: " + str(d["okuma"]["hata"])}
                       for a, d in v["dosyalar"].items() if d["okuma"]["df"] is None]
            eslesmeyen = [a for a, d in v["dosyalar"].items() if not v.get(sec_anahtari(a))]
            try:
                zip_veri, ozet, n = m.paket_olustur(
                    v["matris"]["dersler"], ders_df, ders_dosya, sablon, v["ayar"],
                    {i: v["siralama"].get(i, v["varsayilan_sira"]) for i in ders_df},
                    eslesmeyen, notlar, cak)
                v["cikti"] = {"zip": zip_veri, "ozet": ozet, "n": n}
            except Exception as e:
                st.exception(e)
                v["cikti"] = None

    c = v.get("cikti")
    if c:
        if c["n"]:
            st.success(f"🎉 {c['n']} ders için liste hazırlandı. ZIP'in içinde ayrıca **00_RAPOR.xlsx** var.")
            st.download_button("📥 ZIP olarak indir", c["zip"], "Tanzim_Edilen_Sinav_Listeleri.zip",
                               "application/zip", type="primary", width="stretch")
        else:
            st.error("Hiç liste üretilemedi: hiçbir derse öğrenci listesi bağlı değil ya da salon görevi yok.")
        sorun = [o for o in c["ozet"] if o["Durum"] != "Tamam"]
        if sorun:
            with st.expander(f"Durumu 'Tamam' olmayan dersler ({len(sorun)})"):
                st.dataframe(pd.DataFrame(sorun), hide_index=True, width="stretch")


def main():
    durum_hazirla()
    kenar_cubugu()
    st.title("📋 LİSTECİ")
    st.markdown('<div class="alt-baslik">Sınav görev matrisi + şube listeleri → salon salon imza listeleri</div>',
                unsafe_allow_html=True)
    if not adim_matris():
        return
    if not adim_listeler():
        return
    ders_df, ders_dosya, cak = adim_onizleme()
    adim_uret(ders_df, ders_dosya, cak)


if __name__ == "__main__":
    if sifre_kontrol():
        main()
