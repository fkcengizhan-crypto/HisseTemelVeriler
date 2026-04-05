import sys, time, re, os, queue, threading
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Kütüphaneler eksik: pip install selenium openpyxl beautifulsoup4 webdriver-manager")
    sys.exit(1)

# ══════════════════════════════════════════════════════
# AYARLAR VE GLOBAL DEĞİŞKENLER
# ══════════════════════════════════════════════════════
PARALEL_TARAYICI = 7
SAYFA_BEKLEME   = 5      # DÜZELTİLDİ: 3 → 5 sn (JS yüklenme süresi için)
MAX_DENEME      = 2      # DÜZELTİLDİ: artık kullanılıyor
TXT_DOSYA       = "hisseisimleri.txt"

cekilen_sayisi = 0
sayac_kilidi   = threading.Lock()

OZET_SUTUNLAR = [
    ("Kod",                   "kod"),
    ("Şirket Adı",            "sirket_adi"),
    ("F/K",                   "cari.F/K"),
    ("FD/FAVÖK",              "cari.FD/FAVÖK"),   # DÜZELTİLDİ: parsing'e eklendi
    ("PD/DD",                 "cari.PD/DD"),
    ("Piyasa Değeri (mnTL)",  "cari.Piyasa Değeri"),
    ("Net Kâr (mnTL)",        "mali_ozet.Net Kâr"),
    ("Faaliyet Alanı",        "kunye.Faal Alanı"),
    ("Adres",                 "kunye.Adres"),
]

# ══════════════════════════════════════════════════════
# YARDIMCI FONKSİYONLAR
# ══════════════════════════════════════════════════════

def temizle(x):
    return " ".join(str(x).split()).strip() if x else ""

def sayi(x):
    """Sayısal string'i float'a çevirir. Türkçe format (1.234,56) desteklenir."""
    if x is None:
        return None
    s = str(x).strip()
    # Türkçe format: önce binlik noktaları kaldır, sonra virgülü noktaya çevir
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return None

def deger_al(kayit, yol):
    parcalar = yol.split(".", 1)
    val = kayit.get(parcalar[0])
    if len(parcalar) == 1:
        return val
    if isinstance(val, dict):
        return deger_al(val, parcalar[1])
    return None

def oku_txt(dosya):
    if not os.path.exists(dosya):
        print(f"HATA: '{dosya}' dosyası bulunamadı!")
        return []
    with open(dosya, "r", encoding="utf-8") as f:
        kodlar = [l.strip().upper() for l in f if l.strip()]
    if not kodlar:
        print(f"HATA: '{dosya}' dosyası boş!")
    return kodlar

def chrome_olustur(driver_path: str):
    """driver_path: main()'de bir kez indirilen ChromeDriver binary yolu."""
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    svc = Service(driver_path)
    return webdriver.Chrome(service=svc, options=opts)

# ══════════════════════════════════════════════════════
# SAYFA ÇEKME (RETRY MANTIĞIYLA)
# ══════════════════════════════════════════════════════

def sayfa_cek(driver, kod):
    url = (
        f"https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/"
        f"sirket-karti.aspx?hisse={kod}"
    )
    for deneme in range(1, MAX_DENEME + 1):
        try:
            driver.get(url)
            # Önce herhangi bir tablo bekle
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table"))
            )
            # JS verisinin tam yüklenmesi için ekstra bekleme
            time.sleep(SAYFA_BEKLEME)
            return driver.page_source
        except Exception as e:
            print(f"\n  [UYARI] {kod} - Deneme {deneme}/{MAX_DENEME}: {e}")
            if deneme == MAX_DENEME:
                return driver.page_source  # Son denemede ne varsa al
            time.sleep(3)
    return ""

# ══════════════════════════════════════════════════════
# HTML AYRIŞTIRMA
# ══════════════════════════════════════════════════════

def ayristir(html, kod):
    soup = BeautifulSoup(html, "html.parser")
    kayit = {
        "kod":        kod,
        "sirket_adi": kod,   # varsayılan: kod, aşağıda üstüne yazılır
        "kunye":      {},
        "cari":       {},
        "mali_ozet":  {},
    }

    # ── Şirket adı ──────────────────────────────────────
    # İsyatirim genellikle başlığı <h1> veya belirli class'larda tutar
    for sel in ["h1", "h2", ".company-name", ".card-title"]:
        el = soup.select_one(sel)
        if el:
            ad = temizle(el.get_text())
            if ad and ad.upper() != kod:
                kayit["sirket_adi"] = ad
                break

    # ── Tablo taraması ───────────────────────────────────
    for tablo in soup.find_all("table"):
        for satir in tablo.find_all("tr"):
            hucreler = satir.find_all(["td", "th"])
            if len(hucreler) < 2:
                continue
            k = temizle(hucreler[0].get_text())
            v = temizle(hucreler[1].get_text())
            if not k or not v:
                continue

            # Cari değerler
            if re.search(r"F\s*/\s*K", k, re.I):
                kayit["cari"]["F/K"] = sayi(v)
            elif re.search(r"FD\s*/\s*FAVÖK", k, re.I):         # DÜZELTİLDİ: eklendi
                kayit["cari"]["FD/FAVÖK"] = sayi(v)
            elif re.search(r"PD\s*/\s*DD", k, re.I):
                kayit["cari"]["PD/DD"] = sayi(v)
            elif re.search(r"Piyasa\s*De[ğg]eri", k, re.I):
                # DÜZELTİLDİ: re.sub yerine doğrudan sayi() — zaten temizliyor
                kayit["cari"]["Piyasa Değeri"] = sayi(v)

            # Mali özet
            elif re.search(r"Net\s*Kâr|Net\s*Kar", k, re.I):
                kayit["mali_ozet"]["Net Kâr"] = sayi(v)

            # Künye
            elif re.search(r"Faal\s*Alan", k, re.I):
                kayit["kunye"]["Faal Alanı"] = v
            elif re.search(r"Adres", k, re.I):
                kayit["kunye"]["Adres"] = v

    # ── Debug: hiç veri bulunamadıysa uyar ─────────────
    if not any([kayit["cari"], kayit["mali_ozet"]]):
        # Sayfanın ilk 200 karakterini göster (terminal spam olmadan)
        snippet = soup.get_text()[:200].replace("\n", " ")
        print(f"\n  [DEBUG] {kod} — veri bulunamadı. Sayfa özeti: {snippet}")

    return kayit

# ══════════════════════════════════════════════════════
# WORKER
# ══════════════════════════════════════════════════════

def worker_calis(kodlar_q, sonuclar, toplam, worker_id, driver_path):
    global cekilen_sayisi
    driver = None
    try:
        driver = chrome_olustur(driver_path)
        while True:
            try:
                sira, kod = kodlar_q.get_nowait()
            except queue.Empty:
                break

            try:                                          # DÜZELTİLDİ: try/except eklendi
                html  = sayfa_cek(driver, kod)
                kayit = ayristir(html, kod)
                sonuclar[sira] = kayit
            except Exception as e:
                print(f"\n  [HATA] Worker-{worker_id} | {kod}: {e}")
                sonuclar[sira] = {"kod": kod, "sirket_adi": f"HATA: {e}"}

            with sayac_kilidi:
                cekilen_sayisi += 1
                yuzde = (cekilen_sayisi / toplam) * 100
                sys.stdout.write(
                    f"\r  >> İlerleme: %{yuzde:.1f} | "
                    f"{cekilen_sayisi}/{toplam} hisse tamamlandı ({kod})"
                    f"{'':10}"
                )
                sys.stdout.flush()

            kodlar_q.task_done()
    finally:
        if driver:
            driver.quit()

# ══════════════════════════════════════════════════════
# EXCEL YAZMA
# ══════════════════════════════════════════════════════

def excel_yaz(veri_listesi, dosya_adi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hisse_Verileri"

    # Başlık satırı
    for ci, (baslik, _) in enumerate(OZET_SUTUNLAR, 1):
        hucre = ws.cell(1, ci, baslik)
        hucre.font = Font(bold=True)

    # Veri satırları
    bos_sayisi = 0
    for ri, kayit in enumerate(veri_listesi, 2):
        if not kayit:
            bos_sayisi += 1
            continue
        for ci, (_, yol) in enumerate(OZET_SUTUNLAR, 1):
            ws.cell(ri, ci, deger_al(kayit, yol))

    if bos_sayisi:
        print(f"\n  [UYARI] {bos_sayisi} kayıt None olarak kaldı (hata log'larını inceleyin).")

    wb.save(dosya_adi)
    print(f"  Excel kaydedildi: {dosya_adi}")

# ══════════════════════════════════════════════════════
# ANA AKIŞ
# ══════════════════════════════════════════════════════

def main():
    kodlar = oku_txt(TXT_DOSYA)
    if not kodlar:
        return

    toplam = len(kodlar)
    print(f"\n{toplam} hisse için işlem başladı... ({PARALEL_TARAYICI} paralel tarayıcı)\n")

    # ── ChromeDriver'ı tek seferlik indir (race condition önlemi) ──
    print("  ChromeDriver indiriliyor/kontrol ediliyor...")
    driver_path = ChromeDriverManager().install()
    print(f"  ChromeDriver hazır: {driver_path}\n")

    kodlar_q = queue.Queue()
    for i, k in enumerate(kodlar):
        kodlar_q.put((i, k))

    sonuclar = [None] * toplam
    threads  = []
    for wid in range(1, min(PARALEL_TARAYICI, toplam) + 1):
        t = threading.Thread(
            target=worker_calis,
            args=(kodlar_q, sonuclar, toplam, wid, driver_path),
            daemon=True,
        )
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    tarih = datetime.now().strftime("%Y%m%d_%H%M")
    dosya_adi = f"IsYatirim_{tarih}.xlsx"
    excel_yaz(sonuclar, dosya_adi)
    print(f"\nTamamlandı → {dosya_adi}")

if __name__ == "__main__":
    main()
