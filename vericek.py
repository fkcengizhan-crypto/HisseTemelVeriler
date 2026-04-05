import sys, time, re, os, queue, threading, shutil
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
    print("Eksik: pip install selenium openpyxl beautifulsoup4 webdriver-manager")
    sys.exit(1)

# ══════════════════════════════════════════════════════
# AYARLAR
# ══════════════════════════════════════════════════════
PARALEL_TARAYICI = 7
SAYFA_BEKLEME    = 5
MAX_DENEME       = 2
TXT_DOSYA        = "hisseisimleri.txt"
EXCEL_DOSYA      = "IsYatirim_Guncel.xlsx"

cekilen_sayisi = 0
sayac_kilidi   = threading.Lock()

OZET_SUTUNLAR = [
    ("Kod",                  "kod"),
    ("Şirket Adı",           "sirket_adi"),
    ("F/K",                  "cari.F/K"),
    ("FD/FAVÖK",             "cari.FD/FAVÖK"),
    ("PD/DD",                "cari.PD/DD"),
    ("Piyasa Değeri (mnTL)", "cari.Piyasa Değeri"),
    ("Net Kâr (mnTL)",       "mali_ozet.Net Kâr"),
    ("Faaliyet Alanı",       "kunye.Faal Alanı"),
    ("Adres",                "kunye.Adres"),
]

# ══════════════════════════════════════════════════════
# YARDIMCI
# ══════════════════════════════════════════════════════

def temizle(x):
    return " ".join(str(x).split()).strip() if x else ""

def sayi(x):
    if x is None:
        return None
    s = str(x).strip().replace(".", "").replace(",", ".")
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
        print(f"HATA: '{dosya}' bulunamadı!")
        return []
    with open(dosya, "r", encoding="utf-8") as f:
        kodlar = [l.strip().upper() for l in f if l.strip()]
    if not kodlar:
        print(f"HATA: '{dosya}' boş!")
    return kodlar

# ══════════════════════════════════════════════════════
# CHROME KURULUMU
# ══════════════════════════════════════════════════════

def chrome_ve_driver_bul():
    """
    GitHub Actions'da browser-actions/setup-chrome hem 'chrome' hem
    'chromedriver'ı PATH'e ekler. Sistem PATH'i öncelikli, yoksa
    webdriver-manager ile indir.
    """
    # ChromeDriver
    driver_path = shutil.which("chromedriver")
    if driver_path:
        print(f"  [chromedriver] PATH'ten bulundu: {driver_path}")
    else:
        print("  [chromedriver] PATH'te yok, webdriver-manager ile indiriliyor...")
        driver_path = ChromeDriverManager().install()
        print(f"  [chromedriver] İndirildi: {driver_path}")

    # Chrome binary
    chrome_bin = (
        shutil.which("chrome") or
        shutil.which("google-chrome") or
        shutil.which("google-chrome-stable") or
        shutil.which("chromium-browser") or
        shutil.which("chromium")
    )
    if chrome_bin:
        print(f"  [chrome]       PATH'ten bulundu: {chrome_bin}")
    else:
        print("  [chrome]       PATH'te bulunamadı, Selenium varsayılanı kullanılacak.")

    return driver_path, chrome_bin


def chrome_olustur(driver_path, chrome_bin=None):
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    if chrome_bin:
        opts.binary_location = chrome_bin
    svc = Service(driver_path)
    return webdriver.Chrome(service=svc, options=opts)

# ══════════════════════════════════════════════════════
# SAYFA ÇEKME
# ══════════════════════════════════════════════════════

def sayfa_cek(driver, kod):
    url = (
        "https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/"
        f"sirket-karti.aspx?hisse={kod}"
    )
    for deneme in range(1, MAX_DENEME + 1):
        try:
            driver.get(url)
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table"))
            )
            time.sleep(SAYFA_BEKLEME)
            return driver.page_source
        except Exception as e:
            print(f"\n  [UYARI] {kod} deneme {deneme}/{MAX_DENEME}: {e}")
            if deneme < MAX_DENEME:
                time.sleep(3)
    # Son denemede ne varsa al
    try:
        return driver.page_source
    except Exception:
        return ""

# ══════════════════════════════════════════════════════
# HTML AYRIŞTIRMA
# ══════════════════════════════════════════════════════

def ayristir(html, kod):
    soup = BeautifulSoup(html, "html.parser")
    kayit = {
        "kod":       kod,
        "sirket_adi": kod,
        "kunye":     {},
        "cari":      {},
        "mali_ozet": {},
    }

    # Şirket adı
    for sel in ["h1", "h2", ".company-name", ".card-title"]:
        el = soup.select_one(sel)
        if el:
            ad = temizle(el.get_text())
            if ad and ad.upper() != kod:
                kayit["sirket_adi"] = ad
                break

    # Tablo taraması
    for tablo in soup.find_all("table"):
        for satir in tablo.find_all("tr"):
            hucreler = satir.find_all(["td", "th"])
            if len(hucreler) < 2:
                continue
            k = temizle(hucreler[0].get_text())
            v = temizle(hucreler[1].get_text())
            if not k or not v:
                continue

            if   re.search(r"F\s*/\s*K",        k, re.I): kayit["cari"]["F/K"]           = sayi(v)
            elif re.search(r"FD\s*/\s*FAVÖK",   k, re.I): kayit["cari"]["FD/FAVÖK"]      = sayi(v)
            elif re.search(r"PD\s*/\s*DD",      k, re.I): kayit["cari"]["PD/DD"]         = sayi(v)
            elif re.search(r"Piyasa\s*De[ğg]", k, re.I): kayit["cari"]["Piyasa Değeri"] = sayi(v)
            elif re.search(r"Net\s*K[aâ]r",    k, re.I): kayit["mali_ozet"]["Net Kâr"]  = sayi(v)
            elif re.search(r"Faal\s*Alan",      k, re.I): kayit["kunye"]["Faal Alanı"]   = v
            elif re.search(r"Adres",            k, re.I): kayit["kunye"]["Adres"]        = v

    if not kayit["cari"] and not kayit["mali_ozet"]:
        snippet = soup.get_text()[:300].replace("\n", " ")
        print(f"\n  [DEBUG] {kod} veri YOK — sayfa: {snippet}")

    return kayit

# ══════════════════════════════════════════════════════
# WORKER
# ══════════════════════════════════════════════════════

DEBUG_HTML_KAYDEDILDI = False
debug_kilidi = threading.Lock()

def worker_calis(kodlar_q, sonuclar, toplam, worker_id, driver_path, chrome_bin):
    global cekilen_sayisi, DEBUG_HTML_KAYDEDILDI
    driver = None
    try:
        driver = chrome_olustur(driver_path, chrome_bin)
    except Exception as e:
        print(f"\n  [HATA] Worker-{worker_id} Chrome açılamadı: {e}")
        # Kuyruktaki tüm işleri hata olarak işaretle
        while not kodlar_q.empty():
            try:
                sira, kod = kodlar_q.get_nowait()
                sonuclar[sira] = {"kod": kod, "sirket_adi": f"Chrome hatası: {e}"}
                with sayac_kilidi:
                    global cekilen_sayisi
                    cekilen_sayisi += 1
                kodlar_q.task_done()
            except queue.Empty:
                break
        return

    try:
        while True:
            try:
                sira, kod = kodlar_q.get_nowait()
            except queue.Empty:
                break

            try:
                html  = sayfa_cek(driver, kod)
                # İlk hissenin HTML'ini debug için kaydet
                with debug_kilidi:
                    if not DEBUG_HTML_KAYDEDILDI:
                        DEBUG_HTML_KAYDEDILDI = True
                        with open("debug_sayfa.html", "w", encoding="utf-8") as fh:
                            fh.write(html)
                        print(f"\n  [DEBUG] {kod} HTML kaydedildi: debug_sayfa.html ({len(html)} karakter)")
                kayit = ayristir(html, kod)
                sonuclar[sira] = kayit
            except Exception as e:
                print(f"\n  [HATA] Worker-{worker_id} | {kod}: {e}")
                sonuclar[sira] = {"kod": kod, "sirket_adi": f"HATA: {e}"}

            with sayac_kilidi:
                cekilen_sayisi += 1
                yuzde = (cekilen_sayisi / toplam) * 100
                sys.stdout.write(
                    f"\r  >> %{yuzde:.1f} | {cekilen_sayisi}/{toplam} ({kod}){'':10}"
                )
                sys.stdout.flush()
            kodlar_q.task_done()
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

# ══════════════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════════════

def excel_yaz(veri_listesi, dosya_adi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hisse_Verileri"

    for ci, (baslik, _) in enumerate(OZET_SUTUNLAR, 1):
        ws.cell(1, ci, baslik).font = Font(bold=True)

    bos = 0
    for ri, kayit in enumerate(veri_listesi, 2):
        if not kayit:
            bos += 1
            continue
        for ci, (_, yol) in enumerate(OZET_SUTUNLAR, 1):
            ws.cell(ri, ci, deger_al(kayit, yol))

    wb.save(dosya_adi)
    dolu = len(veri_listesi) - bos
    print(f"\n  Excel kaydedildi: {dosya_adi} ({dolu} dolu / {bos} boş kayıt)")

# ══════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════

def main():
    kodlar = oku_txt(TXT_DOSYA)
    if not kodlar:
        return

    toplam = len(kodlar)
    print(f"\n{toplam} hisse için işlem başladı ({PARALEL_TARAYICI} paralel tarayıcı)\n")

    # ChromeDriver ve Chrome binary'yi TEK SEFERLIK bul/indir
    driver_path, chrome_bin = chrome_ve_driver_bul()
    print()

    kodlar_q = queue.Queue()
    for i, k in enumerate(kodlar):
        kodlar_q.put((i, k))

    sonuclar = [None] * toplam
    threads  = []
    for wid in range(1, min(PARALEL_TARAYICI, toplam) + 1):
        t = threading.Thread(
            target=worker_calis,
            args=(kodlar_q, sonuclar, toplam, wid, driver_path, chrome_bin),
            daemon=True,
        )
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    excel_yaz(sonuclar, EXCEL_DOSYA)
    print(f"Tamamlandı → {EXCEL_DOSYA}")

if __name__ == "__main__":
    main()
