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

def temizle(x):
    return " ".join(str(x).split()).strip() if x else ""

def sayi(x):
    if x is None: return None
    s = str(x).strip().replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try: return float(s)
    except ValueError: return None

def deger_al(kayit, yol):
    parcalar = yol.split(".", 1)
    val = kayit.get(parcalar[0])
    if len(parcalar) == 1: return val
    if isinstance(val, dict): return deger_al(val, parcalar[1])
    return None

def oku_txt(dosya):
    if not os.path.exists(dosya):
        print(f"HATA: '{dosya}' bulunamadi!")
        return []
    with open(dosya, "r", encoding="utf-8") as f:
        kodlar = [l.strip().upper() for l in f if l.strip()]
    if not kodlar:
        print(f"HATA: '{dosya}' bos!")
    return kodlar

def chrome_ve_driver_bul():
    driver_path = shutil.which("chromedriver")
    if driver_path:
        print(f"  [chromedriver] PATH: {driver_path}")
    else:
        print("  [chromedriver] webdriver-manager ile indiriliyor...")
        driver_path = ChromeDriverManager().install()
    chrome_bin = (
        shutil.which("chrome") or
        shutil.which("google-chrome") or
        shutil.which("google-chrome-stable") or
        shutil.which("chromium-browser") or
        shutil.which("chromium")
    )
    if chrome_bin:
        print(f"  [chrome]       PATH: {chrome_bin}")
    else:
        print("  [chrome]       Varsayilan kullanilacak.")
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
    return webdriver.Chrome(service=Service(driver_path), options=opts)

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
            print(f"\n  [UYARI] {kod} deneme {deneme}/{MAX_DENEME}: {type(e).__name__}")
            if deneme < MAX_DENEME:
                time.sleep(3)
    try:
        return driver.page_source
    except Exception:
        return ""

def ayristir(html, kod):
    soup = BeautifulSoup(html, "html.parser")
    kayit = {
        "kod":        kod,
        "sirket_adi": kod,
        "kunye":      {},
        "cari":       {},
        "mali_ozet":  {},
    }
    # Sirket adi - h1'den "|" oncesi
    h1 = soup.find("h1")
    if h1:
        kayit["sirket_adi"] = temizle(h1.get_text()).split("|")[0].strip()

    # Tablo tarama - TAM eslesme kullan (isyatirim basliklar sabit)
    for tablo in soup.find_all("table"):
        for satir in tablo.find_all("tr"):
            hucreler = satir.find_all(["td", "th"])
            if len(hucreler) < 2:
                continue
            k = temizle(hucreler[0].get_text())
            v = temizle(hucreler[1].get_text())
            if not k or not v:
                continue
            if   k == "F/K":           kayit["cari"]["F/K"]           = sayi(v)
            elif k == "FD/FAVOK" or k == "FD/FAVÖK": kayit["cari"]["FD/FAVÖK"] = sayi(v)
            elif k == "PD/DD":         kayit["cari"]["PD/DD"]         = sayi(v)
            elif k == "Piyasa Degeri" or k == "Piyasa Değeri":
                                       kayit["cari"]["Piyasa Değeri"] = sayi(v)
            elif k == "Net Kar" or k == "Net Kâr":
                                       kayit["mali_ozet"]["Net Kâr"]  = sayi(v)
            elif k == "Faal Alani" or k == "Faal Alanı":
                                       kayit["kunye"]["Faal Alanı"]   = v
            elif k == "Adres":         kayit["kunye"]["Adres"]        = v

    if not kayit["cari"]:
        snippet = soup.get_text()[:200].replace("\n", " ")
        print(f"\n  [DEBUG] {kod} veri YOK. Sayfa: {snippet}")

    return kayit

def worker_calis(kodlar_q, sonuclar, toplam, worker_id, driver_path, chrome_bin):
    global cekilen_sayisi
    driver = None
    try:
        driver = chrome_olustur(driver_path, chrome_bin)
    except Exception as e:
        print(f"\n  [HATA] Worker-{worker_id} Chrome acilmadi: {e}")
        while not kodlar_q.empty():
            try:
                sira, kod = kodlar_q.get_nowait()
                sonuclar[sira] = {"kod": kod, "sirket_adi": "Chrome hatasi"}
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
                kayit = ayristir(html, kod)
                sonuclar[sira] = kayit
            except Exception as e:
                print(f"\n  [HATA] Worker-{worker_id} | {kod}: {e}")
                sonuclar[sira] = {"kod": kod, "sirket_adi": "HATA"}
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
            try: driver.quit()
            except: pass

def excel_yaz(veri_listesi, dosya_adi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hisse_Verileri"
    for ci, (baslik, _) in enumerate(OZET_SUTUNLAR, 1):
        ws.cell(1, ci, baslik).font = Font(bold=True)
    dolu = bos = 0
    for ri, kayit in enumerate(veri_listesi, 2):
        if not kayit:
            bos += 1
            continue
        if not kayit.get("cari"):
            bos += 1
            ws.cell(ri, 1, kayit.get("kod", ""))
            ws.cell(ri, 2, kayit.get("sirket_adi", ""))
            continue
        dolu += 1
        for ci, (_, yol) in enumerate(OZET_SUTUNLAR, 1):
            ws.cell(ri, ci, deger_al(kayit, yol))
    wb.save(dosya_adi)
    print(f"\n  Excel: {dosya_adi} — {dolu} veri dolu / {bos} bos")

def main():
    kodlar = oku_txt(TXT_DOSYA)
    if not kodlar:
        return
    toplam = len(kodlar)
    print(f"\n{toplam} hisse isleniyor ({PARALEL_TARAYICI} paralel tarayici)\n")
    driver_path, chrome_bin = chrome_ve_driver_bul()
    print()
    kodlar_q = queue.Queue()
    for i, k in enumerate(kodlar):
        kodlar_q.put((i, k))
    sonuclar = [None] * toplam
    threads = []
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
    print(f"Tamamlandi -> {EXCEL_DOSYA}")

if __name__ == "__main__":
    main()
