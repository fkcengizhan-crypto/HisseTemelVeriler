import sys, time, re, os, queue, threading
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
SAYFA_BEKLEME   = 3
MAX_DENEME      = 2
TXT_DOSYA       = "hisseisimleri.txt"

# Sayaç için kilit ve değişken
cekilen_sayisi = 0
sayac_kilidi = threading.Lock()

OZET_SUTUNLAR = [
    ("Kod", "kod"), ("Şirket Adı", "sirket_adi"), ("F/K", "cari.F/K"),
    ("FD/FAVÖK", "cari.FD/FAVÖK"), ("PD/DD", "cari.PD/DD"),
    ("Piyasa Değeri (mnTL)", "cari.Piyasa Değeri"), ("Net Kâr (mnTL)", "mali_ozet.Net Kâr"),
    ("Faaliyet Alanı", "kunye.Faal Alanı"), ("Adres", "kunye.Adres")
]

# ══════════════════════════════════════════════════════
# FONKSİYONLAR
# ══════════════════════════════════════════════════════

def temizle(x): return " ".join(str(x).split()).strip() if x else ""

def sayi(x):
    if x is None: return None
    s = str(x).strip().replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try: return float(s)
    except: return None

def deger_al(kayit, yol):
    parcalar = yol.split(".", 1)
    val = kayit.get(parcalar[0])
    if len(parcalar) == 1: return val
    if isinstance(val, dict): return deger_al(val, parcalar[1])
    return None

def oku_txt(dosya):
    if not os.path.exists(dosya): return []
    with open(dosya, "r", encoding="utf-8") as f:
        return [l.strip().upper() for l in f if l.strip()]

def chrome_olustur():
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-gpu")
    svc = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=svc, options=opts)

def sayfa_cek(driver, kod):
    url = f"https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/sirket-karti.aspx?hisse={kod}"
    driver.get(url)
    try: WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
    except: pass
    time.sleep(SAYFA_BEKLEME)
    return driver.page_source

def ayristir(html, kod):
    soup = BeautifulSoup(html, "html.parser")
    kayit = {"kod": kod, "sirket_adi": kod, "kunye": {}, "cari": {}, "mali_ozet": {}}
    
    # Basit bir ayrıştırma mantığı (Önceki kodun özeti)
    h1 = soup.find("h1")
    if h1: kayit["sirket_adi"] = temizle(h1.get_text())
    
    for tablo in soup.find_all("table"):
        icerik = str(tablo)
        for s in tablo.find_all("tr"):
            td = s.find_all(["td", "th"])
            if len(td) >= 2:
                k, v = temizle(td[0].get_text()), temizle(td[1].get_text())
                if "F/K" in k: kayit["cari"]["F/K"] = sayi(v)
                elif "PD/DD" in k: kayit["cari"]["PD/DD"] = sayi(v)
                elif "Piyasa Değeri" in k: kayit["cari"]["Piyasa Değeri"] = sayi(re.sub(r"[^\d]", "", v))
                elif "Net Kâr" in k: kayit["mali_ozet"]["Net Kâr"] = sayi(v)
                elif "Faal Alanı" in k: kayit["kunye"]["Faal Alanı"] = v
                elif "Adres" in k: kayit["kunye"]["Adres"] = v
    return kayit

# ══════════════════════════════════════════════════════
# WORKER VE SAYAÇ MANTIĞI
# ══════════════════════════════════════════════════════

def worker_calis(kodlar_q, sonuclar, toplam, worker_id):
    global cekilen_sayisi
    driver = None
    try:
        driver = chrome_olustur()
        while not kodlar_q.empty():
            try: sira, kod = kodlar_q.get_nowait()
            except queue.Empty: break
            
            kayit = ayristir(sayfa_cek(driver, kod), kod)
            sonuclar[sira] = kayit
            
            # Sayaç Güncelleme
            with sayac_kilidi:
                cekilen_sayisi += 1
                yuzde = (cekilen_sayisi / toplam) * 100
                # \r karakteri satırı başa sarar, end="" satır atlamayı engeller
                sys.stdout.write(f"\r  >> İlerleme: %{yuzde:.1f} | {cekilen_sayisi}/{toplam} hisse tamamlandı ({kod}){' ' * 10}")
                sys.stdout.flush()
                
            kodlar_q.task_done()
    finally:
        if driver: driver.quit()

def excel_yaz(veri_listesi, dosya_adi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hisse_Verileri"
    
    # Başlıklar
    for ci, (baslik, _) in enumerate(OZET_SUTUNLAR, 1):
        ws.cell(1, ci, baslik).font = Font(bold=True)
    
    # Veriler
    for ri, kayit in enumerate(veri_listesi, 2):
        if not kayit: continue
        for ci, (_, yol) in enumerate(OZET_SUTUNLAR, 1):
            ws.cell(ri, ci, deger_al(kayit, yol))
    
    wb.save(dosya_adi)

def main():
    kodlar = oku_txt(TXT_DOSYA)
    if not kodlar: return
    toplam = len(kodlar)
    
    print(f"\n{toplam} hisse için işlem başladı...")
    
    kodlar_q = queue.Queue()
    for i, k in enumerate(kodlar): kodlar_q.put((i, k))
    
    sonuclar = [None] * toplam
    threads = []
    for wid in range(1, min(PARALEL_TARAYICI, toplam) + 1):
        t = threading.Thread(target=worker_calis, args=(kodlar_q, sonuclar, toplam, wid))
        t.start()
        threads.append(t)
    
    for t in threads: t.join()
    
    excel_yaz(sonuclar, "IsYatirim_Guncel.xlsx")
    print(f"\n\nİşlem başarıyla tamamlandı. Excel dosyası oluşturuldu.")

if __name__ == "__main__":
    main()