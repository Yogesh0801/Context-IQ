# pip install selenium webdriver-manager requests pdfplumber python-docx
# optional (OCR for posters/images): pip install pytesseract pillow  (+ install Tesseract app)
import os, io, re, json, time
from urllib.parse import urljoin, urlparse, urldefrag, unquote
from collections import deque

import requests
import pdfplumber
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---------------- CONFIG ----------------
START_URL     = "https://www.data.gov.in/catalogs"   # catalog listing, not the home page
ALLOWED_HOST  = "data.gov.in"

# Output location (outside the scraping folder)
OUTPUT_JSON = os.path.join(os.path.expanduser("~"), "scraped_data", "datagov_data.json")
DEBUG_DIR   = os.path.join(os.path.expanduser("~"), "scraped_data", "debug")

MAX_PAGES     = 200
DELAY         = 2
HEADLESS      = False         # False = visible browser (less likely to be blocked)
WAIT_TIMEOUT  = 20            # seconds to wait for links to render
MAX_CHARS     = 200_000       # cap text per file
CSV_MAX_ROWS  = 200           # rows kept per CSV
USE_OCR       = False         # True to read text from images/posters
SAVE_EMPTY    = False         # True = also keep pages with no description/files (for testing)
DEBUG_PAGES   = 3             # save screenshot + html for the first N pages (0 to disable)

# Optional: only queue internal pages whose URL matches this regex. None = queue all.
# Example: PAGE_FILTER = re.compile(r"/(catalog|catalogs|resource|dataset)", re.I)
PAGE_FILTER   = None

TEXT_EXTS     = {"pdf", "doc", "docx", "txt", "csv", "json", "xml"}
IMAGE_EXTS    = {"jpg", "jpeg", "png", "webp"}
SKIP_PATTERNS = re.compile(r"(logout|login|signin|register|javascript:|mailto:|tel:)", re.I)
# ----------------------------------------


def make_driver():
    o = Options()
    if HEADLESS:
        o.add_argument("--headless=new")
    o.add_argument("--no-sandbox")
    o.add_argument("--disable-dev-shm-usage")
    o.add_argument("--window-size=1400,1000")
    o.add_argument("--disable-blink-features=AutomationControlled")
    o.add_experimental_option("excludeSwitches", ["enable-automation"])
    o.add_experimental_option("useAutomationExtension", False)
    o.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=o)
    try:
        d.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
    except Exception:
        pass
    return d


def get_ext(url):
    m = re.search(r"\.([a-zA-Z0-9]{2,5})$", unquote(urlparse(url).path))
    if m:
        return m.group(1).lower()
    q = re.search(r"[?&]format=([a-zA-Z0-9]+)", url)
    return q.group(1).lower() if q else None


def is_file(url):
    e = get_ext(url)
    return e in TEXT_EXTS or (USE_OCR and e in IMAGE_EXTS)


def scroll(driver, n=12):
    last = 0
    for _ in range(n):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.7)
        h = driver.execute_script("return document.body.scrollHeight")
        if h == last:
            break
        last = h


# ---------- page metadata (title, description, tags) ----------
def first_text(driver, selectors):
    for sel in selectors:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                t = el.text.strip()
                if t:
                    return t
        except Exception:
            pass
    return ""


def meta(driver, name):
    for attr in ("name", "property"):
        try:
            el = driver.find_element(By.CSS_SELECTOR, f"meta[{attr}='{name}']")
            v = el.get_attribute("content")
            if v:
                return v.strip()
        except Exception:
            pass
    return ""


def extract_metadata(driver):
    title = first_text(driver, ["h1", ".page-title", "[class*=title]"]) or driver.title.strip()

    description = first_text(driver, [
        ".field--name-body", "[class*=description]", "[class*=abstract]",
        "[class*=summary]", "article p", ".content p",
    ]) or meta(driver, "description") or meta(driver, "og:description")

    tags = set()
    for sel in ["a[rel=tag]", "[class*=tag] a", "[class*=tags] li", "[class*=keyword] a",
                "[class*=sector] a", "[class*=category] a", "[class*=badge]"]:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            t = el.text.strip()
            if t and len(t) < 60:
                tags.add(t)
    for k in meta(driver, "keywords").split(","):
        if k.strip():
            tags.add(k.strip())
    return title, description, sorted(tags)


def extract_links(driver, base):
    pages, files = set(), set()
    for tag, attr in [("a", "href"), ("iframe", "src"), ("embed", "src"),
                      ("source", "src"), ("img", "src")]:
        for el in driver.find_elements(By.CSS_SELECTOR, f"{tag}[{attr}]"):
            try:
                raw = el.get_attribute(attr)
                if not raw:
                    continue
                full = urldefrag(urljoin(base, raw))[0]
                if not full.startswith("http") or SKIP_PATTERNS.search(full):
                    continue
                if is_file(full):
                    files.add(full)
                elif tag == "a" and urlparse(full).netloc.endswith(ALLOWED_HOST):
                    if PAGE_FILTER is None or PAGE_FILTER.search(full):
                        pages.add(full)
            except Exception:
                pass
    return pages, files


# ---------- text extraction from files ----------
def text_from_bytes(ext, data):
    if ext == "pdf":
        parts = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for p in pdf.pages:
                parts.append(p.extract_text() or "")
        return "\n".join(parts)

    if ext == "docx":
        import docx
        d = docx.Document(io.BytesIO(data))
        txt = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for row in t.rows:
                txt.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(txt)

    if ext == "doc":
        return ""   # legacy .doc needs antiword/LibreOffice; skipped

    if ext == "csv":
        lines = data.decode("utf-8", errors="ignore").splitlines()
        return "\n".join(lines[:CSV_MAX_ROWS])

    if ext in {"txt", "json", "xml"}:
        return data.decode("utf-8", errors="ignore")

    if ext in IMAGE_EXTS and USE_OCR:
        import pytesseract
        from PIL import Image
        return pytesseract.image_to_string(Image.open(io.BytesIO(data)))
    return ""


def fetch_file_text(session, url, referer):
    ext = get_ext(url)
    try:
        r = session.get(url, headers={"Referer": referer}, timeout=60)
        if r.status_code != 200:
            return None
        if "text/html" in r.headers.get("Content-Type", "") and ext not in {"txt", "xml"}:
            return None   # login wall / blocked
        text = text_from_bytes(ext, r.content).strip()
        text = re.sub(r"\n{3,}", "\n\n", text)[:MAX_CHARS]
        return {"type": ext, "text": text} if text else None
    except Exception as e:
        print("   extract error:", e)
        return None


def make_session(driver):
    s = requests.Session()
    s.headers["User-Agent"] = driver.execute_script("return navigator.userAgent")
    for c in driver.get_cookies():
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"))
    return s


def save(records):
    os.makedirs(os.path.dirname(os.path.abspath(OUTPUT_JSON)), exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def save_debug(driver, n):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    driver.save_screenshot(os.path.join(DEBUG_DIR, f"debug_{n}.png"))
    with open(os.path.join(DEBUG_DIR, f"debug_{n}.html"), "w", encoding="utf-8") as fh:
        fh.write(driver.page_source)


def main():
    driver = make_driver()
    queue, seen_pages, seen_files = deque([START_URL]), set(), set()
    records = []

    try:
        while queue and len(seen_pages) < MAX_PAGES:
            url = queue.popleft()
            if url in seen_pages:
                continue
            seen_pages.add(url)
            n = len(seen_pages)
            print(f"[{n}/{MAX_PAGES}] {url}")

            try:
                driver.get(url)

                # wait for links to render instead of a blind sleep
                try:
                    WebDriverWait(driver, WAIT_TIMEOUT).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "a[href]")))
                except Exception:
                    print("   no links appeared within", WAIT_TIMEOUT, "seconds")
                time.sleep(DELAY)

                # first run only: pause so you can solve a captcha / log in manually
                # if n == 1: input("   Solve captcha / log in in the browser, then press Enter...")

                scroll(driver)
                title, description, tags = extract_metadata(driver)
                pages, files = extract_links(driver, driver.current_url)

                print("   title:", driver.title)
                print("   final url:", driver.current_url)
                print("   html length:", len(driver.page_source))
                print("   internal links:", len(pages), "| files:", len(files))

                if n <= DEBUG_PAGES:
                    save_debug(driver, n)
            except Exception as e:
                print("   page error:", e)
                continue

            queue.extend(p for p in pages if p not in seen_pages)

            session = make_session(driver)
            file_data = []
            for f in files:
                if f in seen_files:
                    continue
                seen_files.add(f)
                item = fetch_file_text(session, f, url)
                if item:
                    file_data.append(item)

            # keep only pages that actually carry data
            if file_data or description or SAVE_EMPTY:
                parsed = urlparse(url)
                records.append({
                    "title": title,
                    "path": parsed.path + (("?" + parsed.query) if parsed.query else ""),
                    "description": description,
                    "tags": tags,
                    "pdf_data": file_data,
                })
                save(records)   # save progressively so nothing is lost
            time.sleep(DELAY)
    finally:
        driver.quit()
        save(records)
        print(f"\nDone. {len(records)} records saved to {os.path.abspath(OUTPUT_JSON)}")
        print(f"Debug screenshots/HTML (if any): {DEBUG_DIR}")


if __name__ == "__main__":
    main()