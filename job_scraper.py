"""
Career-site job scraper (v3)
Reads career_sites.xlsx (columns: Company, Career URL), visits each site in a
headless browser, searches for jobs in Bangalore, keeps only IT-related ones,
and writes jobs_output.xlsx.

Setup (once):
    pip install playwright pandas openpyxl
    playwright install chromium
Run:
    python job_scraper.py
Tip: set HEADLESS = False below to WATCH the browser and see what it does.
"""
import os
import re
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---------------- SETTINGS ----------------
INPUT_FILE = "career_sites.xlsx"
OUTPUT_FILE = "jobs_output.xlsx"
SEARCH_LOCATION = "Bangalore"      # typed into the site's location box
SEARCH_KEYWORD = "software"        # typed into the site's keyword box
MAX_PAGES = 3                      # result pages to read per site
HEADLESS = True                    # False = show the browser window
SCREENSHOT_DIR = "screenshots"     # one screenshot per site, for debugging

LOCATION_KEYWORDS = ["bangalore", "bengaluru", "bengalooru"]
IT_KEYWORDS = [
    "software", "developer", "engineer", "java", "python", ".net", "devops",
    "cloud", "data", "analyst", "qa", "tester", "testing", "sap", "full stack",
    "frontend", "front end", "backend", "back end", "machine learning", "ai ",
    "cyber", "security", "network", "database", "dba", "sre", "architect",
    "technology", "programmer", "salesforce", "mainframe", "it support",
]
JOB_LINK_HINTS = ["job", "career", "position", "opening", "requisition", "apply", "vacanc"]
# ------------------------------------------

JS_EXTRACT = """
() => {
  // number of distinct "title-like" links inside a node
  const titled = n => new Set(
    [...n.querySelectorAll('a[href]')]
      .filter(x => (x.innerText || '').trim().length >= 10)
      .map(x => x.href)).size;
  const out = [];
  document.querySelectorAll('a[href]').forEach(a => {
    let node = a, ctx = a.innerText || '';
    for (let i = 0; i < 5 && node.parentElement; i++) {
      node = node.parentElement;
      if (titled(node) > 1) break;          // parent holds several jobs: stop
      const t = (node.innerText || '').trim();
      if (t.length > 500) break;
      ctx = t;
    }
    out.push({title: (a.innerText || '').trim(), href: a.href, context: ctx});
  });
  return out;
}
"""

JS_JOB_LINKS = """
() => {
  const re = /(search|find|explore|view|browse|open|current).{0,15}(jobs|roles|positions|openings)|^jobs$|job search|^careers? search$/i;
  return [...document.querySelectorAll('a[href]')]
    .filter(a => re.test((a.innerText || '').trim()))
    .map(a => a.href);
}
"""


def has_any(text, words):
    t = text.lower()
    return any(w in t for w in words)


# ---------- reading the Excel ----------
def load_sites():
    df = pd.read_excel(INPUT_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    print("Columns found in Excel:", list(df.columns))
    low = {c.lower(): c for c in df.columns}

    def pick(options):
        for o in options:
            for lc, orig in low.items():
                if o in lc:
                    return orig
        return None

    url_col = pick(["url", "link", "website", "site"])
    comp_col = pick(["company", "name", "organisation", "organization"])
    en_col = pick(["enabled", "active"])

    if url_col is None:  # no usable header: look for URL-like cells
        raw = pd.read_excel(INPUT_FILE, header=None)
        recs = []
        for _, r in raw.iterrows():
            for v in r.tolist():
                v = str(v).strip()
                if "." in v and " " not in v:
                    if not v.startswith("http"):
                        v = "https://" + v
                    name = re.sub(r"^https?://(www\.)?", "", v).split("/")[0].split(".")[0].title()
                    recs.append({"Company": name, "Career URL": v})
                    break
        return pd.DataFrame(recs)

    out = pd.DataFrame({
        "Company": df[comp_col] if comp_col else df[url_col].astype(str),
        "Career URL": df[url_col].astype(str).str.strip(),
    })
    if en_col:
        out = out[df[en_col].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])]
    out = out[out["Career URL"].str.contains(r"\.", na=False)]
    out["Career URL"] = out["Career URL"].apply(lambda u: u if u.startswith("http") else "https://" + u)
    return out


# ---------- browser helpers ----------
def dismiss_popups(page):
    pat = re.compile(r"^(accept( all)?( cookies)?|i agree|agree|got it|allow all|ok)$", re.I)
    try:
        page.get_by_role("button", name=pat).first.click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass


def first_visible(page, selector):
    try:
        loc = page.locator(selector)
        for i in range(min(loc.count(), 6)):
            el = loc.nth(i)
            if el.is_visible():
                return el
    except Exception:
        pass
    return None


def try_search(page):
    """If the page has a location / keyword search box, fill it and search."""
    loc_box = first_visible(
        page,
        "input[placeholder*='location' i], input[aria-label*='location' i], "
        "input[name*='location' i], input[id*='location' i], input[placeholder*='city' i]",
    )
    kw_box = first_visible(
        page,
        "input[placeholder*='keyword' i], input[placeholder*='job' i], "
        "input[placeholder*='search' i], input[type='search'], input[aria-label*='search' i]",
    )
    did = False
    try:
        if loc_box:
            loc_box.fill(SEARCH_LOCATION, timeout=3000)
            page.wait_for_timeout(1500)
            did = True
        if kw_box:
            kw_box.fill(SEARCH_KEYWORD, timeout=3000)
            kw_box.press("Enter")
            did = True
        elif loc_box:
            loc_box.press("Enter")
        if did:
            page.wait_for_timeout(5000)
    except Exception:
        pass
    return did


def scroll_page(page, rounds=5):
    for _ in range(rounds):
        page.mouse.wheel(0, 4000)
        page.wait_for_timeout(700)


def go_next_page(page):
    pat = re.compile(r"^(next|next page|›|>|»)$", re.I)
    for role in ("link", "button"):
        try:
            el = page.get_by_role(role, name=pat).first
            if el.is_visible() and el.is_enabled():
                el.click(timeout=3000)
                page.wait_for_timeout(3000)
                return True
        except Exception:
            continue
    return False


def collect(page, company, source_url):
    items = page.evaluate(JS_EXTRACT)
    jobs = []
    for it in items:
        title, href, ctx = it["title"], it["href"], it["context"]
        if not href.startswith("http") or len(title) < 6 or len(title) > 150:
            continue
        if (has_any(href, JOB_LINK_HINTS)
                and has_any(ctx, LOCATION_KEYWORDS)
                and has_any(title + " " + ctx, IT_KEYWORDS)):
            location = next((k.title() for k in LOCATION_KEYWORDS if k in ctx.lower()), "")
            jobs.append({
                "Company": company,
                "Job Title": re.sub(r"\s+", " ", title),
                "Location": location,
                "Job Link": href,
                "Source Page": source_url,
                "Fetched On": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
    return jobs, len(items)


def scan_current_page(page, company, source_url):
    try_search(page)
    scroll_page(page)
    jobs, seen = [], 0
    for _ in range(MAX_PAGES):
        j, n = collect(page, company, source_url)
        jobs += j
        seen += n
        if not go_next_page(page):
            break
        scroll_page(page, 2)
    return jobs, seen


def scrape_site(page, company, url):
    page.goto(url, timeout=60000, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    dismiss_popups(page)

    jobs, seen = scan_current_page(page, company, url)

    # Landing page had no jobs: follow "Search jobs" style links and try again
    if not jobs:
        try:
            targets = list(dict.fromkeys(page.evaluate(JS_JOB_LINKS)))[:3]
        except Exception:
            targets = []
        for t in targets:
            try:
                page.goto(t, timeout=60000, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
                dismiss_popups(page)
                j, n = scan_current_page(page, company, t)
                jobs += j
                seen += n
                if jobs:
                    break
            except Exception:
                continue

    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    try:
        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"{re.sub(r'[^A-Za-z0-9]+', '_', company)}.png"))
    except Exception:
        pass
    return jobs, seen, page.url


def main():
    sites = load_sites()
    print(f"{len(sites)} site(s) to scan\n")
    all_jobs, log = [], []
    cols = ["Company", "Job Title", "Location", "Job Link", "Source Page", "Fetched On"]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
            viewport={"width": 1366, "height": 900},
        )
        page = context.new_page()
        for _, row in sites.iterrows():
            company, url = row["Company"], row["Career URL"]
            print(f"Scanning {company} -> {url}")
            try:
                jobs, seen, final_url = scrape_site(page, company, url)
                all_jobs.extend(jobs)
                log.append({"Company": company, "Status": "OK", "Jobs Found": len(jobs),
                            "Links Scanned": seen, "Final URL": final_url})
                print(f"   {len(jobs)} matching jobs ({seen} links scanned)")
            except PWTimeout:
                log.append({"Company": company, "Status": "Timeout", "Jobs Found": 0,
                            "Links Scanned": 0, "Final URL": url})
                print("   timed out")
            except Exception as e:
                log.append({"Company": company, "Status": f"Error: {str(e)[:80]}", "Jobs Found": 0,
                            "Links Scanned": 0, "Final URL": url})
                print(f"   error: {str(e)[:80]}")
        browser.close()

    jobs_df = pd.DataFrame(all_jobs, columns=cols).drop_duplicates(subset=["Job Link"])
    log_df = pd.DataFrame(log)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as xw:
        jobs_df.to_excel(xw, sheet_name="Jobs", index=False)
        log_df.to_excel(xw, sheet_name="Run Log", index=False)
        ws = xw.sheets["Jobs"]
        for col, w in zip("ABCDEF", [20, 55, 14, 70, 45, 18]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        ws2 = xw.sheets["Run Log"]
        for col, w in zip("ABCDE", [20, 30, 12, 14, 70]):
            ws2.column_dimensions[col].width = w
    print(f"\nDone. {len(jobs_df)} jobs saved to {OUTPUT_FILE}")
    print(f"Screenshots of each site are in the '{SCREENSHOT_DIR}' folder.")


if __name__ == "__main__":
    main()
