"""Web scraper for vendor SYSMOD / PTF information.

* scrape_broadcom()  - Broadcom's public mainframe compatibility pages (static HTML, no login):
                       per product: z/OS 2.5/3.1/3.2 status (YES/NO/TBD) + required PTFs per release.
* scrape_generic()   - any other vendor page: tables / PTF-looking ids, pagination, detail pages.
Login / keys: basic, bearer token, API-key header, cookie, or form login (see Fetcher).
Optional render=True uses Playwright for JavaScript sites (IBM SPCR) - install separately.
"""
from __future__ import annotations
import re, time
from urllib.parse import urljoin, urlparse
import requests
from bs4 import BeautifulSoup
import pandas as pd
import kb, docs

ID_RE = docs.ID_RE
UA = "Mozilla/5.0 (compatible; SYSMOD-Compat-Bot)"
LOGIN_HINT = re.compile(r"login|signin|sign-in|sso|oauth", re.I)
NEXT_TXT = {"next", "next >", "next »", "›", "»", ">", "next page"}
IN_SCOPE = set(kb.ORDER)
BROADCOM_BASE = "https://ftpdocs.broadcom.com/phpdocs/0/MSPSaccount/COMPAT/{}.HTML"
# Broadcom page codes seen on the public site (add more: any <code>.HTML under .../COMPAT/)
BROADCOM_SEED = ["AOX", "ASF", "PDS", "VEX", "ANS", "DAT", "LBR", "BRI", "SPL", "SOP", "ADB", "PMO",
                 "AL0", "AL7", "EJT", "BAC", "TAM", "BRR", "SAC", "FHR", "MDN", "BTW", "V4M"]
COMPAT_COLS = ["Vendor", "Product", "Version", "z/OS Release", "Compatible", "PTF Required?", "Source URL", "Status"]

class Fetcher:
    """auth: None or {type: basic|bearer|apikey|cookie|form, ...}.
    apikey: {type:'apikey', header:'X-API-Key', value:'...'}   MFA logins can't be automated -> use cookie."""
    def __init__(self, auth=None, delay=1.0, timeout=30, render=False):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = UA
        self.delay, self.timeout, self.render = delay, timeout, render
        self.auth = auth or {}
        t = self.auth.get("type")
        if t == "basic":
            self.s.auth = (self.auth.get("user", ""), self.auth.get("password", ""))
        elif t == "bearer":
            self.s.headers["Authorization"] = f"Bearer {self.auth.get('token', '')}"
        elif t == "apikey":
            self.s.headers[self.auth.get("header", "X-API-Key")] = self.auth.get("value", "")
        elif t == "cookie":
            for part in self.auth.get("cookie", "").split(";"):
                if "=" in part:
                    k, v = part.strip().split("=", 1)
                    self.s.cookies.set(k, v)
        self._logged = False

    def login(self):
        a = self.auth
        if a.get("type") != "form" or self._logged:
            return
        data = {a.get("user_field", "username"): a.get("user", ""), a.get("pass_field", "password"): a.get("password", "")}
        data.update(a.get("extra", {}))
        try:
            self.s.post(a["login_url"], data=data, timeout=self.timeout)
        except Exception:
            pass
        self._logged = True

    def _render(self, url):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            return "Error: Playwright not installed (pip install playwright && playwright install chromium)", "", url
        with sync_playwright() as p:
            b = p.chromium.launch()
            ctx = b.new_context(user_agent=UA)
            if self.s.cookies:
                host = urlparse(url).hostname
                ctx.add_cookies([{"name": c.name, "value": c.value, "domain": host, "path": "/"} for c in self.s.cookies])
            pg = ctx.new_page(); pg.goto(url, wait_until="networkidle", timeout=60000)
            html = pg.content(); b.close()
        return "OK", html, url

    def get(self, url):
        """-> (status, html, final_url); status: OK / Access denied / Not found in source / Error: ..."""
        time.sleep(self.delay)
        if self.render:
            return self._render(url)
        try:
            r = self.s.get(url, timeout=self.timeout)
        except Exception as e:
            return f"Error: {type(e).__name__}", "", url
        if r.status_code in (401, 403):
            return "Access denied", "", r.url
        if r.status_code == 404:
            return "Not found in source", "", r.url
        if r.status_code >= 400:
            return f"Error: HTTP {r.status_code}", "", r.url
        html = r.text
        path = urlparse(r.url).path + urlparse(r.url).query
        if (LOGIN_HINT.search(path) and not LOGIN_HINT.search(url)) or \
           (re.search(r'type=["\']password["\']', html, re.I) and "<table" not in html.lower()):
            return "Access denied", "", r.url
        return "OK", html, r.url

# ===================================================================== Broadcom
def _rel(label):
    m = re.search(r"z/OS\.?\s*V(\d)R(\d)", label, re.I)
    return f"{m.group(1)}.{m.group(2)}" if m else ""

def parse_broadcom_page(html, url):
    """-> (compat rows, sysmod rows) for one Broadcom compatibility page."""
    soup = BeautifulSoup(html, "html.parser")
    h1 = soup.find("h1")
    product = h1.get_text(" ", strip=True) if h1 else ""
    compat, versions = [], []
    for t in soup.find_all("table"):
        trs = t.find_all("tr")
        rows = [[c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])] for tr in trs]
        if not any(re.match(r"z/os\s*v\dr\d", r[0], re.I) for r in rows if r):
            continue
        head = next((r for r in rows if r and r[0].lower().startswith("compatibility")), [])
        versions = [h for h in head[1:] if h]
        for r in rows:
            rel = _rel(r[0]) if r else ""
            if rel not in IN_SCOPE:
                continue
            for i, cell in enumerate(r[1:]):
                st = re.search(r"\b(YES|NO|TBD|N/A)\b", cell, re.I)
                if not st or (versions and i >= len(versions)):
                    continue
                compat.append({"Vendor": "Broadcom Inc.", "Product": product,
                               "Version": versions[i] if versions else "", "z/OS Release": rel,
                               "Compatible": st.group(1).upper(), "ptf_link": "ptf" in cell.lower()})
    # fix-category blocks -> PTFs per z/OS release
    raw = [l.strip() for l in soup.get_text("\n").splitlines() if l.strip()]
    lines, k = [], 0
    while k < len(raw):                       # "Fix Category:" and its name can land on separate lines
        if raw[k].endswith("Fix Category:") and k + 1 < len(raw):
            lines.append(raw[k] + " " + raw[k + 1]); k += 2
        else:
            lines.append(raw[k]); k += 1
    sysmods, cur, comp, ver = [], "", "", ""
    for i, l in enumerate(lines):
        m = re.search(r"Fix Category:\s*`?([\w./\-]+)", l)
        if m:
            cur = _rel(m.group(1)) if re.search(r"z/OS\.V\d", m.group(1)) else ""
            continue
        if i + 1 < len(lines) and lines[i + 1].startswith("Fix Category") and re.search(r"\s(\d+\.\d+)$", l):
            ver = re.search(r"\s(\d+\.\d+)$", l).group(1)
        c = re.search(r"\(([A-Z0-9]{7})\)", l)
        if c:
            comp = c.group(1)
        if cur in IN_SCOPE:
            for sid in ID_RE.findall(l):
                if sid == comp or f"({sid})" in l:       # a product component / FMID, not a PTF
                    continue
                sysmods.append({"Vendor": "Broadcom Inc.", "Product": product, "Product Version": ver if len(versions) != 1 else versions[0],
                                "z/OS Release": cur, "SYSMOD ID": sid, "Type": "PTF", "FMID": comp, "APAR": "",
                                "Publish Date": "", "Prerequisites": "", "Superseded By": "", "Description": l[:250],
                                "Source URL": url, "Status": "Verified (vendor compatibility page)"})
    sm = pd.DataFrame(sysmods, columns=list(kb.ALIASES)).drop_duplicates(["SYSMOD ID", "z/OS Release", "Product Version"])
    if len(sm):          # one row per PTF, listing every z/OS release it is required for
        agg = {c: "first" for c in sm.columns if c not in ("Product", "Product Version", "SYSMOD ID", "z/OS Release")}
        agg["z/OS Release"] = lambda x: ", ".join(sorted(set(x)))
        sm = sm.groupby(["Vendor", "Product", "Product Version", "SYSMOD ID"], as_index=False, sort=False).agg(agg)[list(kb.ALIASES)]
    cp = pd.DataFrame(compat)
    if len(cp):
        has = {(r["z/OS Release"]) for _, r in sm.iterrows()}
        cp["PTF Required?"] = ["Yes" if (row["ptf_link"] or row["z/OS Release"] in has) else
                               ("No" if row["Compatible"] == "YES" else "") for _, row in cp.iterrows()]
        cp["Source URL"], cp["Status"] = url, "Verified (vendor compatibility page)"
        cp = cp[COMPAT_COLS]
    else:
        cp = pd.DataFrame(columns=COMPAT_COLS)
    return cp, sm

def discover_codes(index_html):
    """Page codes (…/COMPAT/XXX.HTML) found in any index page you give it."""
    return sorted(set(re.findall(r"COMPAT/([A-Za-z0-9]{3})\.HTML", index_html, re.I)) - {"All"})

def scrape_broadcom(codes, fetcher, base=BROADCOM_BASE, progress=None):
    cps, sms, log = [], [], []
    for n, code in enumerate(codes):
        url = base.format(code.strip().upper())
        st, html, final = fetcher.get(url)
        if st != "OK":
            log.append({"URL": url, "Result": st, "Product": "", "Compat rows": 0, "SYSMODs": 0})
        else:
            cp, sm = parse_broadcom_page(html, url)
            cps.append(cp); sms.append(sm)
            log.append({"URL": url, "Result": "OK", "Product": (cp["Product"].iloc[0] if len(cp) else (sm["Product"].iloc[0] if len(sm) else "")),
                        "Compat rows": len(cp), "SYSMODs": len(sm)})
        if progress:
            progress((n + 1) / len(codes))
    cp = pd.concat(cps, ignore_index=True) if cps else pd.DataFrame(columns=COMPAT_COLS)
    sm = pd.concat(sms, ignore_index=True) if sms else pd.DataFrame(columns=list(kb.ALIASES))
    return cp, sm, pd.DataFrame(log)

# ===================================================================== generic sites
def _label(text, names):
    m = re.search(rf"(?:{names})\s*[:\-]\s*([^\n]{{1,300}})", text, re.I)
    return m.group(1).strip() if m else ""

def _page_release(soup):
    head = " ".join(x.get_text(" ") for x in soup.find_all(["title", "h1", "h2"])[:4])
    return ",".join(kb.parse_releases(head))

def parse_tables(soup, base_url):
    rows, page_rel = [], _page_release(soup)
    for t in soup.find_all("table"):
        trs = t.find_all("tr")
        if len(trs) < 2:
            continue
        head = [c.get_text(" ", strip=True) for c in trs[0].find_all(["th", "td"])]
        data = []
        for tr in trs[1:]:
            cells = tr.find_all(["td", "th"])
            if len(cells) == len(head):
                data.append((cells, [c.get_text(" ", strip=True) for c in cells]))
        if not data:
            continue
        m = kb._map_columns(pd.DataFrame([d[1] for d in data], columns=head))
        if "SYSMOD ID" not in m:
            best = max(range(len(head)), key=lambda i: sum(bool(ID_RE.search(r[i])) for _, r in data))
            if not sum(bool(ID_RE.search(r[best])) for _, r in data):
                continue
            m["SYSMOD ID"] = head[best]
        idx = head.index(m["SYSMOD ID"])
        for cells, vals in data:
            rec = {c: (vals[head.index(m[c])] if c in m and m[c] in head else "") for c in kb.ALIASES}
            if not rec["SYSMOD ID"]:
                continue
            a = cells[idx].find("a", href=True)
            rec["Source URL"] = urljoin(base_url, a["href"]) if a else base_url
            rec["z/OS Release"] = rec["z/OS Release"] or ",".join(kb.parse_releases(" ".join(vals))) or page_rel
            rec["Status"] = "Verified"
            rows.append(rec)
    return rows

def parse_text_ids(soup, base_url):
    rows, seen, page_rel = [], set(), _page_release(soup)
    for line in soup.get_text("\n").splitlines():
        for pid in ID_RE.findall(line):
            if pid not in seen:
                seen.add(pid)
                rows.append({**{c: "" for c in kb.ALIASES}, "SYSMOD ID": pid, "Description": line.strip()[:300],
                             "Source URL": base_url, "Status": "Pattern match - verify",
                             "z/OS Release": ",".join(kb.parse_releases(line)) or page_rel})
    return rows

def parse_detail(html):
    text = BeautifulSoup(html, "html.parser").get_text("\n")
    return {"FMID": _label(text, "FMID"), "APAR": _label(text, "APAR(?: ID)?"),
            "Prerequisites": _label(text, "Prerequisites?|Prereqs?"),
            "Superseded By": _label(text, "Superseded by|Superseding PTF|Supersedes"),
            "Publish Date": _label(text, "Publish(?:ed)? date|Release date|Date"),
            "Description": _label(text, "Description|Abstract|Title")}

def next_links(soup, base_url):
    out = []
    for a in soup.find_all("a", href=True):
        txt = a.get_text(" ", strip=True).lower()
        rel = " ".join(a.get("rel", [])) if a.get("rel") else ""
        href = urljoin(base_url, a["href"])
        if "next" in rel or txt in NEXT_TXT or "next" in (a.get("aria-label") or "").lower() \
                or re.search(r"[?&](page|p|start|offset)=\d+", href):
            out.append(href)
    return out

def scrape_generic(urls, fetcher, max_pages=10, detail_limit=50):
    fetcher.login()
    log, rows = [], []
    for start in urls:
        host, queue, seen, pages, details = urlparse(start).netloc, [start], set(), 0, 0
        while queue and pages < max_pages:
            url = queue.pop(0)
            if url in seen:
                continue
            seen.add(url); pages += 1
            status, html, final = fetcher.get(url)
            if status != "OK":
                log.append({"URL": url, "Result": status, "Rows": 0}); continue
            soup = BeautifulSoup(html, "html.parser")
            found = parse_tables(soup, final) or parse_text_ids(soup, final)
            for rec in found:
                u = rec["Source URL"]
                if details < detail_limit and u and u != final and urlparse(u).netloc == host and u not in seen:
                    seen.add(u); details += 1
                    st, h2, _ = fetcher.get(u)
                    if st == "OK":
                        for k, v in parse_detail(h2).items():
                            if v and not rec.get(k):
                                rec[k] = v
                    elif st == "Access denied":
                        rec["Status"] = "Access denied"
            rows += found
            log.append({"URL": url, "Result": "OK", "Rows": len(found)})
            queue += [n for n in next_links(soup, final) if urlparse(n).netloc == host and n not in seen]
    df = pd.DataFrame(rows, columns=list(kb.ALIASES)) if rows else pd.DataFrame(columns=list(kb.ALIASES))
    return df.drop_duplicates(subset=["SYSMOD ID", "Source URL"]).reset_index(drop=True), pd.DataFrame(log)

# ===================================================================== fill the user's workbook
def fill_workbook(path, compat, release="3.2"):
    """Fill 'R3.2 Compatible', 'PTF Required?', link and 'available without login' columns of the ISV workbook
    for every row whose product has scraped compatibility data. Returns xlsx bytes."""
    import io, openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    hdr = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value}
    col = lambda start: next((v for k, v in hdr.items() if k.startswith(start)), None)
    cV, cP, cVer = col("vendor"), col("product"), col("version")
    cC, cT, cL, cI = col("r3.2 compatible"), col("if yes, ptf"), col("link to verify"), col("information available")
    sub = compat[compat["z/OS Release"] == release] if len(compat) else compat
    n = 0
    for r in range(2, ws.max_row + 1):
        vend, prod = ws.cell(r, cV).value, ws.cell(r, cP).value
        if not prod:
            continue
        pt, vt = kb.product_tokens(prod), kb.vendor_tokens(vend)
        ver = ws.cell(r, cVer).value if cVer else ""
        for _, c in sub.iterrows():
            if (kb.vendor_tokens(c["Vendor"]) & vt) and pt <= kb.product_tokens(c["Product"]):
                if ver and c["Version"] and kb._ver_num(ver) != kb._ver_num(c["Version"]):
                    continue
                ws.cell(r, cC).value = {"YES": "Y", "NO": "N"}.get(c["Compatible"], c["Compatible"])
                ws.cell(r, cT).value = c["PTF Required?"]
                ws.cell(r, cL).value = c["Source URL"]
                ws.cell(r, cI).value = "Yes"
                n += 1
                break
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue(), n
