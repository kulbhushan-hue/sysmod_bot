"""Core logic for the SYSMOD Compatibility Bot (pure Python + pandas, no external APIs)."""
from __future__ import annotations
import os, re, difflib
import pandas as pd

ORDER = ["2.5", "3.1", "3.2"]          # z/OS releases in scope (V2R5, V3R1, V3R2)
VENDOR_STOP = {"inc", "ltd", "corp", "corporation", "software", "systems", "international",
               "limited", "co", "llc", "plc", "the", "ltd"}
PROD_STOP = {"for", "the", "and", "of", "z", "os", "zos"}

# canonical column -> accepted header names (lower-case) in uploaded SYSMOD files
ALIASES = {
    "Vendor": ["vendor", "isv"],
    "Product": ["product", "product name", "product / release", "product/release"],
    "Product Version": ["product version", "version"],
    "z/OS Release": ["z/os release", "zos release", "z/os version", "zos version",
                     "target release", "os release", "z/os"],
    "SYSMOD ID": ["sysmod id", "sysmod", "ptf id", "ptf", "usermod", "fix id", "apar/ptf"],
    "Type": ["sysmod type", "type"],
    "FMID": ["fmid"],
    "APAR": ["apar id", "apar"],
    "Publish Date": ["ptf publish date", "publish date", "date", "ptf date"],
    "Prerequisites": ["prerequisites", "prerequisite", "prereq"],
    "Superseded By": ["superseding ptf", "superseded by", "supersedes"],
    "Description": ["full description", "description", "title", "abstract"],
    "Source URL": ["source url", "url", "link"],
    "Status": ["detail extraction status", "status"],
}

# ---------------------------------------------------------------- name handling
def _words(s):
    return re.findall(r"[a-z0-9]+", str(s).lower())

def vendor_tokens(s):
    return frozenset(w for w in _words(s) if w not in VENDOR_STOP)

def product_tokens(s):
    w = _words(s)
    return frozenset([x for x in w if x not in PROD_STOP] or w)

def _ver_num(s):
    m = re.search(r"\d+", str(s or ""))
    return m.group(0) if m else ""

# ---------------------------------------------------------------- release parsing
_REL = re.compile(r"(?<![\w.])(?:z/?os\s*)?(?:v\s*)?([23])\s*(?:\.|r)\s*([0-9])(?:\.0)?(?![\w])", re.I)

def parse_releases(text):
    """Return the in-scope z/OS releases mentioned in text ('2.5', 'V3R1', '2.5 to 3.2' ...)."""
    text = str(text or "")
    found = [f"{m.group(1)}.{m.group(2)}" for m in _REL.finditer(text)]
    found = [r for r in found if r in ORDER]
    if not found:
        return []
    if len(set(found)) >= 2 and re.search(r"\b(to|thru|through)\b|[-–—]", text, re.I):
        idx = [ORDER.index(r) for r in found]
        return ORDER[min(idx):max(idx) + 1]
    return [r for r in ORDER if r in found]

# ---------------------------------------------------------------- loading
def load_catalog(src):
    """Vendor / Product / Version list from the ISV compatibility workbook."""
    df = pd.read_excel(src, sheet_name=0)
    cols = {str(c).strip().lower(): c for c in df.columns}
    pick = lambda k: next((cols[c] for c in cols if c.startswith(k)), None)
    v, p, ver = pick("vendor"), pick("product"), pick("version")
    out = pd.DataFrame({"Vendor": df[v], "Product": df[p], "Version": df[ver] if ver else ""})
    out = out.dropna(subset=["Product"]).fillna("")
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip()
    out = out[out["Product"] != ""].drop_duplicates().reset_index(drop=True)
    return out

def _map_columns(df):
    lc = {str(c).strip().lower(): c for c in df.columns}
    m = {}
    for canon, names in ALIASES.items():
        for n in names:
            if n in lc:
                m[canon] = lc[n]
                break
    return m

def _read_any(name, src):
    if name.lower().endswith(".csv"):
        return {"csv": pd.read_csv(src, dtype=str)}
    return pd.read_excel(src, sheet_name=None, dtype=str)

def load_kb(files):
    """files: list of (name, path_or_filelike). Returns one normalised SYSMOD DataFrame."""
    frames = []
    for name, src in files:
        try:
            sheets = {"docs": src} if isinstance(src, pd.DataFrame) else _read_any(name, src)
        except Exception:
            continue
        for sh, df in sheets.items():
            m = _map_columns(df)
            if "SYSMOD ID" not in m or "Product" not in m:
                continue
            out = pd.DataFrame({c: (df[m[c]] if c in m else "") for c in ALIASES}).fillna("")
            for c in out.columns:
                out[c] = out[c].astype(str).str.strip().str.replace(r"\s+00:00:00$", "", regex=True)
            out["Source File"] = f"{os.path.basename(name)}:{sh}"
            frames.append(out)
    if not frames:
        return pd.DataFrame(columns=list(ALIASES) + ["Source File", "Release Basis", "rels", "ptok", "vtok"])
    kb = pd.concat(frames, ignore_index=True)
    kb = kb[(kb["SYSMOD ID"] != "") & (kb["Status"].str.upper() != "EXAMPLE")]
    kb = kb.drop_duplicates(subset=["Vendor", "Product", "Product Version", "SYSMOD ID", "z/OS Release"])
    rels, basis = [], []
    for rel, desc in zip(kb["z/OS Release"], kb["Description"]):
        r = parse_releases(rel)
        if r:
            rels.append(frozenset(r)); basis.append("stated")
            continue
        r = parse_releases(desc)
        if r:
            rels.append(frozenset(r)); basis.append("parsed from description")
        else:
            rels.append(frozenset()); basis.append("not stated")
    basis = [st[len("Found in document - "):] if st.startswith("Found in document - ") else b
             for st, b in zip(kb["Status"], basis)]
    kb["rels"], kb["Release Basis"] = rels, basis
    kb["ptok"] = [product_tokens(x) for x in kb["Product"]]
    kb["vtok"] = [vendor_tokens(x) for x in kb["Vendor"]]
    return kb.reset_index(drop=True)

def load_kb_folder(folder):
    files = []
    if os.path.isdir(folder):
        for f in sorted(os.listdir(folder)):
            if f.lower().endswith((".xlsx", ".xlsm", ".csv")) and not f.startswith("~"):
                files.append((f, os.path.join(folder, f)))
    return files

# ---------------------------------------------------------------- querying
def query(kb, vendor, product, version, wanted, mode="any", include_unspecified=True):
    """SYSMODs for one catalog product, filtered to the wanted z/OS releases."""
    cols = ["Vendor", "Product", "Product Version", "SYSMOD ID", "Type", "FMID", "APAR",
            "Publish Date", "Prerequisites", "Superseded By", "Description"] + list(wanted) + \
           ["Release Basis", "Source URL", "Status", "Source File"]
    if kb.empty:
        return pd.DataFrame(columns=cols)
    ct, vt, want = product_tokens(product), vendor_tokens(vendor), frozenset(wanted)
    keep = []
    for i, (pt, vtk, kver, rels) in enumerate(zip(kb["ptok"], kb["vtok"], kb["Product Version"], kb["rels"])):
        if not (ct <= pt):
            continue
        if vtk and vt and not (vtk & vt):
            continue
        if version and kver and _ver_num(version) != _ver_num(kver):
            continue
        if not rels:
            ok = include_unspecified
        elif mode == "all":
            ok = want <= rels
        else:
            ok = bool(want & rels)
        if ok:
            keep.append(i)
    res = kb.loc[keep].copy()
    for r in wanted:
        res[r] = [("✓" if r in s else ("?" if not s else "–")) for s in res["rels"]]
    return res[cols].reset_index(drop=True)

def coverage(catalog, kb, wanted):
    """For every catalog row: how many SYSMODs are loaded per release (shows what is still missing)."""
    rows = []
    for _, c in catalog.iterrows():
        r = query(kb, c["Vendor"], c["Product"], c["Version"], wanted, "any", True)
        row = {"Vendor": c["Vendor"], "Product": c["Product"], "Version": c["Version"],
               "SYSMODs loaded": len(r)}
        for rel in wanted:
            row[f"z/OS {rel}"] = int((r[rel] == "✓").sum()) if len(r) else 0
        row["Release not stated"] = int((r[wanted[0]] == "?").sum()) if len(r) else 0
        row["Data status"] = "OK" if len(r) else "No SYSMOD data loaded"
        rows.append(row)
    return pd.DataFrame(rows)

# ---------------------------------------------------------------- chat matching
def find_products(catalog, text):
    """Return (kind, DataFrame) where kind in {'product','vendor','suggest','none'}."""
    qt = set(_words(text)) - {"z", "os", "zos"}   # "z/OS" in the question is the platform, not a product
    prods = catalog[["Vendor", "Product"]].drop_duplicates()
    scored = []
    for v, p in zip(prods["Vendor"], prods["Product"]):
        pt = product_tokens(p)
        frac = len(pt & qt) / max(len(pt), 1)
        vb = 0.05 if (vendor_tokens(v) & qt) else 0
        scored.append((frac + vb, len(pt), v, p))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    full = [s for s in scored if s[0] >= 1.0]
    if full:
        top_len = max(s[1] for s in full)
        best = [s for s in full if s[1] == top_len]
        return "product", pd.DataFrame([(s[2], s[3]) for s in best], columns=["Vendor", "Product"])
    vend = [v for v in prods["Vendor"].unique() if vendor_tokens(v) and vendor_tokens(v) <= qt]
    if vend:
        return "vendor", pd.DataFrame({"Vendor": vend})
    sug = [s for s in scored if s[0] >= 0.5][:5]
    if not sug:
        names = list(prods["Product"])
        close = difflib.get_close_matches(str(text).strip(), names, n=5, cutoff=0.5)
        sug = [(0, 0, prods.loc[prods["Product"] == n, "Vendor"].iloc[0], n) for n in close]
    if sug:
        return "suggest", pd.DataFrame([(s[2], s[3]) for s in sug], columns=["Vendor", "Product"])
    return "none", pd.DataFrame()

def pick_version(catalog, vendor, product, text):
    """If the chat text names one of the product's versions (e.g. 'V13'), return it."""
    vers = catalog[(catalog["Vendor"] == vendor) & (catalog["Product"] == product)]["Version"]
    tl = set(_words(text))
    for v in vers:
        if v and v.lower() in tl:
            return v
    return ""
