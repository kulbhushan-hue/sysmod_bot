"""Search product documents (PDF / DOCX / HTML / TXT / MD) for SYSMODs, PTFs and compatibility statements."""
from __future__ import annotations
import io, os, re
import pandas as pd
import kb

DOC_EXT = (".pdf", ".docx", ".html", ".htm", ".txt", ".md")
ID_RE = re.compile(r"\b(?:[A-Z]{2}\d{5}|(?!ISO)[A-Z]{3}\d{4})\b")   # e.g. UJ12345, BQI1234
KEYWORDS = re.compile(r"compatib|support|requir|minimum|toleration|coexist|PTF|APAR|USERMOD|SYSMOD|"
                      r"maintenance|service level|prerequisite", re.I)

def _chunk(lines, n=60):
    return [(i // n + 1, "\n".join(lines[i:i + n])) for i in range(0, len(lines), n)] or []

def extract_pages(name, raw):
    """-> list of (page_no, text). PDFs keep real page numbers; other formats are cut in 60-line blocks."""
    ext = os.path.splitext(name.lower())[1]
    try:
        if ext == ".pdf":
            from pypdf import PdfReader
            return [(i + 1, p.extract_text() or "") for i, p in enumerate(PdfReader(io.BytesIO(raw)).pages)]
        if ext == ".docx":
            import docx
            d = docx.Document(io.BytesIO(raw))
            lines = [p.text for p in d.paragraphs]
            for t in d.tables:
                lines += [" | ".join(c.text for c in r.cells) for r in t.rows]
            return _chunk([l for l in lines if l.strip()])
        if ext in (".html", ".htm"):
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(raw, "html.parser")
            for t in soup(["script", "style"]):
                t.decompose()
            return _chunk([l.strip() for l in soup.get_text("\n").splitlines() if l.strip()])
        return _chunk([l for l in raw.decode("utf-8", "replace").splitlines() if l.strip()])
    except Exception:
        return []

def load_pages(files):
    """files: list of (name, bytes). -> list of {doc, page, text}."""
    out = []
    for name, raw in files:
        if name.lower().endswith(DOC_EXT):
            for no, text in extract_pages(name, raw):
                if text.strip():
                    out.append({"doc": os.path.basename(name), "page": no, "text": text})
    return out

def folder_files(folder):
    res = []
    if os.path.isdir(folder):
        for f in sorted(os.listdir(folder)):
            if f.lower().endswith(DOC_EXT) and not f.startswith("~"):
                res.append((f, os.path.join(folder, f)))
    return res

def _label(text, pat):
    m = re.search(pat, text, re.I)
    return m.group(1).strip(" .;,") if m else ""

def _most_specific(items):
    """Drop products whose name is a strict subset of another matched product (IMS vs Delta IMS)."""
    return [x for x in items if not any(x[2] < y[2] for y in items)]

def build_doc_kb(pages, catalog):
    """Turn documents into SYSMOD rows (kb schema) for every catalog product named in/near each SYSMOD id."""
    prods = catalog[["Vendor", "Product"]].drop_duplicates()
    plist = [(v, p, kb.product_tokens(p)) for v, p in zip(prods["Vendor"], prods["Product"])
             if not set(kb._words(p)) <= kb.PROD_STOP]          # skip names like "z/OS" that are only stop-words
    doc_rel, rows = {}, []
    for pg in pages:
        d = pg["doc"]
        if d not in doc_rel:
            first = next(x["text"] for x in pages if x["doc"] == d)
            doc_rel[d] = kb.parse_releases(d.replace("_", " ") + " " + first[:1500])
    for pg in pages:
        lines = [l.strip() for l in pg["text"].splitlines() if l.strip()]
        if not any(ID_RE.search(l) for l in lines):
            continue
        name_toks = set(kb._words(os.path.splitext(pg["doc"])[0]))
        ptoks = set(kb._words(pg["text"])) | name_toks
        page_prods = _most_specific([x for x in plist if x[2] <= ptoks])
        page_rel = kb.parse_releases(pg["text"])
        seen = set()
        for i, l in enumerate(lines):
            for sid in ID_RE.findall(l):
                if sid in seen:
                    continue
                seen.add(sid)
                win = lines[max(0, i - 6): i + 7]
                wtoks = set(kb._words(" ".join(win))) | name_toks
                near = _most_specific([x for x in page_prods if x[2] <= wtoks])
                if near:
                    chosen = near
                elif len(page_prods) <= 3:
                    chosen = page_prods
                else:
                    continue
                if not chosen:
                    continue
                wtext = "\n".join(win)
                rel = kb.parse_releases("\n".join(lines[max(0, i - 3): i + 4]))
                if rel:
                    how = "release near SYSMOD"
                elif page_rel:
                    rel, how = page_rel, "release mentioned on page (verify)"
                elif doc_rel[pg["doc"]]:
                    rel, how = doc_rel[pg["doc"]], "release from document title (verify)"
                else:
                    how = "release not found"
                before = l[:l.find(sid)].upper()
                typ = "APAR" if before.rstrip().endswith("APAR") else \
                      ("USERMOD" if "USERMOD" in l.upper() else ("PTF" if "PTF" in l.upper() else ""))
                rec = {"Type": typ, "z/OS Release": ", ".join(rel), "SYSMOD ID": sid,
                       "FMID": _label(wtext, r"FMID[:\s]+([A-Z0-9]{7})\b"),
                       "APAR": _label(wtext, r"APAR[:\s#]+([A-Z]{2}\d{5})") if typ != "APAR" else "",
                       "Prerequisites": _label(wtext, r"(?:prereq\w*|requisite\w*)[:\s]+([^\n|]{1,80})"),
                       "Superseded By": _label(wtext, r"(?:superseded by|superseding)[:\s]+([A-Z]{2}\d{5})"),
                       "Description": l[:300], "Source URL": f"{pg['doc']} p.{pg['page']}",
                       "Status": f"Found in document - {how}"}
                for v, p, _ in chosen:
                    rows.append({"Vendor": v, "Product": p, "Product Version": "", "Publish Date": "", **rec})
    return pd.DataFrame(rows, columns=list(kb.ALIASES)) if rows else pd.DataFrame(columns=list(kb.ALIASES))

def statements(pages, vendor, product, wanted, limit=60):
    """Passages where the product, a wanted z/OS release and a compatibility/PTF keyword occur together."""
    pt, out = kb.product_tokens(product), []
    for pg in pages:
        last = -9
        lines = [l.strip() for l in pg["text"].splitlines() if l.strip()]
        name_toks = set(kb._words(os.path.splitext(pg["doc"])[0]))
        for i in range(len(lines)):
            if i - last < 3:
                continue
            win = lines[max(0, i - 2): i + 3]
            text = " ".join(win)
            rels = [r for r in kb.parse_releases(text) if r in wanted]
            if not rels or not KEYWORDS.search(lines[i]):
                continue
            if pt <= (set(kb._words(text)) | name_toks):
                last = i
                out.append({"Document": pg["doc"], "Page": pg["page"], "z/OS": ", ".join(rels),
                            "Passage": text[:400]})
    df = pd.DataFrame(out, columns=["Document", "Page", "z/OS", "Passage"])
    return df.drop_duplicates(subset=["Document", "Page", "Passage"]).head(limit).reset_index(drop=True)

def search(pages, text, limit=30):
    """Plain keyword search across all documents (ranked by number of query terms in the page)."""
    terms = [w for w in kb._words(text) if len(w) > 1]
    res = []
    for pg in pages:
        low = pg["text"].lower()
        score = sum(low.count(t) for t in terms)
        if score and all(t in low for t in terms[:3]):
            line = next((l for l in pg["text"].splitlines() if any(t in l.lower() for t in terms)), "")
            res.append((score, pg["doc"], pg["page"], line.strip()[:300]))
    res.sort(reverse=True)
    return pd.DataFrame([r[1:] for r in res[:limit]], columns=["Document", "Page", "Matching line"])
