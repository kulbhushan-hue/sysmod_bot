import io, os
import pandas as pd
import streamlit as st
import glob, re
import kb, docs, scraper

HERE = os.path.dirname(os.path.abspath(__file__))
CATALOG_DEFAULT = os.path.join(HERE, "ISV_Compatibility_assessment_-Data.xlsx")
KB_FOLDER = os.path.join(HERE, "kb_data")
DOC_FOLDER = os.path.join(HERE, "product_docs")

st.set_page_config(page_title="SYSMOD Compatibility Bot", page_icon="🖥️", layout="wide")
st.title("🖥️ SYSMOD Compatibility Bot — z/OS 2.5 to 3.2")
st.caption("Vendor/Product list from your ISV workbook. SYSMODs (PTF / APAR / USERMOD) come from the "
           "SYSMOD data files you load. Runs fully offline — no external APIs.")

# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("Inputs")
    cat_up = st.file_uploader("ISV workbook (Vendor / Product / Version)", type=["xlsx"])
    kb_up = st.file_uploader("SYSMOD data files (xlsx / csv)", type=["xlsx", "csv"], accept_multiple_files=True)
    st.caption("Files in the `kb_data/` folder are loaded automatically.")
    doc_up = st.file_uploader("Product documents (PDF, DOCX, HTML, TXT)", type=["pdf", "docx", "html", "htm", "txt", "md"],
                              accept_multiple_files=True)
    st.caption("Documents in the `product_docs/` folder are loaded automatically.")
    st.header("Release range")
    lo, hi = st.select_slider("z/OS releases", options=kb.ORDER, value=(kb.ORDER[0], kb.ORDER[-1]))
    wanted = kb.ORDER[kb.ORDER.index(lo): kb.ORDER.index(hi) + 1]
    mode = st.radio("A SYSMOD is listed if it supports…", ["any", "all"], index=0,
                    format_func=lambda x: "ANY release in range" if x == "any" else "ALL releases in range")
    incl = st.checkbox("Include SYSMODs with no release stated (marked ?)", value=True)

@st.cache_data(show_spinner=False)
def get_catalog(raw):
    return kb.load_catalog(io.BytesIO(raw) if raw else CATALOG_DEFAULT)

@st.cache_data(show_spinner="Reading product documents…")
def get_pages(uploads, folder_sig):
    files = [(n, open(p, "rb").read()) for n, p in docs.folder_files(DOC_FOLDER)] + list(uploads)
    return docs.load_pages(files)

@st.cache_data(show_spinner="Searching documents for SYSMODs…")
def get_doc_kb(pages, catalog):
    return docs.build_doc_kb(pages, catalog)

@st.cache_data(show_spinner=False)
def get_kb(uploads, folder_sig, doc_rows):
    files = kb.load_kb_folder(KB_FOLDER) + [(n, io.BytesIO(b)) for n, b in uploads]
    if len(doc_rows):
        files.append(("documents", doc_rows))
    return kb.load_kb(files)

catalog = get_catalog(cat_up.getvalue() if cat_up else None)
folder_sig = tuple((f, os.path.getmtime(p)) for f, p in kb.load_kb_folder(KB_FOLDER))
doc_sig = tuple((f, os.path.getmtime(p)) for f, p in docs.folder_files(DOC_FOLDER))
pages = get_pages(tuple((f.name, f.getvalue()) for f in (doc_up or [])), doc_sig)
doc_rows = get_doc_kb(pages, catalog)
kbdf = get_kb(tuple((f.name, f.getvalue()) for f in (kb_up or [])), folder_sig, doc_rows)

with st.sidebar:
    st.metric("Catalog products", catalog[["Vendor", "Product"]].drop_duplicates().shape[0])
    st.metric("Document pages read", len(pages))
    st.metric("SYSMODs found", len(kbdf))

def to_excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            from openpyxl.styles import Font, PatternFill
            for c in ws[1]:
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="002D72")
            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 60)
    return buf.getvalue()

def run(vendor, product, version, rels):
    res = kb.query(kbdf, vendor, product, version, rels, mode, incl)
    head = f"**{vendor} — {product}**" + (f" ({version})" if version else "") + \
           f" · z/OS {', '.join(rels)} · {len(res)} SYSMOD(s)"
    notes = []
    if kbdf.empty:
        notes.append("No product documents / SYSMOD data loaded yet — add them in the sidebar or the `product_docs/` folder.")
    elif res.empty:
        notes.append("No SYSMOD for this product matches the selected releases in the loaded data.")
    if res.empty or (res["Release Basis"] != "stated").any():
        notes.append("Rows marked `?` / 'parsed from description' need verification at the source.")
    ps = docs.statements(pages, vendor, product, rels)
    if len(ps):
        notes.append(f"{len(ps)} compatibility passage(s) found in your documents (see below).")
    return res, head + "\n\n" + "\n\n".join(notes)

tab_bot, tab_browse, tab_web, tab_docs, tab_cov = st.tabs(["💬 Bot", "🔎 Browse", "🌐 Online fetch", "📄 Documents", "📊 Coverage"])

# ------------------------------------------------------------------ chat
with tab_bot:
    st.write("Ask e.g. *“CHECK PLUS for DB2 on z/OS 2.5 to 3.2”*, *“BMC Delta IMS V6 3.1”*, or just a vendor name.")
    if "msgs" not in st.session_state:
        st.session_state.msgs = []
    for m in st.session_state.msgs:
        with st.chat_message(m["role"]):
            st.markdown(m["text"])
            if m.get("df") is not None and len(m["df"]):
                st.dataframe(m["df"], width="stretch", hide_index=True)
            if m.get("ps") is not None and len(m["ps"]):
                with st.expander(f"Document passages ({len(m['ps'])})"):
                    st.dataframe(m["ps"], width="stretch", hide_index=True)
    q = st.chat_input("Vendor / product and optional z/OS release(s)…")
    if q:
        st.session_state.msgs.append({"role": "user", "text": q})
        rels = kb.parse_releases(q) or wanted
        kind, hits = kb.find_products(catalog, q)
        df, ps = None, None
        if kind == "product" and len(hits) == 1:
            v, p = hits.iloc[0]
            df, text = run(v, p, kb.pick_version(catalog, v, p, q), rels)
            ps = docs.statements(pages, v, p, rels)
        elif kind == "product":
            text = "That name matches several products — which one?\n\n" + \
                   "\n".join(f"- {v} — {p}" for v, p in hits.itertuples(index=False))
        elif kind == "vendor":
            lines = []
            for v in hits["Vendor"]:
                n = catalog[catalog["Vendor"] == v]["Product"].nunique()
                lines.append(f"- **{v}**: {n} products in your workbook — name one to list its SYSMODs.")
            text = "\n".join(lines)
        elif kind == "suggest":
            text = "I couldn't find that exactly. Did you mean:\n\n" + \
                   "\n".join(f"- {v} — {p}" for v, p in hits.itertuples(index=False))
        else:
            text = "I couldn't match that to a product in the workbook. Try the Browse tab."
        st.session_state.msgs.append({"role": "assistant", "text": text, "df": df, "ps": ps})
        st.rerun()

# ------------------------------------------------------------------ browse
with tab_browse:
    c1, c2, c3 = st.columns(3)
    vendor = c1.selectbox("Vendor", sorted(catalog["Vendor"].unique()))
    prods = sorted(catalog[catalog["Vendor"] == vendor]["Product"].unique())
    product = c2.selectbox("Product", prods)
    vers = [""] + sorted(v for v in catalog[(catalog["Vendor"] == vendor) &
                                            (catalog["Product"] == product)]["Version"].unique() if v)
    version = c3.selectbox("Version", vers, format_func=lambda x: x or "(all)")
    res, text = run(vendor, product, version, wanted)
    st.markdown(text)
    if len(res):
        st.dataframe(res, width="stretch", hide_index=True)
        st.download_button("⬇️ Download Excel", to_excel({"SYSMODs": res}),
                           file_name=f"SYSMODs_{product[:30].replace('/', '_')}.xlsx")
    ps = docs.statements(pages, vendor, product, wanted)
    if len(ps):
        st.subheader("Document passages (product + z/OS release + compatibility wording)")
        st.dataframe(ps, width="stretch", hide_index=True)


# ------------------------------------------------------------------ online fetch
def auth_widgets(key):
    m = st.selectbox("Login / key (only if the site needs it)",
                     ["None", "Username + password (login form)", "HTTP basic", "Bearer token", "API key header", "Session cookie"], key=key + "m")
    c1, c2 = st.columns(2)
    if m == "None":
        return None
    if m.startswith("Username"):
        return {"type": "form", "login_url": c1.text_input("Login page URL", key=key + "a"),
                "user_field": c2.text_input("Username field name", "username", key=key + "b"),
                "pass_field": c2.text_input("Password field name", "password", key=key + "c"),
                "user": c1.text_input("Username", key=key + "d"), "password": c1.text_input("Password", type="password", key=key + "e")}
    if m == "HTTP basic":
        return {"type": "basic", "user": c1.text_input("Username", key=key + "f"), "password": c1.text_input("Password", type="password", key=key + "g")}
    if m == "Bearer token":
        return {"type": "bearer", "token": c1.text_input("Token", type="password", key=key + "h")}
    if m == "API key header":
        return {"type": "apikey", "header": c1.text_input("Header name", "X-API-Key", key=key + "i"),
                "value": c2.text_input("API key", type="password", key=key + "j")}
    st.caption("MFA sites can't be automated: sign in with your browser, then paste the Cookie header here.")
    return {"type": "cookie", "cookie": c1.text_input("Cookie header (name=value; …)", type="password", key=key + "k")}

def save_web(sm, cp, tag):
    os.makedirs(KB_FOLDER, exist_ok=True)
    if len(sm):
        sm.to_csv(os.path.join(KB_FOLDER, f"web_{tag}.csv"), index=False)
    if len(cp):
        cp.to_csv(os.path.join(KB_FOLDER, f"_compat_{tag}.csv"), index=False)
    st.cache_data.clear(); st.rerun()

with tab_web:
    src = st.radio("Source", ["Broadcom compatibility pages (public, no login)", "Any other vendor page"], horizontal=True)
    maxp = st.number_input("Max pages per start URL (other pages)", 10, 200, 10)
    render = st.checkbox("Page needs JavaScript (e.g. IBM Software Product Compatibility Reports) — uses Playwright if installed")
    auth = auth_widgets("w")
    fetcher = lambda: scraper.Fetcher(auth, render=render)
    if src.startswith("Broadcom"):
        codes = st.text_area("Broadcom product page codes (…/COMPAT/<code>.HTML), one per line or comma separated",
                             "\n".join(scraper.BROADCOM_SEED), height=110)
        idx = st.text_input("…or an index page URL to read the codes from (optional)")
        if st.button("🔍 Fetch from Broadcom", type="primary"):
            cl = [c for c in re.split(r"[\s,]+", codes) if c]
            f = fetcher()
            if idx:
                stt, h, _ = f.get(idx)
                cl = sorted(set(cl) | set(scraper.discover_codes(h))) if stt == "OK" else cl
            bar = st.progress(0.0)
            cp, sm, log = scraper.scrape_broadcom(cl, f, progress=bar.progress)
            st.session_state["web"] = ("broadcom", cp, sm, log)
    else:
        urls = st.text_area("Page URL(s), one per line", height=90)
        if st.button("🔍 Fetch", type="primary"):
            ul = [x.strip() for x in urls.splitlines() if x.strip().startswith("http")]
            if ul:
                with st.spinner("Crawling…"):
                    sm, log = scraper.scrape_generic(ul, fetcher(), max_pages=int(maxp))
                st.session_state["web"] = ("generic", pd.DataFrame(columns=scraper.COMPAT_COLS), sm, log)
    if "web" in st.session_state:
        tag, cp, sm, log = st.session_state["web"]
        st.dataframe(log, width="stretch", hide_index=True)
        if len(cp):
            st.subheader(f"z/OS compatibility found ({len(cp)} rows)")
            st.dataframe(cp, width="stretch", hide_index=True)
        if len(sm):
            st.subheader(f"SYSMODs / PTFs found ({len(sm)})")
            st.dataframe(sm, width="stretch", hide_index=True)
        if not len(cp) and not len(sm):
            st.warning("Nothing found. 'Access denied' = login/key needed; JavaScript-only pages need the JavaScript option.")
        else:
            if st.button("💾 Save (use in Bot, Browse and Coverage)"):
                save_web(sm, cp, tag)
    allc = [pd.read_csv(p) for p in glob.glob(os.path.join(KB_FOLDER, "_compat_*.csv"))]
    if allc:
        comp = pd.concat(allc, ignore_index=True)
        st.markdown(f"**Saved compatibility data:** {len(comp)} rows")
        wb_path = CATALOG_DEFAULT
        data, n = scraper.fill_workbook(wb_path, comp, "3.2")
        st.download_button(f"⬇️ Download your workbook filled for z/OS 3.2 ({n} rows)", data,
                           file_name="ISV_Compatibility_assessment_filled.xlsx")

# ------------------------------------------------------------------ documents
with tab_docs:
    if not pages:
        st.info("No documents loaded. Upload product guides / release notes / compatibility documents in the sidebar, "
                "or drop them in the `product_docs/` folder.")
    else:
        summary = pd.DataFrame(pages).groupby("doc").agg(Pages=("page", "count")).reset_index().rename(columns={"doc": "Document"})
        found = doc_rows.groupby("Source URL").size() if len(doc_rows) else pd.Series(dtype=int)
        summary["SYSMOD rows found"] = [int(sum(v for k, v in found.items() if k.startswith(d + " p."))) for d in summary["Document"]]
        st.dataframe(summary, width="stretch", hide_index=True)
        if len(doc_rows):
            st.download_button("⬇️ Download all SYSMODs found in documents", to_excel({"SYSMODs": doc_rows}),
                               file_name="SYSMODs_from_documents.xlsx")
        sq = st.text_input("Search the documents (keywords, e.g. 'DB2 3.1 PTF')")
        if sq:
            r = docs.search(pages, sq)
            st.dataframe(r, width="stretch", hide_index=True) if len(r) else st.write("No matches.")

# ------------------------------------------------------------------ coverage
with tab_cov:
    st.write("Every product in the workbook and how many SYSMODs are loaded for it — "
             "products marked *No SYSMOD data loaded* still need their vendor data.")
    cov = kb.coverage(catalog, kbdf, wanted)
    only_missing = st.checkbox("Show only products with no data")
    show = cov[cov["Data status"] != "OK"] if only_missing else cov
    st.dataframe(show, width="stretch", hide_index=True)
    st.download_button("⬇️ Download coverage Excel", to_excel({"Coverage": cov}), file_name="SYSMOD_coverage.xlsx")
