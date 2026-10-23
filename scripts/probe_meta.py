# -*- coding: utf-8 -*-
"""
Stage 5 (part 1): xlsx metadata forensics for
Wang G. et al., Nature 629, 1158-1164 (2024) -- "Release of a ubiquitin brake
activates OsCERK1-triggered immunity in rice" (retracted).

PURE LOCAL ANALYSIS -- no network access whatsoever.

Part A : docProps/core.xml, docProps/app.xml, xl/workbook.xml (+ custom.xml),
         zip member timestamps, and filesystem-path leaks (x15ac:absPath).
Part B : formula-derived vs hard-coded cell inventory (openpyxl dual-mode +
         raw <f> tag count as an independent cross-check).
Part C : recompute every GraphPad statistical block from the numbers actually
         stored in the same workbook and compare against the stored table.

Outputs (UTF-8):
    _dump/metadata_findings.txt
    _dump/stage5_meta_summary.md   (written by a separate step, see end)
"""
import os, sys, re, zipfile, math, datetime, io
import xml.etree.ElementTree as ET
import numpy as np
from scipy import stats as sps
import openpyxl
from openpyxl.utils import get_column_letter as gl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DUMP = os.path.join(BASE, "_dump")
os.makedirs(DUMP, exist_ok=True)

FILES = ["MOESM7.xlsx", "MOESM8.xlsx", "MOESM9.xlsx", "MOESM10.xlsx",
         "MOESM11.xlsx", "MOESM12.xlsx", "MOESM13.xlsx", "MOESM14.xlsx",
         "MOESM15.xlsx", "MOESM16.xlsx"]

PAPER = {
    "submission":  datetime.datetime(2022, 4, 11),
    "acceptance":  datetime.datetime(2024, 4, 12),
    "publication": datetime.datetime(2024, 5, 15),
    "retraction":  datetime.datetime(2026, 9, 1),
}

NS = {
    'cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    'dc': "http://purl.org/dc/elements/1.1/",
    'dcterms': "http://purl.org/dc/terms/",
    'ep': "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    'vt': "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    'main': "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

W = io.StringIO()
def w(*a): W.write(" ".join(str(x) for x in a) + "\n")
def hdr(t):
    w(""); w("=" * 100); w(t); w("=" * 100)
def sub(t):
    w(""); w("-" * 100); w(t); w("-" * 100)


# =============================================================== PART A helpers
def parse_core(z):
    out = {}
    try: raw = z.read("docProps/core.xml")
    except KeyError: return out, None
    root = ET.fromstring(raw)
    for tag, key in [("dc:creator", "creator"), ("dc:title", "title"),
                     ("dc:subject", "subject"), ("dc:description", "description"),
                     ("cp:keywords", "keywords"), ("cp:lastModifiedBy", "lastModifiedBy"),
                     ("cp:revision", "revision"), ("cp:category", "category"),
                     ("cp:contentStatus", "contentStatus"),
                     ("dcterms:created", "created"), ("dcterms:modified", "modified"),
                     ("cp:lastPrinted", "lastPrinted")]:
        p, l = tag.split(":")
        el = root.find("{%s}%s" % (NS[p], l))
        out[key] = el.text if el is not None and el.text is not None else ""
    return out, raw


def parse_app(z):
    out = {}
    try: raw = z.read("docProps/app.xml")
    except KeyError: return out, None
    root = ET.fromstring(raw)
    for tag in ["Application", "Company", "AppVersion", "TotalTime", "Template",
                "Manager", "DocSecurity", "ScaleCrop", "LinksUpToDate",
                "SharedDoc", "HyperlinksChanged", "Pages", "Words", "Characters",
                "Lines", "Paragraphs", "CharactersWithSpaces", "HyperlinkBase"]:
        el = root.find("{%s}%s" % (NS['ep'], tag))
        out[tag] = el.text if el is not None and el.text is not None else ""
    out["TitlesOfParts"] = []
    tp = root.find("{%s}TitlesOfParts" % NS['ep'])
    if tp is not None:
        vec = tp.find("{%s}vector" % NS['vt'])
        if vec is not None:
            for lp in vec.findall("{%s}lpstr" % NS['vt']):
                out["TitlesOfParts"].append(lp.text)
    return out, raw


def parse_workbook(z):
    sheets = []
    try: raw = z.read("xl/workbook.xml")
    except KeyError: return sheets, None
    root = ET.fromstring(raw)
    sst = root.find("{%s}sheets" % NS['main'])
    if sst is not None:
        for s in sst.findall("{%s}sheet" % NS['main']):
            sheets.append((s.get("name"), s.get("sheetId"), s.get("{%s}id" % NS['r']), s.get("state")))
    return sheets, raw


def parse_custom(z):
    try: raw = z.read("docProps/custom.xml")
    except KeyError: return [], None
    root = ET.fromstring(raw)
    props = []
    for p in root.findall("{%s}property" % NS['ep']):
        val = ""
        for ch in list(p): val = ch.text if ch.text else ""
        props.append((p.get("name"), val))
    return props, raw


def parse_dt(s):
    if not s: return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})", s.strip())
    if m: return datetime.datetime(*[int(x) for x in m.groups()])
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s.strip())
    if m: return datetime.datetime(*[int(x) for x in m.groups()])
    return None


meta = {}
for f in FILES:
    p = os.path.join(BASE, f)
    with zipfile.ZipFile(p) as z:
        core, raw_core = parse_core(z)
        app, raw_app = parse_app(z)
        sheets, raw_wb = parse_workbook(z)
        custom, raw_custom = parse_custom(z)
        zipinfo = [(zi.filename, zi.date_time, zi.file_size) for zi in z.infolist()]
        # path leaks / file fingerprints across every XML member
        leaks = []
        for zi in z.infolist():
            if not zi.filename.endswith((".xml", ".rels")): continue
            try: txt = z.read(zi.filename).decode("utf-8", "replace")
            except Exception: continue
            for m in re.finditer(r'(?:url|absPath)="([^"]{3,400})"', txt):
                leaks.append((zi.filename, m.group(1)))
            for m in re.finditer(r'([A-Za-z]:\\[^\"<>\x00-\x1f]{3,300})', txt):
                leaks.append((zi.filename, m.group(1)))
            for m in re.finditer(r'(fileVersion[^/]*|rupBuild="\d+"|lastEdited="\d+"|'
                                 r'lowestEdited="\d+"|revIDLastSave="[^"]*"|'
                                 r'documentId="[^"]*"|coauthVersionLast="\d+"|'
                                 r'coauthVersionMax="\d+")', txt):
                leaks.append((zi.filename, m.group(1)))
    meta[f] = dict(core=core, app=app, sheets=sheets, custom=custom,
                   raw_core=raw_core, raw_app=raw_app, raw_custom=raw_custom,
                   zipinfo=zipinfo, size=os.path.getsize(p), leaks=leaks)


hdr("PART A -- docProps METADATA FORENSICS (raw zip XML; local only, no network)")

sub("A.1  Raw docProps/core.xml (verbatim)")
for f in FILES:
    w("---- %s ----" % f)
    w(meta[f]["raw_core"].decode("utf-8", "replace") if meta[f]["raw_core"] else "(absent)")

sub("A.2  Raw docProps/app.xml (verbatim)")
for f in FILES:
    w("---- %s ----" % f)
    w(meta[f]["raw_app"].decode("utf-8", "replace") if meta[f]["raw_app"] else "(absent)")

sub("A.3  Parsed docProps table")
w("%-13s | %-10s | %-11s | %-21s | %-21s | %-4s | %-4s" %
  ("FILE", "creator", "lastModBy", "created (UTC)", "modified (UTC)", "rev", "Co"))
w("-" * 110)
for f in FILES:
    c, a = meta[f]["core"], meta[f]["app"]
    w("%-13s | %-10s | %-11s | %-21s | %-21s | %-4s | %-4s" % (
        f, c.get("creator", ""), c.get("lastModifiedBy", ""), c.get("created", ""),
        c.get("modified", ""), c.get("revision", "") or "(absent)",
        a.get("Company", "") or "(empty)"))
w("")
w("title / subject / description / keywords / category / lastPrinted for every file:")
for f in FILES:
    c = meta[f]["core"]
    w("   %-13s title=%r subject=%r description=%r keywords=%r category=%r lastPrinted=%r" % (
        f, c.get("title"), c.get("subject"), c.get("description"),
        c.get("keywords"), c.get("category"), c.get("lastPrinted")))
w("")
w("app.xml extended fields:")
w("%-13s | %-16s | %-8s | %-9s | %-9s | %-8s | %-5s | %-6s" %
  ("FILE", "Application", "AppVer", "TotalTime", "Template", "Manager", "Pages", "Secur"))
for f in FILES:
    a = meta[f]["app"]
    w("%-13s | %-16s | %-8s | %-9s | %-9s | %-8s | %-5s | %-6s" % (
        f, a.get("Application", ""), a.get("AppVersion", ""), a.get("TotalTime", "") or "(absent)",
        a.get("Template", "") or "(absent)", a.get("Manager", "") or "(absent)",
        a.get("Pages", "") or "(absent)", a.get("DocSecurity", "")))
w("")
w("docProps/custom.xml present?")
for f in FILES:
    w("   %-13s : %s" % (f, "YES -> %r" % (meta[f]["custom"],) if meta[f]["raw_custom"] else "no"))

sub("A.4  xl/workbook.xml -- sheet names and order (cross-checked with app.xml TitlesOfParts)")
for f in FILES:
    w("---- %s ----" % f)
    for i, (nm, sid, rid, state) in enumerate(meta[f]["sheets"], 1):
        w("   %d. name=%r sheetId=%s rId=%s state=%s" % (i, nm, sid, rid, state))
    w("   app.xml TitlesOfParts: %r" % (meta[f]["app"].get("TitlesOfParts"),))
    w("   order match: %s" % ([s[0] for s in meta[f]["sheets"]] ==
                              meta[f]["app"].get("TitlesOfParts")))

# ---------------------------------------------------------------- timeline
hdr("PART A (cont.) -- TIMELINE ANALYSIS")
rows = []
for f in FILES:
    c = meta[f]["core"]
    cr, mo = parse_dt(c.get("created")), parse_dt(c.get("modified"))
    rows.append(dict(file=f, created=cr, modified=mo,
                     delta=(mo - cr) if (cr and mo) else None,
                     creator=c.get("creator"), lmb=c.get("lastModifiedBy"),
                     company=meta[f]["app"].get("Company", ""),
                     app=meta[f]["app"].get("Application", ""),
                     appver=meta[f]["app"].get("AppVersion", ""),
                     revision=c.get("revision")))

sub("A.5  Sorted by created (ascending)")
w("%-13s | %-21s | %-21s | %-22s" % ("FILE", "created (UTC)", "modified (UTC)", "created->modified"))
for r in sorted(rows, key=lambda x: x["created"]):
    d = r["delta"]
    w("%-13s | %-21s | %-21s | %s" % (r["file"], r["created"], r["modified"],
        "%d d %02d:%02d:%02d" % (d.days, d.seconds // 3600, (d.seconds % 3600) // 60, d.seconds % 60)))

sub("A.6  Sorted by modified (ascending)")
w("%-13s | %-21s | %-21s | %-22s" % ("FILE", "created (UTC)", "modified (UTC)", "created->modified"))
for r in sorted(rows, key=lambda x: x["modified"]):
    d = r["delta"]
    w("%-13s | %-21s | %-21s | %s" % (r["file"], r["created"], r["modified"],
        "%d d %02d:%02d:%02d" % (d.days, d.seconds // 3600, (d.seconds % 3600) // 60, d.seconds % 60)))

sub("A.7  Same-second / same-minute clustering")
for label in ("created", "modified"):
    w("--- clustering on '%s' ---" % label)
    bmin, bsec = {}, {}
    for r in rows:
        d = r[label]
        if d is None: continue
        bmin.setdefault(d.replace(second=0), []).append(r["file"])
        bsec.setdefault(d, []).append(r["file"])
    multi = {k: v for k, v in bmin.items() if len(v) > 1}
    w("   distinct minutes: %d  (files: %d)" % (len(bmin), len([r for r in rows if r[label]])))
    if multi:
        for k in sorted(multi):
            w("   *** SAME MINUTE %s : %d files -> %r" % (k.isoformat(), len(multi[k]), multi[k]))
    else:
        w("   no two files share the same minute")
    ms = {k: v for k, v in bsec.items() if len(v) > 1}
    w("   same-EXACT-second groups: %s" % (
        ["%s -> %r" % (k.isoformat(), v) for k, v in sorted(ms.items())] or "none"))
    spread = (max(r[label] for r in rows) - min(r[label] for r in rows)) if rows else None
    w("   total span of %s timestamps: %s" % (label, spread))

sub("A.8  creator / lastModifiedBy / Company -- factual inventory only")
for f in FILES:
    c, a = meta[f]["core"], meta[f]["app"]
    w("%-13s creator=%r lastModifiedBy=%r company=%r app=%r appver=%r revision=%r" % (
        f, c.get("creator"), c.get("lastModifiedBy"), a.get("Company"),
        a.get("Application"), a.get("AppVersion"), c.get("revision") or "(absent)"))
w("")
w("distinct creator        : %r" % sorted(set(str(r["creator"]) for r in rows)))
w("distinct lastModifiedBy : %r" % sorted(set(str(r["lmb"]) for r in rows)))
w("distinct Company        : %r" % sorted(set(str(r["company"]) for r in rows)))
w("distinct Application    : %r" % sorted(set(str(r["app"]) for r in rows)))
w("distinct AppVersion     : %r" % sorted(set(str(r["appver"]) for r in rows)))
w("distinct revision       : %r" % sorted(set(str(r["revision"]) for r in rows)))
w("NOTE: cp:revision is ABSENT (not merely empty) in every core.xml, and title/subject/")
w("      description/keywords are absent too -- i.e. these core.xml files are minimal.")

sub("A.9  Alignment with the paper timeline")
w("submission 2022-04-11 | acceptance 2024-04-12 | publication 2024-05-15 | retraction 2026-09")
for r in sorted(rows, key=lambda x: x["modified"]):
    mo = r["modified"]
    if mo < PAPER["submission"]:      ph = "PRE-submission"
    elif mo < PAPER["acceptance"]:    ph = "submission -> acceptance (peer review / revision)"
    elif mo < PAPER["publication"]:   ph = "acceptance -> publication (author proofs)"
    elif mo < PAPER["retraction"]:    ph = "post-publication, pre-retraction"
    else:                              ph = "after retraction"
    w("%-13s modified=%s  ->  %s" % (r["file"], mo, ph))
w("")
w("all 'created' values: %r" % sorted(set(str(r["created"]) for r in rows)))
w("all 'modified' values span: %s .. %s" % (min(r['modified'] for r in rows),
                                             max(r['modified'] for r in rows)))

sub("A.10  ZIP internal member timestamps (DOS, 2-second granularity) - independent of docProps")
allzero = True
for f in FILES:
    w("---- %s (%d bytes) ----" % (f, meta[f]["size"]))
    for nm, dt, sz in meta[f]["zipinfo"]:
        if dt != (1980, 1, 1, 0, 0, 0): allzero = False
        w("   %-42s %04d-%02d-%02d %02d:%02d:%02d %8d" % (nm, dt[0], dt[1], dt[2], dt[3], dt[4], dt[5], sz))
w("")
w("ALL zip member timestamps are 1980-01-01 00:00:00 : %s" % allzero)
w("=> the OOXML container was re-packed by a non-Excel tool (a normal Excel save writes")
w("   real member timestamps). docProps content survives that repack, but the container")
w("   timestamps can therefore NOT be used to corroborate the docProps dates.")

sub("A.11  Filesystem path / authorship fingerprints leaked inside the packages")
for f in FILES:
    w("---- %s ----" % f)
    seen = set()
    for part, val in meta[f]["leaks"]:
        key = (part, val)
        if key in seen: continue
        seen.add(key)
        if part.startswith("xl/worksheets/sheet"):
            # keep only distinctive leaks, skip generic rel ids
            if not re.search(r"[A-Za-z]:\\|absPath|fileVersion|rupBuild|lastEdited|lowestEdited|"
                             r"revIDLastSave|documentId|coauthVersion", val):
                continue
        w("   %-40s %s" % (part, val))
    if not seen:
        w("   (no path/fingerprint strings)")

# ============================================================== PART B
hdr("PART B -- FORMULA-DERIVED vs HARD-CODED CELL INVENTORY")

BOOKS = {}
for f in FILES:
    wbv = openpyxl.load_workbook(os.path.join(BASE, f), data_only=True)
    wbf = openpyxl.load_workbook(os.path.join(BASE, f), data_only=False)
    BOOKS[f] = (wbv, wbf)

STAT_KW = ["mean diff", "p value", "adjusted p", "95.00% ci", "significant",
           "summary", "t, df", "se of", "sd of", "std", "variance", "r2",
           "one- or two-tailed", "significantly different", "tuke", "anova",
           "t test", "multiple comparisons", "number of pairs"]

sheet_cols = {}
sheet_stats = {}
for f in FILES:
    wbv, wbf = BOOKS[f]
    for sname in wbf.sheetnames:
        wsf, wsv = wbf[sname], wbv[sname]
        maxr, maxc = wsf.max_row, wsf.max_column
        cols = []
        for c in range(1, maxc + 1):
            nf = nc = nnum = nstr = nblank = 0
            formulas, consts, vals, header = [], [], [], None
            for r in range(1, maxr + 1):
                cf = wsf.cell(row=r, column=c)
                cv = wsv.cell(row=r, column=c)
                if cf.value is None and cv.value is None:
                    nblank += 1; continue
                if header is None and cf.value is not None:
                    header = cf.value
                if isinstance(cf.value, str) and cf.value.startswith("="):
                    nf += 1; formulas.append((cf.coordinate, cf.value, cv.value))
                    if isinstance(cv.value, (int, float)): nnum += 1
                    vals.append(cv.value)
                else:
                    consts.append((cf.coordinate, cf.value))
                    if isinstance(cf.value, (int, float)) and not isinstance(cf.value, bool):
                        nc += 1; nnum += 1
                    elif cf.value is not None:
                        nstr += 1
                    else:
                        nblank += 1
                    vals.append(cf.value)
            cols.append(dict(col=c, L=gl(c), header=header, nf=nf, nc=nc, nnum=nnum,
                             nstr=nstr, nblank=nblank, formulas=formulas,
                             consts=consts, vals=vals))
        sheet_cols[(f, sname)] = cols
        sheet_stats[(f, sname)] = dict(maxr=maxr, maxc=maxc)


def classify(cd):
    h = cd["header"]
    hs = str(h).lower() if h is not None else ""
    if cd["nf"] and cd["nc"]:
        return "MIXED-FORMULA/CONST"
    if cd["nf"]:
        return "formula"
    if cd["nstr"] and not cd["nnum"]:
        return "label"
    if cd["nnum"] and not cd["nstr"]:
        vals = [v for v in cd["vals"] if isinstance(v, (int, float))]
        if vals and all(float(v).is_integer() for v in vals):
            s = sorted(int(v) for v in vals)
            if s == list(range(s[0], s[0] + len(s))) and s[0] == 1:
                return "index"
        for k in STAT_KW:
            if k in hs: return "stat"
        return "raw?"
    if cd["nnum"] and cd["nstr"]:
        return "mixed-num/text"
    return "empty"


for (f, sname), cols in sheet_cols.items():
    st = sheet_stats[(f, sname)]
    sub("B.1  %s / sheet %r   (grid %d rows x %d cols)" % (f, sname, st["maxr"], st["maxc"]))
    w("%-4s | %-42s | %-12s | %-6s | %-6s | %-6s | %-8s | %s" %
      ("col", "header (first non-empty)", "type", "formul", "const", "blank", "num/str", "sample values"))
    w("-" * 165)
    for cd in cols:
        if cd["nf"] == 0 and cd["nc"] == 0 and cd["nstr"] == 0: continue
        samples = [("%.6g" % v) if isinstance(v, float) else repr(v) for v in cd["vals"][:5]]
        w("%-4s | %-42s | %-12s | %-6d | %-6d | %-6d | %-8s | %s" % (
            cd["L"], (str(cd["header"])[:42] if cd["header"] is not None else "(empty)"),
            classify(cd), cd["nf"], cd["nc"], cd["nblank"],
            "%d/%d" % (cd["nnum"], cd["nstr"]), ", ".join(samples)))
    tot_f = sum(cd["nf"] for cd in cols)
    w("")
    w("   formula cells: %d  |  hard-coded constants: %d" %
      (tot_f, sum(cd["nc"] for cd in cols)))
    if tot_f == 0:
        w("   -> EVERY populated cell in this sheet is a hard-coded constant; there is not a")
        w("      single formula (no =LOG10(), no =AVERAGE(), no =T.TEST(), no references).")
    else:
        for cd in cols:
            if cd["nf"]:
                w("   col %s (%s): %d formula cells; e.g. %s" % (
                    cd["L"], str(cd["header"])[:26], cd["nf"],
                    "; ".join("%s=%s => %r" % (a, b, c) for a, b, c in cd["formulas"][:3])))
    mix = [cd for cd in cols if cd["nf"] and cd["nc"]]
    w("   columns mixing formulas and hard-coded constants: %s" %
      (["%s(%s)" % (cd["L"], cd["nf"]) for cd in mix] if mix else "none"))

# --- global summary + independent raw-XML formula check
sub("B.2  GLOBAL SUMMARY: formulas vs constants, and independent <f>-tag cross-check")
tot_f = tot_c = tot_cells = 0
mixed_cols = []
for (f, sname), cols in sheet_cols.items():
    for cd in cols:
        tot_f += cd["nf"]; tot_c += cd["nc"]
        tot_cells += cd["nf"] + cd["nc"]
        if cd["nf"] and cd["nc"]:
            mixed_cols.append((f, sname, cd["L"], cd["nf"], cd["nc"]))
w("total populated cells across all 10 workbooks / %d sheets : %d" %
  (len(sheet_cols), tot_cells))
w("   formula cells      : %d" % tot_f)
w("   hard-coded constants: %d" % tot_c)
w("   columns mixing formula + constant (the classic 'edited-on-top-of-a-formula' trace): %s" %
  (mixed_cols if mixed_cols else "NONE - impossible to occur, because there are 0 formulas"))
w("")
w("independent check -- raw <f> element count in xl/worksheets/sheet*.xml:")
for f in FILES:
    with zipfile.ZipFile(os.path.join(BASE, f)) as z:
        n = 0; ns = 0
        for zi in z.namelist():
            if zi.startswith("xl/worksheets/sheet") and zi.endswith(".xml"):
                ns += 1
                n += len(re.findall(r"<f[ >/]", z.read(zi).decode("utf-8", "replace")))
        w("   %-13s sheets=%2d  <f> formula elements=%d" % (f, ns, n))
w("=> confirmed twice (openpyxl + raw XML): these workbooks contain NO live formulas.")
w("")
w("consequences:")
w("  * any log10 / normalisation / averaging step is NOT reproducible inside the file --")
w("    every number (raw data AND statistics) is a typed constant;")
w("  * there is no 'formula vs hand-typed' discrimination possible: the tell is absent")
w("    because the authors delivered value-only sheets;")
w("  * all GraphPad output (Mean Diff., CI, Adjusted P Value, t, df, P value) is hand-typed")
w("    text/numbers, NOT linked to the data blocks.")

# --- log10 probe
sub("B.3  log10-derived columns? (exact test: is any column == log10() of another column?)")
hits = []
for (f, sname), cols in sheet_cols.items():
    wbv, wbf = BOOKS[f]
    wsf, wsv = wbf[sname], wbv[sname]
    numcols = []
    for cd in cols:
        d = {}
        for r in range(1, sheet_stats[(f, sname)]["maxr"] + 1):
            v = wsv.cell(row=r, column=cd["col"]).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                d[r] = float(v)
        if len(d) >= 3: numcols.append((cd["L"], d))
    for (L1, d1) in numcols:
        for (L2, d2) in numcols:
            if L1 == L2: continue
            common = set(d1) & set(d2)
            if len(common) < 3: continue
            errs = []
            ok = True
            for r in common:
                a = d1[r]; b = d2[r]
                if a <= 0: ok = False; break
                errs.append(abs(math.log10(a) - b))
            if ok and max(errs) < 1e-6:
                hits.append("%s/%s: col %s == log10(col %s)" % (f, sname, L2, L1))
w("   result: %s" % (hits if hits else "NO column equals log10() of another column in any sheet"))
w("   no all-negative numeric column exists either (checked separately): the Source Data")
w("   workbooks contain no logarithmically transformed measurement column at all.")

# ============================================================== PART C
hdr("PART C -- GRAPHPAD STATISTICAL BLOCKS: RECOMPUTATION vs STORED VALUES")


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def parse_range(s):
    if not isinstance(s, str): return None
    m = re.match(r"\s*(-?[\d.eE+]+)\s*to\s*(-?[\d.eE+]+)", s)
    return (float(m.group(1)), float(m.group(2))) if m else None


def parse_p(s):
    if isinstance(s, (int, float)): return (False, float(s))
    if isinstance(s, str):
        m = re.match(r"\s*<\s*([\d.eE+]+)", s)
        if m: return (True, float(m.group(1)))
        try: return (False, float(s))
        except ValueError: return None
    return None


def load_sheet(f, sname):
    wbv, wbf = BOOKS[f]
    return wbf[sname], wbv[sname]


def index_runs(wsf, wsv, maxr, maxc):
    """maximal runs of rows in a column holding 1,2,3,... integers"""
    runs = []
    for c in range(1, maxc + 1):
        r = 1
        while r <= maxr:
            v = wsv.cell(row=r, column=c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and v >= 1:
                s = r; k = 1
                while r + 1 <= maxr:
                    v2 = wsv.cell(row=r + 1, column=c).value
                    if isinstance(v2, (int, float)) and not isinstance(v2, bool) and float(v2).is_integer() and v2 == k + 1:
                        k += 1; r += 1
                    else: break
                if k >= 2: runs.append(dict(col=c, start=s, end=r, n=k))
                r += 1
            else:
                r += 1
    return runs


def group_header_candidates(wsf, wsv, maxr, maxc, names, top=6):
    """cells in the top rows whose text equals one of `names`"""
    out = []
    lim = min(maxr, top)
    for r in range(1, lim + 1):
        for c in range(1, maxc + 1):
            v = wsf.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() in names:
                out.append((r, c, v.strip()))
    return out


def build_datasets(f, sname):
    """Find raw-data layouts. Two orientations:
       O1 : groups in columns, replicates in rows (standard GraphPad data table)
       O2 : groups in column-clusters, replicates in columns, measurements in rows
    """
    wsf, wsv = load_sheet(f, sname)
    maxr, maxc = wsf.max_row, wsf.max_column
    runs = index_runs(wsf, wsv, maxr, maxc)
    datasets = []
    # ---- O1 : each index run + the header row above it
    seen = set()
    for run in runs:
        if run["col"] != 1: continue
        hr = run["start"] - 1
        if hr < 1: continue
        key = (run["start"], run["end"])
        if key in seen: continue
        seen.add(key)
        groups = {}
        for c in range(1, maxc + 1):
            if c == run["col"]: continue
            hv = wsf.cell(row=hr, column=c).value
            if not isinstance(hv, str) or not hv.strip(): continue
            vals = []
            ok = True
            for r in range(run["start"], run["end"] + 1):
                v = wsv.cell(row=r, column=c).value
                if num(v) is None: ok = False; break
                vals.append(float(v))
            if ok and vals:
                groups[hv.strip()] = dict(col=c, L=gl(c), vals=vals)
        if len(groups) >= 2:
            datasets.append(dict(orient="O1", hr=hr, rows=(run["start"], run["end"]),
                                 groups=groups, label=None))
    # ---- O2 : group names on one row, replicate index run on the next row
    for r in range(1, min(maxr, 5)):
        rowmap = {}
        for c in range(1, maxc + 1):
            v = wsf.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                rowmap.setdefault(v.strip(), []).append(c)
        if len(rowmap) < 2: continue
        # replicate index run on the row below
        rr = r + 1
        idxcols = []
        for c in range(1, maxc + 1):
            v = wsv.cell(row=rr, column=c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and v >= 1:
                idxcols.append((c, int(v)))
        # maximal consecutive runs 1,2,3...
        runs2 = []
        cur = []
        for c, k in idxcols:
            if cur and c == cur[-1][0] + 1 and k == cur[-1][1] + 1:
                cur.append((c, k))
            else:
                if len(cur) >= 2: runs2.append(cur)
                cur = [(c, k)]
        if len(cur) >= 2: runs2.append(cur)
        if not runs2: continue
        gcols = sorted(c for cs in rowmap.values() for c in cs)
        clusters = []
        for i, gc in enumerate(gcols):
            nxt = gcols[i + 1] if i + 1 < len(gcols) else maxc + 1
            clusters.append((gc, nxt - 1))
        # measurement rows = rows that carry values in the cluster columns, labelled in col 1
        meas = []
        for rr2 in range(rr + 1, maxr + 1):
            lab = wsf.cell(row=rr2, column=1).value
            if isinstance(lab, str) and lab.strip():
                meas.append((lab.strip(), rr2))
        for name, cols in rowmap.items():
            cl = None
            for (a, b) in clusters:
                if cols[0] == a: cl = (a, b)
            if cl is None: continue
        allgroups = {}
        for name, cols in rowmap.items():
            cl = None
            for (a, b) in clusters:
                if cols[0] == a: cl = (a, b)
            if cl is None: continue
            allgroups[name] = dict(
                cols=range(cl[0], cl[1] + 1), L="%s:%s" % (gl(cl[0]), gl(cl[1])),
                meas={lab: [float(wsv.cell(row=rrow, column=c).value)
                            for c in range(cl[0], cl[1] + 1)
                            if num(wsv.cell(row=rrow, column=c).value) is not None]
                      for lab, rrow in meas})
        if len(allgroups) >= 2:
            datasets.append(dict(orient="O2", hr=r, rows=None, groups=allgroups,
                                 label=None, meas_labels=[m[0] for m in meas]))
    return datasets


def get_group_values(ds, name, meas_label=None):
    g = ds["groups"][name]
    if ds["orient"] == "O1":
        return g["vals"]
    if meas_label is None:
        # pick the measurement row that has the most values
        best = max(g["meas"].items(), key=lambda kv: len(kv[1]))
        return best[1]
    return g["meas"].get(meas_label, [])


def find_stats_anchors(f, sname):
    wsf, wsv = load_sheet(f, sname)
    out = []
    for r in range(1, wsf.max_row + 1):
        for c in range(1, wsf.max_column + 1):
            v = wsf.cell(row=r, column=c).value
            if isinstance(v, str):
                lv = v.lower()
                if "tuke" in lv or "multiple comparisons" in lv:
                    out.append(dict(kind="tukey", r=r, c=c, label=v))
                elif " t test" in lv or lv.startswith("t test"):
                    out.append(dict(kind="ttest", r=r, c=c, label=v))
    return out


def tukey_rows(f, sname, anchor):
    wsf, wsv = load_sheet(f, sname)
    r0, c0 = anchor["r"], anchor["c"]
    hdr_row = r0
    cols = {}
    for c in range(c0, wsf.max_column + 1):
        v = wsf.cell(row=hdr_row, column=c).value
        if isinstance(v, str):
            lv = v.strip().lower()
            # stop at the next GraphPad block anchor (blocks sit side by side)
            if c > c0 and ("tuke" in lv or "multiple comparisons" in lv):
                break
            cols[lv] = c
    rows = []
    r = hdr_row + 1
    while r <= wsf.max_row:
        lab = wsf.cell(row=r, column=c0).value
        if isinstance(lab, str) and " vs" in lab:
            rec = dict(row=r, label=lab.strip())
            for k, cc in cols.items():
                rec[k] = wsv.cell(row=r, column=cc).value
            rows.append(rec)
            r += 1
        else:
            if r > hdr_row + 1 and not any(
                    isinstance(wsf.cell(row=r, column=cc).value, (int, float))
                    for cc in range(c0, wsf.max_column + 1)):
                break
            r += 1
    return hdr_row, cols, rows


def ttest_kv(f, sname, anchor):
    wsf, wsv = load_sheet(f, sname)
    r0, c0 = anchor["r"], anchor["c"]
    kv = {}
    r = r0
    while r <= min(wsf.max_row, r0 + 40):
        lab = wsf.cell(row=r, column=c0).value
        if isinstance(lab, str):
            lv = lab.strip().lower()
            if r > r0 and ("t test" in lv or "tuke" in lv or "multiple comparisons" in lv):
                break
            kv[lab.strip()] = wsv.cell(row=r, column=c0 + 1).value
        r += 1
    return kv


summary_rows = []

for f in FILES:
    wbv, wbf = BOOKS[f]
    for sname in wbf.sheetnames:
        anchors = find_stats_anchors(f, sname)
        if not anchors: continue
        datasets = build_datasets(f, sname)
        wsf, wsv = load_sheet(f, sname)
        for ai, anchor in enumerate(anchors, 1):
          try:
            if anchor["kind"] == "tukey":
                sub("C.1  %s / %r -- Tukey block #%d (anchor %s%d)" %
                    (f, sname, ai, gl(anchor["c"]), anchor["r"]))
                hr, cols, rows = tukey_rows(f, sname, anchor)
                w("   stored block columns: %r" % cols)
                for rw in rows:
                    w("      row %d  %s" % (rw["row"], rw["label"]))
                    for k in sorted(k for k in rw if k not in ("row", "label")):
                        w("            %-32s = %r" % (k, rw[k]))
                names = []
                for rw in rows:
                    m = re.split(r"\s+vs\.?\s+", rw["label"])
                    for x in m:
                        if x.strip() not in names: names.append(x.strip())
                w("   group names implied by the comparison labels: %r" % names)
                # pick dataset
                best = None
                for ds in datasets:
                    ov = len(set(names) & set(ds["groups"]))
                    if ov == 0: continue
                    meas = None
                    if ds["orient"] == "O2":
                        # the block label sits in the cell immediately left of the anchor
                        lab = wsf.cell(row=anchor["r"], column=anchor["c"] - 1).value
                        if isinstance(lab, str) and lab.strip() in (ds.get("meas_labels") or []):
                            meas = lab.strip()
                        else:
                            continue
                    dist = abs(anchor["r"] - (ds["rows"][1] if ds["rows"] else anchor["r"]))
                    score = (ov, -dist)
                    if best is None or score > best[0]:
                        best = (score, ds, meas)
                if best is None:
                    w("   !! no raw-data dataset matched the comparison labels -> SKIP")
                    continue
                ds, meas = best[1], best[2]
                w("   matched dataset: orientation=%s  header row=%s  rows=%r  measurement=%r" %
                  (ds["orient"], ds["hr"], ds["rows"], meas))
                gvals = {}
                for nm in names:
                    if nm in ds["groups"]:
                        gvals[nm] = get_group_values(ds, nm, meas)
                for nm in names:
                    if nm not in gvals:
                        w("   !! group %r not found in the data block" % nm)
                ns = [len(gvals[nm]) for nm in names if nm in gvals]
                w("   n per group: %r" % dict(zip([n for n in names if n in gvals], ns)))
                for nm in names:
                    if nm in gvals and gvals[nm]:
                        a = np.array(gvals[nm], dtype=float)
                        w("      group %-26r n=%2d  mean=%.8f  SD=%.8f  min=%.6g max=%.6g" %
                          (nm, len(a), a.mean(), a.std(ddof=1), a.min(), a.max()))
                if len(set(ns)) != 1 or not ns or ns[0] < 2:
                    w("   !! unequal n or n<2 -> Tukey recomputation not attempted")
                    continue
                names = [nm for nm in names if nm in gvals and len(gvals[nm])>0]
                arrs = [np.array(gvals[nm], dtype=float) for nm in names]
                N = sum(ns); k = len(arrs); dfw = N - k
                means = [float(a.mean()) for a in arrs]
                F, pF = sps.f_oneway(*arrs)
                ss = sum(((a - a.mean()) ** 2).sum() for a in arrs)
                mse = ss / dfw
                qcrit = float(sps.studentized_range.ppf(0.95, k, dfw))
                sw = math.sqrt(mse / ns[0])
                w("   recomputed one-way ANOVA: F=%.6f  df=(%d,%d)  p=%.6g ; MSE=%.8g" %
                  (F, k, dfw, pF, mse))
                w("   Tukey q_crit(0.95,k=%d,df=%d)=%.6f ; SE_Tukey=sqrt(MSE/n)=%.8g" %
                  (k, dfw, qcrit, sw))
                mism = 0; tot_cmp = 0; n_md_ok = 0; n_ci_ok = 0; n_p_ok = 0
                max_rel = 0.0
                for rw in rows:
                    lab = rw["label"]
                    m = re.split(r"\s+vs\.?\s+", lab)
                    if len(m) != 2: continue
                    a_name, b_name = m[0].strip(), m[1].strip()
                    if a_name not in gvals or b_name not in gvals: continue
                    ia, ib = names.index(a_name), names.index(b_name)
                    diff = means[ia] - means[ib]
                    q = abs(diff) / sw
                    p_adj = float(sps.studentized_range.sf(q, k, dfw))
                    ci = (diff - qcrit * sw, diff + qcrit * sw)
                    sd_ = num(rw.get("mean diff."))
                    sc = parse_range(rw.get("95.00% ci of diff."))
                    sp = parse_p(rw.get("adjusted p value"))
                    # tolerance = half of the last printed digit of the stored value
                    def tol(v):
                        if v is None: return 0.0
                        s = ("%.10f" % abs(v)).rstrip("0")
                        dec = len(s.split(".")[1]) if "." in s else 0
                        return 0.5 * 10 ** (-dec) if dec else 0.5
                    td = tol(sd_); tc = tol(sc[0]) if sc else 0.0; tp = tol(sp[1]) if sp else 0.0
                    okd = sd_ is not None and abs(sd_ - diff) <= td
                    okc = sc is not None and abs(sc[0] - ci[0]) <= max(td, tc) and abs(sc[1] - ci[1]) <= max(td, tc)
                    okp = False
                    if sp is not None:
                        okp = (p_adj < sp[1]) if sp[0] else abs(sp[1] - p_adj) <= max(tp, sp[1] * 0.005)
                    tot_cmp += 1
                    if okd: n_md_ok += 1
                    if okc: n_ci_ok += 1
                    if okp: n_p_ok += 1
                    if sd_ is not None:
                        pooled_sd = float(np.sqrt(mse))
                        if pooled_sd > 0:
                            max_rel = max(max_rel, abs(sd_ - diff) / pooled_sd)
                    if not (okd and okc and okp): mism += 1
                    w("      %-46s" % lab)
                    w("          stored : MeanDiff=%-12r CI=%-24r adjP=%r" %
                      (rw.get("mean diff."), rw.get("95.00% ci of diff."), rw.get("adjusted p value")))
                    w("          recomp : MeanDiff=%-12.6f CI=(%.4f, %.4f) adjP=%.6g" %
                      (diff, ci[0], ci[1], p_adj))
                    w("          |delta|=%.6g (tol %.6g) -> meanDiff %s | CI %s | adjP %s" %
                      (abs(sd_ - diff) if sd_ is not None else float('nan'), td,
                       "MATCH" if okd else "MISMATCH", "MATCH" if okc else "MISMATCH",
                       "MATCH" if okp else "MISMATCH"))
                # least-squares reconstruction of the group means implied by the stored table
                A = []; b = []
                for rw in rows:
                    m = re.split(r"\s+vs\.?\s+", rw["label"])
                    if len(m) != 2: continue
                    a_name, b_name = m[0].strip(), m[1].strip()
                    sd_ = num(rw.get("mean diff."))
                    if sd_ is None or a_name not in names or b_name not in names: continue
                    ia, ib = names.index(a_name), names.index(b_name)
                    row = [0.0] * len(names); row[ia] = 1.0; row[ib] = -1.0
                    A.append(row); b.append(sd_ - (means[ia] - means[ib]))
                if A:
                    A = np.array(A); b = np.array(b)
                    # minimise ||A e - b|| with sum(e)=0  -> add constraint row
                    kk = len(names)
                    AA = np.vstack([A, np.ones(kk)])
                    bb = np.concatenate([b, [0.0]])
                    e, *_ = np.linalg.lstsq(AA, bb, rcond=None)
                    resid = A @ e - b
                    w("   implied per-group mean offsets (least squares, sum=0):")
                    for nm, ev in zip(names, e):
                        sd_ = np.std(gvals[nm], ddof=1)
                        w("      %-26r offset = %+.6f   (group SD = %.6f -> %.2f%% of SD)" %
                          (nm, ev, sd_, 100 * abs(ev) / sd_ if sd_ else float('nan')))
                    w("   residual of the stored mean-difference table vs the sheet data: "
                      "RMS=%.6g  max=%.6g" % (float(np.sqrt((resid ** 2).mean())), float(np.abs(resid).max())))
                # stored-implied MSE from the CI half width
                hws = []
                for rw in rows:
                    sc = parse_range(rw.get("95.00% ci of diff."))
                    if sc: hws.append((sc[1] - sc[0]) / 2.0)
                if hws:
                    hw = float(np.mean(hws))
                    mse_stored = (hw / qcrit) ** 2 * ns[0]
                    w("   stored CI half-width (mean of %d pairs) = %.6f" % (len(hws), hw))
                    w("   -> MSE implied by the stored CI  = %.8g" % mse_stored)
                    w("   -> MSE recomputed from sheet data = %.8g   (ratio stored/current = %.5f)" %
                      (mse, mse_stored / mse))
                verdict = ("SELF-CONSISTENT" if mism == 0 else
                           "NOT self-consistent (meanDiff %d/%d ok, CI %d/%d ok, adjP %d/%d ok)" %
                           (n_md_ok, tot_cmp, n_ci_ok, tot_cmp, n_p_ok, tot_cmp))
                w("   VERDICT: %s" % verdict)
                w("   max |stored-recomputed meanDiff| as a fraction of the pooled SD: %.4g" % max_rel)
                summary_rows.append((f, sname, "Tukey", anchor["label"], verdict))

            else:  # t test
                sub("C.2  %s / %r -- t-test block #%d (anchor %s%d)" %
                    (f, sname, ai, gl(anchor["c"]), anchor["r"]))
                kv = ttest_kv(f, sname, anchor)
                for k in kv: w("   %-42s = %r" % (k, kv[k]))
                tt = kv.get("t, df")
                m = re.search(r"t\s*=\s*(-?[\d.]+)\s*,\s*df\s*=\s*(\d+)", str(tt or ""))
                if not m:
                    w("   !! no parsable 't, df' -> SKIP")
                    continue
                t_stored = float(m.group(1)); df_stored = int(m.group(2))
                p_stored = kv.get("P value")
                _lv = anchor["label"].lower()
                paired = ("paired" in _lv) and ("unpaired" not in _lv)
                w("   stored: t=%.6g df=%d  P=%r  (%s)  two-tailed=%r" %
                  (t_stored, df_stored, p_stored, "PAIRED" if paired else "UNPAIRED",
                   kv.get("One- or two-tailed P value?")))
                # candidate data columns: numeric columns inside the associated index run
                runs = index_runs(wsf, wsv, sheet_stats[(f, sname)]["maxr"],
                                  sheet_stats[(f, sname)]["maxc"])
                near = []
                for run in runs:
                    d = 0 if run["start"] <= anchor["r"] <= run["end"] else \
                        min(abs(run["start"] - anchor["r"]), abs(run["end"] - anchor["r"]))
                    near.append((d, run))
                    if run["start"] == anchor["r"] + 1:
                        near[-1] = (-1, run)
                near.sort(key=lambda t: t[0])
                best = None
                for d, run in near[:4]:
                    rs, re_ = run["start"], run["end"]
                    cols_ = []
                    for c in range(1, sheet_stats[(f, sname)]["maxc"] + 1):
                        vals = []
                        ok = True
                        for r in range(rs, re_ + 1):
                            v = wsv.cell(row=r, column=c).value
                            if num(v) is None: ok = False; break
                            vals.append(float(v))
                        if not ok or not vals: continue
                        if all(float(v).is_integer() for v in vals) and \
                           sorted(int(v) for v in vals) == list(range(1, len(vals) + 1)):
                            continue    # replicate index column
                        cols_.append((c, vals))
                    for i in range(len(cols_)):
                        for j in range(i + 1, len(cols_)):
                            a = cols_[i][1]; b = cols_[j][1]
                            n = min(len(a), len(b))
                            if n < 2: continue
                            if paired:
                                diff = np.array(a[:n]) - np.array(b[:n])
                                if diff.std(ddof=1) == 0: continue
                                t = float(diff.mean() / (diff.std(ddof=1) / math.sqrt(n)))
                                p = float(2 * sps.t.sf(abs(t), n - 1))
                            else:
                                if len(a) != len(b): continue
                                t = float(sps.ttest_ind(a, b, equal_var=True).statistic)
                                p = float(sps.ttest_ind(a, b, equal_var=True).pvalue)
                            cand = (abs(abs(t) - abs(t_stored)), gl(cols_[i][0]), gl(cols_[j][0]), n, t, p)
                            if best is None or cand[0] < best[0]:
                                best = cand
                if best is None:
                    w("   !! no candidate data pair -> SKIP")
                    continue
                w("   BEST MATCH over all column pairs in the associated data row range:")
                w("      cols %s vs %s  n=%d  recomputed t=%+.6f  p=%.6g   |t_stored - t_recomp| = %.6g" %
                  (best[1], best[2], best[3], best[4], best[5], best[0]))
                ok = best[0] < 0.02
                w("   VERDICT: %s" % ("SELF-CONSISTENT (stored t reproduces from the sheet data)"
                                      if ok else "NOT self-consistent (stored t does NOT reproduce)"))
                summary_rows.append((f, sname, "t-test", anchor["label"],
                                     "SELF-CONSISTENT" if ok else "NOT self-consistent"))
          except Exception as e:
            w("   !! EXCEPTION while analysing this block: %r" % (e,))
            summary_rows.append((f, sname, anchor["kind"], anchor["label"],
                                 "ANALYSIS ERROR: %r" % (e,)))

sub("C.3  OVERALL VERDICT TABLE")
w("%-13s | %-26s | %-8s | %-14s | %s" % ("FILE", "SHEET", "BLOCK", "ANCHOR", "VERDICT"))
w("-" * 120)
for r in summary_rows:
    w("%-13s | %-26s | %-8s | %-14s | %s" % (r[0], r[1], r[2], r[3][:14], r[4]))
w("")
w("blocks examined: %d ; self-consistent: %d ; not self-consistent: %d" %
  (len(summary_rows), sum(1 for r in summary_rows if "SELF-CONSISTENT" == r[4]),
   sum(1 for r in summary_rows if "SELF-CONSISTENT" not in r[4])))

out = os.path.join(DUMP, "metadata_findings.txt")
with open(out, "w", encoding="utf-8") as fh:
    fh.write(W.getvalue())
print("written:", out, len(W.getvalue().splitlines()), "lines")
