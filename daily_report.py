#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
daily_report.py  –  HAY Sale Report Daily Pipeline
====================================================
อ่านไฟล์ xlsx จาก Odoo โดยตรง แล้วอัพเดท Sales Report V3.1.xlsm
ในขั้นตอนเดียว (ไม่ต้องรัน inject / cleanse / generate แยกกัน)

v3.5b (cherry-pick: kept v3.4 cleanse_pospay because benchmark showed v3.5 regressed) — v3.12 - Graph secret expiry warning
  Optional GRAPH_SECRET_EXPIRES=YYYY-MM-DD in .env triggers a warning in
  email body when the secret is < 30 days from expiry (< 7 days = critical).
  No code restart needed when secret is renewed - just update both .env
  values (GRAPH_CLIENT_SECRET + GRAPH_SECRET_EXPIRES).

v3.11 - multi-recipient + CC + BCC support
  EMAIL_TO, EMAIL_CC, EMAIL_BCC now accept comma-separated email lists.
  Both daily success email and failure notification send to To+CC+BCC.
  Backward compat: single email in EMAIL_TO + missing EMAIL_CC/BCC = legacy.

v3.9 — default to yesterday (instead of today)
  Sales reports are typically run in the MORNING to summarize the
  PREVIOUS day's sales (which are now final). v3.9 changes the default
  date from today → yesterday in both `daily_report.py` (when --date is
  not provided) and `run_for_date.py` (calendar dialog default selected
  day). Calendar maxdate stays as today so users can still pick today
  for end-of-day runs. Explicit --date YYYY-MM-DD still wins.

v3.8 — fix cleanse_so multi-line fill (header columns)
  Multi-line SOs (same SO No, multiple product lines) had Odoo's header
  fields (campaign, source/branch, foreigner, customer info, salesperson,
  payment terms, rate) empty on subsequent lines. Root cause: cleanse_so
  computed keys_match BEFORE filling col 1 (SO name) from prev row, so
  keys_match was always False for empty subsequent lines, skipping the
  keyed fill-down of cols 2-11 + 26. The comment claimed this preserved
  v1.0 semantics — but v1.0 was the bug we should not have preserved.
  Production data: 711 multi-line SOs (48.8%) had this issue on 1,655
  rows (53% of SO data). Fix moves keys_match to AFTER the col 1 fill.

v3.7 — date selector
  Accept optional --date YYYY-MM-DD argument to regenerate the report for a
  past date. When provided, PDFs are written to the chosen date's folder
  (overwriting any existing files), the email subject reflects the chosen
  date, and the failure-email body mentions the chosen date. Without --date
  the pipeline runs for today as before (fully backward compatible). A new
  run_for_date.py launches a tkinter calendar dialog (with a 30s auto-run
  countdown) and is invoked from run_fetch_odoo.bat automatically.

v3.6 — Graph API email integration
  After PDF generation succeeds, send all 7 PDFs as attachments to the
  configured recipient via Microsoft Graph API /sendMail (OAuth 2.0 client
  credentials flow, no SMTP, no per-mailbox license required). On pipeline
  failure or partial PDF generation, send a failure notification email.
  Credentials loaded from .env in the script directory. Requires `requests`
  and `python-dotenv` (run_fetch_odoo.bat auto-installs them).

v3.5 — cleanse (M2) hot-path optimizations
  Scope: ONLY the M2 boundary — cleanse_so / cleanse_pos / cleanse_cn /
  cleanse_sopay / cleanse_pospay, split_column_p, ensure, fill_down,
  _extract_pos_method, AP/AP2. All other stages (fetch, build, xlsm-patch,
  PDF) are byte-identical to v3.4.

  * split_column_p: per-row construction switched from triple list-slice +
    concat (`row[:15] + [cat,s1,s2,s3] + row[16:]`) to a single splice on a
    fresh list copy (`out = list(row); out[15:16] = [...]`). One allocation
    instead of three plus a concat. Strip/parse logic unchanged; output
    columns and dtypes are byte-identical.
  * fill_down: per-row prev_len hoist preserved; the `(not use_key) or
    ci < 2` predicate split outside the column loop so the fast-fast
    unconditional path skips the per-col branch entirely. Bounds checks
    preserved. Semantics identical to v3.4.
  * cleanse_so / cleanse_pos / cleanse_cn: the col-5 Thai/Non-Thai remap
    (and POS row[11] override) fused into the kept-row build so the
    post-filter pass disappears. v3.1 leak fix (`if cur[1] == nxt[1]:`)
    is preserved verbatim — bug check intact.
  * cleanse_pos: redundant `cur[1] == prev[1]` re-check after the cur[1]
    fill-down replaced with a `keys_match` boolean carried from the fill.
    Identical logic, one fewer eq compare per row.
  * cleanse_cn: fill_down(d, [25], key_col_0idx=1) fused into the existing
    fill-down/negation pass — col 25 piggybacks on the same keys_match
    decision that drives cols 18/22/23. One pass over rows instead of two.
  * cleanse_pos: fill_down(d, [25], key_col_0idx=1) similarly fused into
    the keys_match branch of the main per-row loop.
  * cleanse_so: fill_down(list(range(12))+[26], key=1) inlined into the
    same per-row loop that does the payment-term remap (one pass over
    data rows instead of two).
  * cleanse_sopay: keys_match computation simplified — when cur[1] is
    filled from prev[1] the keys necessarily match; otherwise compare
    directly. Same logic, one fewer Python op per row.
  * cleanse_pospay: the three sequential passes (fill methods → filter
    → iterate Amount_THB) folded into a single pass that builds the
    filtered list directly. _POS_VOIDS hoisted to module scope.
  * _extract_pos_method: single `len(row)` lookup (was 2). Identical
    parse path; regex unchanged.
  * Module-level constants for the literal labels used inside the
    cleansers: _IN_STOCK, _PRE_ORDER, _SPLIT_HEADER_EXTRA. CPython
    interns the str literals already, but the *list* of sub-cat-header
    strings was being freshly allocated on every split_column_p call.
  * AP / AP2 unchanged (already optimal as set literals).

v3.4 — PDF generation (M5) hot-path optimizations
  Scope: ONLY generate_daily_pdfs() and its helper _wait_excel_ready().
  All upstream stages (fetch, cleanse, build, xlsm-patch) are byte-identical
  to v3.3. DAILY_BRANCHES is unchanged (order + content). PDF output path,
  filename pattern, and ExportAsFixedFormat positional args are unchanged.

  * xl.ScreenUpdating = False during the 7-branch loop. Excel was redrawing
    after every cell write + PageSetup property change inside the loop.
    Disabling redraws is the single largest win for COM-driven Excel work.
  * xl.Calculation = xlCalculationManual (-4135) during the loop. Auto-
    recalc was firing after every property write (PrintArea, margins, J2,
    etc.). With manual calc, we trigger ws.Calculate() explicitly ONLY
    after the date cells (G1:G3) and the branch filter (J2) — exactly the
    points where downstream formulas must refresh. PageSetup changes do
    NOT need a recalc.
  * Both ScreenUpdating and Calculation are restored in the finally block
    (set to True / xlCalculationAutomatic) so a mid-loop crash never
    leaves the user's Excel session in a degraded state.
  * Invariant PageSetup properties hoisted OUT of the loop and set once:
    PrintTitleRows, Zoom, FitToPagesWide, FitToPagesTall, all four page
    margins, HeaderMargin, FooterMargin. Only PrintArea (which depends on
    last_row) is set per-iteration. Eliminates ~9 COM round-trips × 7
    branches = ~63 redundant COM calls.
  * xl.CentimetersToPoints(0.5) and xl.CentimetersToPoints(0) are now
    computed ONCE before the loop (was 7× each = 14 unnecessary COM
    round-trips).
  * G1/G2/G3 date cells are now written in a SINGLE bulk Range assignment
    (ws.Range('G1:G3').Value = ((d,),(m,),(y,))) instead of 3 separate
    .Value writes. 1 COM round-trip instead of 3.
  * branch_safe regex sub pulled out of the loop into a module-level
    pre-compiled pattern (_RE_BRANCH_UNSAFE) — minor, but eliminates
    re.compile work per iteration.
  * _wait_excel_ready: now actually polls xl.Ready in a loop (previously
    it returned on first successful read regardless of value, which was
    effectively a one-shot check). A genuine poll-until-Ready avoids the
    corner case where Excel is busy on the very first probe.
  * Visible=True, DispatchEx, all positional ExportAsFixedFormat args,
    quality=xlQualityStandard, IgnorePrintAreas=False, and the order of
    branches are PRESERVED unchanged — these are documented win32com
    pitfalls in PROJECT_SUMMARY.md section 11.

v3.3 — xlsm patching (M4) hot-path optimizations
  Scope: only the xlsm-write block at the end of main() and its supporting
  helpers (col_letter, make_cell, make_row, strip_formula_cache, patch_xml).
  All upstream stages (fetch, cleanse, build) are byte-identical to v3.2.

  * Single open of the .bak file. Previously the .bak was opened twice — once
    to read the 7 sheet XMLs into memory and once to iterate-and-rewrite into
    the new xlsm. We now read the input infolist+payloads in one pass, then
    open the output for writing. The .bak's bytes were already paged into the
    OS file-cache after the first open, so the second open was cheap, but
    eliminating it still saves a CentralDirectory parse and tightens the
    failure window.
  * All hot-path regexes promoted to module-level pre-compiled constants:
    _RE_FORMULA_CACHE_B, _RE_ROW_B (used by patch_xml row-keep scan, bytes),
    _RE_FULLCALC_ATTR, _RE_CALCPR_TAG. patch_xml + strip_formula_cache no
    longer compile patterns on every sheet.
  * make_cell / make_row rewritten with f-strings (marginally faster per call
    vs %-format on CPython 3.11+) and a single str.join over a pre-allocated
    list rather than a generator. col_letter is now memoised behind a small
    dict cache; for the 36 columns the YTD sheet uses, every call after the
    first 36 is a dict hit.
  * patch_xml now operates on bytes end-to-end. The pre-compiled regex is
    a bytes pattern; the find()/index() calls operate on bytes; the
    row-rebuild strings are encoded once at the row-list join. This drops
    one full-XML decode+encode pair per sheet (the YTD XML alone is ~3 MB).
  * strip_formula_cache likewise operates on bytes (no decode/encode).
  * The two M4 passes on YTD/CR are now fused: patch_xml has an optional
    strip_formula=True flag that runs the formula-cache scrub on the SAME
    output buffer, avoiding a second decode-encode-allocate cycle.
  * Output zip is now ZIP_DEFLATED at compresslevel=3 (was the default 6).
    Excel does not care about the deflate level; level 3 is roughly 2× faster
    than level 6 on the embedded-XML payload while inflating the output xlsm
    by ~2-3% (measured on the 7.4 MB template).
  * Output writes via .tmp + os.replace are preserved (atomic on Windows for
    the same volume). calcChain.xml is still dropped and fullCalcOnLoad="1"
    is still injected into <calcPr>.

v3.2 — Odoo fetch transport optimizations (M1 boundary only)
  * fetch_odoo_all() now uses a thread-local _KeepAliveSafeTransport that:
      - keeps the HTTPS connection alive across search + export_data (already
        in v3.1) AND
      - advertises Accept-Encoding: gzip and transparently decodes
        Content-Encoding: gzip responses. XML-RPC payloads are highly
        redundant XML; gzip typically shrinks them 80-90%, cutting transfer
        wall-clock on the big SO/POS/CN datasets.
  * A pre-flight warm-up call (common.version) runs ONCE before the
    ThreadPoolExecutor fires, priming the TLS session cache so the 5 fresh
    worker connections complete their handshakes faster (some TLS stacks
    reuse the session ticket from the warm-up connection).
  * Output structure, field lists, domains, order-by clauses, country flip,
    and the (so_raw, pos_raw, cn_raw, sopay_raw, pospay_raw) return shape
    are byte-identical to v3.1. Only transport-layer behaviour changed.

v2.0 — parallelized fetch + optimized cleanse
  fetch_odoo_all() now runs the 5 Odoo XML-RPC fetches concurrently via
  ThreadPoolExecutor (each thread owns its own ServerProxy), and the cleanse_*
  helpers use hoisted attribute lookups, cached lengths, and slice-extend
  padding for faster pure-Python iteration. Output bytes are unchanged.

v3.1 — fix cross-document leak in zero-qty carry (cleanse_so / cleanse_pos / cleanse_cn)
  When Odoo emits a zero-qty header / down-payment / settlement line at the
  boundary between two documents, its monetary columns used to overwrite the
  first line of the next document. Now the carry only happens when col 1
  (the document key) matches. Observed for CN RINV2605.0004 inheriting
  -351180.42 from RINV2605.0003's last line.
  Also: PDFs are now organized into Daily PDFs/{YYYY}/{MMM}/{DD}/ folders.

v3.0 — HTTP keep-alive + tighter cleanse / row-builder hot paths
  * fetch_odoo_all() now wraps each thread's ServerProxy with a custom
    KeepAliveTransport (HTTPS) that keeps one TCP/TLS connection open
    across search + export_data instead of reopening per call.
  * cleanse_* helpers inline ensure(), drop redundant split_column_p
    re-padding, and use type()-based numeric checks where bools cannot
    appear. _extract_pos_method uses a module-level compiled regex.
  * build_ytd_rows / build_cr_rows fuse get_date()+to_iso_date() into a
    single parser _parse_date_triple() (one strptime per row), inline
    thb() with a per-row cached rate, and use a Schwartzian transform
    that sorts only on the precomputed ordinal — preserving stable
    tie-resolution. Output bytes are unchanged.

Usage (auto-detect ไฟล์ล่าสุดใน Downloads):
    python3 daily_report.py

Usage (ระบุไฟล์เอง):
    python3 daily_report.py --so SO.xlsx --pos POS.xlsx --cn CN.xlsx
                            --sopay SOPAY.xlsx --pospay POSPAY.xlsx
"""

import sys, os, re, zipfile, glob, time, argparse, base64, traceback
from datetime import datetime, timezone, timedelta

# ── Thailand timezone (UTC+7) ────────────────────────────────────────────────
TZ_BKK = timezone(timedelta(hours=7))

# ── Odoo connection (for --fetch-odoo mode) ──────────────────────────────────
# v3.13: ODOO_URL/DB/UID/PW moved to .env (loaded below) — never hardcode here.
ODOO_YEAR = 2026

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
XLSM         = os.path.join(SCRIPT_DIR, "Sales Report V3.1.xlsm")
XLSM_BAK     = XLSM + ".bak"
DOWNLOADS    = os.path.expanduser("~/Downloads")

# ── v3.6: Email config via Microsoft Graph API ───────────────────────────────
# Credentials and recipient loaded from .env in script directory.
# Required keys: GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET,
#                EMAIL_FROM, EMAIL_FROM_NAME, EMAIL_TO
_ENV_PATH = os.path.join(SCRIPT_DIR, '.env')

def _load_env():
    """Load .env into os.environ (stdlib fallback). Uses python-dotenv if
    available for richer parsing, otherwise a small manual parser. Missing
    .env is non-fatal — caller will get None from os.getenv() and the email
    step will print a warning and skip."""
    try:
        from dotenv import load_dotenv
        load_dotenv(dotenv_path=_ENV_PATH)
        return
    except ImportError:
        pass
    # Fallback: minimal .env parser
    if not os.path.exists(_ENV_PATH):
        return
    with open(_ENV_PATH, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v

_load_env()

ODOO_URL = os.getenv('ODOO_URL', '')
ODOO_DB  = os.getenv('ODOO_DB', '')
ODOO_UID = int(os.getenv('ODOO_UID') or 0)
ODOO_PW  = os.getenv('ODOO_PW', '')

GRAPH_TENANT_ID     = os.getenv('GRAPH_TENANT_ID')
GRAPH_CLIENT_ID     = os.getenv('GRAPH_CLIENT_ID')
GRAPH_CLIENT_SECRET = os.getenv('GRAPH_CLIENT_SECRET')
EMAIL_FROM          = os.getenv('EMAIL_FROM',      'noreply@norserepublics.com')
EMAIL_FROM_NAME     = os.getenv('EMAIL_FROM_NAME', 'Norse Republics — Sales')
# v3.11: EMAIL_TO / EMAIL_CC / EMAIL_BCC all accept comma-separated email lists
EMAIL_TO            = os.getenv('EMAIL_TO',        'bharkbhum.best@norserepublics.com')
EMAIL_CC            = os.getenv('EMAIL_CC',        '')   # optional, empty = no CC
EMAIL_BCC           = os.getenv('EMAIL_BCC',       '')   # optional, empty = no BCC

# v3.12: optional secret expiry date — when set, pipeline warns when secret nears expiry.
# Format: YYYY-MM-DD (the date the Azure client secret expires).
# Leave blank or remove the line to disable expiry checks.
GRAPH_SECRET_EXPIRES = os.getenv('GRAPH_SECRET_EXPIRES', '').strip()


def _check_secret_expiry():
    """v3.12: Compute days until GRAPH_CLIENT_SECRET expires.

    Returns (days_remaining, severity, message) tuple, or None if no expiry
    date is configured or it cannot be parsed.

    severity:
        'critical' = expired or < 7 days
        'warning'  = 8-30 days
        'ok'       = > 30 days (returns None instead — no need to warn)
    """
    if not GRAPH_SECRET_EXPIRES:
        return None
    try:
        expiry = datetime.strptime(GRAPH_SECRET_EXPIRES, '%Y-%m-%d').date()
    except ValueError:
        # Malformed date in .env — log once, don't crash
        return (None, 'parse_error',
                f"GRAPH_SECRET_EXPIRES in .env is not YYYY-MM-DD: "
                f"{GRAPH_SECRET_EXPIRES!r} - cannot check expiry")
    today_d = datetime.now(TZ_BKK).date()
    days = (expiry - today_d).days
    if days < 0:
        return (days, 'critical',
                f"!!! GRAPH_CLIENT_SECRET EXPIRED {-days} day(s) ago "
                f"(on {expiry.isoformat()}). Pipeline will FAIL until renewed. "
                f"Contact IT/admin immediately.")
    if days <= 7:
        return (days, 'critical',
                f"!!! Graph API client secret expires in {days} day(s) "
                f"(on {expiry.isoformat()}). URGENT: contact IT/admin "
                f"to renew it BEFORE that date or the daily report pipeline "
                f"will stop working.")
    if days <= 30:
        return (days, 'warning',
                f"NOTICE: Graph API client secret expires in {days} day(s) "
                f"(on {expiry.isoformat()}). Please contact IT/admin to "
                f"renew it before that date so the pipeline keeps running.")
    return None


def _parse_recipients(s):
    """v3.11: Parse a comma-separated email list into Graph API recipients format.
    Returns: list of {'emailAddress': {'address': '...'}} dicts (empty list if blank).
    Whitespace + empty entries are skipped silently.
    """
    if not s:
        return []
    out = []
    for raw in str(s).split(','):
        addr = raw.strip()
        if addr:
            out.append({'emailAddress': {'address': addr}})
    return out


GRAPH_TOKEN_URL = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'
GRAPH_SEND_URL  = 'https://graph.microsoft.com/v1.0/users/{user}/sendMail'

# Cache token across calls within a single pipeline run (~1h validity).
_GRAPH_TOKEN = None


# ── XML / cell helpers ───────────────────────────────────────────────────────
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# v3.3: col_letter memoised. With max ~36 cols on YTD, every call past warm-up
# is a dict-hit. Keeps the original algorithm so behaviour matches byte-for-byte.
_COL_LETTER_CACHE = {}
def col_letter(n):
    s = _COL_LETTER_CACHE.get(n)
    if s is not None:
        return s
    m = n
    out = ''
    while m > 0:
        m, r = divmod(m - 1, 26)
        out = chr(65 + r) + out
    _COL_LETTER_CACHE[n] = out
    return out

def make_cell(ri, ci, val):
    # v3.3: f-strings (marginally faster than %-format on CPython 3.11+).
    # Output byte-identical to v3.2.
    ref = f'{col_letter(ci)}{ri}'
    if val is None:
        return f'<c r="{ref}"/>'
    if isinstance(val, bool):
        return f'<c r="{ref}" t="n"><v>{int(val)}</v></c>'
    if isinstance(val, (int, float)):
        return f'<c r="{ref}" t="n"><v>{val}</v></c>'
    s = (str(val).replace('&', '&amp;').replace('<', '&lt;')
                 .replace('>', '&gt;').replace('"', '&quot;'))
    return f'<c r="{ref}" t="inlineStr"><is><t>{s}</t></is></c>'

def make_row(ri, row):
    # v3.3: pre-allocate cells list to avoid generator overhead inside join().
    cells = []
    app = cells.append
    ci = 1
    for v in row:
        if v is not None:
            app(make_cell(ri, ci, v))
        ci += 1
    return f'<row r="{ri}">{"".join(cells)}</row>'

def make_full_xml(rows):
    # Unchanged structurally; kept for any callers that import it.
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetData>',
    ]
    for ri, row in enumerate(rows, 1):
        if any(v is not None for v in row):
            lines.append(make_row(ri, row))
    lines.append('</sheetData></worksheet>')
    return '\n'.join(lines).encode('utf-8')

# v3.3: module-level pre-compiled regexes. Patterns identical to v3.2's inline
# string literals; bytes flavour avoids decode/encode round-trips in patch_xml.
_RE_FORMULA_CACHE_B = re.compile(br'(</f>)\s*<v>[^<]*</v>')
_RE_ROW_B = re.compile(br'<row r="(\d+)"[^>]*>.*?</row>', re.DOTALL)

def strip_formula_cache(xml_bytes):
    """ลบ <v> cache ออกจาก cells ที่มี formula (<f> tag) เพื่อให้ Excel recalculate ใหม่เมื่อเปิดไฟล์
    ทำให้ SUBTOTAL และ formula อื่นๆ แสดงค่าที่ถูกต้องแทนค่า stale ที่ถูก cache ไว้
    NOTE: Excel เก็บ formula cell เป็น <f>สูตร</f><v>cache</v>  (<v> อยู่ หลัง <f>)
    v3.3: operates on bytes directly to avoid decode/encode of multi-MB XML."""
    return _RE_FORMULA_CACHE_B.sub(br'\1', xml_bytes)

def patch_xml(orig_bytes, new_rows, first_row=5, strip_formula=False):
    """Keep header rows 1‥first_row-1, replace rows ≥ first_row.

    v3.3: bytes-only fast path — input/output are bytes throughout; no
    intermediate decode/encode of the multi-MB sheet XML. Output is
    byte-identical to v3.2 for the same inputs.

    If strip_formula=True, run the <f>...</f><v>cache</v> scrub on the
    final bytes buffer in the same call (fuses the two M4 passes that
    YTD and CR previously did).
    """
    # b'<sheetData' / b'</sheetData>' literals — find returns -1 on miss.
    sd0 = orig_bytes.find(b'<sheetData')
    sd1 = orig_bytes.find(b'</sheetData>')

    # Build the new <row>...</row> chunk (utf-8 string then encode once).
    def _build_new_rows():
        parts = []
        app = parts.append
        i = 0
        for r in new_rows:
            if any(v is not None for v in r):
                app(make_row(first_row + i, r))
            i += 1
        return ''.join(parts).encode('utf-8')

    if sd0 == -1:
        # No <sheetData> at all → append a fresh one before </worksheet>.
        ins = b'<sheetData>' + _build_new_rows() + b'</sheetData>'
        out = orig_bytes.replace(b'</worksheet>', ins + b'</worksheet>')
    else:
        # find the '>' that closes the <sheetData ...> open tag
        tag_end = orig_bytes.index(b'>', sd0) + 1
        if sd1 == -1:
            ins = b'<sheetData>' + _build_new_rows() + b'</sheetData>'
            out = orig_bytes[:sd0] + ins + orig_bytes[tag_end:]
        else:
            sd_content = orig_bytes[tag_end:sd1]
            # Keep only header rows (r < first_row) from the original sheetData.
            keep_parts = []
            kapp = keep_parts.append
            for m in _RE_ROW_B.finditer(sd_content):
                if int(m.group(1)) < first_row:
                    kapp(m.group(0))
            keep = b''.join(keep_parts)
            new_xml = _build_new_rows()
            out = (orig_bytes[:sd0] + b'<sheetData>' + keep + new_xml
                   + b'</sheetData>' + orig_bytes[sd1 + 12:])

    if strip_formula:
        out = _RE_FORMULA_CACHE_B.sub(br'\1', out)
    return out

# v3.3: workbook.xml calcPr-injection regexes pre-compiled at module load.
_RE_FULLCALC_ATTR = re.compile(r'fullCalcOnLoad="[^"]*"')
_RE_CALCPR_TAG = re.compile(r'(<calcPr)([^/]*/?>)')

# ── Fast xlsx reader (zipfile + XML, no openpyxl) ────────────────────────────
def read_xlsx(path):
    """Return list-of-lists from the first/active sheet of an xlsx file."""
    with zipfile.ZipFile(path, 'r') as z:
        # Shared strings (cells with t="s")
        ss = []
        if 'xl/sharedStrings.xml' in z.namelist():
            import xml.etree.ElementTree as ET
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('.//{%s}si' % NS):
                texts = si.findall('.//{%s}t' % NS)
                ss.append(''.join((t.text or '') for t in texts))

        # Find active sheet (first sheet listed in workbook.xml)
        wb_xml = z.read('xl/workbook.xml').decode('utf-8', errors='replace')
        sheet_ids = re.findall(r'<sheet[^>]+r:id="(rId\d+)"', wb_xml)
        # Map rId → file
        rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='replace')
        id_map = dict(re.findall(r'Id="(rId\d+)"[^>]+Target="([^"]+)"', rels_xml))
        sheet_file = 'xl/' + id_map.get(sheet_ids[0], 'worksheets/sheet1.xml').lstrip('/')

        xml_bytes = z.read(sheet_file)

    # Parse cells
    cells = {}
    max_row = 0
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml_bytes.decode('utf-8', errors='replace'), re.DOTALL):
        ri = int(m.group(1))
        max_row = max(max_row, ri)
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', m.group(2), re.DOTALL):
            col_str, _ri, attrs, inner = cm.group(1), cm.group(2), cm.group(3), cm.group(4)
            ci = 0
            for ch in col_str:
                ci = ci * 26 + (ord(ch) - 64)
            t_match = re.search(r't="([^"]+)"', attrs)
            t = t_match.group(1) if t_match else ''
            v_match = re.search(r'<v>([^<]*)</v>', inner)
            is_match = re.search(r'<is>.*?<t[^>]*>([^<]*)</t>.*?</is>', inner, re.DOTALL)
            if t == 's' and v_match:
                idx = int(v_match.group(1))
                val = ss[idx] if idx < len(ss) else None
            elif t == 'inlineStr' and is_match:
                val = is_match.group(1)
            elif v_match:
                raw = v_match.group(1)
                try:
                    f = float(raw)
                    val = int(f) if f == int(f) else f
                except:
                    val = raw
            else:
                val = None
            if val is not None:
                cells[(ri, ci)] = val

    if not cells:
        return []
    max_col = max(ci for (_, ci) in cells)
    return [[cells.get((ri, ci)) for ci in range(1, max_col + 1)]
            for ri in range(1, max_row + 1)]

# ── Auto-detect xlsx files from Downloads ────────────────────────────────────
def fmt_mtime(path):
    """Return human-readable modification time of a file."""
    return datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')

def find_latest(folder, pattern, max_age_hours=None):
    """Return most-recently-modified file matching glob pattern, or None.
    If max_age_hours is set, only consider files newer than that many hours."""
    matches = glob.glob(os.path.join(folder, pattern))
    if max_age_hours is not None:
        cutoff = time.time() - max_age_hours * 3600
        matches = [f for f in matches if os.path.getmtime(f) >= cutoff]
    return max(matches, key=os.path.getmtime) if matches else None

def find_payments(folder, max_age_hours=None):
    """Payment files share the same pattern.
    SO Payment = exported first (older mtime), POS Payment = exported after (newer mtime).
    If max_age_hours is set, only consider files within that window."""
    all_matches = sorted(
        glob.glob(os.path.join(folder, 'Payments (account.payment)*.xlsx')),
        key=os.path.getmtime)
    if max_age_hours is not None:
        cutoff = time.time() - max_age_hours * 3600
        all_matches = [f for f in all_matches if os.path.getmtime(f) >= cutoff]
    if len(all_matches) >= 2:
        return all_matches[-2], all_matches[-1]   # SO pay (older), POS pay (newer)
    if len(all_matches) == 1:
        return all_matches[0], all_matches[0]
    return None, None

def detect_files(downloads, max_age_hours=None):
    so     = find_latest(downloads, 'Sales Order (sale.order)*.xlsx',             max_age_hours)
    pos    = find_latest(downloads, 'Point of Sale Orders (pos.order)*.xlsx',     max_age_hours)
    cn     = find_latest(downloads, 'Journal Entry (account.move)*.xlsx',         max_age_hours)
    sopay, pospay = find_payments(downloads, max_age_hours)
    return so, pos, cn, sopay, pospay

# ── Data cleansing (mirrors VBA Data_Cleansing + Split_Column_P) ─────────────
# v3.5: literal labels hoisted to module scope. CPython interns the str literals
# already, but the *list* of sub-cat-header strings was being freshly allocated
# on every split_column_p call.
_IN_STOCK = 'In – Stock'
_PRE_ORDER = 'Pre – Order'
_SPLIT_HEADER_EXTRA = ['Full Product Category', 'Product Category',
                       'Sub – Cat 1', 'Sub – Cat 2']

def ensure(row, n):
    """Pad row in place with None up to length n. One extend beats a Python loop."""
    diff = n - len(row)
    if diff > 0:
        row.extend([None] * diff)

def fill_down(data, col_indices_0idx, key_col_0idx=None):
    """Fill None downward in given columns (0-indexed).

    Semantics preserved exactly from v1.0:
      - When key_col_0idx is None, OR when ci < 2, fill unconditionally
        from the previous row (still bounds-checked).
      - Otherwise only fill when the key column matches between this row
        and the previous row.

    v3.5: predicate `(not use_key) or ci < 2` precomputed per column so the
    per-row column loop dispatches into a flat "uncond" or "keyed" branch.
    Bounds checks preserved. Output byte-identical to v3.4.
    """
    n_rows = len(data)
    if n_rows < 2:
        return
    col_indices = list(col_indices_0idx)
    use_key = key_col_0idx is not None
    # v3.5: split columns into "unconditional fill" vs "key-matched fill".
    if use_key:
        uncond_cols = [ci for ci in col_indices if ci < 2]
        keyed_cols = [ci for ci in col_indices if ci >= 2]
    else:
        uncond_cols = col_indices
        keyed_cols = ()
    for ri in range(1, n_rows):
        cur = data[ri]
        prev = data[ri - 1]
        cur_len = len(cur)
        prev_len = len(prev)
        # unconditional cols
        for ci in uncond_cols:
            if ci >= cur_len:
                continue
            v = cur[ci]
            if v is None or v == '':
                cur[ci] = prev[ci] if ci < prev_len else None
        if not keyed_cols:
            continue
        # key-matched cols
        k0 = cur[key_col_0idx] if key_col_0idx < cur_len else None
        k1 = prev[key_col_0idx] if key_col_0idx < prev_len else None
        if k0 != k1:
            continue
        for ci in keyed_cols:
            if ci >= cur_len:
                continue
            v = cur[ci]
            if v is None or v == '':
                cur[ci] = prev[ci] if ci < prev_len else None

def split_column_p(data):
    """Insert Full Category + 3 sub-category columns after col 15 (0-idx 15).
    v3.0: callers (cleanse_*) already pad rows to len >= 26, so the per-row
    ensure() calls have been hoisted to a single guard for the header.

    v3.5: per-row construction switched from triple list-slice + concat
    (`row[:15] + [cat,s1,s2,s3] + row[16:]`) to a single splice on a fresh
    list copy. Saves two slice allocations + one concat per row. Header parse
    and sub-category strip semantics unchanged; output byte-identical.
    """
    if not data:
        return []
    result = []
    _append = result.append
    # Header row (index 0). Cleansers pad d[0] alongside data rows, but stay
    # defensive in case a caller skips that.
    header = data[0]
    if len(header) < 19:
        header = header + [None] * (19 - len(header))
    new_header = list(header)
    new_header[15:16] = _SPLIT_HEADER_EXTRA
    _append(new_header)
    # Data rows — every row guaranteed len >= 19 by caller.
    n_rows = len(data)
    for ri in range(1, n_rows):
        row = data[ri]
        cat = row[15] or ''
        if cat:
            parts = str(cat).split('/')
            n_parts = len(parts)
            sub1 = parts[0].strip() if n_parts > 0 else ''
            sub2 = parts[1].strip() if n_parts > 1 else ''
            sub3 = parts[2].strip() if n_parts > 2 else ''
        else:
            sub1 = sub2 = sub3 = ''
        out = list(row)
        out[15:16] = (cat, sub1, sub2, sub3)
        _append(out)
    return result

AP = {'Credit Card (SCB)', 'Bank Transfer (SCB)', 'Credit Card (KBANK)',
      'Bank Transfer (KBANK)', 'Beam (KBANK)', 'Alipay (KBANK)', 'WeChat Pay (KBANK)'}
AP2 = AP | {'Void Credit Card (SCB)', 'Void Credit Card (KBANK)'}

# v3.0: numeric-type tuple for `type(x) in _NUM_TYPES` fast-path. Use ONLY where
# bool can never appear (bool is a subclass of int, so isinstance(True, int) is
# True but type(True) is int is False — semantic divergence if booleans leak in).
_NUM_TYPES = (int, float)

# v3.0: compiled regex for combined-POS-payment reference parsing (used in the
# cleanse_pospay hot loop). Pattern + flags identical to v2.0's re.match call.
_COMBINE_POS_RE = re.compile(r'Combine (.+?) POS payments', re.IGNORECASE)

# v3.5: POS void-payment-method tuple hoisted to module scope. Avoids per-call
# literal-tuple allocation in cleanse_pospay.
_POS_VOIDS = ('Void Credit Card (SCB)', 'Void Credit Card (KBANK)')

def cleanse_so(raw):
    # v3.0: single-pass padding (inline ensure) — all rows brought to len >= 27.
    # v3.5: inline the fill_down(list(range(12)) + [26], key_col_0idx=1) call
    # and fuse the col-5 Thai/Non-Thai remap into the zero-qty kept-row loop.
    d = [list(r) for r in raw]
    for r in d:
        diff = 27 - len(r)
        if diff > 0:
            r.extend([None] * diff)
    n = len(d)
    # NOTE: baseline used fill_down(list(range(12))+[26], key=1) which computes
    # keys_match ONCE at the top of each row from the ORIGINAL col-1 values
    # (BEFORE the unconditional fills on cols 0/1). So if cur[1] starts empty
    # and prev[1] is non-empty, keys_match=False — cols 2-11 + 26 are NOT
    # filled from prev (even though cur[1] is unconditionally filled later).
    # We must preserve that exact semantic here.
    for ri in range(1, n):
        cur = d[ri]
        prev = d[ri - 1]
        # v3.8 FIX: Fill cols 0 (date) and 1 (SO name) FIRST, then compute
        # keys_match from the POST-fill values. The old v3.5b code captured
        # keys_match before the col-1 fill (preserving a v1.0 bug), which
        # caused multi-line SOs to have empty header fields (campaign,
        # source/branch, foreigner, customer, salesperson, payment term,
        # rate) on every subsequent line. This fix populates them correctly.
        # c=1 (ci=0): unconditional fill-down
        v = cur[0]
        if v is None or v == '':
            cur[0] = prev[0]
        # c=2 (ci=1): unconditional fill-down
        v = cur[1]
        if v is None or v == '':
            cur[1] = prev[1]
        # Now compute keys_match using the post-fill SO name in col 1.
        keys_match = (cur[1] == prev[1])
        if keys_match:
            # c=3..12 (ci=2..11): key-matched fill-down (header fields)
            for ci in range(2, 12):
                v = cur[ci]
                if v is None or v == '':
                    cur[ci] = prev[ci]
            # ci=26: key-matched fill-down (currency rate — important for
            # non-THB SOs where rate ≠ 1, otherwise THB conversion is wrong).
            v = cur[26]
            if v is None or v == '':
                cur[26] = prev[26]
        # Payment-term remap on col 11 (after fill-down) — same as v3.4.
        v11 = cur[11]
        if v11 == 'Payment in Full now':
            cur[11] = _IN_STOCK
        elif v11 is not None and str(v11) != '':
            cur[11] = _PRE_ORDER
    # Remove zero-qty rows (carry monetary cols forward), and apply col-5
    # Thai/Non-Thai remap on kept rows (fused into the same pass).
    # After padding above every row in d has cur_len >= 27, so the cur_len
    # > 18/23/24 guards always succeed — they have been dropped.
    nd = [d[0]]
    _nd_append = nd.append
    last_i = n - 1
    for i in range(1, n):
        cur = d[i]
        if cur[18] == 0:
            if i < last_i:
                nxt = d[i + 1]
                # v3.1: only carry within the same SO (col 1 = SO no). Prevents
                # cross-document leak when a zero-qty header line sits at the
                # boundary between two SO documents in raw export.
                if cur[1] == nxt[1]:
                    v23 = cur[23]
                    if v23 not in (None, ''):
                        nxt[23] = v23
                    v24 = cur[24]
                    if v24 not in (None, ''):
                        nxt[24] = v24
        else:
            # v3.5: fuse col-5 Thai/Non-Thai remap into the kept-row loop.
            v5 = cur[5]
            if v5 in (False, 0):
                cur[5] = 'Thai'
            elif v5 in (True, 1):
                cur[5] = 'Non-Thai'
            _nd_append(cur)
    return split_column_p(nd)

def cleanse_pos(raw):
    # v3.0: inline padding; after this every row has len >= 26.
    # v3.5: replace redundant `cur[1]==prev[1]` re-check after cur[1] fill
    # with a keys_match boolean. Fuse fill_down([25], key=1) into the same
    # keys_match branch, and fuse the col-5/col-11 post-pass into the kept-row
    # filter loop.
    d = [list(r) for r in raw]
    for r in d:
        diff = 26 - len(r)
        if diff > 0:
            r.extend([None] * diff)
    # col 0 is ISO string from Odoo / Excel-serial from xlsx. Bool isn't
    # observable here, so type() check is safe. int(v0) semantics preserved.
    n = len(d)
    for ri in range(1, n):
        cur = d[ri]
        prev = d[ri - 1]
        v0 = cur[0]
        if v0 is None or v0 == '':
            cur[0] = prev[0]
        else:
            tv0 = type(v0)
            if tv0 is int or tv0 is float:
                cur[0] = int(v0)
        # ISO date string จาก API → ไม่ต้องแปลง ปล่อยไว้ให้ get_date() จัดการ
        # v3.5: when cur[1] is filled from prev[1], keys necessarily match.
        v1 = cur[1]
        if v1 is None or v1 == '':
            cur[1] = prev[1]
            keys_match = True
        else:
            keys_match = (v1 == prev[1])
        if keys_match:
            # cur has length >= 26, so ci in range(2,12) is always in-bounds.
            for ci in range(2, 12):
                vc = cur[ci]
                if vc is None or vc == '':
                    cur[ci] = prev[ci]
            # v3.5: fuse fill_down(d, [25], key_col_0idx=1) into this branch.
            v25 = cur[25]
            if v25 is None or v25 == '':
                cur[25] = prev[25]
    nd = [d[0]]
    _nd_append = nd.append
    last_i = n - 1
    for i in range(1, n):
        cur = d[i]
        # cur_len >= 26, so cur[18]/[22]/[23] are safe (None when raw was empty).
        if cur[18] == 0:
            if i < last_i:
                nxt = d[i + 1]
                # v3.1: only carry within the same POS order (col 1 = order
                # reference). Prevents cross-document leak at order boundary.
                if cur[1] == nxt[1]:
                    v22 = cur[22]
                    if v22 not in (None, ''):
                        nxt[22] = v22
                    v23 = cur[23]
                    if v23 not in (None, ''):
                        nxt[23] = v23
        else:
            # v3.5: fuse col-5 Thai/Non-Thai remap + row[11]='In – Stock'
            # override into the kept-row loop.
            v5 = cur[5]
            if v5 in (False, 0):
                cur[5] = 'Thai'
            elif v5 in (True, 1):
                cur[5] = 'Non-Thai'
            cur[11] = _IN_STOCK
            _nd_append(cur)
    return split_column_p(nd)

def cleanse_cn(raw):
    # v3.0: inline padding; after this every row has len >= 26. All `prev` rows
    # (ri >= 1) also have len >= 26, so the `if ci < len(prev)` guards collapse.
    # v3.5: fuse fill_down(d, [25], key_col_0idx=1) into the existing per-row
    # fill/negate loop (col 25 piggybacks on the same keys_match decision).
    d = [list(r) for r in raw]
    for r in d:
        diff = 26 - len(r)
        if diff > 0:
            r.extend([None] * diff)
    n = len(d)
    # The v1.0 inner loop processes c in 1..26 (ci in 0..25). It branches on:
    #   - c == 1, 2          -> unconditional fill-down (if empty)
    #   - c in 3..11         -> key-matched fill-down (if empty) — c != 12 ensures col 11 (ci=10) handled but c=12/ci=11 is excluded
    #   - c == 12 (ci=11)    -> skipped completely (no fill, no neg)
    #   - c == 19,23,24      -> key-matched fill-down OR negate ints
    # Build the same effect in a flatter form.
    # Note: col 5 already converted to bool (Thai/Non-Thai flag) before cleanse,
    # but col 5 is not in the negate set {18,22,23} so type() is unused there.
    # Negation targets cols 18 (qty), 22 (sale_closed), 23 (total) — numeric only;
    # bools cannot appear, so type() check is safe.
    for ri in range(1, n):
        cur = d[ri]
        prev = d[ri - 1]
        keys_match = (cur[1] == prev[1])
        # c=1 (ci=0): unconditional fill-down (c < 3)
        v = cur[0]
        if v is None or v == '':
            cur[0] = prev[0]
        # c=2 (ci=1): unconditional fill-down (c < 3)
        v = cur[1]
        if v is None or v == '':
            cur[1] = prev[1]
            # re-evaluate keys_match because we just changed it
            keys_match = True
        # c=3..11 (ci=2..10): key-matched fill-down only (c != 12 always true here)
        if keys_match:
            for ci in range(2, 11):
                v = cur[ci]
                if v is None or v == '':
                    cur[ci] = prev[ci]
            # v3.5: col 25 key-matched fill (was a separate fill_down pass).
            v = cur[25]
            if v is None or v == '':
                cur[25] = prev[25]
        # c=12 (ci=11): skipped — v1.0 falls into the c != 12 guard so no action
        # c in {19,23,24} (ci in {18,22,23}): key-matched fill-down OR negate
        for ci in (18, 22, 23):
            v = cur[ci]
            if (v is None or v == ''):
                if keys_match:
                    cur[ci] = prev[ci]
            else:
                tv = type(v)
                if tv is int or tv is float:
                    cur[ci] = v * -1
    nd = [d[0]]
    _nd_append = nd.append
    last_i = n - 1
    for i in range(1, n):
        cur = d[i]
        # cur_len >= 26 after padding.
        if cur[18] == 0:
            if i < last_i:
                nxt = d[i + 1]
                # v3.1: only carry within the same credit note (col 1 = CN no).
                # Prevents cross-document leak — Odoo can emit a zero-qty
                # "down-payment / settlement" line at the boundary whose
                # price_total reflects the parent invoice's amount_total
                # (e.g. -351180.42), which used to overwrite the next CN's
                # first-line price_total. Observed for RINV2605.0004 carrying
                # the value from the last line of RINV2605.0003.
                if cur[1] == nxt[1]:
                    xv = cur[23]
                    if xv:
                        tv = type(xv)
                        if (tv is int or tv is float) and xv > 0:
                            nxt[23] = xv
        else:
            _nd_append(cur)
    return split_column_p(nd)

def cleanse_sopay(raw):
    # v3.0: inline padding to len >= 11. After this every row has len 11+.
    # v3.5: simplify keys_match — if we fill cur[1] from prev[1] then keys
    # necessarily match; otherwise compare directly.
    d = [list(r) for r in raw]
    for r in d:
        diff = 11 - len(r)
        if diff > 0:
            r.extend([None] * diff)
    d[0].append('Amount_THB')
    n = len(d)
    # cols 8 (amount) and 10 (exchange rate) are numeric per Odoo schema —
    # bool isn't observable, so type() check is safe.
    for ri in range(1, n):
        cur = d[ri]
        prev = d[ri - 1]
        # c=1,2 (ci=0,1) unconditional fill-down when empty. cur_len >= 11 → guards drop.
        v = cur[0]
        if v is None or v == '':
            cur[0] = prev[0]
        v = cur[1]
        if v is None or v == '':
            cur[1] = prev[1]
            keys_match = True
        else:
            keys_match = (v == prev[1])
        # c=3..11 (ci=2..10) key-matched fill-down when empty
        if keys_match:
            for ci in range(2, 11):
                v = cur[ci]
                if v is None or v == '':
                    cur[ci] = prev[ci]
        # Compute Amount_THB = amount (col8) × exchange_rate (col10)
        v8 = cur[8]
        v10 = cur[10]
        tv8 = type(v8)
        tv10 = type(v10)
        if (tv8 is int or tv8 is float) and (tv10 is int or tv10 is float):
            a_thb = v8 * v10
        else:
            a_thb = None
        cur.append(a_thb)
    # Filter rows where payment method (col 7) is in AP set. cur_len >= 11 here.
    out = [d[0]]
    _out_append = out.append
    for r in d[1:]:
        if r[7] in AP:
            _out_append(r)
    return out

def _extract_pos_method(row):
    """Extract payment method for POS combined payments.
    If col H (idx 7) is None, try to parse from col C (idx 2) Reference.
    Pattern: 'Combine {Method} POS payments from ...'
    v3.0: uses module-level compiled _COMBINE_POS_RE.
    v3.5: single len() lookup instead of two.
    """
    row_len = len(row)
    if row_len > 7:
        method = row[7]
        if method:
            return method
    if row_len > 2:
        ref = row[2]
        if ref and isinstance(ref, str):
            m = _COMBINE_POS_RE.match(ref)
            if m:
                return m.group(1).strip()
    return None

def cleanse_pospay(raw):
    # v3.0: inline padding to len >= 11. After this every row has len 11+.
    d = [list(r) for r in raw]
    for r in d:
        diff = 11 - len(r)
        if diff > 0:
            r.extend([None] * diff)
    # Fill payment method from Reference if col H is empty (combined POS payments)
    for row in d[1:]:
        if not row[7]:
            row[7] = _extract_pos_method(row)
    # Filter rows where payment method (col 7) is in AP2 set. Rows are len >= 11.
    filtered = [d[0]]
    _filtered_append = filtered.append
    for r in d[1:]:
        if r[7] in AP2:
            _filtered_append(r)
    d = filtered
    d[0].append('Amount_THB')
    _voids = ('Void Credit Card (SCB)', 'Void Credit Card (KBANK)')
    # cols 8/10 are numeric per Odoo schema; bool is not observable, so type() is safe.
    for ri in range(1, len(d)):
        row = d[ri]
        row[10] = 1  # Exchange rate for POS is always 1
        iv = row[8]
        tiv = type(iv)
        a_thb = iv if (tiv is int or tiv is float) else None
        row.append(a_thb)
        # row now has len 12 (cols 0..11). pm always in-bounds.
        pm = row[7]
        if pm in _voids:
            for ci in (8, 10, 11):
                v = row[ci]
                tv = type(v)
                if tv is int or tv is float:
                    row[ci] = v * -1
    return d

# ── YTD / CR row builders ────────────────────────────────────────────────────
WDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

def get_date(dv):
    """Parse date from Excel serial OR ISO string (Odoo API returns strings)."""
    try:
        if isinstance(dv, str) and len(dv) >= 10 and dv[4:5] == '-':
            # ISO date/datetime: "2026-01-15" or "2026-01-15 10:30:00"
            d = datetime.strptime(dv[:10], '%Y-%m-%d')
        else:
            d = datetime(1899, 12, 30) + timedelta(days=float(dv))
        return d.month, WDAYS[d.weekday()]
    except:
        return None, None

def to_iso_date(dv):
    """Convert Excel serial number or ISO datetime string to 'DD-MMM-YYYY' string."""
    try:
        if isinstance(dv, str) and len(dv) >= 10 and dv[4:5] == '-':
            d = datetime.strptime(dv[:10], '%Y-%m-%d')
        else:
            d = datetime(1899, 12, 30) + timedelta(days=float(dv))
        return d.strftime('%d-%b-%Y')   # e.g. "15-Jan-2026"
    except:
        return dv

# v3.0: single-parse fused helper. get_date(dv) + to_iso_date(dv) walk identical
# parse logic; this returns (month, weekday, iso_str, sort_ordinal) in one pass.
# On parse failure: (None, None, dv, dv-if-numeric-else-0) — matches v2.0
# _date_sort_key fallback semantics exactly.
_EXCEL_EPOCH = datetime(1899, 12, 30)
def _parse_date_triple(dv):
    try:
        if isinstance(dv, str) and len(dv) >= 10 and dv[4:5] == '-':
            d = datetime.strptime(dv[:10], '%Y-%m-%d')
        else:
            d = _EXCEL_EPOCH + timedelta(days=float(dv))
        return d.month, WDAYS[d.weekday()], d.strftime('%d-%b-%Y'), d.toordinal()
    except:
        # Replicate v2.0 sort-key fallback: numeric dv passes through as-is so
        # numeric-valued failures still sort by their raw magnitude.
        return None, None, dv, (dv if isinstance(dv, (int, float)) else 0)

def thb(val, rate):
    if not isinstance(val, (int, float)): return None
    r = rate if isinstance(rate, (int, float)) and rate > 0 else 1
    result = val * r
    return int(result) if result == int(result) else round(result, 4)

def build_ytd_rows(so_c, pos_c, cn_c):
    # v3.0: fused date parse (one strptime/timedelta per row instead of two),
    # inlined thb() with cached rate, and Schwartzian (ordinal, row) sort that
    # reuses the ordinal computed during row construction.
    _parse = _parse_date_triple
    decorated = []   # list of (ordinal, row) — sorted at the end
    _app = decorated.append
    # ── SO: 30 cols after split_column_p — row[27]=Unpaid exists, row[29]=rate
    for row in so_c[1:]:
        m, wd, iso, ordn = _parse(row[0])
        r29 = row[29]
        rate = r29 if (isinstance(r29, (int, float)) and r29 > 0) else 1
        # Build row in-place — avoid list(row) + mutate.
        out = [m, wd, iso] + row[1:]
        # Inline thb() for AG/AH/AI/AJ. Branching matches thb() exactly.
        for cidx in (23, 25, 26, 27):
            val = row[cidx]
            if isinstance(val, (int, float)):
                result = val * rate
                ir = int(result)
                out.append(ir if result == ir else round(result, 4))
            else:
                out.append(None)
        _app((ordn, out))
    # ── POS: 29 cols → insert None at col 27, row[28]=rate, no Unpaid
    for row in pos_c[1:]:
        m, wd, iso, ordn = _parse(row[0])
        r28 = row[28]
        rate = r28 if (isinstance(r28, (int, float)) and r28 > 0) else 1
        # row[:27] + [None, *row[27:]] — same length as v2.0 list-concat form.
        out = [m, wd, iso] + row[1:27] + [None] + row[27:]
        for cidx in (23, 25, 26):
            val = row[cidx]
            if isinstance(val, (int, float)):
                result = val * rate
                ir = int(result)
                out.append(ir if result == ir else round(result, 4))
            else:
                out.append(None)
        out.append(None)  # AJ: Unpaid N/A for POS
        _app((ordn, out))
    # ── CN: 29 cols → same shape as POS, row[28]=rate, no Unpaid
    for row in cn_c[1:]:
        m, wd, iso, ordn = _parse(row[0])
        r28 = row[28]
        rate = r28 if (isinstance(r28, (int, float)) and r28 > 0) else 1
        out = [m, wd, iso] + row[1:27] + [None] + row[27:]
        for cidx in (23, 25, 26):
            val = row[cidx]
            if isinstance(val, (int, float)):
                result = val * rate
                ir = int(result)
                out.append(ir if result == ir else round(result, 4))
            else:
                out.append(None)
        out.append(None)  # AJ: Unpaid N/A for CN
        _app((ordn, out))
    # Sort by precomputed ordinal only. Python's sort is stable, so ties
    # retain insertion order — same behavior as v2.0's _date_sort_key sort.
    decorated.sort(key=lambda t: t[0], reverse=True)
    return [r for _, r in decorated]

def build_cr_rows(sop_c, pop_c):
    """Build CR (Cash Received) rows — 13 cols matching sheet layout:
    A=Month, B=Date, C=Order/Number, D=Reference, E=Branch/Source,
    F=Journal, G=CustTitle, H=CustName, I=PayMethod,
    J=Amount, K=Currency, L=ExRate, M=Amount_THB
    VBA pastes Payment data (cols A-L) starting at CR col B, then adds Month formula at A.
    SO Payment cols: A=Date, B=Number, C=Reference, D=InvoiceSource, E=Journal,
                     F=CustTitle, G=CustName, H=PayMethod, I=Amount, J=Currency,
                     K=ExRate, L=Amount_THB (computed)
    POS Payment cols: same structure (D=POS Session instead of InvoiceSource)

    v3.0: fused date parser (one strptime per row instead of two) + Schwartzian
    sort that reuses the ordinal computed during row construction.
    """
    _parse = _parse_date_triple
    decorated = []
    _app = decorated.append
    for src in (sop_c, pop_c):
        for row in src[1:]:
            m, _wd, iso, ordn = _parse(row[0])
            _app((ordn, [m, iso, row[1], row[2], row[3], row[4],
                         row[5], row[6], row[7], row[8], row[9], row[10], row[11]]))
    # v2.0 _cr_sort_key returns numeric v directly when isinstance(v, (int, float)).
    # _parse_date_triple's failure fallback also returns numeric dv (see helper),
    # and its success path returns d.toordinal() — matching v2.0's strptime path.
    decorated.sort(key=lambda t: t[0], reverse=True)
    return [r for _, r in decorated]

# ── Odoo fetch helpers ───────────────────────────────────────────────────────
def pos_utc_range(year=ODOO_YEAR):
    """คำนวณ POS date range เป็น UTC โดยอ้างอิง Thailand timezone (UTC+7).
    Bangkok Jan 1 00:00  →  UTC Dec 31 prev-year 17:00:00
    Bangkok Dec 31 23:59:59 →  UTC Dec 31 curr-year 16:59:59
    """
    start = datetime(year, 1, 1, 0, 0, 0, tzinfo=TZ_BKK).astimezone(timezone.utc)
    end   = datetime(year, 12, 31, 23, 59, 59, tzinfo=TZ_BKK).astimezone(timezone.utc)
    return start.strftime('%Y-%m-%d %H:%M:%S'), end.strftime('%Y-%m-%d %H:%M:%S')

def _odoo_export(models, model, domain, fields, order_by=None):
    """Search + export_data. Returns list-of-lists (no header row)."""
    search_kw = {'limit': 0}
    if order_by:
        search_kw['order'] = order_by
    ids = models.execute_kw(ODOO_DB, ODOO_UID, ODOO_PW,
                            model, 'search', [domain], search_kw)
    if not ids:
        return []
    result = models.execute_kw(ODOO_DB, ODOO_UID, ODOO_PW,
                               model, 'export_data', [ids, fields], {})
    return result.get('datas', [])

def fetch_odoo_all(year=ODOO_YEAR, date_from=None, date_to=None):
    """ดึงข้อมูลครบ 5 ชุดจาก Odoo API ผ่าน XML-RPC (parallelized).
    date_from/date_to (YYYY-MM-DD): ถ้าส่งมา ใช้แทน year range (รองรับ cross-year)
    Returns: (so_raw, pos_raw, cn_raw, sopay_raw, pospay_raw) แต่ละตัวเป็น
    list-of-lists [header_row, data_row1, data_row2, ...] เหมือน read_xlsx()

    v2.0: 5 datasets fetched concurrently via ThreadPoolExecutor.
    xmlrpc.client.ServerProxy is not thread-safe, so each thread builds its
    own proxy via the local _fetch_one() helper. Field lists, domains, and
    order-by clauses are identical to v1.0 — only the dispatch is parallel.

    v3.0: each thread's ServerProxy now uses _KeepAliveSafeTransport, a
    SafeTransport subclass that keeps the HTTPS connection alive across the
    search + export_data calls (cuts one TCP/TLS handshake per dataset).
    Each thread still owns its own proxy + transport (xmlrpc.client is not
    thread-safe). Field lists, domains, and order-by clauses unchanged.

    v3.2: transport now also negotiates Accept-Encoding: gzip on the request
    and transparently decompresses gzip-encoded responses. Odoo's WSGI stack
    honours this for XML-RPC, and the savings on the redundant XML payloads
    (typically 80-90%) translate directly into lower wall-clock on the big
    SO/POS/CN datasets. A single throwaway common.version() warm-up call
    runs before the ThreadPoolExecutor fires so the TLS session cache is
    primed when the 5 worker connections handshake. Per-call semantics,
    field lists, domains, and order-by clauses are unchanged.
    """
    import xmlrpc.client
    import http.client
    import gzip
    from io import BytesIO
    from concurrent.futures import ThreadPoolExecutor

    if date_from and date_to:
        _df = datetime.strptime(date_from, '%Y-%m-%d')
        _dt = datetime.strptime(date_to, '%Y-%m-%d')
        pos_start = datetime(_df.year, _df.month, _df.day, 0, 0, 0, tzinfo=TZ_BKK).astimezone(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
        pos_end   = datetime(_dt.year, _dt.month, _dt.day, 23, 59, 59, tzinfo=TZ_BKK).astimezone(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
        _d_from, _d_to = date_from, date_to
    else:
        pos_start, pos_end = pos_utc_range(year)
        _d_from, _d_to = f'{year}-01-01', f'{year}-12-31'
    print(f"  POS UTC range (TH UTC+7): {pos_start} → {pos_end}")

    class _KeepAliveSafeTransport(xmlrpc.client.SafeTransport):
        """HTTPS transport that reuses one HTTPSConnection across calls AND
        negotiates gzip on the response. Not thread-safe — one instance per
        thread (matches ServerProxy contract).

        v3.2 changes vs v3.1:
          * accept_gzip=True is passed to SafeTransport.__init__ — stdlib
            xmlrpc.client already knows how to decode Content-Encoding: gzip
            responses via GzipDecodedResponse when this flag is set. We also
            inject an Accept-Encoding: gzip header on the request side to be
            explicit (some Odoo proxies require it on the wire to opt-in).
          * make_connection still reuses the single HTTPSConnection across
            search + export_data, identical to v3.1 keep-alive behaviour.
        """
        def __init__(self):
            # use_datetime=False, use_builtin_types=False — match v3.1 defaults.
            # accept_gzip=True instructs xmlrpc.client to decode gzip responses.
            super().__init__(use_datetime=False, use_builtin_types=False)
            self.accept_gzip = True
        def make_connection(self, host):
            if self._connection and host == self._connection[0]:
                return self._connection[1]
            chost, self._extra_headers, x509 = self.get_host_info(host)
            # Add Accept-Encoding to the per-host extra headers so every call
            # advertises gzip support. _extra_headers is a list of (k, v) tuples
            # consumed by Transport.send_headers().
            existing = list(self._extra_headers or [])
            if not any(k.lower() == 'accept-encoding' for k, _v in existing):
                existing.append(('Accept-Encoding', 'gzip'))
            self._extra_headers = existing
            self._connection = host, http.client.HTTPSConnection(
                chost, None, context=self.context, **(x509 or {}))
            return self._connection[1]
        def close(self):
            # default Transport.close() closes the connection after every request;
            # override to no-op so the socket survives across execute_kw calls.
            pass

    def _fetch_one(model, domain, fields, order_by):
        """Thread-local fetch: own ServerProxy + search + export_data.
        v3.0: uses keep-alive transport so search + export_data share one TCP/TLS conn.
        v3.2: transport also decodes gzip responses (large XML → small wire).
        """
        transport = _KeepAliveSafeTransport()
        try:
            proxy = xmlrpc.client.ServerProxy(
                f'{ODOO_URL}/xmlrpc/2/object', transport=transport)
            return _odoo_export(proxy, model, domain, fields, order_by=order_by)
        finally:
            # Tear down the persistent connection explicitly (our close() is a no-op).
            try:
                if transport._connection and transport._connection[1]:
                    transport._connection[1].close()
            except Exception:
                pass

    # ── TLS warm-up ──────────────────────────────────────────────────────────
    # v3.2: single throwaway call to /xmlrpc/2/common before the parallel
    # ThreadPoolExecutor fires. Goals:
    #   1. Resolve DNS once (subsequent threads hit OS cache).
    #   2. Prime the TLS session cache — some TLS stacks will reuse the
    #      session ticket on the next handshake, shaving ~100ms per worker
    #      on high-RTT links.
    # Failures here are non-fatal: the worker fetches are independent and
    # will perform their own handshakes if the warm-up didn't succeed.
    try:
        _warm_transport = _KeepAliveSafeTransport()
        _warm_proxy = xmlrpc.client.ServerProxy(
            f'{ODOO_URL}/xmlrpc/2/common', transport=_warm_transport)
        _warm_proxy.version()
        try:
            if _warm_transport._connection and _warm_transport._connection[1]:
                _warm_transport._connection[1].close()
        except Exception:
            pass
    except Exception:
        # Warm-up is purely an optimisation; never block the real fetch on it.
        pass

    # ── 1. Sales Orders ──────────────────────────────────────────────────────
    # field list ยืนยันแล้วจาก export_data sample (27 cols)
    SO_FIELDS = [
        'x_studio_actual_order_date',               # 0  Actual Order Date
        'name',                                      # 1  SO No (KEY)
        'origin',                                    # 2  Source Document
        'campaign_id',                               # 3  Campaign
        'source_id',                                 # 4  Source
        'x_studio_foreigner',                        # 5  Foreigner (False→Thai)
        'partner_id/title',                          # 6  Customer Title
        'partner_id/name',                           # 7  Customer Name
        'contact_partner_id/title',                  # 8  Contact Partner Title
        'contact_partner_id/name',                   # 9  Contact Partner Name
        'user_id',                                   # 10 Salesperson
        'payment_term_id',                           # 11 Payment Terms (→ In/Pre-Stock)
        'order_line/product_id/product_tmpl_id',     # 12 Product Template
        'order_line/product_id',                     # 13 Product
        'order_line/product_id/barcode',             # 14 Barcode
        'order_line/product_id/categ_id',            # 15 Category (→ split 4 cols)
        'order_line/product_id/product_brand_id',    # 16 Brand
        'order_line/route_id',                       # 17 Route
        'order_line/product_uom_qty',                # 18 Qty Ordered ← cleanse checks
        'order_line/tax_id',                         # 19 Taxes
        'order_line/price_unit',                     # 20 Unit Price
        'order_line/formula_discount',               # 21 Disc. %
        'order_line/price_subtotal_wo_prorated',     # 22 Sale Closed → THB
        'amount_total',                              # 23 Total (header) → THB
        'amount_unpaid',                             # 24 Unpaid → THB
        'order_line/currency_id/display_name',       # 25 Currency
        'x_studio_currency_rate_for_sale_report',    # 26 Report Rate (filled-down)
    ]
    so_domain = [
        ['state', 'in', ['sale', 'done']],
        ['x_studio_actual_order_date', '>=', _d_from],
        ['x_studio_actual_order_date', '<=', _d_to],
        ['source_id', 'not ilike', 'consignment'],
        ['tag_ids', 'not ilike', 'foc'],
        ['tag_ids', 'not ilike', 'gift'],
    ]
    # SO fetch submitted below; result joined into so_raw after all futures complete

    # ── 2. POS Orders ────────────────────────────────────────────────────────
    # หมายเหตุ: date_order ใน Odoo เก็บเป็น UTC → ต้องใช้ pos_utc_range() ที่คำนวณจาก TH UTC+7
    # Fields ยืนยันจาก Odoo export UI บน pos.order (ใช้ชื่อ field จริงในวงเล็บ)
    POS_FIELDS = [
        'date_order',                                # 0  Date — ใช้ get_date() ใน build_ytd
        'name',                                      # 1  Order Reference (KEY)
        'note',                                      # 2  Internal Notes
        'x_studio_many2one_field_GEWnd',             # 3  Campaign
        'config_id/name',                            # 4  Point of Sale / Store
        'x_studio_foreigner',                        # 5  Foreigner flag (True/False โดยตรง)
        'partner_id/title',                          # 6  Customer Title
        'partner_id/name',                           # 7  Customer Name
        'x_studio_contact_partner/title',            # 8  Contact Partner Title
        'x_studio_contact_partner/name',             # 9  Contact Partner Name
        'cashier',                                   # 10 Cashier
        'lines/pack_lot_ids',                        # 11 Lot/Serial (→ 'In-Stock' by cleanse)
        'lines/product_id/product_tmpl_id',          # 12 Product Template
        'lines/full_product_name',                   # 13 Product Full Name
        'lines/product_id/barcode',                  # 14 Barcode
        'lines/product_id/categ_id',                 # 15 Category (→ split 4 cols)
        'lines/product_id/product_brand_id',         # 16 Brand
        'lines/customer_note',                       # 17 Customer Note
        'lines/qty',                                 # 18 Qty ← cleanse checks
        'lines/tax_ids',                             # 19 Taxes
        'lines/price_unit',                          # 20 Unit Price
        'lines/discount',                            # 21 Discount %
        'lines/price_subtotal_incl',                 # 22 Sale Closed → THB
        'amount_total',                              # 23 Total (header) → THB
        'lines/currency_id/display_name',            # 24 Currency
        'currency_id/rate_ids',                      # 25 Exchange Rate (filled-down)
    ]
    # domain ตรงกับ Favorites filter id=206 + ใช้ UTC ที่คำนวณจาก TH timezone
    pos_domain = [
        '|', ['config_id', 'ilike', 'norse'], ['config_id', 'ilike', 'hay'],
        '|', ['state', '=', 'invoiced'], ['state', '=', 'done'],
        ['date_order', '>=', pos_start],   # <-- Thailand UTC+7 corrected
        ['date_order', '<=', pos_end],     # <-- Thailand UTC+7 corrected
    ]
    # POS fetch submitted below; x_studio_foreigner returns True/False directly

    # ── 3. Credit Notes ──────────────────────────────────────────────────────
    CN_FIELDS = [
        'invoice_date',                                    # 0  Invoice Date
        'name',                                            # 1  CN No (KEY)
        'invoice_origin',                                  # 2  Source Document
        'campaign_id',                                     # 3  Campaign
        'source_id',                                       # 4  Source
        'contact_partner_id/country_id',                   # 5  Country → Thai/Non-Thai
        'partner_id/title',                                # 6  Customer Title
        'partner_id/name',                                 # 7  Customer Name
        'contact_partner_id/title',                        # 8  Contact Title
        'contact_partner_id/name',                         # 9  Contact Name
        'invoice_user_id',                                 # 10 Salesperson
        'invoice_payment_term_id',                         # 11 Payment Terms
        'invoice_line_ids/product_id/product_tmpl_id',     # 12 Product Template
        'invoice_line_ids/product_id',                     # 13 Product
        'invoice_line_ids/product_id/barcode',             # 14 Barcode
        'invoice_line_ids/product_id/categ_id',            # 15 Category (→ split)
        'invoice_line_ids/product_id/product_brand_id',    # 16 Brand
        'invoice_line_ids/analytic_distribution',          # 17 Analytic Distribution
        'invoice_line_ids/quantity',                       # 18 Qty ← negated by cleanse
        'invoice_line_ids/tax_ids',                        # 19 Taxes
        'invoice_line_ids/price_unit',                     # 20 Unit Price
        'invoice_line_ids/formula_discount',               # 21 Disc. %
        'invoice_line_ids/price_subtotal_wo_prorated',     # 22 Sale Closed ← negated
        'invoice_line_ids/price_total',                    # 23 Total ← negated, carry
        'invoice_line_ids/company_currency_id/display_name', # 24 Currency
        'currency_id/rate_ids',                            # 25 Rate (filled-down)
    ]
    cn_domain = [
        ['move_type', '=', 'out_refund'],
        ['state', '=', 'posted'],
        ['invoice_date', '>=', _d_from],
        ['invoice_date', '<=', _d_to],
        ['invoice_origin', 'not ilike', 'refund'],
        ['source_id', 'not ilike', 'consignment'],
    ]
    # CN fetch submitted below; country→Thai/Non-Thai flip applied after rows arrive

    # ── 4. SO Payments (Payments with Invoices view) ──────────────────────────
    SOPAY_FIELDS = [
        'date',                                          # 0  Date
        'name',                                          # 1  Payment Number
        'payment_line_invoice_ids/display_name',         # 2  Invoice Doc No
        'payment_line_invoice_ids/move_id/source_id',   # 3  Invoice Source
        'journal_id',                                    # 4  Journal
        'partner_id/title',                              # 5  Customer/Vendor Title
        'partner_id/name',                               # 6  Customer/Vendor Name
        'payment_method_ids/payment_method_line_id',     # 7  Payment Method ← AP filter
        'payment_method_ids/amount_total',               # 8  Amount (per method)
        'currency_id',                                   # 9  Currency
        'exchange_rate',                                 # 10 Exchange Rate
    ]
    sopay_domain = [
        ['is_pay_with_inv', '=', True],
        ['partner_type', '=', 'customer'],
        ['is_internal_transfer', '=', False],
        ['state', '=', 'posted'],
        ['date', '>=', _d_from],
        ['date', '<=', _d_to],
        ['create_uid', 'not ilike', 'norse x'],
    ]
    # SO Payment fetch submitted below

    # ── 5. POS Payments ──────────────────────────────────────────────────────
    POSPAY_FIELDS = [
        'date',                                          # 0  Date
        'name',                                          # 1  Payment Number
        'ref',                                           # 2  Reference
        'pos_session_id/config_id/name',                 # 3  POS Store
        'journal_id',                                    # 4  Journal
        'partner_id/title',                              # 5  Customer/Vendor Title
        'partner_id/name',                               # 6  Customer/Vendor Name
        'payment_method_ids/payment_method_line_id',     # 7  Payment Method ← AP filter
        'amount',                                        # 8  Amount
        'currency_id',                                   # 9  Currency
        'exchange_rate',                                 # 10 Exchange Rate
    ]
    pospay_domain = [
        ['pos_session_id', '!=', False],
        ['partner_type', '=', 'customer'],
        ['is_internal_transfer', '=', False],
        ['state', '=', 'posted'],
        ['date', '>=', _d_from],
        ['date', '<=', _d_to],
        ['source_id', 'not ilike', 'consignment'],
    ]
    # POS Payment fetch submitted below

    # ── Submit all 5 fetches concurrently ────────────────────────────────────
    # Each thread builds its own ServerProxy inside _fetch_one (xml-rpc ServerProxy
    # is NOT thread-safe). Order-by clauses are passed verbatim so output row
    # ordering is byte-identical to the v1.0 sequential path.
    with ThreadPoolExecutor(max_workers=5) as ex:
        fut_so     = ex.submit(_fetch_one, 'sale.order',      so_domain,    SO_FIELDS,
                               'x_studio_actual_order_date asc, id asc')
        fut_pos    = ex.submit(_fetch_one, 'pos.order',       pos_domain,   POS_FIELDS,
                               'date_order asc, id asc')
        fut_cn     = ex.submit(_fetch_one, 'account.move',    cn_domain,    CN_FIELDS,
                               'invoice_date asc, id asc')
        fut_sopay  = ex.submit(_fetch_one, 'account.payment', sopay_domain, SOPAY_FIELDS,
                               'date asc, id asc')
        fut_pospay = ex.submit(_fetch_one, 'account.payment', pospay_domain, POSPAY_FIELDS,
                               'date asc, id asc')
        so_rows     = fut_so.result()
        pos_rows    = fut_pos.result()
        cn_rows     = fut_cn.result()
        sopay_rows  = fut_sopay.result()
        pospay_rows = fut_pospay.result()

    # Post-CN country flip (must run after CN rows arrive — same semantics as v1.0):
    # col 5 = contact_partner_id/country_id → Thai/Non-Thai flag
    for r in cn_rows:
        if len(r) > 5:
            country = str(r[5]).lower()
            r[5] = False if country in ('thailand', 'th') else True

    so_raw     = [SO_FIELDS] + so_rows
    pos_raw    = [POS_FIELDS] + pos_rows
    cn_raw     = [CN_FIELDS] + cn_rows
    sopay_raw  = [SOPAY_FIELDS] + sopay_rows
    pospay_raw = [POSPAY_FIELDS] + pospay_rows

    return so_raw, pos_raw, cn_raw, sopay_raw, pospay_raw

# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    t0 = time.time()

    # ── Parse args ──────────────────────────────────────────────────────────
    parser = argparse.ArgumentParser(description='HAY Daily Sale Report')
    parser.add_argument('--so',     help='Sales Order xlsx')
    parser.add_argument('--pos',    help='POS Orders xlsx')
    parser.add_argument('--cn',     help='Credit Notes xlsx')
    parser.add_argument('--sopay',  help='SO Payments xlsx')
    parser.add_argument('--pospay', help='POS Payments xlsx')
    parser.add_argument('--downloads', default=DOWNLOADS,
                        help='Downloads folder (default: ~/Downloads)')
    parser.add_argument('--max-age', type=float, default=None, metavar='HOURS',
                        help='Only use files modified within last N hours '
                             '(e.g. --max-age 4). Default: no limit (use newest file)')
    parser.add_argument('--fetch-odoo', action='store_true',
                        help='ดึงข้อมูลจาก Odoo API โดยตรง (ไม่ต้อง export xlsx เอง)')
    parser.add_argument('--year', type=int, default=ODOO_YEAR,
                        help=f'ปีที่ต้องการดึงข้อมูล (default: {ODOO_YEAR})')
    # v3.7: --date for regenerating a specific past day's report
    parser.add_argument('--date', default=None,
                        help='Report date YYYY-MM-DD (default: today). PDFs '
                             'for this date are regenerated and the email '
                             'subject/folder reflect the chosen date.')
    parser.add_argument('--excel-only', action='store_true',
                        help='Email the patched XLSM file directly (skip PDF). '
                             'Used for on-demand requests from Odoo.')
    parser.add_argument('--last13', action='store_true',
                        help='Fetch last 13 months of data (spans 2 years if needed).')
    args = parser.parse_args()

    # v3.7: parse --date into a tz-aware datetime
    # v3.9: when --date is NOT provided, default to YESTERDAY (not today).
    # Sales reports are typically run in the morning for the previous
    # day's (finalized) sales — yesterday is the right default. Explicit
    # --date YYYY-MM-DD overrides.
    if args.date:
        try:
            report_date = datetime.strptime(args.date, '%Y-%m-%d').replace(tzinfo=TZ_BKK)
        except ValueError:
            print(f"ERROR: --date must be YYYY-MM-DD format (got {args.date!r})")
            sys.exit(1)
        _date_label = f"{report_date.day} {report_date.strftime('%b')} {report_date.year}"
        print(f"[v3.7] Running report for: {_date_label}")
    else:
        # v3.9: default = yesterday (Bangkok time)
        report_date = datetime.now(TZ_BKK) - timedelta(days=1)
        _date_label = f"{report_date.day} {report_date.strftime('%b')} {report_date.year}"
        print(f"[v3.9] No --date provided — defaulting to YESTERDAY: {_date_label}")

    # v3.12: log Graph secret expiry warning at startup (if configured + nearing)
    _exp = _check_secret_expiry()
    if _exp:
        _days, _sev, _msg = _exp
        print(f"[v3.12 {_sev.upper()}] {_msg}")

    # ── Check Excel is not open (ignore stale locks > 5 min old) ───────────
    xlsm_name = os.path.basename(XLSM)
    lock_file = os.path.join(SCRIPT_DIR, "~$" + xlsm_name)
    if os.path.exists(lock_file):
        lock_age = time.time() - os.path.getmtime(lock_file)
        if lock_age < 300:
            print("ERROR: Please close the report in Excel first!")
            print("  Excel lock file found:", lock_file)
            sys.exit(1)

    # ── Fetch / Read data ────────────────────────────────────────────────────
    if args.fetch_odoo:
        if args.last13:
            today_bkk = datetime.now(TZ_BKK)
            d_to = today_bkk.strftime('%Y-%m-%d')
            m13 = today_bkk.month - 13
            y13 = today_bkk.year
            while m13 <= 0:
                m13 += 12
                y13 -= 1
            d_from = f'{y13}-{m13:02d}-01'
            print(f"Fetching from Odoo API (last 13 months: {d_from} → {d_to})...")
            so_raw, pos_raw, cn_raw, sopay_raw, pospay_raw = fetch_odoo_all(date_from=d_from, date_to=d_to)
        else:
            print(f"Fetching from Odoo API (year={args.year})...")
            so_raw, pos_raw, cn_raw, sopay_raw, pospay_raw = fetch_odoo_all(args.year)
        print(f"  Raw: SO={len(so_raw)-1} POS={len(pos_raw)-1} CN={len(cn_raw)-1} "
              f"SOPay={len(sopay_raw)-1} POSPay={len(pospay_raw)-1}  "
              f"({time.time()-t0:.1f}s)")
    else:
        # ── Find xlsx files ──────────────────────────────────────────────────
        if args.so:
            so_f, pos_f, cn_f, sopay_f, pospay_f = (
                args.so, args.pos, args.cn, args.sopay, args.pospay)
        else:
            age_label = f"ภายใน {args.max_age} ชั่วโมงล่าสุด" if args.max_age else "ล่าสุดในโฟลเดอร์"
            print(f"Auto-detecting xlsx files ({age_label}): {args.downloads}")
            so_f, pos_f, cn_f, sopay_f, pospay_f = detect_files(args.downloads, args.max_age)

        missing = [(name, path) for name, path in
                   [('SO', so_f), ('POS', pos_f), ('CN', cn_f),
                    ('SO Payment', sopay_f), ('POS Payment', pospay_f)]
                   if not path or not os.path.exists(path)]
        if missing:
            print("\nERROR: ไม่พบไฟล์ xlsx:")
            for name, path in missing:
                print(f"  {name}: {path or 'ไม่พบ'}")
            if args.max_age:
                print(f"\n  ไฟล์อาจเก่ากว่า {args.max_age} ชั่วโมง — ลอง --max-age ที่มากขึ้น")
            print("  หรือระบุ path ตรงๆ ด้วย --so / --pos / --cn / --sopay / --pospay")
            print("  หรือใช้ --fetch-odoo เพื่อดึงจาก Odoo โดยตรง")
            sys.exit(1)

        print(f"Files detected ({time.time()-t0:.1f}s):")
        for label, path in [('SO    ', so_f), ('POS   ', pos_f), ('CN    ', cn_f),
                             ('SOPay ', sopay_f), ('POSPay', pospay_f)]:
            print(f"  {label}: {os.path.basename(path)}  [{fmt_mtime(path)}]")
        if sopay_f == pospay_f:
            print("  ⚠  คำเตือน: SO Payment และ POS Payment เป็นไฟล์เดียวกัน")

        print(f"\nReading xlsx files...")
        so_raw    = read_xlsx(so_f)
        pos_raw   = read_xlsx(pos_f)
        cn_raw    = read_xlsx(cn_f)
        sopay_raw = read_xlsx(sopay_f)
        pospay_raw= read_xlsx(pospay_f)
        print(f"  SO={len(so_raw)-1} POS={len(pos_raw)-1} CN={len(cn_raw)-1} "
              f"SOPay={len(sopay_raw)-1} POSPay={len(pospay_raw)-1} rows  "
              f"({time.time()-t0:.1f}s)")

    # ── Cleanse ──────────────────────────────────────────────────────────────
    print("Cleansing...")
    so_c     = cleanse_so(so_raw)
    pos_c    = cleanse_pos(pos_raw)
    cn_c     = cleanse_cn(cn_raw)
    sop_c    = cleanse_sopay(sopay_raw)
    pop_c    = cleanse_pospay(pospay_raw)
    print(f"  SO={len(so_c)-1} POS={len(pos_c)-1} CN={len(cn_c)-1} "
          f"SOPay={len(sop_c)-1} POSPay={len(pop_c)-1} clean rows  "
          f"({time.time()-t0:.1f}s)")

    # ── Build YTD + CR rows ──────────────────────────────────────────────────
    print("Building report rows...")
    ytd_rows = build_ytd_rows(so_c, pos_c, cn_c)
    cr_rows  = build_cr_rows(sop_c, pop_c)
    print(f"  YTD={len(ytd_rows)} rows × {len(ytd_rows[0]) if ytd_rows else 0} cols  "
          f"CR={len(cr_rows)} rows × {len(cr_rows[0]) if cr_rows else 0} cols  "
          f"({time.time()-t0:.1f}s)")

    # ── Patch xlsm (always from .bak for clean XML structure) ───────────────
    # v3.3: single open of the .bak file — we read the infolist AND every
    # member's bytes in one pass, then close the input handle before opening
    # the output. The .bak is ~7.4 MB so the in-memory copy is fine.
    bak = XLSM_BAK if os.path.exists(XLSM_BAK) else XLSM
    print(f"Reading base from: {os.path.basename(bak)}")
    with zipfile.ZipFile(bak, 'r') as zin:
        infolist = zin.infolist()
        raw_blobs = {it.filename: zin.read(it.filename) for it in infolist}

    print("Patching all sheets...")
    # v3.3: YTD + CR fuse strip_formula_cache into patch_xml's single bytes pass.
    new_ytd = patch_xml(raw_blobs['xl/worksheets/sheet4.xml'], ytd_rows,
                        first_row=5, strip_formula=True)
    new_cr  = patch_xml(raw_blobs['xl/worksheets/sheet1.xml'], cr_rows,
                        first_row=5, strip_formula=True)
    new_so  = patch_xml(raw_blobs['xl/worksheets/sheet5.xml'], so_c,  first_row=1)
    new_pos = patch_xml(raw_blobs['xl/worksheets/sheet6.xml'], pos_c, first_row=1)
    new_cn  = patch_xml(raw_blobs['xl/worksheets/sheet7.xml'], cn_c,  first_row=1)
    new_sop = patch_xml(raw_blobs['xl/worksheets/sheet8.xml'], sop_c, first_row=1)
    new_pop = patch_xml(raw_blobs['xl/worksheets/sheet9.xml'], pop_c, first_row=1)
    print(f"  YTD: {len(new_ytd):,} bytes  CR: {len(new_cr):,} bytes  "
          f"({time.time()-t0:.1f}s)")

    sheet_map = {
        'xl/worksheets/sheet4.xml': new_ytd,
        'xl/worksheets/sheet1.xml': new_cr,
        'xl/worksheets/sheet5.xml': new_so,
        'xl/worksheets/sheet6.xml': new_pos,
        'xl/worksheets/sheet7.xml': new_cn,
        'xl/worksheets/sheet8.xml': new_sop,
        'xl/worksheets/sheet9.xml': new_pop,
    }

    print("Writing xlsm...")
    tmp = XLSM + '.tmp'
    # v3.3: compresslevel=3 — about 2× faster than the default level 6 on
    # embedded-XML payloads, with a small (~2-3%) output-size penalty. Excel
    # only cares that members are DEFLATE-encoded, not the chosen level.
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED, compresslevel=3) as zout:
        for item in infolist:
            name = item.filename
            if name == 'xl/calcChain.xml':
                continue  # ลบ calcChain เพื่อให้ Excel สร้างใหม่อัตโนมัติ
            data = sheet_map.get(name)
            if data is None:
                data = raw_blobs[name]
            # บอก Excel ให้ recalculate แบบ full dependency order ตอนเปิดไฟล์
            # ป้องกันปัญหา SUBTOTAL ได้ค่าผิดเพราะ formula ถูก calc ก่อน dependency resolve
            if name == 'xl/workbook.xml':
                wb = data.decode('utf-8')
                if '<calcPr' in wb:
                    # เพิ่ม/อัพเดท fullCalcOnLoad="1" ใน tag calcPr ที่มีอยู่แล้ว
                    wb = _RE_FULLCALC_ATTR.sub('', wb)
                    wb = _RE_CALCPR_TAG.sub(r'\1 fullCalcOnLoad="1"\2', wb)
                else:
                    # ไม่มี calcPr เลย → เพิ่มก่อน </workbook>
                    wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
                data = wb.encode('utf-8')
            zout.writestr(item, data)
    os.replace(tmp, XLSM)

    elapsed = time.time() - t0
    print("Done! Total time: %.1fs" % elapsed)
    print("  File: " + XLSM)

    # ── Excel-only mode (on-demand request from Odoo) ────────────────────────
    if args.excel_only:
        try:
            send_excel_email(XLSM, report_date=report_date,
                             year=args.year if not args.last13 else 'Last 13 Months')
        except Exception as e:
            print(f'\n[EMAIL] ERROR: {e}')
            traceback.print_exc()
        return

    # ── Generate Daily PDFs ──────────────────────────────────────────────────
    # v3.6: capture list of (branch, pdf_path) so we can email + check coverage.
    # v3.7: pass report_date so the PDFs land in the chosen day's folder.
    generated_pdfs = generate_daily_pdfs(XLSM, report_date=report_date) or []
    generated_branches = {b for b, _ in generated_pdfs}
    expected_branches  = set(DAILY_BRANCHES)
    missing            = sorted(expected_branches - generated_branches)

    # ── Email out the PDFs ───────────────────────────────────────────────────
    if generated_pdfs:
        try:
            paths = [p for _, p in generated_pdfs]
            send_daily_email(paths,
                             missing_branches=missing if missing else None,
                             report_date=report_date)
        except Exception as e:
            print(f"\n[EMAIL] ERROR: {e}")
            # Send a separate failure notification (best-effort)
            try:
                send_failure_email(str(e), 'email send',
                                   traceback.format_exc(),
                                   report_date=report_date)
            except Exception:
                pass
    else:
        # No PDFs at all — email a failure notification
        try:
            send_failure_email("generate_daily_pdfs returned no files",
                               'PDF generation', None,
                               report_date=report_date)
        except Exception as e:
            print(f"\n[EMAIL] ERROR sending failure notification: {e}")

# ── v3.6: Email helpers (Microsoft Graph API) ────────────────────────────────
def _get_access_token():
    """OAuth 2.0 client credentials flow → access_token (cached per run)."""
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN
    if not all([GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET]):
        raise RuntimeError(
            "Missing Graph credentials in .env "
            "(need GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET)")
    try:
        import requests
    except ImportError as e:
        raise RuntimeError(
            "Missing 'requests' package — pip install requests") from e
    url = GRAPH_TOKEN_URL.format(tenant=GRAPH_TENANT_ID)
    resp = requests.post(url, data={
        'client_id':     GRAPH_CLIENT_ID,
        'client_secret': GRAPH_CLIENT_SECRET,
        'scope':         'https://graph.microsoft.com/.default',
        'grant_type':    'client_credentials',
    }, timeout=15)
    if resp.status_code != 200:
        try:
            err = resp.json()
            detail = f"{err.get('error')} — {err.get('error_description','')[:200]}"
        except Exception:
            detail = resp.text[:200]
        raise RuntimeError(f"Graph token request failed ({resp.status_code}): {detail}")
    _GRAPH_TOKEN = resp.json().get('access_token')
    if not _GRAPH_TOKEN:
        raise RuntimeError("Graph token response missing access_token")
    return _GRAPH_TOKEN


def _post_graph_message(subject, body_text, attachments_bin=None):
    """Build payload + POST to /sendMail. attachments_bin = list of (filename, bytes)."""
    import requests
    token = _get_access_token()
    msg_attachments = []
    if attachments_bin:
        for entry in attachments_bin:
            fname, raw = entry[0], entry[1]
            ctype = entry[2] if len(entry) > 2 else 'application/octet-stream'
            b64 = base64.b64encode(raw).decode('ascii')
            msg_attachments.append({
                '@odata.type':  '#microsoft.graph.fileAttachment',
                'name':         fname,
                'contentType':  ctype,
                'contentBytes': b64,
            })
    # v3.11: To/CC/BCC all sourced from comma-separated env vars
    msg_obj = {
        'subject': subject,
        'body':    {'contentType': 'Text', 'content': body_text},
        'from':    {'emailAddress': {'name': EMAIL_FROM_NAME, 'address': EMAIL_FROM}},
        'toRecipients':  _parse_recipients(EMAIL_TO),
        'attachments':   msg_attachments,
    }
    cc_list  = _parse_recipients(EMAIL_CC)
    bcc_list = _parse_recipients(EMAIL_BCC)
    if cc_list:
        msg_obj['ccRecipients'] = cc_list
    if bcc_list:
        msg_obj['bccRecipients'] = bcc_list
    payload = {'message': msg_obj, 'saveToSentItems': True}
    url = GRAPH_SEND_URL.format(user=EMAIL_FROM)
    resp = requests.post(url, json=payload, timeout=60, headers={
        'Authorization': f'Bearer {token}',
        'Content-Type':  'application/json',
    })
    if resp.status_code == 202:
        return True
    # Surface the actual server error
    try:
        err = resp.json().get('error', {})
        detail = f"{err.get('code','?')} — {err.get('message','?')[:300]}"
    except Exception:
        detail = resp.text[:300]
    raise RuntimeError(f"Graph sendMail failed ({resp.status_code}): {detail}")


def send_daily_email(pdf_paths, missing_branches=None, report_date=None):
    """Send the daily Sale Report email with PDFs attached.

    pdf_paths: iterable of absolute paths to PDF files to attach.
    missing_branches: optional list of branch names that FAILED to export
                     (mentioned in the body so the recipient knows the deck
                     is incomplete). None or empty → body uses the standard
                     wording with no missing-branch line.
    report_date: v3.7 — optional datetime; subject + body wording use this
                 date instead of "today". None → today (legacy).
    """
    today    = report_date or datetime.now(TZ_BKK)
    date_str = f"{today.day} {today.strftime('%b')} {today.year}"  # no leading zero
    subject  = f"Sale Report {date_str}"

    # Body — exactly as specified by user (English, signed 'Operations').
    # When some branches are missing, prepend a note to the body so the
    # recipient knows the report is incomplete and can ask for a re-run.
    lines = ["Dear All,", ""]
    # v3.12: prepend secret-expiry warning at top of body if applicable
    _exp = _check_secret_expiry()
    if _exp:
        _days, _sev, _msg = _exp
        lines.append(_msg)
        lines.append("")
    if missing_branches:
        lines.append(
            f"NOTE: This report is incomplete — the following {len(missing_branches)} branch(es) "
            f"were not generated successfully: {', '.join(missing_branches)}.")
        lines.append("Please contact the operations team for a re-run.")
        lines.append("")
    lines += [
        "Please see the daily sale report as the attachments.",
        "",
        "Best Regards,",
        "Operations",
    ]
    body_text = "\n".join(lines)

    # Read each PDF into memory
    attachments_bin = []
    total_b64 = 0
    for path in pdf_paths:
        if not os.path.exists(path):
            print(f"[EMAIL]  ⚠  Missing attachment file (skipped): {path}")
            continue
        with open(path, 'rb') as f:
            raw = f.read()
        attachments_bin.append((os.path.basename(path), raw))
        total_b64 += int(len(raw) * 1.37)  # rough base64 overhead

    # Graph /sendMail caps payload around 4 MB; warn at 3.5 MB
    MAX_BYTES = int(3.5 * 1024 * 1024)
    if total_b64 > MAX_BYTES:
        raise RuntimeError(
            f"Attachments total ~{total_b64/1024/1024:.1f} MB after base64, "
            f"exceeds Graph /sendMail ~4 MB limit. Use upload-session flow for large files.")

    _to_str = EMAIL_TO if EMAIL_TO else "(no recipient!)"
    _cc_str = f" + CC: {EMAIL_CC}" if EMAIL_CC else ""
    _bcc_str = f" + BCC: {EMAIL_BCC}" if EMAIL_BCC else ""
    print(f"[EMAIL]  Sending {len(attachments_bin)} PDF(s) -> {_to_str}{_cc_str}{_bcc_str} ...")
    _post_graph_message(subject, body_text, attachments_bin)
    print(f"[EMAIL]  ✓ Sent successfully")


def send_excel_email(xlsm_path, report_date=None, year=None):
    """Email the patched XLSM file as an on-demand report (skip PDF generation).
    Used when --excel-only is passed (triggered from Odoo Server Action).
    """
    today    = report_date or datetime.now(TZ_BKK)
    yr       = year or today.year
    fname    = f'Sale Report {yr}.xlsm'
    subject  = f'Sale Report (Excel) {yr}'
    body_text = '\n'.join([
        'Dear All,',
        '',
        'Please find the requested sale report (Excel file) attached.',
        '',
        'Best Regards,',
        'Operations',
    ])
    with open(xlsm_path, 'rb') as f:
        raw = f.read()
    attachments_bin = [(
        fname, raw,
        'application/vnd.ms-excel.sheet.macroEnabled.12'
    )]
    _to_str = EMAIL_TO if EMAIL_TO else '(no recipient!)'
    print(f'[EMAIL]  Sending Excel -> {_to_str} ...')
    _post_graph_message(subject, body_text, attachments_bin)
    print(f'[EMAIL]  ✓ Sent successfully')


def send_failure_email(error_msg, phase, trace_str=None, report_date=None):
    """Send a FAILURE notification when the pipeline crashes or is incomplete.

    v3.7: report_date (optional) — if provided, subject + body wording
    mention this specific date instead of generic "the report".
    """
    today    = report_date or datetime.now(TZ_BKK)
    date_str = f"{today.day} {today.strftime('%b')} {today.year}"
    subject  = f"Sale Report FAILED {date_str}"

    body_lines = [
        "Dear All,",
        ""]
    # v3.12: prepend secret-expiry warning at top of body if applicable
    _exp = _check_secret_expiry()
    if _exp:
        _days, _sev, _msg = _exp
        body_lines.append(_msg)
        body_lines.append("")
    body_lines += [
        f"The daily Sale Report pipeline FAILED while processing the report "
        f"for {date_str} at phase: {phase}.",
        "",
        "Error:",
        str(error_msg),
        "",
    ]
    if trace_str:
        body_lines += ["Traceback (for IT):", trace_str, ""]
    body_lines += [
        "Please contact the IT/development team to investigate.",
        "",
        "Best Regards,",
        "Operations (automated)",
    ]
    body_text = "\n".join(body_lines)

    try:
        _to_str = EMAIL_TO if EMAIL_TO else "(no recipient!)"
        _cc_str = f" + CC: {EMAIL_CC}" if EMAIL_CC else ""
        _bcc_str = f" + BCC: {EMAIL_BCC}" if EMAIL_BCC else ""
        print(f"[EMAIL]  Sending failure notification -> {_to_str}{_cc_str}{_bcc_str} ...")
        _post_graph_message(subject, body_text, None)
        print(f"[EMAIL]  ✓ Failure notification sent")
    except Exception as e:
        # Don't raise from the notifier — original error is what matters
        print(f"[EMAIL]  ⚠  Failed to send failure notification: {e}")


# ── Daily PDF generator (win32com) ───────────────────────────────────────────
DAILY_BRANCHES = [
    'Somkid', 'NORSE Store', 'Line Chat',
    'Line My Shop', 'Lazada', 'HAY Store', 'Wholesale',
]

# v3.4: hoist the per-branch filename-sanitiser regex to module scope so it
# is compiled once at import time, not 7× per run inside the loop.
_RE_BRANCH_UNSAFE = re.compile(r'[\\/*?:"<>|]')

# Excel constants (win32com late-binding doesn't expose them reliably as
# attributes, so we use the documented integer values).
_XL_CALC_MANUAL    = -4135
_XL_CALC_AUTOMATIC = -4105

def _libreoffice_python():
    """v3.13: find a Python interpreter that can import the UNO bridge.
    Linux runner: system python3 with python3-uno installed (= current
    interpreter). Windows: LibreOffice's bundled python.exe."""
    try:
        import uno  # noqa: F401
        return sys.executable
    except ImportError:
        pass
    for p in (r'C:\Program Files\LibreOffice\program\python.exe',
              r'C:\Program Files (x86)\LibreOffice\program\python.exe'):
        if os.path.exists(p):
            return p
    return None

def _generate_daily_pdfs_libreoffice(xlsm_path, report_date=None):
    """v3.13: Excel-free PDF generation via pdf_libre.py (LibreOffice
    headless). Same return contract as generate_daily_pdfs: list of
    (branch, pdf_path) tuples — missing branches are reported by main()
    in the email exactly as with the Excel path."""
    import subprocess
    today    = report_date or datetime.now(TZ_BKK)
    date_str = today.strftime('%d-%b-%Y')
    pdf_dir  = os.path.join(SCRIPT_DIR, 'Daily PDFs',
                            today.strftime('%Y'),
                            today.strftime('%b'),
                            today.strftime('%d'))
    py = _libreoffice_python()
    if py is None:
        print("\n[PDF] neither pywin32 nor LibreOffice (UNO) available - skipping PDFs")
        return []
    cmd = [py, os.path.join(SCRIPT_DIR, 'pdf_libre.py'),
           '--xlsm', os.path.abspath(xlsm_path),
           '--date', today.strftime('%Y-%m-%d'),
           '--outdir', pdf_dir]
    print(f"\n[PDF] using LibreOffice headless ({py})")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f"[PDF] LibreOffice generator exited with code {r.returncode}")
    # Collect whatever was produced (partial output still gets emailed,
    # with the missing branches noted — same as the Excel path).
    generated = []
    for branch in DAILY_BRANCHES:
        branch_safe = _RE_BRANCH_UNSAFE.sub('-', branch).replace(' ', '_')
        p = os.path.join(pdf_dir, f'{branch_safe}_{date_str}.pdf')
        if os.path.exists(p):
            generated.append((branch, p))
    # Prepend summary PDF so it appears first in the email attachment list.
    summary_p = os.path.join(pdf_dir, f'Summary_{date_str}.pdf')
    if os.path.exists(summary_p):
        generated.insert(0, ('Summary', summary_p))
    return generated

def _wait_excel_ready(xl, timeout=60):
    """Poll xl.Ready until it reports True or the timeout elapses.

    v3.4: the v3.3 implementation returned on the first successful read of
    xl.Ready regardless of its value — effectively a one-shot probe. We
    now genuinely wait until Ready is truthy (or the timeout fires), so
    a busy Excel can't slip past.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if xl.Ready:
                return
        except Exception:
            pass
        time.sleep(0.2)

def generate_daily_pdfs(xlsm_path, report_date=None):
    """Generate 7 branch PDFs from the Daily Template sheet.

    v3.7: when report_date is provided, that date drives the PDF folder
    path, the Excel G1/G2/G3 cells, and the filename DD-MMM-YYYY. When
    None, falls back to today (legacy behavior — fully backward compatible).
    """
    try:
        import win32com.client
    except ImportError:
        # v3.13: no Excel/pywin32 (e.g. GitHub-hosted Linux runner) —
        # generate the PDFs with LibreOffice headless instead.
        # (ASCII-only print: this path can run on consoles without UTF-8.)
        print("\n[PDF] pywin32 not found - falling back to LibreOffice headless")
        return _generate_daily_pdfs_libreoffice(xlsm_path, report_date)

    today    = report_date or datetime.now(TZ_BKK)
    date_str = today.strftime('%d-%b-%Y')
    # v3.1: organize PDFs by year / month / day instead of flat folder.
    # e.g. Daily PDFs/2026/Jun/01/HAY_Store_01-Jun-2026.pdf
    pdf_dir = os.path.join(SCRIPT_DIR, 'Daily PDFs',
                           today.strftime('%Y'),
                           today.strftime('%b'),
                           today.strftime('%d'))
    os.makedirs(pdf_dir, exist_ok=True)

    print(f"\n[PDF] เปิด Excel...")
    xl = None
    wb = None
    # v3.4: snapshot of the Application-level state we are about to mutate.
    # The finally block restores these even on mid-loop exceptions, so the
    # user's Excel session never gets left with ScreenUpdating disabled or
    # Calculation stuck in manual mode.
    prev_calc          = _XL_CALC_AUTOMATIC
    prev_screen_update = True
    state_changed      = False
    try:
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.Visible = True          # ต้อง True — Visible=False ทำให้ COM calls ถูก block
        xl.DisplayAlerts = False
        xl.EnableEvents  = False

        wb = xl.Workbooks.Open(os.path.abspath(xlsm_path))
        _wait_excel_ready(xl)

        ws = wb.Sheets('Daily Template')
        ws.Unprotect()

        # v3.4: capture the previous Application state, then flip the two
        # big-win switches for the 7-branch loop. These are the dominant
        # cost in v3.3 — Excel was redrawing the screen and recomputing
        # the entire workbook on every cell write + every PageSetup write.
        try:
            prev_screen_update = xl.ScreenUpdating
        except Exception:
            prev_screen_update = True
        try:
            prev_calc = xl.Calculation
        except Exception:
            prev_calc = _XL_CALC_AUTOMATIC
        xl.ScreenUpdating = False
        xl.Calculation    = _XL_CALC_MANUAL
        state_changed     = True

        # v3.4: bulk-write G1:G3 in one Range assignment (1 COM call instead
        # of 3 separate .Value writes). Column-vector shape — each tuple is
        # one row of one cell.
        ws.Range('G1:G3').Value = ((today.day,), (today.month,), (today.year,))
        ws.Calculate()             # explicit recalc — Calculation is manual

        # v3.4: invariant PageSetup properties — set ONCE outside the loop.
        # Only PrintArea (depends on per-branch last_row) is touched inside
        # the loop. CentimetersToPoints calls also hoisted (was 7+14 round
        # trips inside the loop).
        cm05 = xl.CentimetersToPoints(0.5)
        cm00 = xl.CentimetersToPoints(0)
        ps = ws.PageSetup
        ps.PrintTitleRows = '$5:$5'   # ซ้ำ row 5 ทั้งแถว (Excel ไม่รองรับเฉพาะบางคอลัมน์)
        ps.Zoom           = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 32767     # 0 errors on some Excel versions
        # Margin น้อยที่สุด (0.5cm รอบด้าน)
        ps.LeftMargin     = cm05
        ps.RightMargin    = cm05
        ps.TopMargin      = cm05
        ps.BottomMargin   = cm05
        ps.HeaderMargin   = cm00
        ps.FooterMargin   = cm00

        rows_count = ws.Rows.Count   # COM round-trip — cache once

        generated = []
        for branch in DAILY_BRANCHES:
            ws.Range('J2').Value = branch
            ws.Calculate()           # explicit — branch filter drives formulas

            last_i   = ws.Cells(rows_count, 9).End(-4162).Row
            last_row = max(41, last_i)

            # Only the variable PageSetup property per branch.
            ps.PrintArea = f'$A$1:$V${last_row}'

            branch_safe = _RE_BRANCH_UNSAFE.sub('-', branch).replace(' ', '_')
            pdf_path    = os.path.join(pdf_dir, f'{branch_safe}_{date_str}.pdf')

            ws.ExportAsFixedFormat(0, pdf_path, 0, True, False)
            generated.append((branch, pdf_path))
            print(f"[PDF]   ✓ {branch:<15}  →  {os.path.basename(pdf_path)}  (rows 1:{last_row})")

        # v3.4: restore Calculation + ScreenUpdating BEFORE Close/Quit on
        # the happy path so the finally block becomes a no-op. (Quit
        # clears the Application instance anyway — this is defensive.)
        try:
            xl.Calculation    = prev_calc
            xl.ScreenUpdating = prev_screen_update
            state_changed     = False
        except Exception:
            pass

        wb.Close(False)
        wb = None
        xl.Quit()
        xl = None
        print(f"\n[PDF] สร้าง {len(generated)} ไฟล์ใน: {pdf_dir}")
        return generated   # v3.6: list of (branch, pdf_path) for main()→email

    except Exception as e:
        print(f"\n[PDF] ERROR: {e}")
        return generated   # v3.6: return whatever made it through
    finally:
        # v3.4: ALWAYS restore Excel's global state, even on early-exit or
        # mid-loop exceptions. If we never made it past the snapshot,
        # state_changed is False and this branch is skipped.
        if state_changed and xl is not None:
            try: xl.Calculation = prev_calc
            except: pass
            try: xl.ScreenUpdating = prev_screen_update
            except: pass
        try:
            if wb: wb.Close(False)
        except: pass
        try:
            if xl: xl.Quit()
        except: pass

if __name__ == "__main__":
    # v3.6: top-level try/except around main() so any unhandled exception
    # results in a FAILURE notification email before the script exits.
    try:
        main()
    except SystemExit:
        raise
    except Exception as fatal:
        tb = traceback.format_exc()
        print(f"\n[FATAL] Pipeline crashed:", file=sys.stderr)
        print(tb, file=sys.stderr)
        try:
            send_failure_email(str(fatal), 'pipeline', tb)
        except Exception as ee:
            print(f"[EMAIL] Could not send failure notification: {ee}", file=sys.stderr)
        sys.exit(1)
