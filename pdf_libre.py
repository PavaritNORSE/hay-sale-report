#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_libre.py - Excel-free Daily PDF generator (LibreOffice headless / UNO)
===========================================================================
v3.13: replicates daily_report.generate_daily_pdfs() without Microsoft
Excel, so the pipeline can run on cloud machines (GitHub-hosted runners)
that have LibreOffice but no Excel.

Mirrors the Excel COM behavior step-for-step:
  - open the patched xlsm, recalculate all formulas
  - G1/G2/G3 = day/month/year, J2 = branch name (per branch)
  - last_row = last non-empty cell in column I (formulas count), min 41
  - print area $A$1:$V$last_row, repeat row 5 as title on every page
  - fit to 1 page wide / unlimited tall, 0.5cm margins, no header/footer
  - export "Daily Template" sheet only (all other sheets hidden)
  - output {branch}_{DD-MMM-YYYY}.pdf into --outdir

MUST run under a Python that has the `uno` bridge:
  - Linux:   apt install python3-uno  ->  system python3
  - Windows: C:\\Program Files\\LibreOffice\\program\\python.exe

Invoked by daily_report.py automatically when pywin32 is unavailable:
  python3 pdf_libre.py --xlsm "Sales Report V3.1.xlsm" \
                       --date 2026-06-10 --outdir "Daily PDFs/2026/Jun/10"

The xlsm file is never saved - all cell writes are transient (same as the
Excel path, which closes the workbook with SaveChanges=False).
"""

import argparse
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import traceback
from datetime import datetime

_LO_TIMEOUT_SECS = 1200  # 20 minutes — kill soffice if it hangs this long

def _alarm_handler(signum, frame):
    raise TimeoutError(f'[PDF-LO] LibreOffice did not finish within '
                       f'{_LO_TIMEOUT_SECS // 60} minutes — possible hang')

try:
    import uno
    from com.sun.star.beans import PropertyValue
except ImportError:
    sys.stderr.write(
        "ERROR: this script must run under a Python with the UNO bridge.\n"
        "  Linux:   sudo apt install python3-uno   (then use system python3)\n"
        "  Windows: \"C:\\Program Files\\LibreOffice\\program\\python.exe\" pdf_libre.py ...\n")
    sys.exit(1)

# Keep in sync with daily_report.DAILY_BRANCHES
DAILY_BRANCHES = [
    'Somkid', 'NORSE Store', 'Line Chat',
    'Line My Shop', 'Lazada', 'HAY Store', 'Wholesale',
]

SHEET_NAME = 'Daily Template'

# Same filename sanitiser as daily_report.py
_RE_BRANCH_UNSAFE = re.compile(r'[\\/*?:"<>|]')


def find_soffice():
    """Locate the soffice binary (env override > PATH > common paths)."""
    cands = []
    env = os.environ.get('SOFFICE_BIN')
    if env:
        cands.append(env)
    for name in ('soffice', 'soffice.exe', 'libreoffice'):
        w = shutil.which(name)
        if w:
            cands.append(w)
    cands += [
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def _prop(name, value):
    pv = PropertyValue()
    pv.Name = name
    pv.Value = value
    return pv


def _start_soffice(soffice, profile_dir, port):
    """Launch a headless soffice listener with an isolated user profile."""
    profile_url = uno.systemPathToFileUrl(profile_dir)
    args = [
        soffice, '--headless', '--invisible', '--norestore', '--nologo',
        '--nolockcheck', '--nodefault',
        f'-env:UserInstallation={profile_url}',
        f'--accept=socket,host=127.0.0.1,port={port};urp;',
    ]
    return subprocess.Popen(args, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)


def _connect(port, timeout=90):
    """Connect to the soffice UNO socket, retrying until it is up."""
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', local_ctx)
    url = (f'uno:socket,host=127.0.0.1,port={port};urp;'
           'StarOffice.ComponentContext')
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            return resolver.resolve(url)
        except Exception as e:          # NoConnectException until LO is up
            last_err = e
            time.sleep(0.5)
    raise RuntimeError(f'Could not connect to LibreOffice: {last_err}')


# ── Detail-table injection ───────────────────────────────────────────────────
# The Daily Template detail area (I6:V6) holds Excel-365 dynamic-array
# formulas: IFS(J2="All", FILTER(...), J2<>"", FILTER(<date match> *
# (YTD!G = J2), "No Transaction")). LibreOffice mishandles their saved
# spill extents (verified 11-Jun-2026: shrink-below-saved-extent recycles
# rows; collapsed extents stop spilling entirely). Instead of fighting
# LO's spill semantics we DELETE those array formulas and inject the
# filtered YTD rows as plain values — same filter the formula encodes.
# Everything else on the sheet (SUBTOTAL sums in row 2, SUMIFS branch
# summary in column C) references plain column ranges and recalculates
# correctly in LO on top of the injected values.
#
# Template column <- YTD source column (from the FILTER formulas):
#   I<-C  J<-G  K<-H  L<-I  M<-J  N<-N  O<-O  P<-R  Q<-V  R<-X  S<-Z
#   T<-AA U<-AB V<-AC
_YTD_SRC_COLS = [2, 6, 7, 8, 9, 13, 14, 17, 21, 23, 25, 26, 27, 28]
_DETAIL_FIRST_ROW = 5      # template row 6 (0-based)
_DETAIL_COL_FIRST = 8      # template col I (0-based)
_DETAIL_COL_LAST = 21      # template col V (0-based)
_CLEAR_FLAGS = 1 | 2 | 4 | 16   # VALUE | DATETIME | STRING | FORMULA

def _read_ytd_rows(doc, date_str):
    """Read the YTD sheet once. Returns (by_branch, sums) where by_branch
    maps branch -> [14-col detail row, ...] for rows whose date (col C,
    'DD-MMM-YYYY' string) equals date_str, and sums maps branch ->
    [sum YTD!AC, sum YTD!AD] (Actual Sales Closed / Account Receivable,
    mirroring the branch-summary SUMIFS in column C of the template)."""
    ytd = doc.Sheets.getByName('YTD')
    cur = ytd.createCursor()
    cur.gotoEndOfUsedArea(False)
    end_row = cur.RangeAddress.EndRow
    if end_row < 4:
        return {}, {}
    # Cols C..AD (2..29), data rows start at sheet row 5 (0-based 4)
    data = ytd.getCellRangeByPosition(2, 4, 29, end_row).getDataArray()
    by_branch = {}
    sums = {}
    for row in data:
        if row[0] != date_str:          # col C (offset 0 in slice)
            continue
        branch = row[6 - 2]             # col G
        out = tuple(row[c - 2] for c in _YTD_SRC_COLS)
        by_branch.setdefault(branch, []).append(out)
        s = sums.setdefault(branch, [0.0, 0.0])
        for i, col in ((0, 28), (1, 29)):           # AC, AD
            v = row[col - 2]
            if isinstance(v, (int, float)):
                s[i] += v
    return by_branch, sums

def _read_cr_sums(doc, date_str):
    """Sum 'Cash Received'!M per branch label (col E) for rows whose date
    (col B, 'DD-MMM-YYYY' string) equals date_str — the Cash Received
    SUMIFS of the branch summary. Excel coerces those text dates when
    matching a date criterion; LibreOffice does not, so we compute it."""
    cr = doc.Sheets.getByName('Cash Received')
    cur = cr.createCursor()
    cur.gotoEndOfUsedArea(False)
    end_row = cur.RangeAddress.EndRow
    if end_row < 4:
        return {}
    # Cols B..M (1..12), data rows start at sheet row 5 (0-based 4)
    data = cr.getCellRangeByPosition(1, 4, 12, end_row).getDataArray()
    sums = {}
    for row in data:
        if row[0] != date_str:          # col B
            continue
        label = row[4 - 1]              # col E
        v = row[12 - 1]                 # col M
        if isinstance(v, (int, float)):
            sums[label] = sums.get(label, 0.0) + v
    return sums

# Branch-summary blocks in the left panel: label cell row -> value rows
# label+1..label+3 hold Actual Sales Closed / Cash Received / Account
# Receivable (template rows, 1-based): B3/C4-C6, B8/C9-C11, ... B33/C34-C36.
_PANEL_LABEL_ROWS = [3, 8, 13, 18, 23, 28, 33]

def _write_panel(sheet, branch, ytd_sums, cr_sums):
    """Overwrite the C-column SUMIFS results: zero everywhere except the
    block whose label equals the current branch (same IF(OR(J2=label,...))
    guard the template formulas encode). Totals C39:C41 stay as formulas."""
    for label_row in _PANEL_LABEL_ROWS:
        label = sheet.getCellByPosition(1, label_row - 1).getString()
        if label == branch:
            asc, ar = ytd_sums.get(branch, (0.0, 0.0))
            vals = (asc, cr_sums.get(branch, 0.0), ar)
        else:
            vals = (0.0, 0.0, 0.0)
        for off, v in enumerate(vals, start=1):
            sheet.getCellByPosition(2, label_row - 1 + off).setValue(v)

def _clear_detail_area(sheet, n_rows):
    """Clear the detail area I6:V{6+n_rows-1}, expanding over any array
    formula so LO does not refuse a partial-array edit."""
    for col in range(_DETAIL_COL_FIRST, _DETAIL_COL_LAST + 1):
        cell = sheet.getCellByPosition(col, _DETAIL_FIRST_ROW)
        cur = sheet.createCursorByRange(cell)
        try:
            cur.collapseToCurrentArray()
        except Exception:
            pass
        cur.clearContents(_CLEAR_FLAGS)
    if n_rows > 1:
        sheet.getCellRangeByPosition(
            _DETAIL_COL_FIRST, _DETAIL_FIRST_ROW + 1,
            _DETAIL_COL_LAST, _DETAIL_FIRST_ROW + n_rows - 1
        ).clearContents(_CLEAR_FLAGS)

def _write_detail_rows(sheet, rows):
    """Inject filtered rows (or the 'No Transaction' placeholder) as plain
    values into I6:V{...}. Returns the number of rows written."""
    if not rows:
        rows = [tuple(['No Transaction'] * 14)]
    rng = sheet.getCellRangeByPosition(
        _DETAIL_COL_FIRST, _DETAIL_FIRST_ROW,
        _DETAIL_COL_LAST, _DETAIL_FIRST_ROW + len(rows) - 1)
    rng.setDataArray(tuple(rows))
    return len(rows)


def _last_used_row_col_I(sheet):
    """1-based last row with any content in column I (formulas count,
    mirroring Excel End(xlUp) semantics). 0 if the column is empty."""
    cur = sheet.createCursor()
    cur.gotoEndOfUsedArea(False)
    end_row = cur.RangeAddress.EndRow            # 0-based
    for r in range(end_row, -1, -1):
        if sheet.getCellByPosition(8, r).getType().value != 'EMPTY':
            return r + 1
    return 0


def _generate_summary_pdf(desktop, by_branch, ytd_sums, cr_sums, report_date, outdir):
    """Create a one-page summary PDF: all branches × (Orders, Sales Closed,
    Cash Received, AR) in a single table.  Saved as Summary_DD-MMM-YYYY.pdf
    in outdir.  Returns the path on success, None on failure."""
    date_str = report_date.strftime('%d-%b-%Y')
    sdoc = None
    try:
        sdoc = desktop.loadComponentFromURL(
            'private:factory/scalc', '_blank', 0, ())
        sheet = sdoc.Sheets.getByIndex(0)

        # ── number format #,##0 ───────────────────────────────────────────
        locale = uno.createUnoStruct('com.sun.star.lang.Locale')
        fmts   = sdoc.getNumberFormats()
        num_key = fmts.queryKey('#,##0', locale, False)
        if num_key == -1:
            num_key = fmts.addNew('#,##0', locale)

        # ── helpers ───────────────────────────────────────────────────────
        def _c(r, c):
            return sheet.getCellByPosition(c, r)

        def _style(r, c, bold=False, bg=None, fg=None, height=None):
            cell = _c(r, c)
            if bold:   cell.CharWeight = 150      # BOLD
            if bg is not None: cell.CellBackColor = bg
            if fg is not None: cell.CharColor = fg
            if height:  cell.CharHeight = height

        # ── title (row 0) ─────────────────────────────────────────────────
        _c(0, 0).setString(f'Daily Sales Summary  —  {date_str}')
        _style(0, 0, bold=True, height=13)

        # ── column headers (row 2) ────────────────────────────────────────
        BLUE, WHITE = 0x2F5496, 0xFFFFFF
        GRAY        = 0xD9E1F2   # alternating row tint
        headers = ['Branch', 'Orders', 'Sales Closed', 'Cash Received', 'AR']
        for col, h in enumerate(headers):
            _c(2, col).setString(h)
            _style(2, col, bold=True, bg=BLUE, fg=WHITE)

        # ── column widths (1/100 mm) ──────────────────────────────────────
        for col, w in enumerate([4500, 2200, 3800, 3800, 3200]):
            sheet.Columns.getByIndex(col).Width = w

        # ── data rows (rows 3-9) ──────────────────────────────────────────
        total_orders = 0
        total_asc = total_cr = total_ar = 0.0

        for i, branch in enumerate(DAILY_BRANCHES):
            r      = 3 + i
            orders = len(by_branch.get(branch, []))
            asc, ar = ytd_sums.get(branch, (0.0, 0.0))
            cr     = cr_sums.get(branch, 0.0)

            _c(r, 0).setString(branch)
            for col, v in enumerate([orders, asc, cr, ar], 1):
                cell = _c(r, col)
                cell.setValue(v)
                cell.NumberFormat = num_key

            if i % 2 == 1:                        # alternating tint
                for col in range(5):
                    _c(r, col).CellBackColor = GRAY

            total_orders += orders
            total_asc    += asc
            total_cr     += cr
            total_ar     += ar

        # ── total row ─────────────────────────────────────────────────────
        tr = 3 + len(DAILY_BRANCHES)
        _c(tr, 0).setString('TOTAL')
        for col, v in enumerate([total_orders, total_asc, total_cr, total_ar], 1):
            cell = _c(tr, col)
            cell.setValue(v)
            cell.NumberFormat = num_key
        for col in range(5):
            _style(tr, col, bold=True, bg=BLUE, fg=WHITE)

        # ── page setup: fit to 1 page, portrait ──────────────────────────
        ps = sdoc.StyleFamilies.getByName('PageStyles').getByName(sheet.PageStyle)
        ps.ScaleToPagesX = 1
        ps.ScaleToPagesY = 1
        ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 1000
        ps.HeaderIsOn = False
        ps.FooterIsOn = False

        area = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        area.Sheet      = 0
        area.StartColumn, area.EndColumn = 0, 4
        area.StartRow,    area.EndRow    = 0, tr
        sheet.setPrintAreas((area,))

        # ── export ────────────────────────────────────────────────────────
        pdf_path = os.path.join(outdir, f'Summary_{date_str}.pdf')
        sdoc.storeToURL(
            uno.systemPathToFileUrl(os.path.abspath(pdf_path)),
            (_prop('FilterName', 'calc_pdf_Export'), _prop('Overwrite', True)))
        print(f'[PDF-LO]   OK Summary          -> Summary_{date_str}.pdf')
        return pdf_path

    except Exception as e:
        print(f'[PDF-LO]   WARN summary PDF failed: {e}')
        traceback.print_exc()
        return None
    finally:
        if sdoc is not None:
            try:
                sdoc.close(False)
            except Exception:
                pass


def generate(xlsm_path, report_date, outdir):
    soffice = find_soffice()
    if not soffice:
        print('[PDF-LO] ERROR: soffice binary not found '
              '(install LibreOffice or set SOFFICE_BIN)')
        return 1

    os.makedirs(outdir, exist_ok=True)
    date_str = report_date.strftime('%d-%b-%Y')
    port = 2002 + (os.getpid() % 500)
    profile_dir = tempfile.mkdtemp(prefix='lo_profile_')
    proc = None
    doc = None
    exit_code = 0
    print(f'[PDF-LO] soffice: {soffice}')
    if hasattr(signal, 'SIGALRM'):
        signal.signal(signal.SIGALRM, _alarm_handler)
        signal.alarm(_LO_TIMEOUT_SECS)
    try:
        proc = _start_soffice(soffice, profile_dir, port)
        ctx = _connect(port)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext(
            'com.sun.star.frame.Desktop', ctx)

        load_props = (
            _prop('Hidden', True),
            _prop('MacroExecutionMode', 0),    # never run VBA macros
            _prop('UpdateDocMode', 0),         # no link updates
        )
        url = uno.systemPathToFileUrl(os.path.abspath(xlsm_path))
        print(f'[PDF-LO] opening {os.path.basename(xlsm_path)} ...')
        doc = desktop.loadComponentFromURL(url, '_blank', 0, load_props)
        if doc is None:
            print('[PDF-LO] ERROR: failed to open the workbook')
            return 1

        sheets = doc.Sheets
        sheet = sheets.getByName(SHEET_NAME)
        try:
            if sheet.isProtected():
                sheet.unprotect('')
        except Exception:
            pass

        # Export must contain ONLY the Daily Template sheet -> hide the rest.
        # (calc_pdf_Export skips hidden sheets; nothing is ever saved.)
        for name in sheets.ElementNames:
            if name != SHEET_NAME:
                sheets.getByName(name).IsVisible = False

        # Date cells G1/G2/G3 (day / month / year), same as the Excel path.
        sheet.getCellByPosition(6, 0).setValue(report_date.day)
        sheet.getCellByPosition(6, 1).setValue(report_date.month)
        sheet.getCellByPosition(6, 2).setValue(report_date.year)
        doc.calculateAll()

        # Column C (index 2) holds branch-summary figures. LO may render the
        # xlsm column width slightly narrower than Excel, causing "###" for
        # large values. Enforce a minimum of 3 cm to prevent this.
        col_c = sheet.Columns.getByIndex(2)
        if col_c.Width < 3000:
            col_c.Width = 3000

        # Replace the dynamic-array detail table with injected values
        # (see the note above _YTD_SRC_COLS for why).
        by_branch, ytd_sums = _read_ytd_rows(doc, date_str)
        cr_sums = _read_cr_sums(doc, date_str)
        _clear_detail_area(sheet, 500)     # wipe formulas + any stale cache
        prev_rows = 1

        # Page style = Excel's PageSetup: fit 1 page wide / unlimited tall,
        # 0.5cm margins, no header/footer. Orientation + paper size come
        # from the template file itself (Excel didn't set them either).
        style = doc.StyleFamilies.getByName('PageStyles') \
                                 .getByName(sheet.PageStyle)
        style.ScaleToPagesX = 1
        style.ScaleToPagesY = 0          # 0 = as many pages tall as needed
        style.LeftMargin = style.RightMargin = 500      # 1/100 mm = 0.5cm
        style.TopMargin = style.BottomMargin = 500
        style.HeaderIsOn = False
        style.FooterIsOn = False

        sheet_idx = sheet.RangeAddress.Sheet
        title_rows = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        title_rows.Sheet = sheet_idx
        title_rows.StartColumn, title_rows.EndColumn = 0, 21
        title_rows.StartRow, title_rows.EndRow = 4, 4    # row 5 (0-based 4)

        generated = 0
        for branch in DAILY_BRANCHES:
            sheet.getCellByPosition(9, 1).setString(branch)   # J2
            if prev_rows > 1:
                _clear_detail_area(sheet, prev_rows)
            prev_rows = _write_detail_rows(sheet, by_branch.get(branch, []))
            _write_panel(sheet, branch, ytd_sums, cr_sums)
            doc.calculateAll()

            # Auto-fit all columns A–V so no cell shows ### regardless of value.
            for _ci in range(22):
                sheet.Columns.getByIndex(_ci).OptimalWidth = True

            last_row = max(41, _last_used_row_col_I(sheet))

            area = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
            area.Sheet = sheet_idx
            area.StartColumn, area.EndColumn = 0, 21          # A..V
            area.StartRow, area.EndRow = 0, last_row - 1
            sheet.setPrintAreas((area,))
            # setPrintAreas can reset title settings -> reapply every loop
            sheet.setTitleRows(title_rows)
            sheet.setPrintTitleRows(True)

            branch_safe = _RE_BRANCH_UNSAFE.sub('-', branch).replace(' ', '_')
            pdf_path = os.path.join(outdir, f'{branch_safe}_{date_str}.pdf')
            export_props = (
                _prop('FilterName', 'calc_pdf_Export'),
                _prop('Overwrite', True),
            )
            doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(pdf_path)),
                           export_props)
            generated += 1
            print(f'[PDF-LO]   OK {branch:<15} -> '
                  f'{os.path.basename(pdf_path)}  (rows 1:{last_row})')

        print(f'[PDF-LO] generated {generated}/{len(DAILY_BRANCHES)} '
              f'files in: {outdir}')
        if generated < len(DAILY_BRANCHES):
            exit_code = 1

        _generate_summary_pdf(desktop, by_branch, ytd_sums, cr_sums,
                               report_date, outdir)
        return exit_code

    except TimeoutError as e:
        print(e)
        return 1
    except Exception as e:
        print(f'[PDF-LO] ERROR: {e}')
        traceback.print_exc()
        return 1
    finally:
        if hasattr(signal, 'SIGALRM'):
            signal.alarm(0)
        try:
            if doc is not None:
                doc.close(False)
        except Exception:
            pass
        if proc is not None:
            try:
                proc.terminate()
                proc.wait(timeout=20)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        shutil.rmtree(profile_dir, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(
        description='Generate the 7 branch Daily PDFs via LibreOffice')
    ap.add_argument('--xlsm', required=True, help='path to the patched xlsm')
    ap.add_argument('--date', required=True, help='report date YYYY-MM-DD')
    ap.add_argument('--outdir', required=True, help='PDF output folder')
    args = ap.parse_args()

    try:
        report_date = datetime.strptime(args.date, '%Y-%m-%d')
    except ValueError:
        ap.error(f'--date must be YYYY-MM-DD (got {args.date!r})')

    if not os.path.exists(args.xlsm):
        print(f'[PDF-LO] ERROR: xlsm not found: {args.xlsm}')
        sys.exit(1)

    sys.exit(generate(args.xlsm, report_date, args.outdir))


if __name__ == '__main__':
    main()
