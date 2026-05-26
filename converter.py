#!/usr/bin/env python3
"""
PDF to Excel Converter — BCA Rekening Koran
Mengekstrak data dari PDF mutasi rekening ke satu sheet Excel dengan kolom:
  NO. REKENING | PERIODE | MATA UANG | TANGGAL | KETERANGAN | CBG | MUTASI | SALDO
"""

import os, sys, re, argparse, logging, datetime
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)-8s] %(funcName)s — %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# Untuk menonaktifkan log DEBUG dari library pihak ketiga (camelot, pdfminer, dll.)
for _noisy in ("camelot", "pdfminer", "PIL", "matplotlib", "ghostscript"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)

# ── dependency guard ─────────────────────────────────────────────────────────
try:
    import camelot
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "camelot-py[cv]", "openpyxl", "pandas"])
    import camelot, pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ── style constants ──────────────────────────────────────────────────────────
C_DARK   = "0D2B45"
C_BLUE   = "1F4E79"
C_LBLUE  = "D6E4F0"
C_WHITE  = "FFFFFF"

def _fill(hex_color): return PatternFill("solid", fgColor=hex_color)
def _font(bold=False, color=C_WHITE, size=10):
    return Font(bold=bold, color=color, size=size)
def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)
def _align(h="left", wrap=True):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)


# ── helpers ──────────────────────────────────────────────────────────────────
def clean(s: str) -> str:
    """Collapse whitespace in a cell value."""
    return re.sub(r"\s+", " ", str(s)).strip()


def parse_amount(val_str: str) -> float:
    """Parse number string that might be in US (1,234.56) or Indonesian (1.234,56) format."""
    val_str = re.sub(r"[^\d,.-]", "", val_str).strip()
    if not val_str:
        return 0.0

    last_comma = val_str.rfind(',')
    last_dot = val_str.rfind('.')

    if last_comma != -1 and last_dot != -1:
        if last_dot > last_comma:
            # US format: 1,234.56
            val_str = val_str.replace(',', '')
        else:
            # Indonesian format: 1.234,56
            val_str = val_str.replace('.', '').replace(',', '.')
    elif last_comma != -1:
        # Only comma is present
        if last_comma == len(val_str) - 3:
            # 100,00 -> decimal
            val_str = val_str.replace(',', '.')
        else:
            # 1,000 -> thousands
            val_str = val_str.replace(',', '')
    elif last_dot != -1:
        # Only dot is present
        if last_dot == len(val_str) - 3:
            # 100.00 -> decimal
            pass
        else:
            # 1.000 -> thousands
            val_str = val_str.replace('.', '')

    try:
        return float(val_str)
    except ValueError:
        return 0.0


def extract_meta(df: pd.DataFrame) -> dict:
    """
    Parse metadata from Table 1 (header info block).
    Looks for patterns like  'NO. RE KE NIN G  :  862077 65 90'
    """
    log.debug("extract_meta: shape tabel header = %s", df.shape)
    text = " ".join(clean(v) for row in df.values for v in row)
    # normalise extra spaces that camelot introduces inside words
    text = re.sub(r"\s+", " ", text)
    log.debug("extract_meta: raw text (200 char pertama) = %r", text[:200])

    def find_raw(patterns):
        """Return the raw matched group (with internal spaces intact)."""
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                log.debug("find_raw: pattern %r → match %r", pat, m.group(1).strip())
                return m.group(1).strip()
            log.debug("find_raw: pattern %r → tidak cocok", pat)
        return ""

    def find_compact(patterns):
        """Return matched group with ALL spaces removed (for numbers)."""
        raw = find_raw(patterns)
        result = re.sub(r"\s+", "", raw)
        log.debug("find_compact: raw=%r → compact=%r", raw, result)
        return result

    rekening = find_compact([
        r"NO\.\s*RE\s*KE\s*NIN\s*G\s*:\s*([\d\s]+?)(?=[A-Z]|$)",
        r"REKENING\s*:\s*([\d\s]+?)(?=[A-Z]|$)",
    ])
    log.debug("extract_meta: NO. REKENING = %r", rekening)

    # PERIODE value e.g. "APRIL 2 026" → remove spaces in digits only → "APRIL 2026"
    periode_raw = find_raw([
        r"PE\s*RIOD\s*E\s*:\s*([A-Z0-9][A-Z0-9\s]+?)(?=MATA|GUNUNGANY|KOTA|INDONESIA|$)",
        r"PERIODE\s*:\s*([A-Z0-9 ]+?)(?=MATA|$)",
    ])
    log.debug("extract_meta: periode_raw = %r", periode_raw)
    # Collapse spaces: keep word boundaries but squash spaces inside digit runs
    # e.g. "APRIL 2 026" → join → "APRIL 2026"
    periode = periode_raw.strip()
    # Remove spaces that are sandwiched between digits
    periode = re.sub(r"(?<=\d)\s+(?=\d)", "", periode)
    # Collapse multiple spaces elsewhere to single space
    periode = re.sub(r"\s+", " ", periode).strip()
    log.debug("extract_meta: PERIODE (setelah normalisasi) = %r", periode)

    # MATA UANG e.g. "ID R" → remove spaces → "IDR"
    mata_uang_raw = find_raw([
        r"MATA\s*U\s*ANG\s*:\s*([A-Z][A-Z\s]{1,5})(?=\s|$)",
        r"MATA\s*UANG\s*:\s*([A-Z]+)",
    ])
    mata_uang = re.sub(r"\s+", "", mata_uang_raw)
    log.debug("extract_meta: MATA UANG = %r (raw: %r)", mata_uang, mata_uang_raw)

    result = {
        "NO. REKENING": rekening,
        "PERIODE":      periode,
        "MATA UANG":    mata_uang,
    }
    log.info("Metadata diekstrak: %s", result)
    return result


def parse_transactions(df: pd.DataFrame) -> list[dict]:
    """
    Parse Table 2 into clean transaction rows.

    Camelot splits multi-line cells across separate DataFrame rows.
    We group them: a new transaction starts when column-0 has a date (DD/MM).
    """
    log.debug("parse_transactions: shape tabel = %s, kolom = %d", df.shape, df.shape[1])
    DATE_RE = re.compile(r"^\d{2}/\d{2}$")

    rows   = df.values.tolist()
    groups = []   # list of raw row-groups per transaction
    cur    = None

    for i, row in enumerate(rows):
        date_val = clean(row[0])
        if DATE_RE.match(date_val):
            if cur is not None:
                groups.append(cur)
            log.debug("  baris %d: tanggal baru ditemukan → %r", i, date_val)
            cur = [row]
        elif cur is not None:
            log.debug("  baris %d: continuation, ditambah ke grup saat ini", i)
            cur.append(row)
        else:
            log.debug("  baris %d: dilewati (sebelum transaksi pertama): %r", i, [clean(v) for v in row[:3]])

    if cur is not None:
        groups.append(cur)

    log.debug("parse_transactions: total grup transaksi = %d", len(groups))

    transactions = []
    for gi, grp in enumerate(groups):
        # col indices (from raw table): 0=TANGGAL, 1=KETERANGAN, 2=KETERANGAN(cont),
        #                               3=CBG, 4=MUTASI, 5=SALDO(extra), 6=SALDO
        tanggal    = clean(grp[0][0])
        # merge all keterangan lines (cols 1+2 of all rows in group)
        ket_parts  = []
        for r in grp:
            for c in [1, 2]:
                v = clean(r[c]) if c < len(r) else ""
                if v:
                    ket_parts.append(v)
        keterangan = " ".join(ket_parts)

        cbg    = clean(grp[0][3]) if len(grp[0]) > 3 else ""
        mutasi = clean(grp[0][4]) if len(grp[0]) > 4 else ""
        saldo  = ""
        saldo_src = ""
        # saldo is in col 6 if present, else col 5
        for c in [6, 5]:
            v = clean(grp[0][c]) if c < len(grp[0]) else ""
            if v and re.search(r"[\d,.]", v):
                saldo = v
                saldo_src = f"col[{c}]"
                break

        # strip trailing " DB" from mutasi and tag as debit in CBG if needed
        db_tagged = False
        if mutasi.endswith(" DB"):
            mutasi = mutasi[:-3].strip()
            db_tagged = True

        # Parse mutasi to float
        if mutasi:
            mutasi = parse_amount(mutasi)

        log.debug(
            "  grup[%d] tanggal=%r ket=%r mutasi=%r%s",
            gi, tanggal, keterangan[:40], mutasi,
            " [DB tagged]" if db_tagged else "",
        )

        transactions.append({
            "TANGGAL":    tanggal,
            "KETERANGAN": keterangan,
            "CBG":        cbg,
            "MUTASI":     mutasi,
            "SALDO":      saldo,
        })

    log.info("parse_transactions selesai: %d transaksi dari tabel shape=%s", len(transactions), df.shape)
    return transactions


# ── Excel writer ─────────────────────────────────────────────────────────────
COLUMNS = ["TANGGAL", "KETERANGAN", "MUTASI"]

NUMERIC_COLS = {"MUTASI"}   # right-align these

def write_excel(records: list[dict], output_path: str, pdf_name: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Mutasi Rekening"

    # ── title row ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    tc = ws.cell(row=1, column=1, value=f"MUTASI REKENING — {pdf_name.upper()}")
    tc.font      = _font(bold=True, size=13)
    tc.fill      = _fill(C_DARK)
    tc.alignment = _align(h="center")
    ws.row_dimensions[1].height = 28

    # ── header row ──
    for ci, col in enumerate(COLUMNS, start=1):
        c = ws.cell(row=2, column=ci, value=col)
        c.font      = _font(bold=True)
        c.fill      = _fill(C_BLUE)
        c.alignment = _align(h="center", wrap=True)
        c.border    = _border()
    ws.row_dimensions[2].height = 24

    # ── data rows ──
    for ri, rec in enumerate(records, start=3):
        # ── Pengecekan Highlight Kuning ──
        ket_upper = str(rec.get("KETERANGAN", "")).upper()
        if any(w in ket_upper for w in ["DOMPET", "VISIONET", "AIRPAY"]):
            fill = _fill("FFFF00") # Kuning terang
        else:
            fill = _fill(C_LBLUE) if ri % 2 == 0 else _fill(C_WHITE)

        for ci, col in enumerate(COLUMNS, start=1):
            val  = rec.get(col, "")
            cell = ws.cell(row=ri, column=ci, value=val)

            # ── Format TANGGAL jadi Date DD/MM ──
            if col == "TANGGAL" and isinstance(val, str) and "/" in val:
                try:
                    d, m = val.split("/")
                    periode_str = str(rec.get("PERIODE", ""))
                    year_match = re.search(r'\d{4}', periode_str)
                    year = int(year_match.group()) if year_match else datetime.date.today().year
                    dt = datetime.datetime(year, int(m), int(d))
                    cell.value = dt
                    cell.number_format = "DD/MM"
                except Exception:
                    pass

            # ── Format MUTASI jadi Number ──
            if col == "MUTASI" and isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"

            cell.fill      = fill
            cell.border    = _border()
            cell.font      = _font(color="000000", size=10)
            
            h_align = "right" if col in NUMERIC_COLS else "left"
            if col == "TANGGAL":
                h_align = "center"
            cell.alignment = _align(h=h_align)

    # ── freeze & column widths ──
    ws.freeze_panes = "A3"

    col_widths = {
        "TANGGAL":      12,
        "KETERANGAN":   60,
        "MUTASI":       20,
    }
    for ci, col in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(col, 14)

    wb.save(output_path)
    log.info("✅  Saved: %s", os.path.abspath(output_path))


# ── main ─────────────────────────────────────────────────────────────────────
def convert(pdf_path: str, output_path: str):
    if not os.path.isfile(pdf_path):
        log.error("File tidak ditemukan: %s", pdf_path)
        sys.exit(1)

    file_size = os.path.getsize(pdf_path)
    log.info("=" * 60)
    log.info("Mulai konversi: %s  (%.1f KB)", pdf_path, file_size / 1024)
    log.info("Output target : %s", os.path.abspath(output_path))
    log.info("=" * 60)

    # ── Step 1: baca PDF dengan camelot ──────────────────────────────
    log.info("[1/4] Membaca tabel dari PDF (flavor=stream) …")
    tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
    log.debug("      camelot stream → %d tabel ditemukan", len(tables))

    if len(tables) == 0:
        log.warning("Stream tidak menemukan tabel → coba lattice …")
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
        log.debug("      camelot lattice → %d tabel ditemukan", len(tables))

    if len(tables) < 2:
        log.error("Dibutuhkan minimal 2 tabel (header + transaksi). Ditemukan: %d", len(tables))
        sys.exit(1)

    # Log ringkasan setiap tabel
    for i, t in enumerate(tables):
        log.debug(
            "  tabel[%d]: halaman=%s, shape=%s, akurasi_parsing=%.1f%%",
            i, t.page, t.df.shape, t.parsing_report.get("accuracy", 0),
        )
    log.info("      Ditemukan %d tabel.", len(tables))

    # ── Step 2: ekstrak metadata ──────────────────────────────────────
    log.info("[2/4] Mengekstrak metadata dari tabel-0 …")
    meta = extract_meta(tables[0].df)
    # Periksa apakah ada field kosong
    empty_fields = [k for k, v in meta.items() if not v]
    if empty_fields:
        log.warning("Field metadata kosong (regex tidak cocok): %s", empty_fields)
    else:
        log.info("      Metadata OK: %s", meta)

    # ── Step 3: parse transaksi ───────────────────────────────────────
    log.info("[3/4] Mem-parse transaksi dari %d tabel (halaman 2+) …", len(tables) - 1)
    all_trx = []
    for idx, tbl in enumerate(tables[1:], start=1):
        log.debug("  memproses tabel[%d] (halaman %s) …", idx, tbl.page)
        trx = parse_transactions(tbl.df)
        log.debug("  → %d transaksi dari tabel[%d]", len(trx), idx)
        all_trx.extend(trx)

    log.info("      Total transaksi: %d baris", len(all_trx))

    if len(all_trx) == 0:
        log.warning("Tidak ada transaksi yang berhasil diekstrak! Periksa format PDF.")

    # ── Step 4: tulis Excel ───────────────────────────────────────────
    log.info("[4/4] Menulis ke Excel …")
    records = []
    for trx in all_trx:
        row = {**meta, **trx}   # meta first, trx overrides nothing (different keys)
        records.append(row)

    write_excel(records, output_path, Path(pdf_path).stem)
    log.info("=" * 60)
    log.info("Selesai! Total %d baris ditulis ke: %s", len(records), os.path.abspath(output_path))
    log.info("=" * 60)


def main():
    p = argparse.ArgumentParser(
        description="Konversi PDF mutasi rekening → Excel (1 sheet)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("pdf", help="File PDF input")
    p.add_argument("-o", "--output", default=None,
                   help="File .xlsx output (default: ./Hasil/<nama_pdf>.xlsx)")
    p.add_argument("--output-dir", default="Hasil",
                   help="Folder output (default: Hasil)")
    args = p.parse_args()

    if args.output:
        out = args.output
    else:
        out_dir = Path(args.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        out = str(out_dir / Path(args.pdf).with_suffix(".xlsx").name)

    convert(args.pdf, out)


if __name__ == "__main__":
    main()
