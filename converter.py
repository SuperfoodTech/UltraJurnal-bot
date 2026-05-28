#!/usr/bin/env python3
"""
PDF to Excel Converter — BCA Rekening Koran
Mengekstrak data dari PDF mutasi rekening ke satu sheet Excel dengan kolom:
  NO. REKENING | PERIODE | MATA UANG | TANGGAL | KETERANGAN | CBG | MUTASI | SALDO
"""

import os, sys, re, argparse, logging, datetime, time
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
    """
    log.debug("extract_meta: shape tabel header = %s", df.shape)

    rekening = ""
    periode = ""
    mata_uang = "IDR"

    for row in df.values:
        row_cells = [clean(cell) for cell in row if cell]
        row_str = " ".join(row_cells).upper()

        if "REKENING" in row_str or "RE KE NIN G" in row_str:
            digits = re.findall(r"\d+", row_str)
            if digits:
                rekening = "".join(digits)

        if "PERIODE" in row_str or "PE RIOD E" in row_str:
            m = re.search(r"(?:PERIODE|PE\s*RIOD\s*E)\s*:\s*(.+)", row_str)
            if m:
                periode = m.group(1).strip()
                # Remove spaces between digits: e.g. "2 026" -> "2026"
                periode = re.sub(r"(?<=\d)\s+(?=\d)", "", periode)
                # Collapse extra whitespace
                periode = re.sub(r"\s+", " ", periode).strip()
                # Normalize common month names spaces from Camelot
                months_map = {
                    "JANU ARI": "JANUARI",
                    "FEBRU ARI": "FEBRUARI",
                    "AGUS TUS": "AGUSTUS",
                    "SEP TEMBER": "SEPTEMBER",
                    "OK TOBER": "OKTOBER",
                    "NO VEMBER": "NOVEMBER",
                    "DE SEMBER": "DESEMBER"
                }
                for k, v in months_map.items():
                    periode = periode.replace(k, v)

        if re.search(r"MATA\s*U\s*ANG", row_str):
            m = re.search(r"MATA\s*U\s*ANG\s*:\s*(.+)", row_str)
            if m:
                mata_uang = re.sub(r"\s+", "", m.group(1)).strip()

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
        # Stop grouping if we hit a summary row anywhere in the table
        row_str = " ".join([clean(v) for v in row]).upper()
        if any(keyword in row_str for keyword in ["SALDO AWAL", "MUTASI CR", "MUTASI DB", "SALDO AKHIR"]):
            log.debug("  baris %d: keyword summary terdeteksi -> skip & stop grouping", i)
            if cur is not None:
                groups.append(cur)
            cur = None
            continue

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
        # col indices (from explicit columns): 0=TANGGAL, 1=KETERANGAN, 2=CBG, 3=MUTASI, 4=SALDO
        tanggal    = clean(grp[0][0])
        # merge all keterangan lines (col 1 of all rows in group)
        ket_parts  = []
        for r in grp:
            v = clean(r[1]) if len(r) > 1 else ""
            if v:
                ket_parts.append(v)
        keterangan = " ".join(ket_parts)

        cbg    = clean(grp[0][2]) if len(grp[0]) > 2 else ""
        mutasi = clean(grp[0][3]) if len(grp[0]) > 3 else ""
        saldo  = clean(grp[0][4]) if len(grp[0]) > 4 else ""

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

    t0 = time.perf_counter()
    file_size = os.path.getsize(pdf_path)
    log.info("=" * 60)
    log.info("Mulai konversi: %s  (%.1f KB)", pdf_path, file_size / 1024)
    log.info("Output target : %s", os.path.abspath(output_path))
    log.info("=" * 60)

    # ── Step 1: baca metadata dari halaman 1 (flavor=stream) ─────────────────────────
    log.info("[1/4] Membaca metadata dari halaman 1 (flavor=stream) …")
    meta_tables = camelot.read_pdf(pdf_path, pages="1", flavor="stream")
    if len(meta_tables) == 0:
        log.error("Tidak dapat mengekstrak metadata dari halaman 1.")
        sys.exit(1)
    
    # ── Step 2: ekstrak metadata ──────────────────────────────────────
    log.info("[2/4] Mengekstrak metadata dari tabel-0 …")
    meta = extract_meta(meta_tables[0].df)
    empty_fields = [k for k, v in meta.items() if not v]
    if empty_fields:
        log.warning("Field metadata kosong (regex tidak cocok): %s", empty_fields)
    else:
        log.info("      Metadata OK: %s", meta)

    # ── Step 3: baca semua tabel transaksi dengan table_areas dan columns ────────────
    log.info("[3/4] Membaca tabel transaksi dari semua halaman (flavor=stream) …")
    tables = camelot.read_pdf(
        pdf_path,
        pages="all",
        flavor="stream",
        table_areas=["0,650,595,30"],
        columns=["80,290,350,480"]
    )
    log.info("      Ditemukan %d tabel transaksi.", len(tables))

    all_trx = []
    for idx, tbl in enumerate(tables):
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

    elapsed = time.perf_counter() - t0
    elapsed_min = int(elapsed // 60)
    elapsed_sec = elapsed % 60
    if elapsed_min > 0:
        elapsed_str = f"{elapsed_min}m {elapsed_sec:.1f}s"
    else:
        elapsed_str = f"{elapsed_sec:.1f}s"

    log.info("=" * 60)
    log.info("Selesai! Total %d baris ditulis ke: %s", len(records), os.path.abspath(output_path))
    log.info("⏱  Durasi konversi: %s", elapsed_str)
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
