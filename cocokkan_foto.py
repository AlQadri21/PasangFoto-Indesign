#!/usr/bin/env python3
"""Match names in a CSV (students / teachers) to photo files and write a
Data Merge-ready CSV with an image column (default header: @Foto).

Core logic lives in proses(); this file is also a small command-line tool:
    python cocokkan_foto.py data.csv "D:\\Foto\\XII-IPA-1"
    python cocokkan_foto.py            (interactive: asks for the paths)
The Windows app (aplikasi_foto.py) uses the same proses() function.
"""
import argparse
import csv
import datetime
import difflib
import functools
import io
import re
import sys
import unicodedata
from pathlib import Path

csv.field_size_limit(2 ** 31 - 1)   # default 128 KB per cell made long text cells fail to load
MAX_FILE_MB = 80                    # bigger data files could eat all memory; refuse with a clear message
MAX_NAME_WORDS = 15                 # a "name" longer than this is not a name; keeps matching fast

PHOTO_EXT = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".psd", ".webp"}
OK_SCORE = 0.85       # >= this and unambiguous -> filled automatically
CHECK_SCORE = 0.60    # >= this but doubtful   -> flagged "CEK"
TIE_MARGIN = 0.03     # runner-up this close to the best -> ambiguous
PREFILTER = 0.4       # min share of common letter pairs before a full comparison is attempted
TOPK = 15             # per name, at most this many (x2 rankings) photos get a full comparison

# Academic / honorific titles that often appear in one source but not the other
TITLES = {
    "spd", "mpd", "spdi", "mpdi", "spdsd", "drs", "dra", "dr", "ir", "prof",
    "skom", "mkom", "ssi", "msi", "se", "sh", "mh", "st", "mt", "mm", "sag",
    "shi", "mag", "ssos", "sst", "amd", "skm", "skep", "ns", "kes", "gr",
    "hj", "h", "hi",
}

# Common abbreviations -> full form ("Muh. Ali" == "Muhammad Ali")
ALIASES = {"muh": "muhammad", "mhd": "muhammad", "muhd": "muhammad",
           "moh": "muhammad", "mohd": "muhammad", "mohammad": "muhammad"}


class DataError(Exception):
    """User-facing problem with the input files (message is in Indonesian)."""


# ---------- reading ----------

def read_text(path):
    raw = Path(path).read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


EXCEL_EXT = {".xlsx", ".xlsm"}


def find_header_index(rows):
    """Index of the header row: the first of the top rows with 2+ filled cells, one of them
    containing 'nama'. Falls back to the first row. Rows above it (titles) are skipped."""
    for i, r in enumerate(rows[:15]):
        if sum(1 for c in r if str(c).strip()) >= 2 and any("nama" in str(c).lower() for c in r):
            return i
    return 0


def is_excel(path):
    return Path(path).suffix.lower() in EXCEL_EXT


def _check_size(path):
    mb = Path(path).stat().st_size / 1048576
    if mb > MAX_FILE_MB:
        raise DataError(f"File terlalu besar ({mb:.0f} MB). Batas {MAX_FILE_MB} MB. "
                        "Pecah per kelas / per sheet dulu.")


def _open_workbook(path):
    _check_size(path)
    import openpyxl  # imported lazily: only needed for Excel files
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        raise DataError("File Excel tidak bisa dibuka. Tutup dulu file itu di Excel, lalu coba lagi.")
    except Exception as e:
        raise DataError(f"File Excel tidak bisa dibaca: {e}")


# Reading an .xlsx is slow (about 1 s), and the app needs the same sheet several times
# (sheet list, columns, processing), so parsed results are kept until the file changes.
_CACHE_MAX = 8
_names_cache, _table_cache = {}, {}


def _file_key(path):
    p = Path(path)
    st = p.stat()
    return (str(p.resolve()), st.st_mtime_ns, st.st_size)


def list_sheets(path):
    """Names of the visible sheets of an .xlsx file."""
    key = _file_key(path)
    if key not in _names_cache:
        if len(_names_cache) >= _CACHE_MAX:
            _names_cache.clear()
        wb = _open_workbook(path)
        try:
            _names_cache[key] = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        finally:
            wb.close()
    return list(_names_cache[key])


def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return (v.date() if isinstance(v, datetime.datetime) else v).isoformat()  # unambiguous
    if isinstance(v, float) and v.is_integer():
        return str(int(v))                                                        # 1.0 -> 1
    return str(v).strip()


def read_excel_table(path, sheet=None):
    names = list_sheets(path)
    if not names:
        raise DataError("File Excel tidak punya sheet yang terlihat.")
    sheet = sheet or names[0]
    key = _file_key(path) + (sheet,)
    if key not in _table_cache:
        if len(_table_cache) >= _CACHE_MAX:
            _table_cache.clear()
        _table_cache[key] = _read_excel_table_uncached(path, sheet)
    header, rows, nums = _table_cache[key]
    return list(header), [list(r) for r in rows], list(nums)      # copies: callers may modify


def _read_excel_table_uncached(path, sheet=None):
    wb = _open_workbook(path)
    try:
        names = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        if not names:
            raise DataError("File Excel tidak punya sheet yang terlihat.")
        sheet = sheet or names[0]
        if sheet not in names:
            raise DataError(f"Sheet '{sheet}' tidak ada. Sheet tersedia: {', '.join(names)}")
        ws = wb[sheet]
        try:
            ws.reset_dimensions()          # some exporters write a wrong sheet size
        except AttributeError:
            pass
        numbered = [(i, [_cell_text(v) for v in r])
                    for i, r in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1)]
    finally:
        wb.close()
    numbered = [(i, r) for i, r in numbered if any(c for c in r)]   # i = real Excel row number
    if not numbered:
        raise DataError(f"Sheet '{sheet}' kosong.")
    width = max(len(r) for _, r in numbered)
    numbered = [(i, r + [""] * (width - len(r))) for i, r in numbered]
    numbered = numbered[find_header_index([r for _, r in numbered]):]   # skip title rows above the table
    head, body = numbered[0][1], numbered[1:]
    keep = [j for j in range(width) if head[j] or any(r[j] for _, r in body)]
    header = [head[j] or f"Kolom{j + 1}" for j in keep]
    return header, [[r[j] for j in keep] for _, r in body], [i for i, _ in body]


def read_table(path, sheet=None, with_rows=False):
    """Returns (header, rows), or (header, rows, row_numbers) when with_rows=True.
    row_numbers are the real row/line numbers in the file (blank rows are skipped)."""
    if is_excel(path):
        header, rows, nums = read_excel_table(path, sheet)
    else:
        _check_size(path)
        text = read_text(path)
        first = text.splitlines()[0] if text else ""
        delim = max(",;\t", key=first.count)
        rdr = csv.reader(io.StringIO(text, newline=""), delimiter=delim)  # keeps line breaks inside cells
        # line_num = last line of the row; subtract line breaks inside cells to get its first line
        table = [(rdr.line_num - sum(c.count("\n") for c in r), r) for r in rdr if any(c.strip() for c in r)]
        if not table:
            raise DataError("CSV kosong.")
        table = table[find_header_index([r for _, r in table]):]
        header, rows, nums = table[0][1], [r for _, r in table[1:]], [n for n, _ in table[1:]]
    return (header, rows, nums) if with_rows else (header, rows)


def pick_name_column(header, kolom=None):
    if kolom:
        if kolom not in header:
            raise DataError(f"Kolom '{kolom}' tidak ada. Kolom tersedia: {', '.join(header)}")
        return header.index(kolom)
    cand = [i for i, h in enumerate(header) if "nama" in h.lower()]
    if not cand:
        raise DataError(f"Tidak ketemu kolom nama. Pilih manual. Kolom tersedia: {', '.join(header)}")
    return cand[0]


def read_corrections(path):
    kh, kr = read_table(path)
    low = [h.strip().lower() for h in kh]
    if "nama" not in low or "file" not in low:
        raise DataError("File koreksi harus punya kolom 'nama' dan 'file'.")
    ni, fi = low.index("nama"), low.index("file")
    return {r[ni].strip().lower(): r[fi].strip() for r in kr if len(r) > max(ni, fi)}


# ---------- name matching ----------

def tokens(text):
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c)).lower()
    text = text.replace(".", "")                 # S.Pd -> spd
    text = re.sub(r"[^a-z]+", " ", text)         # digits, _, -, (), etc. -> space
    toks = [ALIASES.get(t, t) for t in text.split()][:MAX_NAME_WORDS]
    return [t for t in toks if t not in TITLES] or toks


@functools.lru_cache(maxsize=300_000)       # the same words are compared over and over
def ratio(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def similarity(a, b):
    """a, b = token lists. Returns 0..1."""
    ca, cb = "".join(a), "".join(b)
    if not ca or not cb:
        return 0.0
    if ca == cb:                                  # also covers "no spaces"
        return 1.0
    best = ratio(ca, cb)                          # typos
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    # incomplete name / reordered words
    tok = sum(max(ratio(s, w) for w in long_) for s in short) / len(short)
    tok *= 0.8 + 0.2 * len(short) / len(long_)
    if len(short) < 2 and len(long_) > 1:         # single word is too weak
        tok = min(tok, 0.75)
    best = max(best, tok)
    sc, lc = sorted((ca, cb), key=len)
    if len(sc) >= 8 and sc in lc:                 # one name contained in the other
        best = max(best, 0.87)
    return best


def find_photos(folder, recursive):
    it = Path(folder).rglob("*") if recursive else Path(folder).glob("*")
    # Skip hidden/junk files: "._Name.jpg" are macOS metadata twins created when copying from a
    # Mac drive; they are not real images and would tie with the real photo.
    return sorted(p for p in it
                  if p.is_file() and p.suffix.lower() in PHOTO_EXT
                  and not p.name.startswith((".", "~$"))
                  and "__MACOSX" not in p.parts)


def match(names, photos, forced):
    """Returns list of (photo_index|None, score, runner_up_score, note) per name.
    forced: {name.lower(): file name, stem or full path} manual corrections."""
    ptoks = [tokens(p.stem) for p in photos]
    by_path = {str(p.resolve()).lower(): i for i, p in enumerate(photos)}
    by_name = {p.name.lower(): i for i, p in enumerate(photos)}
    by_stem = {p.stem.lower(): i for i, p in enumerate(photos)}

    result = [None] * len(names)
    used = set()
    for i, n in enumerate(names):                 # manual corrections first
        f = forced.get(n.strip().lower())
        if f:
            key = f.lower()
            j = by_path.get(key)
            if j is None and ("\\" in f or "/" in f):
                j = by_path.get(str(Path(f).resolve()).lower())
            if j is None:
                j = by_name.get(key, by_stem.get(key))
            if j is not None and j not in used:
                result[i] = (j, 1.0, 0.0, "manual")
                used.add(j)

    # Cheap pre-filter: pairs sharing almost no letter pairs cannot be similar names, so skip the
    # expensive comparison (makes 5000 photos x 40 names run in seconds instead of minutes).
    def bigrams(toks):
        c = "".join(toks)
        return {c[i:i + 2] for i in range(len(c) - 1)} or {c}

    pbg = [bigrams(pt) for pt in ptoks]
    scores = []
    for n in names:
        nt = tokens(n)
        nb = bigrams(nt)
        cand = []                                    # (dice, overlap, photo) that pass the pre-filter
        for j, pb in enumerate(pbg):
            inter = len(nb & pb)
            if inter:
                ov = inter / min(len(nb), len(pb))
                if ov >= PREFILTER:
                    cand.append((2 * inter / (len(nb) + len(pb)), ov, j))
        if len(cand) > 2 * TOPK:                     # many look alike: only the closest ones can win
            keep = {c[2] for c in sorted(cand, reverse=True)[:TOPK]}
            keep |= {c[2] for c in sorted(cand, key=lambda c: (c[1], c[0]), reverse=True)[:TOPK]}
            cand = [c for c in cand if c[2] in keep]
        row = [0.0] * len(photos)
        for _, _, j in cand:
            row[j] = similarity(nt, ptoks[j])
        scores.append(row)
    pairs = sorted(((s, i, j) for i, row in enumerate(scores)
                    for j, s in enumerate(row) if s >= CHECK_SCORE),
                   key=lambda p: (-p[0], p[1], p[2]))     # best score first; ties -> earlier row
    for s, i, j in pairs:                         # best pairs first, one photo per person
        if result[i] is None and j not in used:
            runner = max((x for k, x in enumerate(scores[i]) if k != j), default=0.0)
            result[i] = (j, s, runner, "")
            used.add(j)

    # Two people compete for one photo (e.g. two students with the same name, one photo):
    # the winner is then ambiguous too, so treat the loser's score as a runner-up.
    rival = {}
    for i, r in enumerate(result):
        if r is None:
            for j, s in enumerate(scores[i]):
                if s >= CHECK_SCORE:
                    rival[j] = max(rival.get(j, 0.0), s)
    out = []
    for r in result:
        if r is None:
            out.append((None, 0.0, 0.0, ""))
        else:
            j, s, runner, note = r
            if note != "manual":
                runner = max(runner, rival.get(j, 0.0))
            out.append((j, s, runner, note))
    return out


# ---------- birth-date cleanup ----------

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
BLN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
_EN_FULL = ["January", "February", "March", "April", "May", "June", "July",
            "August", "September", "October", "November", "December"]
_MONTH_KEYS = {"jan": 1, "feb": 2, "peb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
               "jun": 6, "jul": 7, "agu": 8, "aug": 8, "agt": 8, "sep": 9, "okt": 10,
               "oct": 10, "nov": 11, "nop": 11, "des": 12, "dec": 12}
_MONTH_NAMES = {m: [BULAN[m - 1].lower(), _EN_FULL[m - 1].lower(), BLN[m - 1].lower()]
                for m in range(1, 13)}
_MONTH_NAMES[2].append("pebruari")
_MONTH_NAMES[11].append("nopember")

DEFAULT_DATE_FORMAT = "{d} {bulan} {yyyy}"
# (template, shown to the user as an example of 21 Aug 2000). Tokens:
# {d} {dd} {mm} {yyyy} {bulan} {BULAN} {bln} {BLN}
DATE_TEMPLATES = [
    "{d} {bulan} {yyyy}",
    "{d} {BULAN} {yyyy}",
    "{d} {bln} {yyyy}",
    "{d}-{bln}-{yyyy}",
    "{dd}-{mm}-{yyyy}",
    "{dd}/{mm}/{yyyy}",
]

_SEP = r"\s*[-/.,\s]*\s*"
RX_ISO = re.compile(r"(?<!\d)(\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})(?!\d)")
RX_NUM = re.compile(r"(?<!\d)(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4}|\d{2})(?!\d)")
RX_TXT = re.compile(r"(?<!\d)(\d{1,2})\s*(?:st|nd|rd|th)?" + _SEP +
                    r"([A-Za-z]{3,12})\.?" + _SEP + r"(\d{4}|\d{2})(?!\d)", re.I)
RX_TXT2 = re.compile(r"([A-Za-z]{3,12})\.?\s*(\d{1,2})\s*,?\s*(\d{4})(?!\d)")


def format_date(y, m, d, template=DEFAULT_DATE_FORMAT):
    return (template.replace("{yyyy}", f"{y:04d}").replace("{dd}", f"{d:02d}")
            .replace("{mm}", f"{m:02d}").replace("{d}", str(d))
            .replace("{bulan}", BULAN[m - 1]).replace("{BULAN}", BULAN[m - 1].upper())
            .replace("{bln}", BLN[m - 1]).replace("{BLN}", BLN[m - 1].upper()))


def template_example(template):
    return format_date(2000, 8, 21, template)


def month_from_word(word):
    w = word.lower().strip(".")
    m = _MONTH_KEYS.get(w[:3])
    if m and max(ratio(w, n) for n in _MONTH_NAMES[m]) >= 0.6:
        return m
    best = max(((ratio(w, n), mm) for mm, names in _MONTH_NAMES.items() for n in names),
               default=(0, None))
    return best[1] if best[0] >= 0.75 else None


def parse_date(text):
    """Find a date inside text. Returns (y, m, d, (start, end), doubtful) or None.
    Numeric dates are read day-month-year (Indonesian convention)."""
    t = text or ""
    doubtful = False
    if (mt := RX_ISO.search(t)):
        y, m, d = int(mt[1]), int(mt[2]), int(mt[3])
    elif (mt := RX_NUM.search(t)):
        d, m, y = int(mt[1]), int(mt[2]), mt[3]
        if m > 12 and d <= 12:                     # looks like month/day/year
            d, m, doubtful = m, d, True
        y = int(y)
        if len(mt[3]) == 2:
            y, doubtful = (2000 + y if y <= 30 else 1900 + y), True
    elif (mt := RX_TXT.search(t)) and month_from_word(mt[2]):
        d, m, y = int(mt[1]), month_from_word(mt[2]), mt[3]
        y = int(y)
        if len(mt[3]) == 2:
            y, doubtful = (2000 + y if y <= 30 else 1900 + y), True
    elif (mt := RX_TXT2.search(t)) and month_from_word(mt[1]):
        m, d, y = month_from_word(mt[1]), int(mt[2]), int(mt[3])
    else:
        return None
    try:
        datetime.date(y, m, d)
    except ValueError:
        return None
    if not 1930 <= y <= datetime.date.today().year:
        doubtful = True
    return y, m, d, mt.span(), doubtful


def convert_date(text, template=DEFAULT_DATE_FORMAT):
    """Returns (new_text, status): status is '' (empty cell), 'OK', 'RAGU' or 'GAGAL'.
    Only the date part is replaced; other text in the cell is kept."""
    if not (text or "").strip():
        return text, ""
    p = parse_date(text)
    if p is None:
        return text, "GAGAL"
    y, m, d, (s, e), doubtful = p
    new = (text[:s] + format_date(y, m, d, template) + text[e:]).strip()
    return new, "RAGU" if doubtful else "OK"


RX_ISO_FULL = re.compile(r"\s*(\d{4})-(\d{2})-(\d{2})\s*")
RX_NUM_FULL = re.compile(r"\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*\d{2,4}\s*")


def is_swap_candidate(text):
    """A real Excel date cell (read as YYYY-MM-DD) whose day <= 12 and day != month."""
    m = RX_ISO_FULL.fullmatch(text or "")
    return bool(m) and int(m[3]) <= 12 and int(m[3]) != int(m[2])


def detect_swapped_dates(values):
    """Excel in an m/d locale turns d/m/y text into a real date only when day <= 12 (read as m/d,
    so day and month end up swapped); dates with day > 12 stay text. Signature in one column:
    several numeric text dates whose first number is always > 12, next to real date cells with
    day <= 12. Returns True when that pattern is present (a hint, not proof)."""
    txt = [m for m in (RX_NUM_FULL.fullmatch(v or "") for v in values) if m]
    if len(txt) < 5 or any(int(m[1]) <= 12 for m in txt):
        return False
    return any(is_swap_candidate(v) for v in values)


def pick_date_column(header):
    """Index of the birth-date column or None."""
    low = [h.lower() for h in header]
    for test in (lambda h: "lahir" in h and ("tanggal" in h or "tgl" in h),
                 lambda h: "tanggal" in h or "tgl" in h,
                 lambda h: "ttl" in h):
        for i, h in enumerate(low):
            if test(h):
                return i
    return None


# ---------- main routine ----------

def clean_cell(v):
    """Tabs / line breaks inside a cell would break the merge file: turn them into spaces."""
    return re.sub(r"[\t\r\n]+", " ", v).strip()


def write_rows(path, encoding, delimiter, rows):
    try:
        with open(path, "w", encoding=encoding, newline="") as f:
            csv.writer(f, delimiter=delimiter).writerows(rows)
    except PermissionError:
        raise DataError(f"File hasil tidak bisa ditulis (sedang terbuka di Excel/InDesign?):\n{path}\n"
                        "Tutup dulu file itu, lalu klik Proses lagi.")


def proses(csv_path, foto_dir, kolom=None, kolom_foto="@Foto", keluar=None,
           koreksi=None, forced=None, isi_ragu=False, subfolder=True, utf16=True,
           kolom_tanggal=None, format_tanggal=DEFAULT_DATE_FORMAT, tgl_manual=None,
           sheet=None, tukar_hari_bulan=False):
    """Run the whole job and write the _merge file (the only file written).

    kolom_tanggal: header of the birth-date column to normalise (None = leave as is).
    tgl_manual: {csv row number: typed date text} manual date edits.
    Returns a dict: counts, tgl_counts, report (list of dicts), merge_path,
    outdir, kolom, n_photos, n_rows.  Raises DataError for bad input."""
    csv_path = Path(csv_path)
    if not csv_path.is_file():
        raise DataError(f"File data tidak ditemukan: {csv_path}")
    if not Path(foto_dir).is_dir():
        raise DataError(f"Folder foto tidak ditemukan: {foto_dir}")

    stem = csv_path.stem
    if is_excel(csv_path):
        sheet = sheet or list_sheets(csv_path)[0]
        stem += " - " + re.sub(r'[\\/:*?"<>|]+', "_", sheet)   # one output set per sheet
    header, rows, rownums = read_table(csv_path, sheet, with_rows=True)
    rows = [r + [""] * (len(header) - len(r)) for r in rows]
    col = pick_name_column(header, kolom)

    photos = find_photos(foto_dir, subfolder)
    if not photos:
        raise DataError("Tidak ada file foto (jpg/png/tif/psd/webp) di folder itu.")

    dcol = None
    if kolom_tanggal:
        if kolom_tanggal not in header:
            raise DataError(f"Kolom tanggal '{kolom_tanggal}' tidak ada. Kolom tersedia: {', '.join(header)}")
        dcol = header.index(kolom_tanggal)
    tgl_manual = tgl_manual or {}

    all_forced = read_corrections(koreksi) if koreksi else {}
    all_forced.update({k.strip().lower(): v for k, v in (forced or {}).items()})

    names = [r[col] for r in rows]
    res = match(names, photos, all_forced)

    fcol = header.index(kolom_foto) if kolom_foto in header else None
    out_header = header + ([kolom_foto] if fcol is None else [])
    fcol = len(header) if fcol is None else fcol
    out_rows, report = [], []
    data_rows = []        # full-sheet view for the app: (row number, cells, photo status, date status)
    data_header = header + ["Foto terpilih", "Status foto"] + (["Status tanggal"] if dcol is not None else [])
    counts = {"OK": 0, "CEK": 0, "TIDAK ADA": 0, "DIABAIKAN": 0}
    tgl_counts = {"OK": 0, "RAGU": 0, "GAGAL": 0, "MANUAL": 0, "CURIGA": 0, "DITUKAR": 0}
    # CURIGA = real Excel date that may have day/month swapped; DITUKAR = we swapped it (user's choice)
    suspect_swap = (dcol is not None and is_excel(csv_path)
                    and detect_swapped_dates([r[dcol] for r in rows]))
    # A row with no photo and nothing but a name is probably a title / footer line ("BIODATA ANGKATAN 29"),
    # but only when there are a few of them. If many rows look like that (e.g. a sheet that only has a
    # name column) they are real people without a photo and must stay visible as "Tidak ada".
    blank_rows = {i for i, (r, x) in enumerate(zip(rows, res))
                  if x[0] is None and not any(c.strip() for k, c in enumerate(r) if k != col)}
    ignorable = blank_rows if len(blank_rows) <= max(3, len(rows) // 20) else set()
    for i, (n, row, (j, s, runner, note)) in enumerate(zip(rownums, rows, res)):   # n = real row number
        if i in ignorable:
            counts["DIABAIKAN"] += 1
            data_rows.append((n, list(row) + ["", "DIABAIKAN"] + ([""] if dcol is not None else []),
                              "DIABAIKAN", ""))
            report.append(dict(baris=n, nama=row[col], status="DIABAIKAN", file="", path="",
                               skor="", catatan="bukan siswa? (kolom lain kosong & tidak ada foto), "
                                                "tidak dimasukkan ke _merge",
                               tgl_asli="", tgl_baru="", tgl_status=""))
            continue
        tgl_asli = tgl_baru = tgl_status = ""
        if dcol is not None:
            tgl_asli = row[dcol]
            if n in tgl_manual:                    # typed by the user in the app
                parsed = convert_date(tgl_manual[n], format_tanggal)
                tgl_baru, tgl_status = parsed[0], "MANUAL"
            else:
                src = tgl_asli
                swap_cell = suspect_swap and is_swap_candidate(tgl_asli)
                if swap_cell and tukar_hari_bulan:
                    y_, m_, d_ = RX_ISO_FULL.fullmatch(tgl_asli).groups()
                    src = f"{y_}-{d_}-{m_}"                     # new month = old day, new day = old month
                tgl_baru, tgl_status = convert_date(src, format_tanggal)
                if swap_cell and tgl_status == "OK":
                    tgl_status = "DITUKAR" if tukar_hari_bulan else "CURIGA"
            if tgl_status:
                tgl_counts[tgl_status if tgl_status in tgl_counts else "GAGAL"] += 1
            row = list(row)
            row[dcol] = tgl_baru
        if j is None:
            status = "TIDAK ADA"
        elif note == "manual" or (s >= OK_SCORE and runner < s - TIE_MARGIN):
            status = "OK"
        else:
            status = "CEK"
        counts[status] += 1
        fill = j is not None and (status == "OK" or isi_ragu)
        data_rows.append((n, list(row[:len(header)]) + [photos[j].name if j is not None else "", status]
                          + ([tgl_status] if dcol is not None else []), status, tgl_status))
        row = [clean_cell(c) for c in row] + [""] * (len(out_header) - len(row))
        row[fcol] = str(photos[j].resolve()) if fill else ""
        out_rows.append(row)
        report.append(dict(baris=n, nama=row[col], status=status,
                           file=photos[j].name if j is not None else "",
                           path=str(photos[j].resolve()) if j is not None else "",
                           skor=f"{s:.0%}" if j is not None else "", catatan=note,
                           tgl_asli=tgl_asli, tgl_baru=tgl_baru, tgl_status=tgl_status))

    used = {j for j, *_ in res if j is not None}
    unused = [p for k, p in enumerate(photos) if k not in used]
    for p in unused:
        report.append(dict(baris="", nama="", status="FOTO TIDAK TERPAKAI", file=p.name,
                           path=str(p.resolve()), skor="", catatan="",
                           tgl_asli="", tgl_baru="", tgl_status=""))

    outdir = Path(keluar) if keluar else csv_path.parent / "hasil_merge"
    outdir.mkdir(parents=True, exist_ok=True)
    # InDesign's Data Merge only loads this reliably as UTF-16 tab-delimited (.txt)
    # (confirmed by the user: UTF-8 CSV errors in InDesign), so UTF-16 is the default.
    if utf16:
        merge_path = outdir / f"{stem}_merge.txt"
        write_rows(merge_path, "utf-16", "\t", [[clean_cell(h) for h in out_header]] + out_rows)
    else:
        merge_path = outdir / f"{stem}_merge.csv"
        write_rows(merge_path, "utf-8-sig", ",", [[clean_cell(h) for h in out_header]] + out_rows)
    return dict(counts=counts, tgl_counts=tgl_counts, tgl_suspect=suspect_swap, report=report,
                data_header=data_header, data_rows=data_rows, unused=len(unused), merge_path=merge_path,
                outdir=outdir, kolom=header[col], sheet=sheet,
                n_photos=len(photos), n_rows=len(rows))


# ---------- command line ----------

def clean_path(p):
    return p.strip().strip('"').strip("'")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("csv", nargs="?", help="CSV dari database")
    ap.add_argument("foto", nargs="?", help="folder foto")
    ap.add_argument("--kolom", help="nama kolom berisi nama orang (default: kolom pertama yang mengandung 'nama')")
    ap.add_argument("--kolom-foto", default="@Foto", help="header kolom gambar (default @Foto)")
    ap.add_argument("--keluar", help="folder hasil (default: <folder csv>\\hasil_merge)")
    ap.add_argument("--koreksi", help="CSV koreksi manual dengan kolom: nama,file")
    ap.add_argument("--isi-ragu", action="store_true", help="isi juga baris CEK")
    ap.add_argument("--tanpa-subfolder", action="store_true", help="jangan cari foto di subfolder")
    ap.add_argument("--utf8", action="store_true",
                    help="tulis CSV UTF-8 biasa (default: .txt UTF-16 tab-delimited, yang dibaca InDesign)")
    ap.add_argument("--sheet", help="nama sheet kalau file berformat .xlsx (default: sheet pertama)")
    ap.add_argument("--tukar-hari-bulan", action="store_true",
                    help="tukar hari<->bulan pada sel tanggal Excel yang diduga tertukar (lihat peringatan)")
    ap.add_argument("--kolom-tanggal", help="rapikan kolom tanggal lahir ini ('auto' = deteksi otomatis)")
    ap.add_argument("--format-tanggal", default=DEFAULT_DATE_FORMAT,
                    help="template tanggal, default '{d} {bulan} {yyyy}' = 21 Agustus 2000. "
                         "Token: {d} {dd} {mm} {yyyy} {bulan} {BULAN} {bln} {BLN}")
    args = ap.parse_args()

    interactive = args.csv is None
    if interactive:
        args.csv = clean_path(input("Seret file CSV ke sini lalu Enter: "))
        args.foto = clean_path(input("Seret FOLDER foto ke sini lalu Enter: "))
    if not args.foto:
        sys.exit("Folder foto belum diisi.")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    try:
        kt = args.kolom_tanggal
        if kt == "auto":
            header, _ = read_table(args.csv, args.sheet)
            i = pick_date_column(header)
            kt = header[i] if i is not None else None
        r = proses(args.csv, args.foto, kolom=args.kolom, kolom_foto=args.kolom_foto,
                   keluar=args.keluar, koreksi=args.koreksi, isi_ragu=args.isi_ragu,
                   subfolder=not args.tanpa_subfolder, utf16=not args.utf8,
                   kolom_tanggal=kt, format_tanggal=args.format_tanggal, sheet=args.sheet,
                   tukar_hari_bulan=args.tukar_hari_bulan)
    except DataError as e:
        sys.exit(str(e))

    c = r["counts"]
    if is_excel(args.csv):
        print(f"Sheet      : {r['sheet']}")
    print(f"Kolom nama : {r['kolom']}")
    print(f"Jumlah foto: {r['n_photos']} | Jumlah baris CSV: {r['n_rows']}")
    print(f"\nOK: {c['OK']} | CEK: {c['CEK']} | TIDAK ADA FOTO: {c['TIDAK ADA']} "
          f"| Diabaikan: {c['DIABAIKAN']} | Foto tidak terpakai: {r['unused']}")
    for x in r["report"]:
        if x["status"] in ("CEK", "TIDAK ADA", "DIABAIKAN"):
            print(f"  [{x['status']}] baris {x['baris']}: {x['nama']}  ->  {x['file'] or '-'} {x['skor']}")
        elif x["status"] == "FOTO TIDAK TERPAKAI":
            print(f"  [FOTO TIDAK TERPAKAI] {x['file']}")
    if kt:
        t = r["tgl_counts"]
        print(f"\nTanggal ({kt}): dirapikan {t['OK']} | RAGU {t['RAGU']} | GAGAL dibaca {t['GAGAL']}"
              f" | CURIGA tertukar {t['CURIGA']} | DITUKAR {t['DITUKAR']}")
        if r["tgl_suspect"] and not args.tukar_hari_bulan:
            print("  ! Pola khas Excel menukar hari/bulan terdeteksi. Cek tanggal berstatus CURIGA ke database;"
                  " kalau memang tertukar, jalankan lagi dengan --tukar-hari-bulan")
        for x in r["report"]:
            if x["tgl_status"] in ("RAGU", "GAGAL", "CURIGA", "DITUKAR"):
                print(f"  [{x['tgl_status']}] baris {x['baris']}: {x['tgl_asli']!r} -> {x['tgl_baru']!r}")
    print(f"\nUntuk InDesign : {r['merge_path']}")
    if interactive:
        input("\nTekan Enter untuk menutup...")


if __name__ == "__main__":
    main()
