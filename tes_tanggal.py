"""Quick checks for the birth-date cleanup. Run: python tes_tanggal.py"""
from cocokkan_foto import DATE_TEMPLATES, convert_date, pick_date_column, template_example

T = "{d} {bulan} {yyyy}"
OK = [  # (input, expected output, expected status)
    ("21-agustus-2000", "21 Agustus 2000", "OK"),
    ("11-Apr-2009", "11 April 2009", "OK"),
    ("5-Agu-2009", "5 Agustus 2009", "OK"),
    ("25-Des-2009", "25 Desember 2009", "OK"),
    ("10-Mei-2009", "10 Mei 2009", "OK"),
    ("21 Agu 2000", "21 Agustus 2000", "OK"),
    ("21/08/2000", "21 Agustus 2000", "OK"),
    ("21.08.2000", "21 Agustus 2000", "OK"),
    ("2000-08-21", "21 Agustus 2000", "OK"),
    ("01-02-2009", "1 Februari 2009", "OK"),
    ("Agustus 21, 2000", "21 Agustus 2000", "OK"),
    ("21 AGUSTUS2000", "21 Agustus 2000", "OK"),
    ("21 Agusuts 2000", "21 Agustus 2000", "OK"),            # typo
    ("21 Agt 2000", "21 Agustus 2000", "OK"),
    ("3 Sept 2008", "3 September 2008", "OK"),
    ("3 Nopember 2008", "3 November 2008", "OK"),
    ("21 August 2000", "21 Agustus 2000", "OK"),
    ("Makassar, 11 April 2009", "Makassar, 11 April 2009", "OK"),   # rest of text kept
    ("  21-agustus-2000  ", "21 Agustus 2000", "OK"),
    ("21-08-00", "21 Agustus 2000", "RAGU"),                  # 2-digit year
    ("12/31/2000", "31 Desember 2000", "RAGU"),               # month/day order guessed
    ("21-08-1800", "21 Agustus 1800", "RAGU"),                # implausible year
    ("31-feb-2000", "31-feb-2000", "GAGAL"),                  # impossible date
    ("abc", "abc", "GAGAL"),
    ("Makassar", "Makassar", "GAGAL"),
    ("", "", ""),
]
bad = 0
for src, want, status in OK:
    got = convert_date(src, T)
    if got != (want, status):
        bad += 1
        print(f"SALAH: {src!r} -> {got!r}, seharusnya {(want, status)!r}")

assert [template_example(t) for t in DATE_TEMPLATES] == [
    "21 Agustus 2000", "21 AGUSTUS 2000", "21 Agu 2000", "21-Agu-2000",
    "21-08-2000", "21/08/2000"]
assert pick_date_column(["NAMA", "TEMPAT LAHIR", "TANGGAL LAHIR"]) == 2
assert pick_date_column(["Nama", "Tgl Lahir"]) == 1
assert pick_date_column(["Nama", "Kelas"]) is None
print("SEMUA LOLOS" if not bad else f"{bad} tes gagal")
raise SystemExit(1 if bad else 0)
