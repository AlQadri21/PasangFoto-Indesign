"""Regression checks for cocokkan_foto.py. Run: python tes_inti.py"""
import csv
import datetime
import os
import sys
import tempfile
from pathlib import Path

import openpyxl

from cocokkan_foto import DataError, proses, read_table, write_rows

fails = []


def check(name, cond, info=""):
    print(("  ok   " if cond else "  GAGAL") + f" {name}" + (f"  -> {info}" if not cond else ""))
    if not cond:
        fails.append(name)


def touch(folder, *names):
    folder.mkdir(parents=True, exist_ok=True)
    for n in names:
        (folder / n).write_bytes(b"x")


def read_utf16(path):
    raw = Path(path).read_bytes()
    text = raw.decode("utf-16")
    return raw, [r for r in csv.reader(text.splitlines(), delimiter="\t")]


with tempfile.TemporaryDirectory(dir=os.environ.get("TES_TMP")) as t:
    t = Path(t)

    # 1. Excel: title rows above the header, blank row in the middle -> real row numbers
    print("Excel: nomor baris asli")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "K1"
    ws["A1"] = "JUDUL DI ATAS"
    ws.append([])
    ws.append(["NAMA", "TEMPAT LAHIR", "TANGGAL LAHIR", "MOTTO"])        # header on row 3
    ws.append(["Budi Santoso", "Makassar   ", datetime.datetime(2009, 4, 11), "satu\ndua\ttiga"])  # row 4
    ws.append([])                                                          # blank row 5
    ws.append(["Siti Aisyah", "Gowa", "5-Agu-2009", '"kutipan"'])         # row 6
    wb.save(t / "k.xlsx")
    touch(t / "foto", "Budi Santoso.jpg", "Siti Aisyah.jpg", "._Budi Santoso.jpg", ".DS_Store.jpg")
    h, rows, nums = read_table(t / "k.xlsx", with_rows=True)
    check("header terbaca dari baris 3", h == ["NAMA", "TEMPAT LAHIR", "TANGGAL LAHIR", "MOTTO"], h)
    check("nomor baris = baris Excel asli", nums == [4, 6], nums)

    r = proses(t / "k.xlsx", t / "foto", kolom_tanggal="TANGGAL LAHIR", keluar=t / "out")
    check("2 baris, 0 CEK (file '._' diabaikan)", r["counts"]["OK"] == 2 and r["counts"]["CEK"] == 0, r["counts"])
    check("foto = 2 (bukan 4)", r["n_photos"] == 2, r["n_photos"])
    check("hasil memakai baris asli", [x["baris"] for x in r["report"][:2]] == [4, 6],
          [x["baris"] for x in r["report"][:2]])
    check("hanya file _merge yang ditulis (tanpa laporan)",
          sorted(p.name for p in (t / "out").iterdir()) == [r["merge_path"].name],
          sorted(p.name for p in (t / "out").iterdir()))

    # 2. UTF-16 default output
    print("Keluaran UTF-16")
    check("default = .txt", r["merge_path"].suffix == ".txt", r["merge_path"].name)
    raw, tab = read_utf16(r["merge_path"])
    check("ada BOM UTF-16", raw[:2] in (b"\xff\xfe", b"\xfe\xff"))
    check("header + kolom @Foto", tab[0] == ["NAMA", "TEMPAT LAHIR", "TANGGAL LAHIR", "MOTTO", "@Foto"], tab[0])
    check("spasi ujung dibuang", tab[1][1] == "Makassar", tab[1][1])
    check("tab/enter dalam sel -> spasi", tab[1][3] == "satu dua tiga", tab[1][3])
    check("tanggal Excel asli dirapikan", tab[1][2] == "11 April 2009", tab[1][2])
    check("tanggal teks dirapikan", tab[2][2] == "5 Agustus 2009", tab[2][2])
    check("kutip dalam sel selamat", tab[2][3] == '"kutipan"', tab[2][3])
    check("path foto ada", tab[1][4].endswith("Budi Santoso.jpg") and ":" in tab[1][4], tab[1][4])
    r8 = proses(t / "k.xlsx", t / "foto", keluar=t / "out8", utf16=False)
    check("opsi UTF-8 -> .csv", r8["merge_path"].suffix == ".csv", r8["merge_path"].name)

    # 3. CSV with a line break inside a quoted cell
    print("CSV")
    (t / "m.csv").write_bytes(b'NAMA,MOTTO\r\nBudi Santoso,"baris satu\r\nbaris dua"\r\n\r\nSiti Aisyah,ok\r\n')
    h, rows, nums = read_table(t / "m.csv", with_rows=True)
    check("sel multi-baris utuh", rows[0][1] == "baris satu\r\nbaris dua", repr(rows[0][1]))
    check("nomor baris CSV = baris awal (2 dan 5)", nums == [2, 5], nums)
    rm = proses(t / "m.csv", t / "foto", keluar=t / "outm")
    _, tab = read_utf16(rm["merge_path"])
    check("enter di CSV -> spasi", tab[1][1] == "baris satu baris dua", tab[1][1])

    # 4. Same name twice, only one photo -> the winner must not be silently trusted
    print("Nama kembar, satu foto")
    (t / "d.csv").write_text("NAMA,KELAS\nMuhammad Rizki,A\nMuhammad Rizki,B\nDewi Lestari,A\n", encoding="utf-8")
    touch(t / "fd", "Muhammad Rizki.jpg", "Dewi Lestari.jpg")
    rd = proses(t / "d.csv", t / "fd", keluar=t / "outd")
    st = [x["status"] for x in rd["report"] if x["nama"]]
    check("kembar tidak ada yang otomatis OK", st[0] == "CEK" and st[1] == "TIDAK ADA", st)
    check("siswa lain tetap OK", st[2] == "OK", st)
    check("baris CEK kosong di merge", read_utf16(rd["merge_path"])[1][1][2] == "", read_utf16(rd["merge_path"])[1][1])

    # 5. Dua nama kembar, dua foto (…, ... (2)) -> keduanya CEK (tidak bisa dibedakan)
    touch(t / "fd2", "Muhammad Rizki.jpg", "Muhammad Rizki (2).jpg")
    rd2 = proses(t / "d.csv", t / "fd2", keluar=t / "outd2")
    st2 = [x["status"] for x in rd2["report"] if x["nama"]][:2]
    check("dua kembar + dua foto -> CEK", st2 == ["CEK", "CEK"], st2)

    # 6. Koreksi manual menang
    rf = proses(t / "d.csv", t / "fd2", keluar=t / "outd3",
                forced={"muhammad rizki": str(t / "fd2" / "Muhammad Rizki (2).jpg")})
    first = [x for x in rf["report"] if x["nama"]][0]
    check("koreksi manual -> OK & foto (2)", first["status"] == "OK" and "(2)" in first["file"], first)

    # 7. Penulisan gagal -> pesan Indonesia, bukan traceback
    print("Gagal simpan")
    try:
        write_rows(t / "foto", "utf-16", "\t", [["a"]])        # a folder: cannot be opened for writing
        check("PermissionError jadi DataError", False, "tidak error")
    except DataError as e:
        check("PermissionError jadi DataError", "Tutup dulu" in str(e))

    # 7b. Excel swapped day/month signature (real date cells day<=12 + text dates all day>12)
    print("Deteksi hari/bulan tertukar")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SWAP"
    ws.append(["NAMA", "TANGGAL LAHIR"])
    ws.append(["Aisyah Dzakirah", datetime.datetime(2009, 6, 9)])      # typed 06-09-2009, stored 9 Jun
    ws.append(["Diva Audrey", datetime.datetime(2009, 5, 1)])
    ws.append(["Syaikha Ummul", datetime.datetime(2009, 2, 2)])         # day == month: swap is a no-op
    for i, tx in enumerate(["26-08-2009", "23-01-2009", "14-03-2009", "30-05-2009", "22-01-2009"]):
        ws.append([f"Siswa {i}", tx])
    wb.save(t / "swap.xlsx")
    touch(t / "fs", "Aisyah Dzakirah.jpg")
    kw = dict(kolom_tanggal="TANGGAL LAHIR", keluar=t / "outs")
    rs = proses(t / "swap.xlsx", t / "fs", **kw)
    check("pola terdeteksi", rs["tgl_suspect"] is True)
    tg = {x["nama"]: (x["tgl_baru"], x["tgl_status"]) for x in rs["report"] if x["nama"]}
    check("tanpa opsi: tanggal apa adanya + CURIGA", tg["Aisyah Dzakirah"] == ("9 Juni 2009", "CURIGA"), tg["Aisyah Dzakirah"])
    check("hari == bulan tidak dicurigai", tg["Syaikha Ummul"] == ("2 Februari 2009", "OK"), tg["Syaikha Ummul"])
    check("teks dd-mm-yyyy tidak disentuh", tg["Siswa 0"] == ("26 Agustus 2009", "OK"), tg["Siswa 0"])
    rt = proses(t / "swap.xlsx", t / "fs", tukar_hari_bulan=True, **kw)
    tg2 = {x["nama"]: (x["tgl_baru"], x["tgl_status"]) for x in rt["report"] if x["nama"]}
    check("dengan opsi tukar: 9 Jun -> 6 September", tg2["Aisyah Dzakirah"] == ("6 September 2009", "DITUKAR"), tg2["Aisyah Dzakirah"])
    check("dengan opsi tukar: 1 Mei -> 5 Januari", tg2["Diva Audrey"] == ("5 Januari 2009", "DITUKAR"), tg2["Diva Audrey"])
    check("dengan opsi tukar: hari==bulan tetap", tg2["Syaikha Ummul"] == ("2 Februari 2009", "OK"), tg2["Syaikha Ummul"])
    # no false alarm: only real dates (like IPA 1/2), or text dates that include day <= 12
    ws.title = "X"
    wb2 = openpyxl.Workbook(); w2 = wb2.active; w2.append(["NAMA", "TANGGAL LAHIR"])
    for i in range(6): w2.append([f"A{i}", datetime.datetime(2009, 6, i + 1)])
    wb2.save(t / "nofalse.xlsx")
    check("hanya sel tanggal asli -> tidak ada alarm",
          proses(t / "nofalse.xlsx", t / "fs", **kw)["tgl_suspect"] is False)
    wb3 = openpyxl.Workbook(); w3 = wb3.active; w3.append(["NAMA", "TANGGAL LAHIR"])
    w3.append(["B0", datetime.datetime(2009, 6, 5)])
    for i, tx in enumerate(["05-06-2009", "26-08-2009", "14-03-2009", "30-05-2009", "22-01-2009"]): w3.append([f"B{i+1}", tx])
    wb3.save(t / "nofalse2.xlsx")
    check("teks ber-hari<=12 ada -> tidak ada alarm",
          proses(t / "nofalse2.xlsx", t / "fs", **kw)["tgl_suspect"] is False)

    # 7c. Safeguards found by the stress test
    print("Pengaman hasil uji beban")
    import cocokkan_foto as ck
    (t / "h.csv").write_text("NAMA,MOTTO\nBudi Santoso," + "kata " * 40000 + "\n", encoding="utf-8")
    _, hrows = read_table(t / "h.csv")
    check("sel teks 200 KB terbaca (batas lama 128 KB)", len(hrows) == 1 and len(hrows[0][1]) > 150000)
    (t / "n.csv").write_text("NAMA\n" + "\n".join(f"Orang Nomor {chr(65 + i)}" for i in range(20)) + "\n", encoding="utf-8")
    rn = proses(t / "n.csv", t / "foto", keluar=t / "outn")
    check("data hanya-nama: semua 'Tidak ada', tidak ada yang diabaikan diam-diam",
          rn["counts"]["TIDAK ADA"] == 20 and rn["counts"]["DIABAIKAN"] == 0, rn["counts"])
    (t / "f.csv").write_text("NAMA,KELAS\nBudi Santoso,A\nSiti Aisyah,A\nBIODATA ANGKATAN 29,\n", encoding="utf-8")
    rf2 = proses(t / "f.csv", t / "foto", keluar=t / "outf")
    check("baris judul/footer tetap diabaikan", rf2["counts"]["DIABAIKAN"] == 1 and rf2["counts"]["OK"] == 2, rf2["counts"])
    old = ck.MAX_FILE_MB
    ck.MAX_FILE_MB = 0
    try:
        read_table(t / "f.csv")
        check("file melebihi batas ukuran ditolak", False, "tidak error")
    except DataError as e:
        check("file melebihi batas ukuran ditolak dengan pesan jelas", "terlalu besar" in str(e))
    finally:
        ck.MAX_FILE_MB = old
    check("nama super panjang dipotong (cepat)", len(ck.tokens("kata " * 500)) == ck.MAX_NAME_WORDS)

    # 8. Input salah
    for label, args, kw in (("sheet tidak ada", (t / "k.xlsx", t / "foto"), dict(sheet="X")),
                            ("folder foto salah", (t / "k.xlsx", t / "nope"), {}),
                            ("kolom nama salah", (t / "k.xlsx", t / "foto"), dict(kolom="ZZZ"))):
        try:
            proses(*args, keluar=t / "oe", **kw)
            check(label + " -> DataError", False, "tidak error")
        except DataError:
            check(label + " -> DataError", True)

print("\nSEMUA LOLOS" if not fails else f"\n{len(fails)} GAGAL: {fails}")
sys.exit(1 if fails else 0)
