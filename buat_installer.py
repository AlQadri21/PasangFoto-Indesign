"""Builds Aplikasi\\PasangkanFoto_Setup.exe (one-file installer).

Run after you rebuilt the app folder (Aplikasi\\PasangkanFoto_Cepat):   python buat_installer.py
"""
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
APP_DIR = HERE / "Aplikasi" / "PasangkanFoto_Cepat"
OUT_DIR = HERE / "Aplikasi"

if not (APP_DIR / "PasangkanFoto_Cepat.exe").is_file():
    sys.exit(f"Folder aplikasi belum ada: {APP_DIR}\nBangun dulu versi folder (lihat PANDUAN.md).")

work = Path(tempfile.mkdtemp(prefix="pf_setup_"))
try:
    payload = work / "payload.zip"
    print("Mengemas aplikasi…")
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for f in sorted(APP_DIR.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(APP_DIR))
    print(f"  paket: {payload.stat().st_size / 1048576:.1f} MB")
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed",
           "--name", "PasangkanFoto_Setup", "--icon", str(HERE / "ikon.ico"),
           "--add-data", f"{payload};.", "--paths", str(HERE),
           "--distpath", str(OUT_DIR), "--workpath", str(work / "build"), "--specpath", str(work),
           str(HERE / "installer.py")]
    print("Membangun installer…")
    subprocess.run(cmd, check=True)
    exe = OUT_DIR / "PasangkanFoto_Setup.exe"
    print(f"Selesai: {exe} ({exe.stat().st_size / 1048576:.1f} MB)")
finally:
    shutil.rmtree(work, ignore_errors=True)
