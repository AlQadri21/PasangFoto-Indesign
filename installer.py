"""Installer for Pasangkan Foto: per-user (no administrator rights needed).

Double-click for the wizard, or:  PasangkanFoto_Setup.exe --silent [--dir PARENT] [--no-desktop] [--log FILE]
The application files travel inside this exe as payload.zip (made by buat_installer.py).
"""
import argparse
import datetime
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import winreg
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

try:
    from ikon_data import ICON_PNG_B64
except ImportError:
    ICON_PNG_B64 = ""

APP_NAME = "Pasangkan Foto"
EXE_NAME = "PasangkanFoto_Cepat.exe"
VERSION = "1.3"
REG_BASE = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
REG_NAME = "PasangkanFoto"
NO_WINDOW = 0x08000000           # CREATE_NO_WINDOW
DETACHED = 0x00000008 | 0x00000200   # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
FONT, FONT_BOLD = "Segoe UI", "Segoe UI Semibold"
C = dict(bg="#f3f5f8", card="#ffffff", line="#dbe0e8", text="#1c2733", muted="#667487",
         head="#1e40af", accent="#2563eb", accent_dark="#1d4ed8")


class InstallError(Exception):
    """Problem the user can understand and act on (message is in Indonesian)."""


def payload_path():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "payload.zip"


def default_parent():
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / "Programs"


def final_target(parent):
    """The app always lives in its own folder, so an install never touches unrelated files."""
    p = Path(parent)
    return p if p.name.lower() == APP_NAME.lower() else p / APP_NAME


def ps(command, timeout=60):
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-Command", command], capture_output=True, text=True, creationflags=NO_WINDOW,
                       timeout=timeout, encoding="utf-8", errors="replace")
    return r.returncode, r.stdout.strip(), r.stderr.strip()


def q(s):
    return "'" + str(s).replace("'", "''") + "'"


def desktop_dir():
    try:                                            # asks Windows, so a OneDrive-redirected Desktop works too
        code, out, _ = ps("[Environment]::GetFolderPath('Desktop')")
        if code == 0 and out:
            return Path(out)
    except (OSError, subprocess.SubprocessError):
        pass
    return Path.home() / "Desktop"


def start_menu_dir():
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs"


def make_shortcut(lnk, target, workdir, icon):
    lnk = Path(lnk)
    lnk.parent.mkdir(parents=True, exist_ok=True)
    code, _, err = ps("$s=(New-Object -ComObject WScript.Shell).CreateShortcut(%s);"
                      "$s.TargetPath=%s;$s.WorkingDirectory=%s;$s.IconLocation=%s;$s.Description=%s;$s.Save()"
                      % (q(lnk), q(target), q(workdir), q(icon), q(APP_NAME)))
    if code != 0 or not lnk.exists():
        raise InstallError(f"Shortcut gagal dibuat: {err[:200]}")


def running_instances(target):
    """Paths of PasangkanFoto processes that run from the install folder."""
    _, out, _ = ps("Get-Process -Name PasangkanFoto_Cepat -ErrorAction SilentlyContinue | ForEach-Object { $_.Path }")
    t = str(Path(target)).lower().rstrip("\\") + "\\"
    return [p for p in out.splitlines() if p.strip() and p.lower().startswith(t)]


def stop_instances(target):
    ps("Get-Process -Name PasangkanFoto_Cepat -ErrorAction SilentlyContinue | Where-Object { $_.Path -like %s } "
       "| Stop-Process -Force" % q(str(Path(target)) + "\\*"))
    time.sleep(1.0)


def uninstall_script(target, shortcuts, reg_key):
    """uninstall.cmd, written next to the app. Windows' 'Apps' list runs it."""
    # stop only the copy that runs from the install folder (other copies of the app are left alone)
    kill = ("powershell -NoProfile -NonInteractive -Command \"Get-Process -Name PasangkanFoto_Cepat "
            f"-ErrorAction SilentlyContinue | Where-Object {{ $_.Path -like '{str(target).replace(chr(39), chr(39) * 2)}\\*' }} "
            "| Stop-Process -Force\" >nul 2>&1")
    lines = ["@echo off", "setlocal", kill]
    lines += [f'del /q "{s}" >nul 2>&1' for s in shortcuts]
    lines += [f'reg delete "HKCU\\{reg_key}" /f >nul 2>&1',
              'set "CL=%TEMP%\\pf_cleanup.cmd"',
              '> "%CL%" echo @echo off',
              '>> "%CL%" echo ping 127.0.0.1 -n 3 ^>nul',
              f'>> "%CL%" echo rmdir /s /q "{target}"',
              '>> "%CL%" echo del "%%~f0"',
              'start "" /min cmd /c ""%CL%""',          # /c (not /K) so no console window stays open afterwards
              "endlocal", ""]
    return "\r\n".join(lines)


def install(parent, desktop=True, close_running=False, progress=lambda done, total: None,
            start_menu=None, desktop_folder=None, reg_base=REG_BASE):
    """Install (or update) the app. Returns a dict describing what was created."""
    target = final_target(parent)
    zpath = payload_path()
    if not zpath.is_file():
        raise InstallError("Paket aplikasi tidak ada di dalam installer (file installer rusak). Unduh/salin ulang.")

    if running_instances(target):
        if not close_running:
            raise InstallError("Aplikasi sedang berjalan. Tutup dulu, lalu pasang lagi.")
        stop_instances(target)

    if target.exists() and any(target.iterdir()) and not (target / EXE_NAME).exists():
        raise InstallError(f"Folder {target} sudah berisi file lain. Pilih folder tujuan yang berbeda.")

    staging = target.with_name(target.name + ".baru")
    try:
        if staging.exists():
            shutil.rmtree(staging)
        staging.mkdir(parents=True)
    except OSError as e:
        raise InstallError(f"Tidak bisa menulis ke folder tujuan:\n{e}\n\nPilih folder lain (mis. di drive D:).")
    try:
        with zipfile.ZipFile(zpath) as zf:
            infos = zf.infolist()
            total = sum(i.file_size for i in infos) or 1
            done = 0
            for info in infos:
                zf.extract(info, staging)
                done += info.file_size
                progress(done, total)
        if not (staging / EXE_NAME).is_file():
            raise InstallError("Paket aplikasi tidak lengkap (file utama tidak ada).")
        if target.exists():
            try:
                shutil.rmtree(target)            # only reached for an earlier install of this app
            except OSError:
                raise InstallError("Versi lama tidak bisa diganti karena ada file yang sedang dipakai. "
                                   "Tutup aplikasi lalu pasang lagi.")
        staging.rename(target)
    except OSError as e:
        raise InstallError(f"Gagal menyalin file: {e}")
    finally:
        shutil.rmtree(staging, ignore_errors=True)

    exe = target / EXE_NAME
    shortcuts = []
    sm = Path(start_menu) if start_menu else start_menu_dir()
    make_shortcut(sm / f"{APP_NAME}.lnk", exe, target, exe)
    shortcuts.append(sm / f"{APP_NAME}.lnk")
    if desktop:
        dd = Path(desktop_folder) if desktop_folder else desktop_dir()
        make_shortcut(dd / f"{APP_NAME}.lnk", exe, target, exe)
        shortcuts.append(dd / f"{APP_NAME}.lnk")

    reg_key = f"{reg_base}\\{REG_NAME}"
    script = target / "uninstall.cmd"
    script.write_text(uninstall_script(target, shortcuts, reg_key), encoding="mbcs", errors="replace")
    size_kb = int(sum(f.stat().st_size for f in target.rglob("*") if f.is_file()) / 1024)
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, reg_key, 0, winreg.KEY_WRITE) as k:
        for name, val in (("DisplayName", APP_NAME), ("DisplayVersion", VERSION), ("DisplayIcon", str(exe)),
                          ("InstallLocation", str(target)), ("UninstallString", f'cmd.exe /c ""{script}""'),
                          ("InstallDate", datetime.date.today().strftime("%Y%m%d"))):
            winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)
        for name, val in (("EstimatedSize", size_kb), ("NoModify", 1), ("NoRepair", 1)):
            winreg.SetValueEx(k, name, 0, winreg.REG_DWORD, val)
    return dict(target=str(target), exe=str(exe), shortcuts=[str(s) for s in shortcuts],
                uninstall=str(script), reg_key=reg_key, size_kb=size_kb)


def existing_install():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"{REG_BASE}\\{REG_NAME}") as k:
            return Path(winreg.QueryValueEx(k, "InstallLocation")[0])
    except OSError:
        return None


# ----------------------------------------------------------------------------- wizard
class Wizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title(f"Pasang {APP_NAME}")
        self.s = self.winfo_fpixels("1i") / 96.0
        px = self.px
        w, h = px(560), px(430)
        self.geometry(f"{w}x{h}+{max((self.winfo_screenwidth() - w) // 2, 0)}+{max((self.winfo_screenheight() - h) // 3, 0)}")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self._icon = None
        if ICON_PNG_B64:
            try:
                self._icon = tk.PhotoImage(data=ICON_PNG_B64)
                self.iconphoto(True, self._icon)
            except tk.TclError:
                pass
        for n in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            tkfont.nametofont(n).configure(family=FONT, size=10)
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("TButton", background="#eef1f6", bordercolor=C["line"], lightcolor="#eef1f6", darkcolor="#eef1f6",
                     padding=(px(14), px(6)), relief="flat", focuscolor="#eef1f6")
        st.map("TButton", background=[("disabled", "#f5f6f8"), ("active", "#e2e8f2")], foreground=[("disabled", "#a3adba")])
        st.configure("Accent.TButton", background=C["accent"], foreground="white", bordercolor=C["accent"],
                     lightcolor=C["accent"], darkcolor=C["accent"], padding=(px(26), px(9)), font=(FONT_BOLD, 11),
                     focuscolor=C["accent"])
        st.map("Accent.TButton", background=[("disabled", "#9db8f0"), ("active", C["accent_dark"])],
               foreground=[("disabled", "white")])
        st.configure("W.TCheckbutton", background=C["bg"], focuscolor=C["bg"])
        st.map("W.TCheckbutton", background=[("active", C["bg"])], indicatorbackground=[("selected", C["accent"]), ("!selected", "white")],
               indicatorforeground=[("selected", "white")])
        st.configure("TProgressbar", troughcolor="#dde3ec", background=C["accent"], bordercolor=C["bg"],
                     lightcolor=C["accent"], darkcolor=C["accent"])

        head = tk.Frame(self, bg=C["head"])
        head.pack(fill="x")
        hi = tk.Frame(head, bg=C["head"])
        hi.pack(fill="x", padx=px(22), pady=px(14))
        if self._icon:
            self._small = self._icon.subsample(2)
            tk.Label(hi, image=self._small, bg=C["head"]).pack(side="left", padx=(0, px(12)))
        tt = tk.Frame(hi, bg=C["head"])
        tt.pack(side="left")
        tk.Label(tt, text=f"Pasang {APP_NAME}", bg=C["head"], fg="white", font=(FONT_BOLD, 16)).pack(anchor="w")
        tk.Label(tt, text=f"Versi {VERSION}", bg=C["head"], fg="#c7d7fe", font=(FONT, 9)).pack(anchor="w")

        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=px(26), pady=px(18))
        old = existing_install()
        tk.Label(body, text="Perbarui versi yang sudah terpasang" if old else "Folder tujuan", bg=C["bg"], fg=C["text"],
                 font=(FONT_BOLD, 10)).pack(anchor="w")
        row = tk.Frame(body, bg=C["bg"])
        row.pack(fill="x", pady=(px(6), px(14)))
        self.parent = old.parent if old else default_parent()
        self.path_var = tk.StringVar(value=str(final_target(self.parent)))
        ttk.Entry(row, textvariable=self.path_var, state="readonly").pack(side="left", fill="x", expand=True, ipady=px(3))
        self.btn_change = ttk.Button(row, text="Ubah…", command=self.change_dir)
        self.btn_change.pack(side="left", padx=(px(8), 0))
        self.desk_var = tk.BooleanVar(value=True)
        self.open_var = tk.BooleanVar(value=True)
        self.chk1 = ttk.Checkbutton(body, text="Buat shortcut di Desktop", variable=self.desk_var, style="W.TCheckbutton")
        self.chk1.pack(anchor="w")
        self.chk2 = ttk.Checkbutton(body, text="Buka aplikasi setelah selesai", variable=self.open_var, style="W.TCheckbutton")
        self.chk2.pack(anchor="w", pady=(px(4), 0))
        tk.Label(body, text="Tidak perlu akses administrator. Bisa dihapus lewat Settings › Apps.", bg=C["bg"],
                 fg=C["muted"], font=(FONT, 9)).pack(anchor="w", pady=(px(12), 0))

        self.status = tk.StringVar()
        tk.Label(body, textvariable=self.status, bg=C["bg"], fg=C["text"], font=(FONT, 10), anchor="w",
                 justify="left", wraplength=px(500)).pack(fill="x", pady=(px(14), px(4)))
        self.bar = ttk.Progressbar(body, mode="determinate", maximum=100)
        self.btn = ttk.Button(body, text="Pasang", style="Accent.TButton", command=self.start)
        self.btn.pack(side="bottom", anchor="e")
        self.installing = False
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.deiconify()

    def px(self, n):
        return int(round(n * self.s))

    def on_close(self):
        if self.installing:
            return                              # never leave a half-copied install behind
        self.destroy()

    def change_dir(self):
        p = filedialog.askdirectory(title="Pilih folder tujuan (aplikasi dipasang di subfolder 'Pasangkan Foto')",
                                    initialdir=str(self.parent) if Path(self.parent).exists() else None)
        if p:
            self.parent = Path(os.path.normpath(p))
            self.path_var.set(str(final_target(self.parent)))

    def start(self, close_running=False):
        self.installing = True
        for w in (self.btn, self.btn_change):
            w.config(state="disabled")
        self.status.set("Memasang…")
        self.bar.pack(fill="x", pady=(0, self.px(8)))
        q_ = queue.Queue()
        parent, want_desktop = self.parent, self.desk_var.get()     # read Tk values here, never in the worker thread

        def work():
            try:
                q_.put(("ok", install(parent, desktop=want_desktop, close_running=close_running,
                                      progress=lambda d, t: q_.put(("p", d * 100 / t)))))
            except InstallError as e:
                q_.put(("err", str(e)))
            except BaseException as e:           # anything unexpected is shown, never swallowed
                q_.put(("err", f"Terjadi kesalahan tak terduga: {e!r}"))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            try:
                while True:
                    kind, val = q_.get_nowait()
                    if kind == "p":
                        self.bar["value"] = val
                    else:
                        return self.finished(kind, val)
            except queue.Empty:
                self.after(80, poll)

        self.after(80, poll)

    def finished(self, kind, val):
        self.installing = False
        if kind == "ok":
            self.result = val
            self.bar["value"] = 100
            self.status.set("✔  Terpasang. Shortcut ada di Start Menu" + (" dan Desktop." if len(val["shortcuts"]) > 1 else "."))
            self.btn.config(text="Selesai", state="normal", command=self.done)
            return
        self.bar.pack_forget()
        self.btn_change.config(state="normal")
        self.btn.config(state="normal")
        self.status.set("")
        if "sedang berjalan" in val and messagebox.askyesno(f"Pasang {APP_NAME}", val + "\n\nTutup aplikasi otomatis?"):
            return self.start(close_running=True)
        if "sedang berjalan" not in val:
            messagebox.showerror(f"Pasang {APP_NAME}", val)

    def done(self):
        if self.open_var.get() and self.result:
            try:
                subprocess.Popen([self.result["exe"]], cwd=self.result["target"], creationflags=DETACHED)
            except OSError:
                pass
        self.destroy()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--silent", action="store_true", help="pasang tanpa jendela")
    ap.add_argument("--dir", help="folder induk tujuan (aplikasi masuk ke subfolder 'Pasangkan Foto')")
    ap.add_argument("--no-desktop", action="store_true")
    ap.add_argument("--close-running", action="store_true")
    ap.add_argument("--log", help="tulis hasil (JSON) ke file ini")
    ap.add_argument("--start-menu-dir", help=argparse.SUPPRESS)       # test hooks
    ap.add_argument("--desktop-dir", help=argparse.SUPPRESS)
    ap.add_argument("--reg-base", default=REG_BASE, help=argparse.SUPPRESS)
    a = ap.parse_args()
    if a.silent:
        res = {"ok": False}
        try:
            res = {"ok": True, **install(a.dir or default_parent(), desktop=not a.no_desktop,
                                         close_running=a.close_running, start_menu=a.start_menu_dir,
                                         desktop_folder=a.desktop_dir, reg_base=a.reg_base)}
        except InstallError as e:
            res["error"] = str(e)
        except BaseException as e:
            res["error"] = repr(e)
        if a.log:
            Path(a.log).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        sys.exit(0 if res["ok"] else 1)
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    Wizard().mainloop()


if __name__ == "__main__":
    main()
