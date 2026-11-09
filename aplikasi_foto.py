"""Pasangkan Foto - Data Merge InDesign (Windows app, tkinter)."""
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

from cocokkan_foto import (DATE_TEMPLATES, PHOTO_EXT, RX_ISO_FULL, DataError, is_excel,
                           list_sheets, parse_date, pick_date_column, pick_name_column, proses,
                           read_table, template_example)

try:
    from ikon_data import ICON_PNG_B64
except ImportError:          # icon is optional
    ICON_PNG_B64 = ""

APP_NAME = "Pasangkan Foto"
APP_TITLE = "Pasangkan Foto – Data Merge InDesign"
NO_DATE = "(tidak dirapikan)"
FONT = "Segoe UI"
FONT_BOLD = "Segoe UI Semibold"

# ---- palette -------------------------------------------------------------
C = dict(
    bg="#f3f5f8", card="#ffffff", line="#dbe0e8", text="#1c2733", muted="#667487",
    head="#1e40af", accent="#2563eb", accent_dark="#1d4ed8", sel="#d6e4ff", status_bg="#e9edf3",
)
ROW_COLORS = {                     # only rows that need attention get a colour
    "CEK": "#fff4cc",
    "TIDAK ADA": "#fde4e1",
    "DIABAIKAN": "#eef0f3",
    "FOTO TIDAK TERPAKAI": "#e6eefb",
    "TGL": "#ffe8cf",              # photo OK but the date needs a look
}
ZEBRA = "#f7f9fc"
STATUS_LABEL = {
    "OK": "✔ OK", "CEK": "⚠ Cek", "TIDAK ADA": "✖ Tidak ada",
    "DIABAIKAN": "– Diabaikan", "FOTO TIDAK TERPAKAI": "◻ Tak terpakai",
}
BANNER = {                          # kind -> (background, border, text colour)
    "warn": ("#fdecea", "#f3b9b3", "#8a1c13"),
    "info": ("#e8f0fe", "#b6cbf7", "#1e3a8a"),
}
TGL_CHECK = ("RAGU", "GAGAL", "CURIGA", "DITUKAR")   # date statuses that need a human look
ALL_COLS = ("baris", "nama", "status", "file", "tgl_asli", "tgl_baru", "catatan")
DATE_COLS = ("tgl_asli", "tgl_baru")
MAX_UNUSED_ROWS = 60

# PASANGKAN_SETTINGS lets automated tests use a throw-away settings file instead of the user's real one.
SETTINGS_FILE = Path(os.environ.get("PASANGKAN_SETTINGS")
                     or Path(os.environ.get("APPDATA", str(Path.home()))) / "PasangkanFoto" / "settings.json")


def load_settings():
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}       # a damaged settings file must never stop the app


def clip(v, n):
    """Very long cell text freezes Tk when drawn; show only the first n characters."""
    v = str(v)
    return v if len(v) <= n else v[:n] + "…"


def save_settings(data):
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass                       # remembering folders is a convenience, never an error


def _preload_excel():
    """Importing openpyxl takes ~2 s; do it in the background right at start-up."""
    try:
        import openpyxl  # noqa: F401
        from openpyxl.reader import excel  # noqa: F401
    except Exception:
        pass


def short_path(p, n=64):
    p = os.path.normpath(p) if p else ""
    return p if len(p) <= n else "…" + p[-(n - 1):]


class DataView(tk.Toplevel):
    """Separate window: the whole sheet (all columns / rows) with search, sort and a detail panel."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        px = app.px
        self.title("Semua data")
        self.configure(bg=C["bg"])
        self.geometry(f"{px(1000)}x{px(620)}")
        if app._icon:
            self.iconphoto(False, app._icon)
        self.search_var = tk.StringVar()
        self.info = tk.StringVar()
        self.sort = None                     # (column id, descending)
        self.drows = {}

        card = ttk.Frame(self, style="Card.TFrame", padding=(px(12), px(10), px(12), px(10)))
        card.pack(fill="both", expand=True, padx=px(12), pady=px(12))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        bar = ttk.Frame(card, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, px(8)))
        ttk.Label(bar, text="Cari", style="Card.TLabel").pack(side="left", padx=(0, px(8)))
        ttk.Entry(bar, textvariable=self.search_var, width=30).pack(side="left")
        ttk.Button(bar, text="Hapus", command=lambda: self.search_var.set("")).pack(side="left", padx=px(6))
        ttk.Label(bar, textvariable=self.info, style="Muted.TLabel").pack(side="right")
        self.search_var.trace_add("write", lambda *a: self.refresh())

        body = ttk.Frame(card, style="Card.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(body, show="headings", selectmode="browse")
        vs = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        for tag, color in ROW_COLORS.items():
            self.tree.tag_configure(tag, background=color)
        self.tree.tag_configure("odd", background=ZEBRA)
        self.tree.tag_configure("even", background="white")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_detail())
        self.empty = tk.Label(body, text="Belum ada data", bg="white", fg=C["muted"], font=(FONT, 12))

        frame = tk.Frame(card, bg=C["line"])
        frame.grid(row=2, column=0, sticky="ew", pady=(px(8), 0))
        inner = tk.Frame(frame, bg="white")
        inner.pack(fill="x", padx=1, pady=1)
        self.detail = tk.Text(inner, height=5, wrap="word", state="disabled", font=(FONT, 10), relief="flat",
                              bg="white", fg=C["text"], padx=px(10), pady=px(6), borderwidth=0)
        dsb = ttk.Scrollbar(inner, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="x", expand=True)
        dsb.pack(side="right", fill="y")
        self.detail.tag_configure("h", font=(FONT_BOLD, 10), foreground=C["accent"])
        self.detail.tag_configure("k", font=(FONT_BOLD, 10), foreground=C["muted"])
        self.detail.tag_configure("hint", foreground=C["muted"])
        self.refresh()

    def refresh(self):
        px = self.app.px
        hdr, rows = self.app.data_source()
        self.tree.delete(*self.tree.get_children())
        self.drows = {}
        if hdr is None:
            self.tree["columns"] = ()
            self.info.set("")
            self.empty.place(relx=.5, rely=.4, anchor="center")
            self.show_detail()
            return
        self.empty.place_forget()
        cols = ["_baris"] + [f"c{i}" for i in range(len(hdr))]
        self.tree["displaycolumns"] = ("#all",)   # drop old column list first, else Tk rejects the new ids
        self.tree["columns"] = cols
        sort_col, desc = self.sort or (None, False)
        arrow = " ▼" if desc else " ▲"
        self.tree.heading("_baris", text="Baris" + arrow * (sort_col == "_baris"), anchor="center",
                          command=lambda: self.sort_data("_baris"))
        self.tree.column("_baris", width=px(60), stretch=False, anchor="center")
        sample = rows[:200]
        for i, h in enumerate(hdr):
            cid = f"c{i}"
            longest = max([len(h)] + [len(str(r[1][i])) for r in sample])
            self.tree.heading(cid, text=h + arrow * (sort_col == cid), anchor="w",
                              command=lambda c=cid: self.sort_data(c))
            self.tree.column(cid, width=min(max(px(longest * 7 + 24), px(80)), px(340)),
                             stretch=False, anchor="w")

        if sort_col:
            def key(item):
                v = item[0] if sort_col == "_baris" else item[1][int(sort_col[1:])]
                try:
                    return (0, float(v))                      # numbers
                except (TypeError, ValueError):
                    pass
                text = str(v).strip()
                p = parse_date(text)
                if p and p[3][1] - p[3][0] >= 0.8 * len(text):  # cell is (almost only) a date
                    return (1, (p[0], p[1], p[2]))            # chronological
                return (2, text.lower())
            rows = sorted(rows, key=key, reverse=desc)
        q = self.search_var.get().strip().lower()
        shown = 0
        for n, cells, tag in rows:
            if q and not (q in str(n) or any(q in str(c).lower() for c in cells)):
                continue
            self.drows[str(n)] = cells
            use = tag if tag in ROW_COLORS else ("odd" if shown % 2 else "even")
            self.tree.insert("", "end", iid=str(n), tags=(use,),
                             values=[n] + [clip(str(c).replace("\n", " ⏎ "), 300) for c in cells])
            shown += 1
        self.info.set(f"{shown} dari {len(rows)} baris × {len(hdr)} kolom")
        self.show_detail()

    def sort_data(self, cid):
        col, desc = self.sort or (None, False)
        self.sort = (cid, not desc if col == cid else False)
        self.refresh()

    def show_detail(self):
        hdr, _ = self.app.data_source()
        sel = self.tree.selection()
        self.detail.config(state="normal")
        self.detail.delete("1.0", "end")
        if sel and hdr and sel[0] in self.drows:
            self.detail.insert("end", f"Baris {sel[0]}\n", "h")
            for h, v in zip(hdr, self.drows[sel[0]]):
                if str(v).strip():
                    self.detail.insert("end", f"{h}:  ", "k")
                    self.detail.insert("end", clip(v, 5000) + "\n")
        else:
            self.detail.insert("end", "Pilih satu baris untuk melihat isi lengkapnya", "hint")
        self.detail.config(state="disabled")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()           # stay invisible until fully built: a half-built window looks "Not Responding"
        self.title(APP_TITLE)
        self.s = self.winfo_fpixels("1i") / 96.0          # screen scaling (100% = 1.0)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(self.px(1040), int(sw * 0.96)), min(self.px(680), int(sh * 0.92))
        self._size = (w, h)
        self.geometry(f"{w}x{h}")
        self.minsize(min(self.px(860), w), min(self.px(520), h))
        self.settings = load_settings()

        self._icon = None
        if ICON_PNG_B64:
            try:
                self._icon = tk.PhotoImage(data=ICON_PNG_B64)
                self.iconphoto(True, self._icon)
            except tk.TclError:
                self._icon = None

        self.csv_var = tk.StringVar()
        self.foto_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.kolom_var = tk.StringVar()
        self.tgl_kolom_var = tk.StringVar(value=NO_DATE)
        self.tgl_fmt_var = tk.StringVar()
        self.tukar_var = tk.BooleanVar(value=False)       # shown only inside the date warning banner
        self.filter_var = tk.BooleanVar(value=False)
        self.warn_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Siap")
        self.empty_var = tk.StringVar()
        self.raw = None           # (header, rows, row numbers) of the loaded sheet, before processing
        self.forced = {}          # name.lower() -> full photo path (manual photo corrections)
        self.tgl_manual = {}      # file row number -> typed date text (manual date edits)
        self.result = None
        self.rows = {}            # tree iid -> report dict
        self.data_win = None
        self._jobs = 0            # background jobs in flight
        self._gen = 0             # bumped whenever another file / sheet is chosen (stale results are dropped)
        self._loading = False     # the data file is being read
        self._running = False     # a processing run is in flight
        self._run_pending = False # something changed during a run: run once more afterwards
        self.cols = []

        self._style()
        self._build()
        self.csv_var.trace_add("write", lambda *a: self._paths_changed())
        self.foto_var.trace_add("write", lambda *a: self._paths_changed())
        self._paths_changed()
        self.refresh_tree()
        self.set_banner("")
        self.reset_chips()
        self.update_buttons()
        self.bind_all("<Control-Return>", lambda e: self.run())
        self.bind_all("<Control-o>", lambda e: self.pick_csv())
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._center()
        self.deiconify()
        # load the Excel reader (~2 s) only once the window is up, so it never competes with building the UI
        self.after(300, lambda: threading.Thread(target=_preload_excel, daemon=True).start())

    # ---------- helpers ----------
    def px(self, n):
        return int(round(n * self.s))

    def _center(self):
        w, h = self._size          # (a hidden window reports a width of 1, so use the size we asked for)
        x = max((self.winfo_screenwidth() - w) // 2, 0)
        y = max((self.winfo_screenheight() - h) // 3, 0)
        self.geometry(f"+{x}+{y}")

    def _close(self):
        self._save_settings()
        self.destroy()

    def _save_settings(self):
        self.settings["template"] = max(self.tpl_box.current(), 0)
        save_settings(self.settings)

    def ready(self):
        return bool(self.csv_var.get().strip() and self.foto_var.get().strip())

    def _paths_changed(self):
        f, d = self.csv_var.get().strip(), self.foto_var.get().strip()
        self.file_lbl.config(text=short_path(f) if f else "Belum dipilih", fg=C["text"] if f else C["muted"])
        self.foto_lbl.config(text=short_path(d) if d else "Belum dipilih", fg=C["text"] if d else C["muted"])
        self.btn_run.config(state="normal" if self.ready() and not self._jobs else "disabled")

    # ---------- background jobs ----------
    def _bump(self):
        self._gen += 1
        return self._gen

    def _async(self, work, done):
        """Run work() in a thread so the window never freezes; done(ok, value) runs back on the UI thread."""
        self._jobs += 1
        self._busy_ui(True)
        q = queue.Queue()

        def target():
            try:
                q.put((True, work()))
            except BaseException as e:           # reported to the user in done()
                q.put((False, e))

        threading.Thread(target=target, daemon=True).start()

        def poll():
            try:
                ok, val = q.get_nowait()
            except queue.Empty:
                self.after(50, poll)
                return
            self._jobs -= 1
            if not self._jobs:
                self._busy_ui(False)
            done(ok, val)

        self.after(50, poll)

    def _busy_ui(self, busy):
        if busy:
            self.config(cursor="watch")
            self.btn_run.config(state="disabled")
            self.status_var.set("Memproses…")
            self.progress.pack(side="right", padx=self.px(16), pady=self.px(4))
            self.progress.start(14)
        else:
            self.config(cursor="")
            self.progress.stop()
            self.progress.pack_forget()
            self._paths_changed()

    def wait_idle(self, timeout=180):
        """Process events until all background work is done (used by the automated tests)."""
        t0 = time.time()
        while (self._jobs or self._run_pending) and time.time() - t0 < timeout:
            self.update()
            time.sleep(0.01)
        self.update()

    # ---------- styling ----------
    def _style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            tkfont.nametofont(name).configure(family=FONT, size=10)
        self.option_add("*TCombobox*Listbox.font", (FONT, 10))
        self.option_add("*TCombobox*Listbox.selectBackground", C["accent"])
        self.configure(bg=C["bg"])
        card, line = C["card"], C["line"]

        st.configure(".", background=C["bg"], foreground=C["text"], font=(FONT, 10), bordercolor=line)
        st.configure("TFrame", background=C["bg"])
        st.configure("Card.TFrame", background=card)
        st.configure("TLabel", background=C["bg"], foreground=C["text"])
        st.configure("Card.TLabel", background=card)
        st.configure("Muted.TLabel", background=card, foreground=C["muted"], font=(FONT, 9))

        st.configure("TEntry", fieldbackground="white", bordercolor=line, lightcolor=line, darkcolor=line, padding=6)
        st.map("TEntry", bordercolor=[("focus", C["accent"])], lightcolor=[("focus", C["accent"])],
               darkcolor=[("focus", C["accent"])])
        st.configure("TCombobox", fieldbackground="white", background="#eef1f6", bordercolor=line,
                     lightcolor=line, darkcolor=line, arrowcolor=C["muted"], padding=5)
        st.map("TCombobox", fieldbackground=[("readonly", "white"), ("disabled", "#eef0f3")],
               selectbackground=[("readonly", "white")], selectforeground=[("readonly", C["text"])],
               foreground=[("disabled", "#9aa6b6")], bordercolor=[("focus", C["accent"])])

        btn_bg = "#eef1f6"
        st.configure("TButton", background=btn_bg, foreground=C["text"], bordercolor=line, lightcolor=btn_bg,
                     darkcolor=btn_bg, padding=(self.px(14), self.px(6)), relief="flat", focuscolor=btn_bg)
        bg_map = [("disabled", "#f5f6f8"), ("pressed", "#d6deea"), ("active", "#e2e8f2")]
        st.map("TButton", background=bg_map, lightcolor=bg_map, darkcolor=bg_map,
               foreground=[("disabled", "#a3adba")], bordercolor=[("disabled", "#e6e9ef")])
        acc = C["accent"]
        st.configure("Accent.TButton", background=acc, foreground="white", bordercolor=acc, lightcolor=acc,
                     darkcolor=acc, padding=(self.px(16), self.px(6)), font=(FONT_BOLD, 10), focuscolor=acc)
        acc_map = [("disabled", "#9db8f0"), ("pressed", "#1e40af"), ("active", C["accent_dark"])]
        st.map("Accent.TButton", background=acc_map, lightcolor=acc_map, darkcolor=acc_map,
               foreground=[("disabled", "white")], bordercolor=acc_map)

        st.configure("Card.TCheckbutton", background=card, foreground=C["text"], focuscolor=card)
        st.map("Card.TCheckbutton", background=[("active", card)],
               indicatorbackground=[("selected", acc), ("!selected", "white")],
               indicatorforeground=[("selected", "white")],
               upperbordercolor=[("selected", acc), ("!selected", "#9aa6b6")],
               lowerbordercolor=[("selected", acc), ("!selected", "#9aa6b6")])

        st.configure("Treeview", rowheight=self.px(30), background="white", fieldbackground="white",
                     foreground=C["text"], borderwidth=0, font=(FONT, 10))
        st.map("Treeview", background=[("selected", C["sel"])], foreground=[("selected", C["text"])])
        st.configure("Treeview.Heading", background="#eef1f6", foreground=C["text"], relief="flat",
                     font=(FONT_BOLD, 10), padding=(self.px(8), self.px(8)), borderwidth=0)
        st.map("Treeview.Heading", background=[("active", "#e2e8f2")])
        st.configure("TScrollbar", background="#c9d0da", troughcolor=card, bordercolor=card,
                     arrowcolor=C["muted"], relief="flat", arrowsize=self.px(14))
        st.map("TScrollbar", background=[("active", "#aab4c3")])

    # ---------- layout ----------
    def _build(self):
        pad = self.px(16)

        header = tk.Frame(self, bg=C["head"])
        header.pack(fill="x")
        hi = tk.Frame(header, bg=C["head"])
        hi.pack(fill="x", padx=pad, pady=self.px(7))
        if self._icon:
            self._icon_small = self._icon.subsample(2)
            tk.Label(hi, image=self._icon_small, bg=C["head"]).pack(side="left", padx=(0, self.px(10)))
        tk.Label(hi, text=APP_NAME, bg=C["head"], fg="white", font=(FONT_BOLD, 15)).pack(side="left")

        sb = tk.Frame(self, bg=C["status_bg"])           # packed before the body so it never gets squeezed out
        sb.pack(fill="x", side="bottom")
        tk.Label(sb, textvariable=self.status_var, bg=C["status_bg"], fg=C["muted"], anchor="w",
                 font=(FONT, 9)).pack(side="left", padx=pad, pady=self.px(4))
        self.progress = ttk.Progressbar(sb, mode="indeterminate", length=self.px(150))   # shown only while busy

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=pad, pady=(self.px(12), 0))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1, minsize=self.px(240))

        # -- row 0: two pickers
        outer = tk.Frame(body, bg=C["line"])
        outer.grid(row=0, column=0, sticky="ew")
        f = ttk.Frame(outer, style="Card.TFrame", padding=(self.px(16), self.px(12), self.px(16), self.px(12)))
        f.pack(fill="both", expand=True, padx=1, pady=1)
        f.columnconfigure(1, weight=1)
        vp = dict(pady=self.px(4))

        ttk.Button(f, text="Pilih file data…", width=18, command=self.pick_csv).grid(row=0, column=0, sticky="w", **vp)
        self.file_lbl = tk.Label(f, bg=C["card"], anchor="w", font=(FONT, 10))
        self.file_lbl.grid(row=0, column=1, sticky="ew", padx=self.px(12))
        ttk.Label(f, text="Sheet", style="Card.TLabel").grid(row=0, column=2, sticky="e", padx=(0, self.px(8)))
        self.sheet_box = ttk.Combobox(f, textvariable=self.sheet_var, state="disabled", width=16)
        self.sheet_box.grid(row=0, column=3, sticky="w", **vp)
        self.sheet_box.bind("<<ComboboxSelected>>", lambda ev: self.change_sheet())

        ttk.Button(f, text="Pilih folder foto…", width=18, command=self.pick_foto).grid(row=1, column=0, sticky="w", **vp)
        self.foto_lbl = tk.Label(f, bg=C["card"], anchor="w", font=(FONT, 10))
        self.foto_lbl.grid(row=1, column=1, sticky="ew", padx=self.px(12))
        ttk.Label(f, text="Format tanggal", style="Card.TLabel").grid(row=1, column=2, sticky="e", padx=(self.px(12), self.px(8)))
        fmt = ttk.Frame(f, style="Card.TFrame")
        fmt.grid(row=1, column=3, sticky="w", **vp)
        self.tpl_box = ttk.Combobox(fmt, textvariable=self.tgl_fmt_var, state="readonly", width=16,
                                    values=[template_example(t) for t in DATE_TEMPLATES])
        try:
            tpl = min(max(int(self.settings.get("template", 0)), 0), len(DATE_TEMPLATES) - 1)
        except (TypeError, ValueError):
            tpl = 0                                  # damaged settings: fall back to the default format
        self.tpl_box.current(tpl)
        self.tpl_box.pack(side="left")
        self.tpl_box.bind("<<ComboboxSelected>>", lambda ev: self.rerun_if_ready())
        self.btn_cols = ttk.Button(fmt, text="⚙", width=3, command=self.columns_dialog)
        self.btn_cols.pack(side="left", padx=(self.px(6), 0))

        # -- row 1: counters + reprocess / open result
        act = ttk.Frame(body)
        act.grid(row=1, column=0, sticky="ew", pady=(self.px(12), 0))
        self.chip_vals = {}
        for key, caption, fg, bg in (("ok", "Cocok", "#1a7f37", "#eaf6ee"), ("cek", "Cek", "#a16207", "#fff6d6"),
                                     ("tdk", "Tanpa foto", "#b3261e", "#fde8e6"),
                                     ("tcek", "Tanggal cek", "#c2410c", "#ffeedd")):
            chip = tk.Frame(act, bg=bg)
            chip.pack(side="left", padx=(0, self.px(8)), ipady=self.px(5))
            tk.Frame(chip, bg=fg, width=self.px(4)).pack(side="left", fill="y")
            v = tk.Label(chip, text="–", bg=bg, fg=fg, font=(FONT_BOLD, 13))
            v.pack(side="left", padx=(self.px(10), self.px(5)))
            tk.Label(chip, text=caption, bg=bg, fg=C["muted"], font=(FONT, 10)).pack(side="left", padx=(0, self.px(12)))
            self.chip_vals[key] = v
        self.btn_open = ttk.Button(act, text="Buka hasil", style="Accent.TButton", command=self.open_out, state="disabled")
        self.btn_open.pack(side="right")
        self.btn_run = ttk.Button(act, text="↻  Proses ulang", command=self.run, state="disabled")
        self.btn_run.pack(side="right", padx=(0, self.px(8)))

        # -- row 2: banner (only for the date warning)
        self.banner = tk.Frame(body)
        self.banner.grid(row=2, column=0, sticky="ew", pady=(self.px(10), 0))
        self.banner_in = tk.Frame(self.banner)
        self.banner_in.pack(fill="x", padx=1, pady=1)
        self.banner_lbl = tk.Label(self.banner_in, textvariable=self.warn_var, justify="left", anchor="w",
                                   padx=self.px(12), pady=self.px(8), font=(FONT, 10))
        self.banner_lbl.pack(side="left", fill="x", expand=True)
        self.banner_chk = tk.Checkbutton(self.banner_in, text="Tukar hari↔bulan", variable=self.tukar_var,
                                         command=self.rerun_if_ready, font=(FONT_BOLD, 10), relief="flat",
                                         bd=0, highlightthickness=0, cursor="hand2", padx=self.px(12))
        self.banner_chk.pack(side="right")          # shown/hidden by set_banner()

        # -- row 3: results
        o2 = tk.Frame(body, bg=C["line"])
        o2.grid(row=3, column=0, sticky="nsew", pady=(self.px(10), self.px(10)))
        tab = ttk.Frame(o2, style="Card.TFrame", padding=(self.px(12), self.px(10), self.px(12), self.px(10)))
        tab.pack(fill="both", expand=True, padx=1, pady=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        bar = ttk.Frame(tab, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, self.px(8)))
        self.btn_manual = ttk.Button(bar, text="Ganti foto", command=self.manual_pick, state="disabled")
        self.btn_manual.pack(side="left")
        self.btn_tgl = ttk.Button(bar, text="Edit tanggal", command=self.edit_date, state="disabled")
        self.btn_tgl.pack(side="left", padx=self.px(6))
        self.btn_undo = ttk.Button(bar, text="Reset", command=self.undo_manual, state="disabled")
        self.btn_undo.pack(side="left")
        ttk.Button(bar, text="Semua data", command=self.open_data).pack(side="right")
        ttk.Checkbutton(bar, text="Hanya yang perlu dicek", variable=self.filter_var, style="Card.TCheckbutton",
                        command=self.refresh_tree).pack(side="right", padx=(0, self.px(14)))

        holder = ttk.Frame(tab, style="Card.TFrame")
        holder.grid(row=1, column=0, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(holder, columns=ALL_COLS, show="headings", selectmode="browse")
        for c, w, t, anchor in (("baris", 54, "Baris", "center"), ("nama", 190, "Nama", "w"),
                                ("status", 104, "Foto", "w"), ("file", 210, "File foto", "w"),
                                ("tgl_asli", 138, "Tanggal asli", "w"), ("tgl_baru", 118, "Tanggal baru", "w"),
                                ("catatan", 150, "Catatan", "w")):
            self.tree.heading(c, text=t, anchor="w" if anchor == "w" else "center")
            self.tree.column(c, width=self.px(w), anchor=anchor, stretch=(c in ("nama", "file", "catatan")))
        for tag, color in ROW_COLORS.items():
            self.tree.tag_configure(tag, background=color)
        self.tree.tag_configure("odd", background=ZEBRA)
        self.tree.tag_configure("even", background="white")
        vs = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        self.empty = tk.Label(holder, textvariable=self.empty_var, bg="white", fg=C["muted"],
                              font=(FONT, 12), justify="center")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_buttons())
        self.tree.bind("<Double-1>", self.on_double)

    # ---------- counters / banner ----------
    def reset_chips(self):
        for v in self.chip_vals.values():
            v.config(text="–")

    def set_chips(self, c, t, has_date):
        self.chip_vals["ok"].config(text=str(c["OK"]))
        self.chip_vals["cek"].config(text=str(c["CEK"]))
        self.chip_vals["tdk"].config(text=str(c["TIDAK ADA"]))
        self.chip_vals["tcek"].config(
            text=str(t["RAGU"] + t["GAGAL"] + t["CURIGA"] + t["DITUKAR"]) if has_date else "–")

    def set_banner(self, text, kind="warn", swap=False):
        """Empty text hides the banner. The 'Tukar hari↔bulan' checkbox lives inside it (swap=True)."""
        self.warn_var.set(text)
        if not text:
            self.banner.grid_remove()
            return
        if swap:
            self.banner_chk.pack(side="right")
        else:
            self.banner_chk.pack_forget()
        bg, border, fg = BANNER[kind]
        self.banner.config(bg=border)
        self.banner_in.config(bg=bg)
        self.banner_lbl.config(bg=bg, fg=fg)
        self.banner_chk.config(bg=bg, fg=fg, activebackground=bg, activeforeground=fg, selectcolor="white")
        self.banner.grid()

    # ---------- all-data window ----------
    def data_source(self):
        """(header, [(row number, cells, colour tag)]): processed result if present, else raw sheet."""
        if self.result:
            rows = [(n, cells, "TGL" if st == "OK" and ts in TGL_CHECK else st)
                    for n, cells, st, ts in self.result["data_rows"]]
            return self.result["data_header"], rows
        if self.raw:
            hdr, rows, nums = self.raw
            return hdr, [(n, r, "OK") for n, r in zip(nums, rows)]
        return None, []

    def open_data(self):
        if self.data_win is not None and self.data_win.winfo_exists():
            self.data_win.lift()
            self.data_win.refresh()
        else:
            self.data_win = DataView(self)
        return self.data_win

    def refresh_data(self):
        if self.data_win is not None and self.data_win.winfo_exists():
            self.data_win.refresh()

    # ---------- inputs ----------
    def pick_csv(self):
        p = filedialog.askopenfilename(
            title="Pilih file data", initialdir=self.settings.get("data_dir") or None,
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv *.txt"), ("Semua file", "*.*")])
        if not p:
            return
        self.settings["data_dir"] = str(Path(p).parent)
        self.csv_var.set(os.path.normpath(p))
        self.load_file()

    def pick_foto(self):
        p = filedialog.askdirectory(title="Pilih folder foto", initialdir=self.settings.get("foto_dir") or None)
        if p:
            self.settings["foto_dir"] = p
            self.foto_var.set(os.path.normpath(p))
            self.auto_run()

    def auto_run(self):
        """Results appear by themselves as soon as both the data file and the photo folder are chosen."""
        if self.ready():
            self.run()

    def load_file(self):
        """New data file chosen: reset corrections, then read it in the background."""
        self.forced.clear()
        self.tgl_manual.clear()
        self.result = None
        self.raw = None
        self.set_banner("")
        self.reset_chips()
        self.status_var.set("Siap")
        self.refresh_tree()
        self.refresh_data()
        self.sheet_var.set("")
        self.sheet_box.config(state="disabled", values=[])
        self.btn_open.config(state="disabled")
        path = self.csv_var.get().strip()
        if path:
            self._read_data(path, None)

    def change_sheet(self):
        self.forced.clear()
        self.tgl_manual.clear()
        self.result = None                 # old sheet's result must not be shown for the new sheet
        self.set_banner("")
        self.reset_chips()
        self.refresh_tree()
        self._read_data(self.csv_var.get().strip(), self.sheet_var.get() or None)

    def _read_data(self, path, sheet):
        """Read sheet names + table in a worker thread, then detect the columns and process."""
        gen = self._bump()
        self._loading = True

        def work():
            names = list_sheets(path) if is_excel(path) else None
            # No sheet chosen yet: take the first one that has data (a cover / empty first sheet is common).
            candidates = [sheet] if sheet else (names or [None])
            error = None
            for use in candidates:
                try:
                    header, data, nums = read_table(path, use, with_rows=True)
                    return names, (use, header, data, nums), None
                except DataError as e:
                    error = e
            return names, None, (error, candidates[0])

        self._async(work, lambda ok, val: self._data_read(gen, ok, val))

    def _data_read(self, gen, ok, val):
        if gen != self._gen:
            return                          # another file / sheet was chosen meanwhile
        self._loading = False
        if not ok:
            self.status_var.set("Gagal")
            messagebox.showerror(APP_TITLE, f"File data tidak bisa dibaca:\n{val}")
            return
        names, loaded, failed = val
        if names is not None:                # always offer the sheet list, even if this sheet is unusable
            self.sheet_box.config(state="readonly", values=names)
            self.sheet_var.set((loaded[0] if loaded else failed[1]) or "")
        else:
            self.sheet_box.config(state="disabled", values=[])
        if loaded is None:
            self.status_var.set("Gagal")
            messagebox.showerror(APP_TITLE, f"File data tidak bisa dibaca:\n{failed[0]}"
                                 + ("\n\nCoba pilih sheet lain." if names and len(names) > 1 else ""))
            return
        sheet, header, data, nums = loaded
        self.raw = (header, [r + [""] * (len(header) - len(r)) for r in data], nums)
        self.cols = header
        self.refresh_data()
        di = pick_date_column(header)
        self.tgl_kolom_var.set(header[di] if di is not None else NO_DATE)
        try:
            self.kolom_var.set(header[pick_name_column(header)])
        except DataError:
            self.kolom_var.set(header[0])
            self.columns_dialog()                   # could not tell which column holds the names: ask
        self.auto_run()

    def columns_dialog(self):
        """Small dialog to change which columns hold the name and the birth date."""
        if not getattr(self, "cols", None):
            return
        win = tk.Toplevel(self)
        win.title("Kolom")
        win.configure(bg=C["bg"])
        win.transient(self)
        win.resizable(False, False)
        if self._icon:
            win.iconphoto(False, self._icon)
        f = ttk.Frame(win, style="Card.TFrame", padding=self.px(18))
        f.pack(padx=self.px(14), pady=self.px(14))
        ttk.Label(f, text="Kolom nama", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=self.px(5))
        nama = tk.StringVar(value=self.kolom_var.get())
        ttk.Combobox(f, textvariable=nama, state="readonly", values=self.cols, width=28).grid(
            row=0, column=1, padx=(self.px(12), 0))
        ttk.Label(f, text="Kolom tanggal lahir", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=self.px(5))
        tgl = tk.StringVar(value=self.tgl_kolom_var.get())
        ttk.Combobox(f, textvariable=tgl, state="readonly", values=[NO_DATE] + self.cols, width=28).grid(
            row=1, column=1, padx=(self.px(12), 0))

        def ok():
            self.kolom_var.set(nama.get())
            self.tgl_kolom_var.set(tgl.get())
            win.destroy()
            self.rerun_if_ready()

        ttk.Button(f, text="OK", style="Accent.TButton", command=ok).grid(
            row=2, column=1, sticky="e", pady=(self.px(14), 0))
        win.grab_set()
        self.wait_window(win)

    def date_column(self):
        v = self.tgl_kolom_var.get()
        return None if v in ("", NO_DATE) else v

    # ---------- processing ----------
    def rerun_if_ready(self):
        if self.result:
            self.run()

    def run(self):
        csv_path, foto = self.csv_var.get().strip(), self.foto_var.get().strip()
        if not csv_path or not foto:
            messagebox.showwarning(APP_TITLE, "Pilih file data dan folder foto dulu.")
            return
        if self._loading:
            return                          # the file is still being read; it starts the run itself
        if self._running:
            self._run_pending = True        # run once more when the current run is done
            return
        self._running = True
        gen = self._gen
        idx = max(self.tpl_box.current(), 0)
        kw = dict(csv_path=csv_path, foto_dir=foto, kolom=self.kolom_var.get() or None,
                  forced=dict(self.forced), isi_ragu=False, subfolder=True, utf16=True,
                  kolom_tanggal=self.date_column(), format_tanggal=DATE_TEMPLATES[idx],
                  tgl_manual=dict(self.tgl_manual), sheet=self.sheet_var.get() or None,
                  tukar_hari_bulan=self.tukar_var.get())
        self._async(lambda: proses(**kw), lambda ok, val: self._run_done(gen, ok, val))

    def _run_done(self, gen, ok, val):
        self._running = False
        try:
            if gen != self._gen:
                return                      # result of a file / sheet that is no longer shown
            if not ok:
                self.status_var.set("Gagal")
                if isinstance(val, (DataError, OSError)):
                    messagebox.showerror(APP_TITLE, str(val))
                else:
                    messagebox.showerror(APP_TITLE, f"Terjadi kesalahan tak terduga:\n{val!r}")
                return
            self._apply_result(val)
        finally:
            if self._run_pending:
                self._run_pending = False
                self.run()

    def _apply_result(self, result):
        self.result = result
        c, t = result["counts"], result["tgl_counts"]
        has_date = bool(self.date_column())
        self.set_chips(c, t, has_date)

        swap = has_date and bool(result["tgl_suspect"])
        if has_date and t["CURIGA"]:
            self.set_banner(f"⚠  {t['CURIGA']} tanggal KEMUNGKINAN TERTUKAR hari↔bulan (baris oranye). "
                            "Cek ke database dulu.", "warn", swap)
        elif has_date and t["DITUKAR"]:
            self.set_banner(f"ℹ  {t['DITUKAR']} tanggal ditukar hari↔bulan (baris oranye).", "info", swap)
        else:
            self.set_banner("", "warn", swap)

        self.status_var.set(f"Tersimpan: {result['merge_path']}")
        self.btn_open.config(state="normal")
        self._save_settings()
        self.refresh_tree()
        self.refresh_data()

    @staticmethod
    def needs_check(r):
        return r["status"] != "OK" or r["tgl_status"] in TGL_CHECK

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        if not self.result:
            f, d = self.csv_var.get().strip(), self.foto_var.get().strip()
            self.empty_var.set("Pilih file data dan folder foto\nHasil muncul otomatis" if not f and not d
                               else "Pilih folder foto" if f else "Pilih file data")
            self.empty.place(relx=.5, rely=.4, anchor="center")
            self.update_buttons()
            return
        has_date = bool(self.date_column())
        self.tree["displaycolumns"] = [c for c in ALL_COLS if has_date or c not in DATE_COLS]
        only_todo = self.filter_var.get()
        shown = unused_shown = unused_hidden = 0
        for i, r in enumerate(self.result["report"]):
            if only_todo and not self.needs_check(r):
                continue
            if r["status"] == "FOTO TIDAK TERPAKAI":          # parent folders can hold thousands of them
                if unused_shown >= MAX_UNUSED_ROWS:
                    unused_hidden += 1
                    continue
                unused_shown += 1
            notes = []
            if r["catatan"] == "manual":
                notes.append("foto manual")
            elif r["status"] == "DIABAIKAN":
                notes.append("bukan siswa?")
            ts = r["tgl_status"]
            if ts == "GAGAL":
                notes.append("tanggal tidak terbaca")
            elif ts == "RAGU":
                notes.append("tanggal ragu")
            elif ts == "CURIGA":
                notes.append("hari/bulan tertukar?")
            elif ts == "DITUKAR":
                notes.append("hari↔bulan ditukar")
            elif ts == "MANUAL":
                notes.append("tanggal manual")
            if r["status"] == "OK" and ts in TGL_CHECK:
                tag = "TGL"
            elif r["status"] == "OK":
                tag = "odd" if shown % 2 else "even"
            else:
                tag = r["status"]
            asli = r["tgl_asli"]
            if RX_ISO_FULL.fullmatch(asli or ""):
                asli = f"{asli}  (Excel)"
            iid = str(i)
            self.tree.insert("", "end", iid=iid, tags=(tag,), values=(
                r["baris"], clip(r["nama"], 120), STATUS_LABEL.get(r["status"], r["status"]), r["file"],
                clip(asli, 60), clip(r["tgl_baru"], 60), "; ".join(notes)))
            self.rows[iid] = r
            shown += 1
        if unused_hidden:
            self.tree.insert("", "end", tags=("FOTO TIDAK TERPAKAI",), values=(
                "", f"+ {unused_hidden} foto lain tidak terpakai", STATUS_LABEL["FOTO TIDAK TERPAKAI"], "", "", "", ""))
        if shown:
            self.empty.place_forget()
        else:
            self.empty_var.set("✔  Semua beres")
            self.empty.place(relx=.5, rely=.4, anchor="center")
        self.update_buttons()

    def selected(self):
        sel = self.tree.selection()
        return self.rows.get(sel[0]) if sel else None

    def update_buttons(self):
        r = self.selected()
        person = bool(r and r["nama"] != "")
        self.btn_manual.config(state="normal" if person else "disabled")
        self.btn_tgl.config(state="normal" if person and self.date_column() else "disabled")
        has_fix = person and (r["nama"].strip().lower() in self.forced or r["baris"] in self.tgl_manual)
        self.btn_undo.config(state="normal" if has_fix else "disabled")

    def on_double(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        col = self.tree.column(self.tree.identify_column(event.x), "id")
        if col in DATE_COLS and self.date_column():
            self.edit_date()
        else:
            self.view_photo()

    # ---------- manual corrections ----------
    def manual_pick(self):
        r = self.selected()
        if not r or not r["nama"]:
            return
        foto = self.foto_var.get().strip()
        p = filedialog.askopenfilename(
            title=f"Foto untuk: {r['nama']}", initialdir=foto,
            filetypes=[("Foto", " ".join("*" + e for e in sorted(PHOTO_EXT)))])
        if not p:
            return
        p = str(Path(p).resolve())
        if not p.lower().startswith(str(Path(foto).resolve()).lower()):
            messagebox.showwarning(APP_TITLE, "Foto harus berada di dalam folder foto yang dipilih.")
            return
        self.forced[r["nama"].strip().lower()] = p
        self.run()

    def edit_date(self):
        r = self.selected()
        if not r or not r["nama"] or not self.date_column():
            return
        prompt = f"Tanggal lahir: {r['nama']}\n(contoh: 21 Agustus 2000)"
        init = r["tgl_baru"]
        while True:
            s = simpledialog.askstring(APP_TITLE, prompt, initialvalue=init, parent=self)
            if s is None:
                return
            if parse_date(s):
                break
            messagebox.showwarning(APP_TITLE, "Tanggal tidak terbaca.\nContoh: 21 Agustus 2000 atau 21-08-2000")
            init = s
        self.tgl_manual[r["baris"]] = s
        self.run()

    def undo_manual(self):
        r = self.selected()
        if not r:
            return
        a = self.forced.pop(r["nama"].strip().lower(), None)
        b = self.tgl_manual.pop(r["baris"], None)
        if a is not None or b is not None:
            self.run()

    # ---------- misc ----------
    def view_photo(self):
        r = self.selected()
        if r and r["path"] and Path(r["path"]).is_file():
            os.startfile(r["path"])

    def open_out(self):
        if self.result:
            os.startfile(self.result["outdir"])


def selftest(data, foto, out):
    """Headless check of the packaged exe:  PasangkanFoto.exe --selftest <data> <foto folder> <result.json>
    Reads the data file, matches the photos, writes the merge file and reports what happened."""
    res = {"ok": False}
    try:
        names = list_sheets(data) if is_excel(data) else None
        header, _ = read_table(data, names[0] if names else None)
        di = pick_date_column(header)
        r = proses(data, foto, kolom_tanggal=header[di] if di is not None else None,
                   keluar=Path(out).parent / "selftest_hasil")
        merge = Path(r["merge_path"])
        res.update(ok=True, sheets=names, counts=r["counts"], tgl=r["tgl_counts"], n_photos=r["n_photos"],
                   merge_file=merge.name, merge_utf16=merge.read_bytes()[:2] in (b"\xff\xfe", b"\xfe\xff"),
                   frozen=bool(getattr(sys, "frozen", False)))
    except BaseException as e:           # report everything: this is how a packaging problem shows up
        import traceback
        res.update(error=repr(e), trace=traceback.format_exc()[-1200:])
    Path(out).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0 if res["ok"] else 1


def main():
    if len(sys.argv) == 5 and sys.argv[1] == "--selftest":
        sys.exit(selftest(*sys.argv[2:5]))
    try:   # sharper text on high-DPI screens
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
