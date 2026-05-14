#!/usr/bin/env python3
"""
FS Incident Log — iOS GUI
ISO 26262 Powertrain  |  Archive · Form · .docx Export
"""

import sys, json, sqlite3
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ── AUTO-INSTALL ───────────────────────────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    import subprocess
    print("Installing python-docx...")
    for _flags in [[], ["--break-system-packages"]]:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "python-docx", "-q"] + _flags
            )
            break
        except subprocess.CalledProcessError:
            continue
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

# ── THEME ──────────────────────────────────────────────────────────────────────
BG        = "#F2F2F7"   # iOS system background
CARD      = "#FFFFFF"
SIDEBAR   = "#EFEFF4"
PRIMARY   = "#007AFF"   # iOS system blue
LABEL     = "#1C1C1E"
SECONDARY = "#8E8E93"
TERTIARY  = "#C7C7CC"
SEP       = "#E5E5EA"
GREEN     = "#34C759"
ORANGE    = "#FF9500"
RED       = "#FF3B30"
NAVY      = "#1D2D44"
LBL_BG    = "#EBF4FF"   # label cell light blue

if sys.platform == "win32":
    _F = "Segoe UI"
elif sys.platform == "darwin":
    _F = "SF Pro Display"
else:
    _F = "Helvetica"

def fnt(size=11, weight="normal"):
    return (_F, size, weight)

# ── DATABASE ───────────────────────────────────────────────────────────────────
DB_FILE = Path(__file__).parent / "fs_incidents.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS incidents (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at       TEXT DEFAULT (datetime('now','localtime')),
    updated_at       TEXT DEFAULT (datetime('now','localtime')),
    ticket_id        TEXT DEFAULT '',
    incident_date    TEXT DEFAULT '',
    asil             TEXT DEFAULT '',
    project          TEXT DEFAULT '',
    ecu              TEXT DEFAULT '',
    function_topic   TEXT DEFAULT '',
    safety_goal      TEXT DEFAULT '',
    sw_appeared      TEXT DEFAULT '',
    sw_fixed         TEXT DEFAULT '',
    reported_by      TEXT DEFAULT '',
    cr_id            TEXT DEFAULT '',
    powertrain       TEXT DEFAULT '[]',
    op_mode          TEXT DEFAULT '[]',
    speed_range      TEXT DEFAULT '',
    thermal_soc      TEXT DEFAULT '',
    reproducibility  TEXT DEFAULT '',
    observed         TEXT DEFAULT '',
    trigger          TEXT DEFAULT '',
    side_effect      TEXT DEFAULT '',
    customer_visible TEXT DEFAULT '',
    severity         TEXT DEFAULT '',
    signals          TEXT DEFAULT '[]',
    detection        TEXT DEFAULT '',
    safe_state       TEXT DEFAULT '',
    ftti             TEXT DEFAULT '',
    rc_layer         TEXT DEFAULT '[]',
    rc_type          TEXT DEFAULT '[]',
    rc_statement     TEXT DEFAULT '',
    resolution       TEXT DEFAULT '[]',
    fix_applied      TEXT DEFAULT '',
    engineer_notes   TEXT DEFAULT '',
    sm_changed       TEXT DEFAULT '',
    vv_status        TEXT DEFAULT '',
    incident_status  TEXT DEFAULT 'Open',
    summary          TEXT DEFAULT ''
)
"""

class Database:
    def __init__(self):
        self.con = sqlite3.connect(str(DB_FILE))
        self.con.row_factory = sqlite3.Row
        self.con.execute(SCHEMA)
        self.con.commit()

    def all(self, q=""):
        if q:
            like = f"%{q}%"
            return self.con.execute(
                "SELECT * FROM incidents WHERE ticket_id LIKE ? OR ecu LIKE ? "
                "OR project LIKE ? OR summary LIKE ? OR incident_status LIKE ? "
                "ORDER BY updated_at DESC", (like,)*5
            ).fetchall()
        return self.con.execute(
            "SELECT * FROM incidents ORDER BY updated_at DESC"
        ).fetchall()

    def get(self, id_):
        return self.con.execute(
            "SELECT * FROM incidents WHERE id=?", (id_,)
        ).fetchone()

    def upsert(self, data: dict) -> int:
        data = dict(data)
        id_ = data.pop("id", None)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["updated_at"] = now
        for k, v in list(data.items()):
            if isinstance(v, list):
                data[k] = json.dumps(v, ensure_ascii=False)
        if id_:
            sets = ", ".join(f'"{c}"=?' for c in data)
            self.con.execute(
                f'UPDATE incidents SET {sets} WHERE id=?',
                list(data.values()) + [id_]
            )
            self.con.commit()
            return id_
        data["created_at"] = now
        cols = '","'.join(data)
        ph   = ",".join("?" * len(data))
        self.con.execute(
            f'INSERT INTO incidents ("{cols}") VALUES ({ph})',
            list(data.values())
        )
        self.con.commit()
        return self.con.execute("SELECT last_insert_rowid()").fetchone()[0]

    def delete(self, id_):
        self.con.execute("DELETE FROM incidents WHERE id=?", (id_,))
        self.con.commit()


# ── SCROLLABLE FRAME ───────────────────────────────────────────────────────────
class ScrollFrame(tk.Frame):
    def __init__(self, parent, bg=BG, **kw):
        super().__init__(parent, bg=bg, **kw)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb    = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner  = tk.Frame(self.canvas, bg=bg)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self._win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._scroll)

    def _scroll(self, e):
        delta = -1 * (e.delta // 120) if sys.platform != "darwin" else -1 * e.delta
        self.canvas.yview_scroll(delta, "units")

    def to_top(self):
        self.canvas.yview_moveto(0)


# ── OPTION GROUP (checkboxes / radio) ─────────────────────────────────────────
class OptionGroup(tk.Frame):
    def __init__(self, parent, options, multi=True, cols=3, **kw):
        super().__init__(parent, bg=CARD, **kw)
        self._multi   = multi
        self._options = options
        if multi:
            self._vars = {o: tk.BooleanVar() for o in options}
            for i, opt in enumerate(options):
                r, c = divmod(i, cols)
                tk.Checkbutton(
                    self, text=opt, variable=self._vars[opt],
                    bg=CARD, fg=LABEL, activebackground=CARD,
                    selectcolor=PRIMARY, font=fnt(10),
                    relief="flat", bd=0
                ).grid(row=r, column=c, sticky="w", padx=12, pady=3)
        else:
            self._var = tk.StringVar()
            for i, opt in enumerate(options):
                r, c = divmod(i, cols)
                tk.Radiobutton(
                    self, text=opt, variable=self._var, value=opt,
                    bg=CARD, fg=LABEL, activebackground=CARD,
                    selectcolor=PRIMARY, font=fnt(10),
                    relief="flat", bd=0
                ).grid(row=r, column=c, sticky="w", padx=12, pady=3)

    def get_value(self):
        if self._multi:
            return [o for o, v in self._vars.items() if v.get()]
        return self._var.get()

    def set_value(self, val):
        if self._multi:
            for o, v in self._vars.items():
                v.set(o in (val or []))
        else:
            self._var.set(val or "")


# ── SIGNAL TABLE ───────────────────────────────────────────────────────────────
class SignalTable(tk.Frame):
    COLS   = ["Role", "Signal / Function", "Expected", "Observed", "Note"]
    WIDTHS = [14,      22,                  16,          16,         20]

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=CARD, **kw)
        self._rows = []
        hdr = tk.Frame(self, bg=PRIMARY)
        hdr.pack(fill="x")
        for col, w in zip(self.COLS, self.WIDTHS):
            tk.Label(hdr, text=col, bg=PRIMARY, fg="#FFFFFF",
                     font=fnt(9, "bold"), width=w, anchor="w",
                     padx=6, pady=4).pack(side="left")
        self._body = tk.Frame(self, bg=CARD)
        self._body.pack(fill="x")
        btn_row = tk.Frame(self, bg=CARD, pady=4)
        btn_row.pack(fill="x")
        tk.Label(btn_row, text="+ Add Row", bg=CARD, fg=PRIMARY,
                 font=fnt(10, "bold"), cursor="hand2", padx=10
                 ).pack(side="left").bind("<Button-1>", lambda e: self._add_row())
        # make label clickable without separate bind call
        for w in btn_row.winfo_children():
            w.bind("<Button-1>", lambda e: self._add_row())
        self._add_row()

    def _add_row(self, prefill=None):
        row_bg = LBL_BG if len(self._rows) % 2 == 0 else CARD
        f = tk.Frame(self._body, bg=row_bg)
        f.pack(fill="x")
        entries = []
        keys = ["role", "signal", "expected", "observed", "note"]
        for k, w in zip(keys, self.WIDTHS):
            e = tk.Entry(f, font=fnt(10), bg=row_bg, fg=LABEL,
                         relief="flat", highlightthickness=0, width=w)
            e.pack(side="left", padx=2, pady=3)
            if prefill and k in prefill:
                e.insert(0, prefill[k])
            entries.append((k, e))
        x = tk.Label(f, text="✕", bg=row_bg, fg=SECONDARY,
                     cursor="hand2", font=fnt(10), padx=4)
        x.pack(side="left")
        x.bind("<Button-1>", lambda e, fr=f, row=entries: self._del_row(fr, row))
        self._rows.append(entries)

    def _del_row(self, frame, entries):
        if entries in self._rows:
            self._rows.remove(entries)
        frame.destroy()

    def get_value(self):
        out = []
        for row in self._rows:
            d = {k: e.get().strip() for k, e in row}
            if any(d.values()):
                out.append(d)
        return out

    def set_value(self, signals):
        for row in list(self._rows):
            for _, e in row:
                e.master.destroy()
        self._rows.clear()
        for s in (signals or []):
            self._add_row(prefill=s)
        if not self._rows:
            self._add_row()


# ── WIDGET HELPERS ─────────────────────────────────────────────────────────────
def sec_hdr(parent, title):
    f = tk.Frame(parent, bg=NAVY)
    tk.Label(f, text=f"  {title}", bg=NAVY, fg="#FFFFFF",
             font=fnt(11, "bold"), pady=8, anchor="w").pack(fill="x")
    return f

def sub_hdr(parent, title):
    f = tk.Frame(parent, bg=PRIMARY)
    tk.Label(f, text=f"  {title}", bg=PRIMARY, fg="#FFFFFF",
             font=fnt(10, "bold"), pady=5, anchor="w").pack(fill="x")
    return f

def lbl_bar(parent, text):
    f = tk.Frame(parent, bg=LBL_BG)
    tk.Label(f, text=text, bg=LBL_BG, fg=PRIMARY,
             font=fnt(9, "bold"), padx=12, pady=5, anchor="w").pack(fill="x")
    return f

def hsep(parent):
    tk.Frame(parent, bg=SEP, height=1).pack(fill="x")


# ── MAIN APPLICATION ───────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.db         = Database()
        self.current_id = None
        self._fields    = {}

        self.title("FS Incident Log")
        self.geometry("1300x840")
        self.minsize(1000, 660)
        self.configure(bg=BG)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Vertical.TScrollbar",
                        background=SEP, troughcolor=BG,
                        borderwidth=0, arrowsize=12)

        self._build_ui()
        self.refresh_sidebar()
        self._new_incident()

    # ── LAYOUT ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg=NAVY, height=52)
        top.pack(fill="x")
        top.pack_propagate(False)
        tk.Label(top, text="FS Incident Log", bg=NAVY, fg="#FFFFFF",
                 font=fnt(15, "bold"), padx=20).pack(side="left", pady=12)
        tk.Label(top, text="ISO 26262 · Powertrain",
                 bg=NAVY, fg=SECONDARY, font=fnt(10)).pack(side="left")

        btn_f = tk.Frame(top, bg=NAVY)
        btn_f.pack(side="right", padx=16, pady=10)
        self._mk_btn(btn_f, "Export .docx", self._export, PRIMARY).pack(side="right", padx=4)
        self._mk_btn(btn_f, "Save",          self._save,   "#2C3E50").pack(side="right", padx=4)
        self._mk_btn(btn_f, "Delete",         self._delete, RED).pack(side="right", padx=4)
        self._mk_btn(btn_f, "+ New",           self._new_incident, "#2C3E50").pack(side="right", padx=4)

        # Body
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        # Sidebar
        sb = tk.Frame(body, bg=SIDEBAR, width=265)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        self._build_sidebar(sb)

        tk.Frame(body, bg=SEP, width=1).pack(side="left", fill="y")

        # Form
        self._scroll = ScrollFrame(body, bg=BG)
        self._scroll.pack(side="left", fill="both", expand=True)
        self._build_form(self._scroll.inner)

    def _mk_btn(self, parent, text, cmd, bg):
        f = tk.Frame(parent, bg=bg, cursor="hand2")
        lbl = tk.Label(f, text=text, bg=bg, fg="#FFFFFF",
                       font=fnt(10, "bold"), padx=14, pady=5, cursor="hand2")
        lbl.pack()
        lbl.bind("<Button-1>", lambda e: cmd())
        f.bind("<Button-1>",   lambda e: cmd())
        return f

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    def _build_sidebar(self, parent):
        hdr = tk.Frame(parent, bg=SIDEBAR, pady=10)
        hdr.pack(fill="x", padx=12)
        tk.Label(hdr, text="ARCHIVE", bg=SIDEBAR, fg=SECONDARY,
                 font=fnt(9, "bold")).pack(side="left")

        # Search box
        sf = tk.Frame(parent, bg=SEP, padx=1, pady=1)
        sf.pack(fill="x", padx=10, pady=(0, 8))
        si = tk.Frame(sf, bg=CARD)
        si.pack(fill="x")
        tk.Label(si, text="⌕", bg=CARD, fg=SECONDARY, font=fnt(12)).pack(side="left", padx=6)
        self._search = tk.StringVar()
        self._search.trace("w", lambda *a: self.refresh_sidebar())
        tk.Entry(si, textvariable=self._search, font=fnt(10),
                 bg=CARD, fg=LABEL, relief="flat",
                 highlightthickness=0).pack(side="left", fill="x", expand=True, pady=7)

        # List canvas
        lf = tk.Frame(parent, bg=SIDEBAR)
        lf.pack(fill="both", expand=True)
        self._lc = tk.Canvas(lf, bg=SIDEBAR, highlightthickness=0)
        lsb = ttk.Scrollbar(lf, orient="vertical", command=self._lc.yview)
        self._li = tk.Frame(self._lc, bg=SIDEBAR)
        self._lc.configure(yscrollcommand=lsb.set)
        self._lc.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        self._lw = self._lc.create_window((0, 0), window=self._li, anchor="nw")
        self._li.bind("<Configure>", lambda e: self._lc.configure(
            scrollregion=self._lc.bbox("all")))
        self._lc.bind("<Configure>", lambda e: self._lc.itemconfig(
            self._lw, width=e.width))

    def refresh_sidebar(self):
        rows = self.db.all(self._search.get() if hasattr(self, "_search") else "")
        for w in self._li.winfo_children():
            w.destroy()
        for row in rows:
            self._sidebar_item(row)
        if not rows:
            tk.Label(self._li, text="No incidents yet", bg=SIDEBAR,
                     fg=SECONDARY, font=fnt(10), pady=20).pack()

    def _sidebar_item(self, row):
        id_  = row["id"]
        sel  = self.current_id == id_
        bg   = PRIMARY if sel else SIDEBAR
        fg   = "#FFFFFF" if sel else LABEL
        sfg  = "#FFFFFF" if sel else SECONDARY

        item = tk.Frame(self._li, bg=bg, cursor="hand2")
        item.pack(fill="x")

        top = tk.Frame(item, bg=bg)
        top.pack(fill="x", padx=12, pady=(8, 2))
        tk.Label(top, text=row["ticket_id"] or "—",
                 bg=bg, fg=fg, font=fnt(11, "bold"), anchor="w").pack(side="left")
        dt = (row["incident_date"] or row["created_at"] or "")[:10]
        tk.Label(top, text=dt, bg=bg, fg=sfg, font=fnt(9)).pack(side="right")

        bot = tk.Frame(item, bg=bg)
        bot.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(bot, text=row["ecu"] or row["project"] or "—",
                 bg=bg, fg=sfg, font=fnt(9)).pack(side="left")
        status = row["incident_status"] or "Open"
        sc = "#FFFFFF" if sel else {
            "Closed": GREEN, "Open": ORANGE, "In Progress": PRIMARY
        }.get(status, SECONDARY)
        tk.Label(bot, text=status, bg=bg, fg=sc,
                 font=fnt(8, "bold")).pack(side="right")

        tk.Frame(self._li, bg=SEP, height=1).pack(fill="x")

        for widget in self._all_children(item):
            widget.bind("<Button-1>", lambda e, i=id_: self._load(i))
        item.bind("<Button-1>", lambda e, i=id_: self._load(i))

    def _all_children(self, widget):
        result = [widget]
        for child in widget.winfo_children():
            result.extend(self._all_children(child))
        return result

    # ── FORM BUILD ────────────────────────────────────────────────────────────
    def _build_form(self, f):
        pad = {"fill": "x"}

        def card():
            c = tk.Frame(f, bg=CARD)
            c.pack(fill="x", pady=(0, 12))
            return c

        def entry(parent, key, label, hint=""):
            lbl_bar(parent, label).pack(fill="x")
            ef = tk.Frame(parent, bg=CARD, padx=10, pady=4)
            ef.pack(fill="x")
            e = tk.Entry(ef, font=fnt(11), bg=CARD, fg=LABEL,
                         relief="flat", highlightthickness=1,
                         highlightbackground=SEP, highlightcolor=PRIMARY)
            e.pack(fill="x", ipady=6)
            if hint:
                e.insert(0, hint)
                e.configure(fg=TERTIARY)
                def _fi(ev, widget=e, h=hint):
                    if widget.get() == h:
                        widget.delete(0, "end")
                        widget.configure(fg=LABEL)
                def _fo(ev, widget=e, h=hint):
                    if not widget.get().strip():
                        widget.insert(0, h)
                        widget.configure(fg=TERTIARY)
                e.bind("<FocusIn>",  _fi)
                e.bind("<FocusOut>", _fo)
                e._hint = hint
            else:
                e._hint = ""
            hsep(parent)
            self._fields[key] = e

        def textarea(parent, key, label, height=4):
            lbl_bar(parent, label).pack(fill="x")
            tf = tk.Frame(parent, bg=CARD, padx=10, pady=6)
            tf.pack(fill="x")
            t = tk.Text(tf, font=fnt(11), bg=CARD, fg=LABEL,
                        insertbackground=PRIMARY, relief="flat",
                        highlightthickness=1, highlightbackground=SEP,
                        highlightcolor=PRIMARY, wrap="word",
                        height=height, padx=8, pady=6)
            t.pack(fill="x")
            hsep(parent)
            self._fields[key] = t

        def options(parent, key, opts, multi=True, cols=3):
            og = OptionGroup(parent, opts, multi=multi, cols=cols)
            og.pack(fill="x", padx=2, pady=4)
            hsep(parent)
            self._fields[key] = og

        def radio_row(parent, key, label, opts):
            lbl_bar(parent, label).pack(fill="x")
            og = OptionGroup(parent, opts, multi=False, cols=len(opts))
            og.pack(fill="x", padx=2, pady=4)
            hsep(parent)
            self._fields[key] = og

        # ── 1. IDENTIFICATION ─────────────────────────────────────────────────
        sec_hdr(f, "1.  IDENTIFICATION").pack(**pad, pady=(0, 0))
        c = card()

        # 2×2 quad
        quad = tk.Frame(c, bg=CARD)
        quad.pack(fill="x")
        for i, (key, label, hint) in enumerate([
            ("ticket_id",     "TICKET / ID",        "e.g. JIRA-1234"),
            ("incident_date", "DATE",                str(date.today())),
            ("asil",          "ASIL",                "A / B / C / D / QM"),
            ("project",       "PROJECT / PLATFORM",  "e.g. MEB, PPE, C3"),
        ]):
            cell = tk.Frame(quad, bg=CARD)
            cell.grid(row=i//2, column=i%2, sticky="nsew", padx=1, pady=1)
            quad.columnconfigure(i%2, weight=1)
            lbl_bar(cell, label).pack(fill="x")
            ef = tk.Frame(cell, bg=CARD, padx=10, pady=4)
            ef.pack(fill="x")
            e = tk.Entry(ef, font=fnt(11), bg=CARD, fg=TERTIARY,
                         relief="flat", highlightthickness=1,
                         highlightbackground=SEP, highlightcolor=PRIMARY)
            e.insert(0, hint)
            e._hint = hint
            def _fi(ev, w=e, h=hint):
                if w.get() == h: w.delete(0, "end"); w.configure(fg=LABEL)
            def _fo(ev, w=e, h=hint):
                if not w.get().strip(): w.insert(0, h); w.configure(fg=TERTIARY)
            e.bind("<FocusIn>",  _fi)
            e.bind("<FocusOut>", _fo)
            e.pack(fill="x", ipady=6)
            self._fields[key] = e
        hsep(c)

        for key, label, hint in [
            ("ecu",            "ECU / SYSTEM",             "e.g. HCP1, GE3, DCDC"),
            ("function_topic", "FUNCTION / TOPIC",          "e.g. MonMTqO, DrvDmdasil"),
            ("safety_goal",    "SAFETY GOAL REF.",          "SG-xx  |  FSC-xx  |  SwSR-xx"),
            ("sw_appeared",    "SW VERSION — APPEARED IN",  "e.g. HCP1_SW_23.40.1"),
            ("sw_fixed",       "SW VERSION — FIXED IN",     "e.g. HCP1_SW_23.44.0  |  Pending"),
            ("reported_by",    "REPORTED BY",               "Name, role, date"),
            ("cr_id",          "CHANGE REQUEST (CR)",       "StarTeam / Jira CR ID"),
        ]:
            entry(c, key, label, hint)

        # ── 2. CONTEXT ────────────────────────────────────────────────────────
        sec_hdr(f, "2.  CONTEXT").pack(**pad, pady=(12, 0))
        c = card()
        sub_hdr(c, "Powertrain Type").pack(fill="x")
        options(c, "powertrain", ["ICE","MHEV","HEV","PHEV","BEV","FCEV"], cols=6)
        sub_hdr(c, "Operating Mode").pack(fill="x")
        options(c, "op_mode",
            ["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"], cols=3)
        entry(c, "speed_range", "SPEED RANGE",   "e.g. standstill / <30 km/h / highway")
        entry(c, "thermal_soc", "THERMAL / SOC", "e.g. cold start, warm, low SOC")
        radio_row(c, "reproducibility", "REPRODUCIBILITY", ["Always","Intermittent","Rare"])

        # ── 3. INCIDENT ───────────────────────────────────────────────────────
        sec_hdr(f, "3.  INCIDENT").pack(**pad, pady=(12, 0))
        c = card()
        textarea(c, "observed",    "OBSERVED BEHAVIOR",  height=4)
        textarea(c, "trigger",     "TRIGGER CONDITION",  height=4)
        textarea(c, "side_effect", "SIDE EFFECT",        height=3)

        # Customer visible + Severity side by side
        row2 = tk.Frame(c, bg=CARD)
        row2.pack(fill="x")
        for i, (key, label, opts) in enumerate([
            ("customer_visible", "CUSTOMER VISIBLE?", ["Yes","No"]),
            ("severity",         "SEVERITY",           ["Safety-relevant","Functional","Comfort"]),
        ]):
            cell = tk.Frame(row2, bg=CARD)
            cell.grid(row=0, column=i, sticky="nsew", padx=1)
            row2.columnconfigure(i, weight=1)
            lbl_bar(cell, label).pack(fill="x")
            og = OptionGroup(cell, opts, multi=False, cols=len(opts))
            og.pack(fill="x", padx=2, pady=4)
            self._fields[key] = og

        # ── 4. SIGNAL CONFLICT ────────────────────────────────────────────────
        sec_hdr(f, "4.  SIGNAL CONFLICT").pack(**pad, pady=(12, 0))
        c = card()
        lbl_bar(c, "CONFLICT TABLE").pack(fill="x")
        sf = tk.Frame(c, bg=CARD, padx=10, pady=8)
        sf.pack(fill="x")
        self._signal_table = SignalTable(sf)
        self._signal_table.pack(fill="x")
        self._fields["signals"] = self._signal_table
        hsep(c)

        # ── 5. SAFETY RESPONSE ────────────────────────────────────────────────
        sec_hdr(f, "5.  SAFETY RESPONSE").pack(**pad, pady=(12, 0))
        c = card()
        textarea(c, "detection",  "DETECTION FUNCTION", height=3)
        textarea(c, "safe_state", "SAFE STATE ENTERED",  height=3)
        radio_row(c, "ftti", "FTTI RESPECTED?",
            ["Yes","No — explain in analysis","Not assessed"])

        # ── 6. ROOT CAUSE & FIX ───────────────────────────────────────────────
        sec_hdr(f, "6.  ROOT CAUSE & FIX").pack(**pad, pady=(12, 0))
        c = card()
        sub_hdr(c, "Root Cause Layer").pack(fill="x")
        options(c, "rc_layer",
            ["SW Logic","Signal Definition","Calibration",
             "Interface / Communication","Timing / Scheduling","Architecture"], cols=3)
        sub_hdr(c, "Root Cause Type").pack(fill="x")
        options(c, "rc_type",
            ["Incorrect rule","Missing condition","Wrong threshold","State machine error",
             "Signal misinterpretation","Missing debounce","Race condition","Incomplete spec"], cols=4)
        textarea(c, "rc_statement", "ROOT CAUSE STATEMENT", height=3)
        sub_hdr(c, "Resolution Category").pack(fill="x")
        options(c, "resolution",
            ["Logic refinement","State reclassification","Timing / debounce",
             "Architectural change","Calibration correction","Signal definition fix",
             "Spec update only","No fix — accepted risk"], cols=4)
        textarea(c, "fix_applied", "FIX APPLIED", height=3)

        # ── 7. ENGINEER SPACE ─────────────────────────────────────────────────
        sec_hdr(f, "7.  ENGINEER SPACE").pack(**pad, pady=(12, 0))
        c = card()
        textarea(c, "engineer_notes", "NOTES & ANALYSIS", height=8)

        # ── 8. STATUS ─────────────────────────────────────────────────────────
        sec_hdr(f, "8.  STATUS").pack(**pad, pady=(12, 0))
        c = card()
        radio_row(c, "sm_changed", "SAFETY MECHANISM CHANGED?",
            ["No — false trigger removed only","Yes — safety behavior modified"])
        radio_row(c, "vv_status", "V&V STATUS",
            ["Pending","In Progress","Complete","Waived"])
        radio_row(c, "incident_status", "INCIDENT STATUS",
            ["Open","In Progress","Closed","Monitoring"])

        # ── 9. TECHNICAL SUMMARY ──────────────────────────────────────────────
        sec_hdr(f, "9.  TECHNICAL SUMMARY").pack(**pad, pady=(12, 0))
        c = card()
        textarea(c, "summary",
            "ONE-SENTENCE SUMMARY\n[ Root cause ] — [ Safety reaction ] — [ Functional effect ]",
            height=3)

        tk.Frame(f, bg=BG, height=40).pack()

    # ── FORM DATA ─────────────────────────────────────────────────────────────
    def _get_entry_val(self, widget):
        val = widget.get().strip()
        hint = getattr(widget, "_hint", "")
        return "" if val == hint else val

    def _get_form(self) -> dict:
        data = {}
        for key, w in self._fields.items():
            if isinstance(w, tk.Entry):
                data[key] = self._get_entry_val(w)
            elif isinstance(w, tk.Text):
                data[key] = w.get("1.0", "end-1c").strip()
            elif isinstance(w, (OptionGroup, SignalTable)):
                data[key] = w.get_value()
        return data

    def _set_form(self, data: dict):
        for key, w in self._fields.items():
            raw = data.get(key, "")
            # Decode JSON strings
            if isinstance(raw, str) and raw.startswith(("[", "{")):
                try:
                    raw = json.loads(raw)
                except Exception:
                    pass
            if isinstance(w, tk.Entry):
                w.delete(0, "end")
                hint = getattr(w, "_hint", "")
                if raw:
                    w.insert(0, str(raw))
                    w.configure(fg=LABEL)
                elif hint:
                    w.insert(0, hint)
                    w.configure(fg=TERTIARY)
            elif isinstance(w, tk.Text):
                w.delete("1.0", "end")
                if raw:
                    w.insert("1.0", str(raw))
            elif isinstance(w, (OptionGroup, SignalTable)):
                w.set_value(raw)

    def _clear_form(self):
        empty = {k: "" for k in self._fields}
        empty["incident_date"] = str(date.today())
        self._set_form(empty)
        # Reset lists
        for k in ["powertrain","op_mode","rc_layer","rc_type","resolution","signals"]:
            if k in self._fields:
                self._fields[k].set_value([])

    # ── ACTIONS ───────────────────────────────────────────────────────────────
    def _new_incident(self):
        self.current_id = None
        self._clear_form()
        self._scroll.to_top()
        self.refresh_sidebar()

    def _load(self, id_):
        row = self.db.get(id_)
        if not row:
            return
        self.current_id = id_
        self._set_form(dict(row))
        self._scroll.to_top()
        self.refresh_sidebar()

    def _save(self):
        data = self._get_form()
        data["id"] = self.current_id
        new_id = self.db.upsert(data)
        self.current_id = new_id
        self.refresh_sidebar()
        self.title("FS Incident Log  ✓  Saved")
        self.after(2000, lambda: self.title("FS Incident Log"))

    def _delete(self):
        if not self.current_id:
            messagebox.showinfo("Nothing selected", "Open an incident first.")
            return
        if messagebox.askyesno("Delete", "Delete this incident permanently?"):
            self.db.delete(self.current_id)
            self.current_id = None
            self._clear_form()
            self.refresh_sidebar()

    def _export(self):
        if not self.current_id:
            self._save()
        data     = self._get_form()
        # Ensure lists are decoded
        for k in ["powertrain","op_mode","rc_layer","rc_type","resolution"]:
            if isinstance(data.get(k), str):
                try: data[k] = json.loads(data[k])
                except: data[k] = []
        sigs = data.get("signals", [])
        if isinstance(sigs, str):
            try: sigs = json.loads(sigs)
            except: sigs = []

        doc_data = {
            "ticket_id":        data.get("ticket_id",""),
            "date":             data.get("incident_date", str(date.today())),
            "asil":             data.get("asil",""),
            "project":          data.get("project",""),
            "ecu":              data.get("ecu",""),
            "function":         data.get("function_topic",""),
            "safety_goal":      data.get("safety_goal",""),
            "sw_appeared":      data.get("sw_appeared",""),
            "sw_fixed":         data.get("sw_fixed",""),
            "reported_by":      data.get("reported_by",""),
            "cr_id":            data.get("cr_id",""),
            "powertrain":       data.get("powertrain",[]),
            "op_mode":          data.get("op_mode",[]),
            "speed_range":      data.get("speed_range",""),
            "thermal_soc":      data.get("thermal_soc",""),
            "reproducibility":  data.get("reproducibility",""),
            "observed":         data.get("observed",""),
            "trigger":          data.get("trigger",""),
            "side_effect":      data.get("side_effect",""),
            "customer_visible": data.get("customer_visible",""),
            "severity":         data.get("severity",""),
            "signals":          sigs,
            "detection":        data.get("detection",""),
            "safe_state":       data.get("safe_state",""),
            "ftti":             data.get("ftti",""),
            "rc_layer":         data.get("rc_layer",[]),
            "rc_type":          data.get("rc_type",[]),
            "rc_statement":     data.get("rc_statement",""),
            "resolution":       data.get("resolution",[]),
            "fix_applied":      data.get("fix_applied",""),
            "engineer_notes":   data.get("engineer_notes",""),
            "attachments":      [],
            "sm_changed":       data.get("sm_changed",""),
            "vv_status":        data.get("vv_status",""),
            "incident_status":  data.get("incident_status",""),
            "summary":          data.get("summary",""),
        }

        ticket  = doc_data["ticket_id"] or "incident"
        default = f"FS_Incident_{ticket}_{date.today().strftime('%Y%m%d')}.docx"
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document","*.docx")],
            initialfile=default
        )
        if not path:
            return
        try:
            build_doc(doc_data, path)
            self.title("FS Incident Log  ✓  Exported")
            self.after(2500, lambda: self.title("FS Incident Log"))
        except Exception as ex:
            messagebox.showerror("Export failed", str(ex))


# ── DOCX BUILDER ───────────────────────────────────────────────────────────────
DN  = "1D2D44"   # navy
DB  = "0A84FF"   # blue
DLB = "F0F4FF"   # label bg
DVB = "FFFFFF"   # value bg
DBR = "E2E8F0"   # border
DTX = "1C1C1E"   # text
DSC = "6E6E73"   # secondary
DPH = "AEAEB2"   # placeholder
DGR = "34C759"   # green
DOR = "FF9500"   # orange
DSF = "F5F5F7"   # surface

def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def _set_bg(cell, color):
    tc = cell._tc; p = tc.get_or_add_tcPr()
    for s in p.findall(qn("w:shd")): p.remove(s)
    s = OxmlElement("w:shd")
    s.set(qn("w:val"),"clear"); s.set(qn("w:color"),"auto")
    s.set(qn("w:fill"), color.lstrip("#").upper()); p.append(s)

def _borders(cell, top=None, bot=None, left=None, right=None):
    tc = cell._tc; p = tc.get_or_add_tcPr()
    for b in p.findall(qn("w:tcBorders")): p.remove(b)
    bd = OxmlElement("w:tcBorders")
    for edge, c in [("top",top),("bottom",bot),("start",left),("end",right)]:
        el = OxmlElement(f"w:{edge}")
        if c:
            el.set(qn("w:val"),"single"); el.set(qn("w:sz"),"4")
            el.set(qn("w:color"), c.lstrip("#").upper())
        else:
            el.set(qn("w:val"),"nil")
        bd.append(el)
    sh = p.find(qn("w:shd"))
    if sh is not None: sh.addprevious(bd)
    else: p.append(bd)

def _margins(cell, t=80, b=80, l=120, r=120):
    tc = cell._tc; p = tc.get_or_add_tcPr()
    for m in p.findall(qn("w:tcMar")): p.remove(m)
    mar = OxmlElement("w:tcMar")
    for edge, val in [("top",t),("bottom",b),("start",l),("end",r)]:
        el = OxmlElement(f"w:{edge}"); el.set(qn("w:w"),str(val)); el.set(qn("w:type"),"dxa"); mar.append(el)
    p.append(mar)

def _run(para, text, bold=False, italic=False, color=DTX, size=10):
    run = para.add_run(text)
    run.bold = bold; run.italic = italic
    run.font.name = "Arial"; run.font.size = Pt(size)
    r, g, b = _rgb(color); run.font.color.rgb = RGBColor(r, g, b)
    return run

def _cp(cell, text, bold=False, italic=False, color=DTX, size=10,
        align=WD_ALIGN_PARAGRAPH.LEFT, sb=60, sa=60, clear=True):
    if clear: cell.paragraphs[0].clear(); para = cell.paragraphs[0]
    else:     para = cell.add_paragraph()
    para.alignment = align
    para.paragraph_format.space_before = Pt(sb/20)
    para.paragraph_format.space_after  = Pt(sa/20)
    if text: _run(para, text, bold=bold, italic=italic, color=color, size=size)
    return para

def _merge_row(tbl, fill, text, bold=True, color="FFFFFF", size=10.5):
    row = tbl.add_row(); cell = row.cells[0]
    for i in range(1, len(row.cells)): cell = cell.merge(row.cells[i])
    _set_bg(cell, fill); _borders(cell, top=fill, bot=fill, left=fill, right=fill)
    _margins(cell, 70, 70, 140, 140); _cp(cell, text, bold=bold, color=color, size=size)

def _lv(tbl, label, value):
    row = tbl.add_row(); c0, c1 = row.cells[0], row.cells[1]
    for i in range(2, len(row.cells)): c1 = c1.merge(row.cells[i])
    _set_bg(c0, DLB); _borders(c0, top=DBR, bot=DBR, left=DBR, right=DBR)
    _set_bg(c1, DVB); _borders(c1, bot=DBR)
    _margins(c0, 80, 80, 130, 130); _margins(c1, 80, 80, 130, 130)
    _cp(c0, label, bold=True, color=DB, size=9.5)
    _cp(c1, value, italic=not bool(value),
        color=DPH if not value else DTX, size=10)

def _cb(tbl, items, selected):
    row = tbl.add_row(); cell = row.cells[0]
    for i in range(1, len(row.cells)): cell = cell.merge(row.cells[i])
    _set_bg(cell, DVB); _borders(cell, bot=DBR)
    _margins(cell, 60, 60, 130, 130)
    para = cell.paragraphs[0]; para.clear()
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after  = Pt(3)
    for j, item in enumerate(items):
        if j > 0: _run(para, "     ", size=10, color=DTX)
        tick = "☑" if item in selected else "☐"
        _run(para, f"{tick}  {item}", size=10,
             color=DB if item in selected else DSC,
             bold=(item in selected))

def _ta(tbl, text, min_lines=5):
    row = tbl.add_row(); cell = row.cells[0]
    for i in range(1, len(row.cells)): cell = cell.merge(row.cells[i])
    _set_bg(cell, DVB); _borders(cell, bot=DBR)
    _margins(cell, 100, 100, 140, 140)
    para = cell.paragraphs[0]; para.clear()
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after  = Pt(4)
    if text:
        for i, line in enumerate(text.split("\n")):
            if i == 0: _run(para, line, size=10, color=DTX)
            else:
                np = cell.add_paragraph()
                np.paragraph_format.space_before = Pt(2)
                np.paragraph_format.space_after  = Pt(2)
                _run(np, line, size=10, color=DTX)
    used = max(1, len(text.split("\n")) if text else 0)
    for _ in range(max(0, min_lines - used)):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)

def build_doc(data, path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21); sec.page_height = Cm(29.7)
    sec.left_margin = Cm(1.5); sec.right_margin = Cm(1.5)
    sec.top_margin  = Cm(1.2); sec.bottom_margin = Cm(1.2)

    hp = doc.sections[0].header.paragraphs[0]
    hp.clear(); hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _run(hp, "FUNCTIONAL SAFETY INCIDENT LOG  —  ISO 26262 Powertrain", bold=True, color=DN, size=8)
    _run(hp, "          CONFIDENTIAL — INTERNAL USE ONLY", italic=True, color=DSC, size=8)
    hp.paragraph_format.space_after = Pt(6)

    tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(0); tp.paragraph_format.space_after = Pt(4)
    _run(tp, "FS INCIDENT LOG", bold=True, color=DN, size=20)

    sp = doc.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp.paragraph_format.space_before = Pt(0); sp.paragraph_format.space_after = Pt(14)
    _run(sp, "ISO 26262  —  Powertrain  —  v3.0", italic=True, color=DB, size=10)

    CW = int(18*567); C1 = int(3.5*567); C2 = CW - C1

    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"; tbl.autofit = False
    tg = tbl._tbl.find(qn("w:tblGrid"))
    if tg is not None:
        for gc in tg.findall(qn("w:gridCol")): tg.remove(gc)
        for w in [C1, C2]:
            gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(w)); tg.append(gc)
    tpr = tbl._tbl.find(qn("w:tblPr"))
    if tpr is None: tpr = OxmlElement("w:tblPr"); tbl._tbl.insert(0, tpr)
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), str(CW)); tw.set(qn("w:type"), "dxa")
    tpr.insert(0, tw)

    def S(title): _merge_row(tbl, DN, f"  {title}", bold=True, color="FFFFFF", size=10.5)
    def U(title): _merge_row(tbl, DB, f"  {title}", bold=True, color="FFFFFF", size=9.5)

    # 1 — IDENTIFICATION
    S("1.  IDENTIFICATION")
    r = tbl.add_row(); ca = r.cells[0].merge(r.cells[1])
    _set_bg(ca, DVB); _borders(ca, top=DBR, bot=DBR, left=DBR, right=DBR)
    _margins(ca, 0, 0, 0, 0)
    inn = OxmlElement("w:tbl")
    ip = OxmlElement("w:tblPr"); iw = OxmlElement("w:tblW")
    iw.set(qn("w:w"), str(CW)); iw.set(qn("w:type"), "dxa"); ip.append(iw)
    ib = OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"),"single")
        e.set(qn("w:sz"),"4"); e.set(qn("w:color"), DBR); ib.append(e)
    ip.append(ib); inn.append(ip)
    ig = OxmlElement("w:tblGrid"); qw = CW//4
    for _ in range(4):
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(qw)); ig.append(gc)
    inn.append(ig)
    ir = OxmlElement("w:tr")
    for lbl_t, val in [("TICKET / ID",data["ticket_id"]),("DATE",data["date"]),
                        ("ASIL",data["asil"]),("PROJECT / PLATFORM",data["project"])]:
        tc = OxmlElement("w:tc"); tp2 = OxmlElement("w:tcPr")
        tw2 = OxmlElement("w:tcW"); tw2.set(qn("w:w"),str(qw)); tw2.set(qn("w:type"),"dxa"); tp2.append(tw2)
        sh = OxmlElement("w:shd"); sh.set(qn("w:val"),"clear"); sh.set(qn("w:color"),"auto")
        sh.set(qn("w:fill"), DVB); tp2.append(sh)
        m2 = OxmlElement("w:tcMar")
        for edge in ["top","bottom","left","right"]:
            me = OxmlElement(f"w:{edge}"); me.set(qn("w:w"),"100"); me.set(qn("w:type"),"dxa"); m2.append(me)
        tp2.append(m2); tc.append(tp2)
        lp = OxmlElement("w:p"); lr = OxmlElement("w:r"); lrp = OxmlElement("w:rPr")
        lb = OxmlElement("w:b"); lc = OxmlElement("w:color"); lc.set(qn("w:val"), DB)
        ls = OxmlElement("w:sz"); ls.set(qn("w:val"),"17")
        lrp.append(lb); lrp.append(lc); lrp.append(ls); lr.append(lrp)
        lt = OxmlElement("w:t"); lt.text = lbl_t; lr.append(lt); lp.append(lr); tc.append(lp)
        vp = OxmlElement("w:p"); vr = OxmlElement("w:r"); vrp = OxmlElement("w:rPr")
        vc = OxmlElement("w:color"); vc.set(qn("w:val"), DTX if val else DPH)
        vs = OxmlElement("w:sz"); vs.set(qn("w:val"),"20")
        if not val: vi = OxmlElement("w:i"); vrp.append(vi)
        vrp.append(vc); vrp.append(vs); vr.append(vrp)
        vt = OxmlElement("w:t"); vt.text = val or "—"; vr.append(vt); vp.append(vr); tc.append(vp)
        ir.append(tc)
    inn.append(ir); ca._tc.append(inn)

    _lv(tbl, "ECU / SYSTEM",        data["ecu"])
    _lv(tbl, "FUNCTION / TOPIC",    data["function"])
    _lv(tbl, "SAFETY GOAL REF.",    data["safety_goal"])
    _lv(tbl, "SW — APPEARED IN",    data["sw_appeared"])
    _lv(tbl, "SW — FIXED IN",       data["sw_fixed"])
    _lv(tbl, "REPORTED BY",         data["reported_by"])
    _lv(tbl, "CHANGE REQUEST (CR)", data["cr_id"])

    # 2 — CONTEXT
    S("2.  CONTEXT")
    U("Powertrain Type")
    _cb(tbl, ["ICE","MHEV","HEV","PHEV","BEV","FCEV"], data["powertrain"])
    U("Operating Mode")
    _cb(tbl, ["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"], data["op_mode"])
    _lv(tbl, "SPEED RANGE",   data["speed_range"])
    _lv(tbl, "THERMAL / SOC", data["thermal_soc"])
    rr = tbl.add_row(); rc0, rc1 = rr.cells[0], rr.cells[1]
    _set_bg(rc0, DLB); _borders(rc0, top=DBR, bot=DBR, left=DBR, right=DBR)
    _set_bg(rc1, DVB); _borders(rc1, bot=DBR)
    _margins(rc0,80,80,130,130); _margins(rc1,60,60,130,130)
    _cp(rc0,"REPRODUCIBILITY",bold=True,color=DB,size=9.5)
    pr = rc1.paragraphs[0]; pr.clear()
    pr.paragraph_format.space_before = Pt(3); pr.paragraph_format.space_after = Pt(3)
    for j, opt in enumerate(["Always","Intermittent","Rare"]):
        if j > 0: _run(pr,"     ",size=10,color=DTX)
        sel = data["reproducibility"] == opt
        _run(pr,f"{'☑' if sel else '☐'}  {opt}",size=10,
             color=DB if sel else DSC,bold=sel)

    # 3 — INCIDENT
    S("3.  INCIDENT")
    _lv(tbl,"OBSERVED BEHAVIOR",""); _ta(tbl, data["observed"])
    _lv(tbl,"TRIGGER CONDITION",""); _ta(tbl, data["trigger"])
    _lv(tbl,"SIDE EFFECT","");       _ta(tbl, data["side_effect"])
    rcv = tbl.add_row(); cc0, cc1 = rcv.cells[0], rcv.cells[1]
    for cx, lbl_t, items, sel_ in [
        (cc0,"CUSTOMER VISIBLE?",["Yes","No"],
         [data["customer_visible"]] if data["customer_visible"] else []),
        (cc1,"SEVERITY",["Safety-relevant","Functional","Comfort"],
         [data["severity"]] if data["severity"] else []),
    ]:
        _set_bg(cx, DVB); _borders(cx, top=DBR, bot=DBR, left=DBR, right=DBR)
        _margins(cx,80,80,130,130); _cp(cx,lbl_t,bold=True,color=DB,size=9.5)
        p2 = cx.add_paragraph()
        p2.paragraph_format.space_before=Pt(3); p2.paragraph_format.space_after=Pt(3)
        for j, opt in enumerate(items):
            if j > 0: _run(p2,"     ",size=10,color=DTX)
            sel = opt in sel_
            _run(p2,f"{'☑' if sel else '☐'}  {opt}",size=10,
                 color=DB if sel else DSC,bold=sel)

    # 4 — SIGNAL CONFLICT
    S("4.  SIGNAL CONFLICT")
    sr2 = tbl.add_row(); ss0, ss1 = sr2.cells[0], sr2.cells[1]
    _set_bg(ss0,DLB); _borders(ss0,top=DBR,bot=DBR,left=DBR,right=DBR)
    _margins(ss0,80,80,130,130); _cp(ss0,"CONFLICT TABLE",bold=True,color=DB,size=9.5)
    _set_bg(ss1,DVB); _borders(ss1,top=DBR,bot=DBR,left=DBR,right=DBR)
    _margins(ss1,80,80,100,100)
    sc = ["Role","Signal / Function","Expected","Observed","Note"]
    sw = [int(C2*.14),int(C2*.24),int(C2*.18),int(C2*.18),
          C2-int(C2*.14)-int(C2*.24)-int(C2*.18)*2]
    st = OxmlElement("w:tbl"); stp = OxmlElement("w:tblPr")
    stw = OxmlElement("w:tblW"); stw.set(qn("w:w"),str(C2)); stw.set(qn("w:type"),"dxa"); stp.append(stw)
    stb = OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"),"single")
        e.set(qn("w:sz"),"3"); e.set(qn("w:color"),DBR); stb.append(e)
    stp.append(stb); st.append(stp)
    sg = OxmlElement("w:tblGrid")
    for w in sw: gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"),str(w)); sg.append(gc)
    st.append(sg)

    def _str(values, is_hdr=False, rfill=DVB):
        tr = OxmlElement("w:tr")
        for val, w in zip(values, sw):
            tc = OxmlElement("w:tc"); tp3 = OxmlElement("w:tcPr")
            tw3 = OxmlElement("w:tcW"); tw3.set(qn("w:w"),str(w)); tw3.set(qn("w:type"),"dxa"); tp3.append(tw3)
            sh3 = OxmlElement("w:shd"); sh3.set(qn("w:val"),"clear"); sh3.set(qn("w:color"),"auto")
            sh3.set(qn("w:fill"),DB if is_hdr else rfill); tp3.append(sh3)
            m3 = OxmlElement("w:tcMar")
            for edge in ["top","bottom","left","right"]:
                me = OxmlElement(f"w:{edge}"); me.set(qn("w:w"),"80"); me.set(qn("w:type"),"dxa"); m3.append(me)
            tp3.append(m3); tc.append(tp3)
            wp3 = OxmlElement("w:p"); wr3 = OxmlElement("w:r"); wrp3 = OxmlElement("w:rPr")
            if is_hdr: wb = OxmlElement("w:b"); wrp3.append(wb)
            wc3 = OxmlElement("w:color"); wc3.set(qn("w:val"),"FFFFFF" if is_hdr else DTX); wrp3.append(wc3)
            ws3 = OxmlElement("w:sz"); ws3.set(qn("w:val"),"17"); wrp3.append(ws3); wr3.append(wrp3)
            wt3 = OxmlElement("w:t"); wt3.text = val or ""; wr3.append(wt3); wp3.append(wr3); tc.append(wp3); tr.append(tc)
        return tr

    st.append(_str(sc, is_hdr=True))
    rfills = [DLB,DVB,"EBF5EE",DVB,"FDF8EC",DVB]
    rlbls  = ["Driver Intent","","Physical State","","System State",""]
    sigs   = data["signals"]
    for ri in range(6):
        if ri < len(sigs):
            s = sigs[ri]; vals = [s.get("role",""),s.get("signal",""),s.get("expected",""),s.get("observed",""),s.get("note","")]
        else:
            vals = [rlbls[ri],"","","",""]
        st.append(_str(vals, rfill=rfills[ri]))
    for s in sigs[6:]:
        st.append(_str([s.get("role",""),s.get("signal",""),s.get("expected",""),s.get("observed",""),s.get("note","")]))
    ss1._tc.append(st)

    # 5 — SAFETY RESPONSE
    S("5.  SAFETY RESPONSE")
    _lv(tbl,"DETECTION FUNCTION",""); _ta(tbl, data["detection"])
    _lv(tbl,"SAFE STATE ENTERED","");  _ta(tbl, data["safe_state"])
    rf = tbl.add_row(); rf0, rf1 = rf.cells[0], rf.cells[1]
    _set_bg(rf0,DLB); _borders(rf0,top=DBR,bot=DBR,left=DBR,right=DBR)
    _set_bg(rf1,DVB); _borders(rf1,bot=DBR)
    _margins(rf0,80,80,130,130); _margins(rf1,60,60,130,130)
    _cp(rf0,"FTTI RESPECTED?",bold=True,color=DB,size=9.5)
    pf = rf1.paragraphs[0]; pf.clear()
    pf.paragraph_format.space_before=Pt(3); pf.paragraph_format.space_after=Pt(3)
    for j, opt in enumerate(["Yes","No — explain in analysis","Not assessed"]):
        if j > 0: _run(pf,"          ",size=10,color=DTX)
        sel = data["ftti"]==opt
        _run(pf,f"{'☑' if sel else '☐'}  {opt}",size=10,color=DB if sel else DSC,bold=sel)

    # 6 — ROOT CAUSE & FIX
    S("6.  ROOT CAUSE & FIX")
    U("Root Cause Layer")
    _cb(tbl,["SW Logic","Signal Definition","Calibration","Interface / Communication","Timing / Scheduling","Architecture"],data["rc_layer"])
    U("Root Cause Type")
    _cb(tbl,["Incorrect rule","Missing condition","Wrong threshold","State machine error","Signal misinterpretation","Missing debounce","Race condition","Incomplete spec"],data["rc_type"])
    _lv(tbl,"ROOT CAUSE STATEMENT",""); _ta(tbl, data["rc_statement"], min_lines=4)
    U("Resolution Category")
    _cb(tbl,["Logic refinement","State reclassification","Timing / debounce","Architectural change","Calibration correction","Signal definition fix","Spec update only","No fix — accepted risk"],data["resolution"])
    _lv(tbl,"FIX APPLIED",""); _ta(tbl, data["fix_applied"])

    # 7 — ENGINEER SPACE
    S("7.  ENGINEER SPACE")
    nr = tbl.add_row(); nc = nr.cells[0].merge(nr.cells[1])
    _set_bg(nc,DVB); _borders(nc,top=DBR,bot=DBR,left=DBR,right=DBR)
    _margins(nc,120,120,160,160); nc.paragraphs[0].clear()
    nc.paragraphs[0].paragraph_format.space_before=Pt(4)
    nc.paragraphs[0].paragraph_format.space_after=Pt(4)
    if data["engineer_notes"]:
        first=True
        for line in data["engineer_notes"].split("\n"):
            if first: _run(nc.paragraphs[0],line,size=10,color=DTX); first=False
            else:
                np2=nc.add_paragraph(); np2.paragraph_format.space_before=Pt(2); np2.paragraph_format.space_after=Pt(2)
                _run(np2,line,size=10,color=DTX)
        for _ in range(8):
            p=nc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
    else:
        _run(nc.paragraphs[0],"Write freely — narrative, insights, hypotheses, observations, anything relevant.",italic=True,color=DPH,size=9.5)
        for _ in range(12):
            p=nc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)

    # 8 — STATUS
    S("8.  STATUS")
    for sl, sv, sc2 in [
        ("SAFETY MECHANISM CHANGED?", data["sm_changed"] or "—",
         "006400" if data["sm_changed"]=="No — false trigger removed only" else
         ("8B0000" if data["sm_changed"]=="Yes — safety behavior modified" else DTX)),
        ("V&V STATUS",       data["vv_status"]       or "—", DTX),
        ("INCIDENT STATUS",  data["incident_status"] or "—",
         DGR if data["incident_status"]=="Closed" else
         (DOR if data["incident_status"]=="Open" else DTX)),
    ]:
        rr2=tbl.add_row(); ra,rb=rr2.cells[0],rr2.cells[1]
        _set_bg(ra,DLB); _borders(ra,top=DBR,bot=DBR,left=DBR,right=DBR)
        _set_bg(rb,DSF); _borders(rb,top=DBR,bot=DBR,left=DBR,right=DBR)
        _margins(ra,80,80,130,130); _margins(rb,80,80,130,130)
        _cp(ra,sl,bold=True,color=DB,size=9.5)
        _cp(rb,sv,bold=(bool(sv) and sv!="—"),color=sc2,size=10)

    # 9 — TECHNICAL SUMMARY
    smr=tbl.add_row(); smc=smr.cells[0].merge(smr.cells[1])
    _set_bg(smc,DN); _borders(smc,top=DN,bot=DN,left=DN,right=DN)
    _margins(smc,120,120,160,160); smc.paragraphs[0].clear()
    lp2=smc.paragraphs[0]; lp2.paragraph_format.space_before=Pt(5); lp2.paragraph_format.space_after=Pt(4)
    _run(lp2,"TECHNICAL SUMMARY",bold=True,color="FFFFFF",size=9.5)
    sp2=smc.add_paragraph(); sp2.paragraph_format.space_before=Pt(4); sp2.paragraph_format.space_after=Pt(6)
    if data["summary"]: _run(sp2,data["summary"],color="D0E8FF",size=10.5)
    else: _run(sp2,"[ Root cause ]  —  [ Safety reaction ]  —  [ Functional effect ]",italic=True,color=DSC,size=10)

    doc.save(path)


# ── ENTRY ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
