#!/usr/bin/env python3
"""FS Incident Log — iOS GUI  |  ISO 26262 Powertrain"""

import sys, json, sqlite3
from datetime import date, datetime
from pathlib import Path
from tkinter import messagebox, filedialog
import tkinter as tk

# ── AUTO-INSTALL ───────────────────────────────────────────────────────────────
def _pip(pkg):
    import subprocess
    for flags in [[], ["--break-system-packages"]]:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", pkg, "-q"] + flags,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
        except Exception:
            continue

try:
    import customtkinter as ctk
except ImportError:
    print("Installing customtkinter..."); _pip("customtkinter")
    import customtkinter as ctk

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Installing python-docx..."); _pip("python-docx")
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── iOS PALETTE ────────────────────────────────────────────────────────────────
BG      = "#F2F2F7"
CARD    = "#FFFFFF"
SIDEBAR = "#EFEFF4"
NAVY    = "#1D2D44"
BLUE    = "#007AFF"
LABEL   = "#1C1C1E"
SEC     = "#8E8E93"
SEP     = "#E5E5EA"
GREEN   = "#34C759"
ORANGE  = "#FF9500"
RED     = "#FF3B30"
LBLUE   = "#EBF4FF"

F = "Segoe UI" if sys.platform == "win32" else ("SF Pro Display" if sys.platform == "darwin" else "Helvetica")

def fnt(size=12, weight="normal"):
    return ctk.CTkFont(family=F, size=size, weight=weight)

# ── DATABASE ───────────────────────────────────────────────────────────────────
DB_FILE = Path(__file__).parent / "fs_incidents.db"

SCHEMA = """CREATE TABLE IF NOT EXISTS incidents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    updated_at TEXT DEFAULT (datetime('now','localtime')),
    ticket_id TEXT DEFAULT '', incident_date TEXT DEFAULT '',
    asil TEXT DEFAULT '', project TEXT DEFAULT '', ecu TEXT DEFAULT '',
    function_topic TEXT DEFAULT '', safety_goal TEXT DEFAULT '',
    sw_appeared TEXT DEFAULT '', sw_fixed TEXT DEFAULT '',
    reported_by TEXT DEFAULT '', cr_id TEXT DEFAULT '',
    powertrain TEXT DEFAULT '[]', op_mode TEXT DEFAULT '[]',
    speed_range TEXT DEFAULT '', thermal_soc TEXT DEFAULT '',
    reproducibility TEXT DEFAULT '', observed TEXT DEFAULT '',
    trigger TEXT DEFAULT '', side_effect TEXT DEFAULT '',
    customer_visible TEXT DEFAULT '', severity TEXT DEFAULT '',
    signals TEXT DEFAULT '[]', detection TEXT DEFAULT '',
    safe_state TEXT DEFAULT '', ftti TEXT DEFAULT '',
    rc_layer TEXT DEFAULT '[]', rc_type TEXT DEFAULT '[]',
    rc_statement TEXT DEFAULT '', resolution TEXT DEFAULT '[]',
    fix_applied TEXT DEFAULT '', engineer_notes TEXT DEFAULT '',
    sm_changed TEXT DEFAULT '', vv_status TEXT DEFAULT '',
    incident_status TEXT DEFAULT 'Open', summary TEXT DEFAULT '')"""

class Database:
    def __init__(self):
        self.con = sqlite3.connect(str(DB_FILE))
        self.con.row_factory = sqlite3.Row
        self.con.execute(SCHEMA); self.con.commit()

    def all(self, q=""):
        if q:
            like = f"%{q}%"
            return self.con.execute(
                "SELECT * FROM incidents WHERE ticket_id LIKE ? OR ecu LIKE ? "
                "OR project LIKE ? OR summary LIKE ? OR incident_status LIKE ? "
                "ORDER BY updated_at DESC", (like,)*5).fetchall()
        return self.con.execute("SELECT * FROM incidents ORDER BY updated_at DESC").fetchall()

    def get(self, id_):
        return self.con.execute("SELECT * FROM incidents WHERE id=?", (id_,)).fetchone()

    def upsert(self, data: dict) -> int:
        data = dict(data); id_ = data.pop("id", None)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["updated_at"] = now
        for k, v in list(data.items()):
            if isinstance(v, list): data[k] = json.dumps(v, ensure_ascii=False)
        if id_:
            sets = ", ".join(f'"{c}"=?' for c in data)
            self.con.execute(f'UPDATE incidents SET {sets} WHERE id=?', list(data.values()) + [id_])
            self.con.commit(); return id_
        data["created_at"] = now
        cols = '","'.join(data); ph = ",".join("?"*len(data))
        self.con.execute(f'INSERT INTO incidents ("{cols}") VALUES ({ph})', list(data.values()))
        self.con.commit()
        return self.con.execute("SELECT last_insert_rowid()").fetchone()[0]

    def delete(self, id_):
        self.con.execute("DELETE FROM incidents WHERE id=?", (id_,)); self.con.commit()


# ── OPTION GROUP ───────────────────────────────────────────────────────────────
class OptionGroup(ctk.CTkFrame):
    def __init__(self, parent, opts, multi=True, cols=3, **kw):
        super().__init__(parent, fg_color=CARD, corner_radius=0, **kw)
        self._multi = multi
        self._opts  = opts
        if multi:
            self._vars = {o: tk.BooleanVar() for o in opts}
            for i, o in enumerate(opts):
                r, c = divmod(i, cols)
                ctk.CTkCheckBox(self, text=o, variable=self._vars[o],
                                fg_color=BLUE, hover_color=NAVY,
                                checkmark_color="white", text_color=LABEL,
                                font=fnt(11)).grid(row=r, column=c, sticky="w", padx=12, pady=4)
        else:
            self._var = tk.StringVar()
            for i, o in enumerate(opts):
                r, c = divmod(i, cols)
                ctk.CTkRadioButton(self, text=o, variable=self._var, value=o,
                                   fg_color=BLUE, hover_color=NAVY,
                                   text_color=LABEL, font=fnt(11)
                                   ).grid(row=r, column=c, sticky="w", padx=12, pady=4)

    def get_value(self):
        if self._multi: return [o for o, v in self._vars.items() if v.get()]
        return self._var.get()

    def set_value(self, val):
        if self._multi:
            for o, v in self._vars.items(): v.set(o in (val or []))
        else:
            self._var.set(val or "")


# ── SIGNAL TABLE ───────────────────────────────────────────────────────────────
class SignalTable(ctk.CTkFrame):
    COLS   = ["Role", "Signal / Function", "Expected", "Observed", "Note"]
    WIDTHS = [110, 165, 120, 120, 155]

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=CARD, corner_radius=0, **kw)
        self._rows = []
        hdr = ctk.CTkFrame(self, fg_color=BLUE, corner_radius=0)
        hdr.pack(fill="x")
        for col, w in zip(self.COLS, self.WIDTHS):
            ctk.CTkLabel(hdr, text=col, text_color="white", font=fnt(9, "bold"),
                         width=w, anchor="w").pack(side="left", padx=4, pady=4)
        self._body = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0)
        self._body.pack(fill="x")
        add_btn = ctk.CTkButton(self, text="+ Add Row", width=110, height=28,
                                fg_color=LBLUE, text_color=BLUE,
                                hover_color=SEP, corner_radius=6,
                                border_width=0, font=fnt(10, "bold"),
                                command=self._add_row)
        add_btn.pack(anchor="w", padx=8, pady=6)
        self._add_row()

    def _add_row(self, prefill=None):
        bg = LBLUE if len(self._rows) % 2 == 0 else CARD
        f = ctk.CTkFrame(self._body, fg_color=bg, corner_radius=0)
        f.pack(fill="x", pady=1)
        keys = ["role", "signal", "expected", "observed", "note"]
        entries = []
        for k, w in zip(keys, self.WIDTHS):
            e = ctk.CTkEntry(f, width=w, height=28, fg_color=bg,
                             border_color=SEP, text_color=LABEL,
                             font=fnt(10), corner_radius=4, border_width=1)
            e.pack(side="left", padx=2, pady=2)
            if prefill and k in prefill: e.insert(0, prefill[k])
            entries.append((k, e))
        ctk.CTkButton(f, text="✕", width=26, height=26, fg_color="transparent",
                      text_color=SEC, hover_color=SEP, corner_radius=4,
                      font=fnt(11), border_width=0,
                      command=lambda fr=f, row=entries: self._del(fr, row)
                      ).pack(side="left", padx=2)
        self._rows.append(entries)

    def _del(self, frame, entries):
        if entries in self._rows: self._rows.remove(entries)
        frame.destroy()

    def get_value(self):
        out = []
        for row in self._rows:
            d = {k: e.get().strip() for k, e in row}
            if any(d.values()): out.append(d)
        return out

    def set_value(self, signals):
        for row in list(self._rows):
            for _, e in row: e.master.destroy()
        self._rows.clear()
        for s in (signals or []): self._add_row(prefill=s)
        if not self._rows: self._add_row()


# ── HELPERS ────────────────────────────────────────────────────────────────────
def sec_hdr(parent, title):
    f = ctk.CTkFrame(parent, fg_color=NAVY, corner_radius=0)
    ctk.CTkLabel(f, text=f"  {title}", text_color="white",
                 font=fnt(11, "bold"), anchor="w").pack(fill="x", padx=8, pady=8)
    return f

def sub_hdr(parent, title):
    f = ctk.CTkFrame(parent, fg_color=BLUE, corner_radius=0)
    ctk.CTkLabel(f, text=f"  {title}", text_color="white",
                 font=fnt(10, "bold"), anchor="w").pack(fill="x", padx=8, pady=5)
    return f

def lbl(parent, text):
    f = ctk.CTkFrame(parent, fg_color=LBLUE, corner_radius=0)
    ctk.CTkLabel(f, text=text, text_color=BLUE,
                 font=fnt(9, "bold"), anchor="w").pack(fill="x", padx=12, pady=5)
    return f

def hsep(parent):
    ctk.CTkFrame(parent, fg_color=SEP, height=1, corner_radius=0).pack(fill="x")


# ── APP ────────────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = Database(); self.current_id = None; self._fields = {}
        self.title("FS Incident Log")
        self.geometry("1300x860"); self.minsize(1000, 660)
        self.configure(fg_color=BG)
        self._build_ui()
        self.refresh_sidebar()
        self._new_incident()

    # ── LAYOUT ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        top = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, height=52)
        top.pack(fill="x"); top.pack_propagate(False)
        ctk.CTkLabel(top, text="FS Incident Log", text_color="white",
                     font=fnt(15, "bold")).pack(side="left", padx=20, pady=12)
        ctk.CTkLabel(top, text="ISO 26262 · Powertrain",
                     text_color=SEC, font=fnt(10)).pack(side="left")
        btn_f = ctk.CTkFrame(top, fg_color="transparent")
        btn_f.pack(side="right", padx=16, pady=8)
        for text, cmd, color in [
            ("Export .docx", self._export, BLUE),
            ("Save",          self._save,   "#2A3F54"),
            ("Delete",        self._delete, RED),
            ("+ New",         self._new_incident, "#2A3F54"),
        ]:
            ctk.CTkButton(btn_f, text=text, command=cmd,
                          fg_color=color, hover_color=NAVY,
                          corner_radius=8, height=32, border_width=0,
                          font=fnt(10, "bold")).pack(side="right", padx=4)

        body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        body.pack(fill="both", expand=True)

        # Sidebar
        sb = ctk.CTkFrame(body, fg_color=SIDEBAR, corner_radius=0, width=268)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)
        self._build_sidebar(sb)

        ctk.CTkFrame(body, fg_color=SEP, width=1, corner_radius=0).pack(side="left", fill="y")

        # Form
        self._form = ctk.CTkScrollableFrame(body, fg_color=BG, corner_radius=0,
                                            scrollbar_button_color=SEP,
                                            scrollbar_button_hover_color=SEC)
        self._form.pack(side="left", fill="both", expand=True)
        self._build_form(self._form)

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    def _build_sidebar(self, parent):
        ctk.CTkLabel(parent, text="ARCHIVE", text_color=SEC,
                     font=fnt(9, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
        self._search = ctk.CTkEntry(parent, placeholder_text="Search...",
                                    fg_color=CARD, border_color=SEP,
                                    text_color=LABEL, font=fnt(10),
                                    corner_radius=8, height=32, border_width=1)
        self._search.pack(fill="x", padx=10, pady=(0, 8))
        self._search.bind("<KeyRelease>", lambda e: self.refresh_sidebar())
        self._list = ctk.CTkScrollableFrame(parent, fg_color=SIDEBAR, corner_radius=0,
                                            scrollbar_button_color=SEP)
        self._list.pack(fill="both", expand=True)

    def refresh_sidebar(self):
        q = self._search.get() if hasattr(self, "_search") else ""
        for w in self._list.winfo_children(): w.destroy()
        rows = self.db.all(q)
        for row in rows: self._sidebar_item(row)
        if not rows:
            ctk.CTkLabel(self._list, text="No incidents yet",
                         text_color=SEC, font=fnt(10)).pack(pady=20)

    def _sidebar_item(self, row):
        id_  = row["id"]; sel = self.current_id == id_
        bg   = BLUE if sel else CARD
        fg   = "white" if sel else LABEL
        sfg  = "white" if sel else SEC

        item = ctk.CTkFrame(self._list, fg_color=bg, corner_radius=8, cursor="hand2")
        item.pack(fill="x", padx=6, pady=3)

        top = ctk.CTkFrame(item, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(top, text=row["ticket_id"] or "—",
                     text_color=fg, font=fnt(11, "bold"), anchor="w").pack(side="left")
        dt = (row["incident_date"] or row["created_at"] or "")[:10]
        ctk.CTkLabel(top, text=dt, text_color=sfg, font=fnt(9)).pack(side="right")

        bot = ctk.CTkFrame(item, fg_color="transparent")
        bot.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(bot, text=row["ecu"] or row["project"] or "—",
                     text_color=sfg, font=fnt(9)).pack(side="left")
        status = row["incident_status"] or "Open"
        sc = "white" if sel else {"Closed": GREEN, "Open": ORANGE, "In Progress": BLUE}.get(status, SEC)
        ctk.CTkLabel(bot, text=status, text_color=sc, font=fnt(8, "bold")).pack(side="right")

        for w in self._descendants(item):
            w.bind("<Button-1>", lambda e, i=id_: self._load(i))
        item.bind("<Button-1>", lambda e, i=id_: self._load(i))

    def _descendants(self, w):
        result = []
        for child in w.winfo_children():
            result.append(child); result.extend(self._descendants(child))
        return result

    # ── FORM ──────────────────────────────────────────────────────────────────
    def _build_form(self, parent):

        def card(pady_top=0):
            f = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=10)
            f.pack(fill="x", padx=16, pady=(pady_top, 0))
            return f

        def entry(p, key, label, hint=""):
            lbl(p, label).pack(fill="x")
            e = ctk.CTkEntry(p, placeholder_text=hint, fg_color=CARD,
                             border_color=SEP, text_color=LABEL,
                             font=fnt(11), corner_radius=0, height=36, border_width=0)
            e.pack(fill="x", padx=10, pady=4)
            hsep(p); self._fields[key] = e

        def textarea(p, key, label, height=90):
            lbl(p, label).pack(fill="x")
            t = ctk.CTkTextbox(p, fg_color=CARD, border_color=SEP, text_color=LABEL,
                               font=fnt(11), corner_radius=6, height=height,
                               border_width=1, wrap="word")
            t.pack(fill="x", padx=10, pady=4)
            hsep(p); self._fields[key] = t

        def opts(p, key, options, multi=True, cols=3):
            og = OptionGroup(p, options, multi=multi, cols=cols)
            og.pack(fill="x", padx=2, pady=4)
            hsep(p); self._fields[key] = og

        def radio(p, key, label, options):
            lbl(p, label).pack(fill="x")
            og = OptionGroup(p, options, multi=False, cols=len(options))
            og.pack(fill="x", padx=2, pady=4)
            hsep(p); self._fields[key] = og

        # ── 1. IDENTIFICATION ─────────────────────────────────────────────────
        sec_hdr(parent, "1.  IDENTIFICATION").pack(fill="x", padx=16, pady=(16, 0))
        c = card()
        quad = ctk.CTkFrame(c, fg_color=CARD, corner_radius=0)
        quad.pack(fill="x")
        for i, (key, label, hint) in enumerate([
            ("ticket_id",     "TICKET / ID",        "e.g. JIRA-1234"),
            ("incident_date", "DATE",                str(date.today())),
            ("asil",          "ASIL",                "A / B / C / D / QM"),
            ("project",       "PROJECT / PLATFORM",  "e.g. MEB, PPE, C3"),
        ]):
            cell = ctk.CTkFrame(quad, fg_color=CARD, corner_radius=0)
            cell.grid(row=i//2, column=i%2, sticky="nsew", padx=1, pady=1)
            quad.columnconfigure(i%2, weight=1)
            lbl(cell, label).pack(fill="x")
            e = ctk.CTkEntry(cell, placeholder_text=hint, fg_color=CARD,
                             border_color=SEP, text_color=LABEL, font=fnt(11),
                             corner_radius=0, height=36, border_width=0)
            e.pack(fill="x", padx=10, pady=4)
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
        sec_hdr(parent, "2.  CONTEXT").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        sub_hdr(c, "Powertrain Type").pack(fill="x")
        opts(c, "powertrain", ["ICE","MHEV","HEV","PHEV","BEV","FCEV"], cols=6)
        sub_hdr(c, "Operating Mode").pack(fill="x")
        opts(c, "op_mode", ["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"], cols=3)
        entry(c, "speed_range", "SPEED RANGE",   "e.g. standstill / <30 km/h / highway")
        entry(c, "thermal_soc", "THERMAL / SOC", "e.g. cold start, warm, low SOC")
        radio(c, "reproducibility", "REPRODUCIBILITY", ["Always","Intermittent","Rare"])

        # ── 3. INCIDENT ───────────────────────────────────────────────────────
        sec_hdr(parent, "3.  INCIDENT").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        textarea(c, "observed",    "OBSERVED BEHAVIOR", height=90)
        textarea(c, "trigger",     "TRIGGER CONDITION", height=90)
        textarea(c, "side_effect", "SIDE EFFECT",       height=70)
        row2 = ctk.CTkFrame(c, fg_color=CARD, corner_radius=0)
        row2.pack(fill="x")
        for i, (key, label, options) in enumerate([
            ("customer_visible", "CUSTOMER VISIBLE?", ["Yes","No"]),
            ("severity",         "SEVERITY",           ["Safety-relevant","Functional","Comfort"]),
        ]):
            cell = ctk.CTkFrame(row2, fg_color=CARD, corner_radius=0)
            cell.grid(row=0, column=i, sticky="nsew", padx=1)
            row2.columnconfigure(i, weight=1)
            lbl(cell, label).pack(fill="x")
            og = OptionGroup(cell, options, multi=False, cols=len(options))
            og.pack(fill="x", padx=2, pady=4)
            self._fields[key] = og

        # ── 4. SIGNAL CONFLICT ────────────────────────────────────────────────
        sec_hdr(parent, "4.  SIGNAL CONFLICT").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        lbl(c, "CONFLICT TABLE").pack(fill="x")
        sf = ctk.CTkFrame(c, fg_color=CARD, corner_radius=0)
        sf.pack(fill="x", padx=10, pady=8)
        self._signal_table = SignalTable(sf)
        self._signal_table.pack(fill="x")
        self._fields["signals"] = self._signal_table
        hsep(c)

        # ── 5. SAFETY RESPONSE ────────────────────────────────────────────────
        sec_hdr(parent, "5.  SAFETY RESPONSE").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        textarea(c, "detection",  "DETECTION FUNCTION", height=70)
        textarea(c, "safe_state", "SAFE STATE ENTERED",  height=70)
        radio(c, "ftti", "FTTI RESPECTED?", ["Yes","No — explain in analysis","Not assessed"])

        # ── 6. ROOT CAUSE & FIX ───────────────────────────────────────────────
        sec_hdr(parent, "6.  ROOT CAUSE & FIX").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        sub_hdr(c, "Root Cause Layer").pack(fill="x")
        opts(c, "rc_layer", ["SW Logic","Signal Definition","Calibration",
             "Interface / Communication","Timing / Scheduling","Architecture"], cols=3)
        sub_hdr(c, "Root Cause Type").pack(fill="x")
        opts(c, "rc_type", ["Incorrect rule","Missing condition","Wrong threshold",
             "State machine error","Signal misinterpretation","Missing debounce",
             "Race condition","Incomplete spec"], cols=4)
        textarea(c, "rc_statement", "ROOT CAUSE STATEMENT", height=70)
        sub_hdr(c, "Resolution Category").pack(fill="x")
        opts(c, "resolution", ["Logic refinement","State reclassification","Timing / debounce",
             "Architectural change","Calibration correction","Signal definition fix",
             "Spec update only","No fix — accepted risk"], cols=4)
        textarea(c, "fix_applied", "FIX APPLIED", height=70)

        # ── 7. ENGINEER SPACE ─────────────────────────────────────────────────
        sec_hdr(parent, "7.  ENGINEER SPACE").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        textarea(c, "engineer_notes", "NOTES & ANALYSIS", height=160)

        # ── 8. STATUS ─────────────────────────────────────────────────────────
        sec_hdr(parent, "8.  STATUS").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        radio(c, "sm_changed",      "SAFETY MECHANISM CHANGED?",
              ["No — false trigger removed only","Yes — safety behavior modified"])
        radio(c, "vv_status",       "V&V STATUS",       ["Pending","In Progress","Complete","Waived"])
        radio(c, "incident_status", "INCIDENT STATUS",  ["Open","In Progress","Closed","Monitoring"])

        # ── 9. TECHNICAL SUMMARY ──────────────────────────────────────────────
        sec_hdr(parent, "9.  TECHNICAL SUMMARY").pack(fill="x", padx=16, pady=(12, 0))
        c = card()
        textarea(c, "summary", "ONE-SENTENCE SUMMARY\n[ Root cause ] — [ Safety reaction ] — [ Functional effect ]", height=70)
        ctk.CTkFrame(parent, fg_color=BG, height=40, corner_radius=0).pack()

    # ── DATA ACCESS ───────────────────────────────────────────────────────────
    def _get_form(self) -> dict:
        data = {}
        for key, w in self._fields.items():
            if isinstance(w, ctk.CTkEntry):
                data[key] = w.get().strip()
            elif isinstance(w, ctk.CTkTextbox):
                data[key] = w.get("1.0", "end-1c").strip()
            elif isinstance(w, (OptionGroup, SignalTable)):
                data[key] = w.get_value()
        return data

    def _set_form(self, data: dict):
        for key, w in self._fields.items():
            raw = data.get(key, "")
            if isinstance(raw, str) and raw.startswith(("[","{")):
                try: raw = json.loads(raw)
                except: pass
            if isinstance(w, ctk.CTkEntry):
                w.delete(0, "end")
                if raw: w.insert(0, str(raw))
            elif isinstance(w, ctk.CTkTextbox):
                w.delete("1.0", "end")
                if raw: w.insert("1.0", str(raw))
            elif isinstance(w, (OptionGroup, SignalTable)):
                w.set_value(raw)

    def _clear_form(self):
        for key, w in self._fields.items():
            if isinstance(w, ctk.CTkEntry): w.delete(0, "end")
            elif isinstance(w, ctk.CTkTextbox): w.delete("1.0", "end")
            elif isinstance(w, OptionGroup): w.set_value([] if w._multi else "")
            elif isinstance(w, SignalTable): w.set_value([])
        if "incident_date" in self._fields:
            self._fields["incident_date"].insert(0, str(date.today()))

    # ── ACTIONS ───────────────────────────────────────────────────────────────
    def _new_incident(self):
        self.current_id = None; self._clear_form(); self.refresh_sidebar()

    def _load(self, id_):
        row = self.db.get(id_)
        if not row: return
        self.current_id = id_; self._set_form(dict(row)); self.refresh_sidebar()

    def _save(self):
        data = self._get_form(); data["id"] = self.current_id
        self.current_id = self.db.upsert(data)
        self.refresh_sidebar()
        self.title("FS Incident Log  ✓  Saved")
        self.after(2000, lambda: self.title("FS Incident Log"))

    def _delete(self):
        if not self.current_id:
            messagebox.showinfo("Nothing selected", "Open an incident first."); return
        if messagebox.askyesno("Delete", "Delete this incident permanently?"):
            self.db.delete(self.current_id)
            self.current_id = None; self._clear_form(); self.refresh_sidebar()

    def _export(self):
        if not self.current_id: self._save()
        data = self._get_form()
        for k in ["powertrain","op_mode","rc_layer","rc_type","resolution"]:
            if isinstance(data.get(k), str):
                try: data[k] = json.loads(data[k])
                except: data[k] = []
        sigs = data.get("signals", [])
        if isinstance(sigs, str):
            try: sigs = json.loads(sigs)
            except: sigs = []

        doc_data = {
            "ticket_id": data.get("ticket_id",""),       "date": data.get("incident_date", str(date.today())),
            "asil":      data.get("asil",""),            "project":      data.get("project",""),
            "ecu":       data.get("ecu",""),             "function":     data.get("function_topic",""),
            "safety_goal":  data.get("safety_goal",""),  "sw_appeared":  data.get("sw_appeared",""),
            "sw_fixed":     data.get("sw_fixed",""),     "reported_by":  data.get("reported_by",""),
            "cr_id":        data.get("cr_id",""),        "powertrain":   data.get("powertrain",[]),
            "op_mode":      data.get("op_mode",[]),      "speed_range":  data.get("speed_range",""),
            "thermal_soc":  data.get("thermal_soc",""),  "reproducibility": data.get("reproducibility",""),
            "observed":     data.get("observed",""),     "trigger":      data.get("trigger",""),
            "side_effect":  data.get("side_effect",""),  "customer_visible": data.get("customer_visible",""),
            "severity":     data.get("severity",""),     "signals":      sigs,
            "detection":    data.get("detection",""),    "safe_state":   data.get("safe_state",""),
            "ftti":         data.get("ftti",""),         "rc_layer":     data.get("rc_layer",[]),
            "rc_type":      data.get("rc_type",[]),      "rc_statement": data.get("rc_statement",""),
            "resolution":   data.get("resolution",[]),   "fix_applied":  data.get("fix_applied",""),
            "engineer_notes": data.get("engineer_notes",""), "attachments": [],
            "sm_changed":   data.get("sm_changed",""),   "vv_status":    data.get("vv_status",""),
            "incident_status": data.get("incident_status",""), "summary": data.get("summary",""),
        }
        ticket  = doc_data["ticket_id"] or "incident"
        default = f"FS_Incident_{ticket}_{date.today().strftime('%Y%m%d')}.docx"
        path = filedialog.asksaveasfilename(defaultextension=".docx",
                                            filetypes=[("Word Document","*.docx")],
                                            initialfile=default)
        if not path: return
        try:
            build_doc(doc_data, path)
            self.title("FS Incident Log  ✓  Exported")
            self.after(2500, lambda: self.title("FS Incident Log"))
        except Exception as ex:
            messagebox.showerror("Export failed", str(ex))


# ── DOCX BUILDER ───────────────────────────────────────────────────────────────
DN="1D2D44"; DB="0A84FF"; DLB="F0F4FF"; DVB="FFFFFF"; DBR="E2E8F0"
DTX="1C1C1E"; DSC="6E6E73"; DPH="AEAEB2"; DGR="34C759"; DOR="FF9500"; DSF="F5F5F7"

def _rgb(h):
    h=h.lstrip("#"); return tuple(int(h[i:i+2],16) for i in (0,2,4))

def _bg(cell,color):
    tc=cell._tc; p=tc.get_or_add_tcPr()
    for s in p.findall(qn("w:shd")): p.remove(s)
    s=OxmlElement("w:shd"); s.set(qn("w:val"),"clear"); s.set(qn("w:color"),"auto")
    s.set(qn("w:fill"),color.lstrip("#").upper()); p.append(s)

def _bdr(cell,top=None,bot=None,left=None,right=None):
    tc=cell._tc; p=tc.get_or_add_tcPr()
    for b in p.findall(qn("w:tcBorders")): p.remove(b)
    bd=OxmlElement("w:tcBorders")
    for edge,c in [("top",top),("bottom",bot),("start",left),("end",right)]:
        el=OxmlElement(f"w:{edge}")
        if c: el.set(qn("w:val"),"single"); el.set(qn("w:sz"),"4"); el.set(qn("w:color"),c.lstrip("#").upper())
        else: el.set(qn("w:val"),"nil")
        bd.append(el)
    sh=p.find(qn("w:shd"))
    if sh is not None: sh.addprevious(bd)
    else: p.append(bd)

def _mar(cell,t=80,b=80,l=120,r=120):
    tc=cell._tc; p=tc.get_or_add_tcPr()
    for m in p.findall(qn("w:tcMar")): p.remove(m)
    mar=OxmlElement("w:tcMar")
    for edge,val in [("top",t),("bottom",b),("start",l),("end",r)]:
        el=OxmlElement(f"w:{edge}"); el.set(qn("w:w"),str(val)); el.set(qn("w:type"),"dxa"); mar.append(el)
    p.append(mar)

def _run(para,text,bold=False,italic=False,color=DTX,size=10):
    run=para.add_run(text); run.bold=bold; run.italic=italic
    run.font.name="Arial"; run.font.size=Pt(size)
    r,g,b=_rgb(color); run.font.color.rgb=RGBColor(r,g,b); return run

def _cp(cell,text,bold=False,italic=False,color=DTX,size=10,
        align=WD_ALIGN_PARAGRAPH.LEFT,sb=60,sa=60,clear=True):
    if clear: cell.paragraphs[0].clear(); para=cell.paragraphs[0]
    else: para=cell.add_paragraph()
    para.alignment=align
    para.paragraph_format.space_before=Pt(sb/20); para.paragraph_format.space_after=Pt(sa/20)
    if text: _run(para,text,bold=bold,italic=italic,color=color,size=size)
    return para

def _mrow(tbl,fill,text,bold=True,color="FFFFFF",size=10.5):
    row=tbl.add_row(); cell=row.cells[0]
    for i in range(1,len(row.cells)): cell=cell.merge(row.cells[i])
    _bg(cell,fill); _bdr(cell,top=fill,bot=fill,left=fill,right=fill)
    _mar(cell,70,70,140,140); _cp(cell,text,bold=bold,color=color,size=size)

def _lv(tbl,label,value):
    row=tbl.add_row(); c0,c1=row.cells[0],row.cells[1]
    for i in range(2,len(row.cells)): c1=c1.merge(row.cells[i])
    _bg(c0,DLB); _bdr(c0,top=DBR,bot=DBR,left=DBR,right=DBR)
    _bg(c1,DVB); _bdr(c1,bot=DBR)
    _mar(c0,80,80,130,130); _mar(c1,80,80,130,130)
    _cp(c0,label,bold=True,color=DB,size=9.5)
    _cp(c1,value,italic=not bool(value),color=DPH if not value else DTX,size=10)

def _cb(tbl,items,selected):
    row=tbl.add_row(); cell=row.cells[0]
    for i in range(1,len(row.cells)): cell=cell.merge(row.cells[i])
    _bg(cell,DVB); _bdr(cell,bot=DBR); _mar(cell,60,60,130,130)
    para=cell.paragraphs[0]; para.clear()
    para.paragraph_format.space_before=Pt(3); para.paragraph_format.space_after=Pt(3)
    for j,item in enumerate(items):
        if j>0: _run(para,"     ",size=10,color=DTX)
        tick="☑" if item in selected else "☐"
        _run(para,f"{tick}  {item}",size=10,color=DB if item in selected else DSC,bold=(item in selected))

def _ta(tbl,text,min_lines=5):
    row=tbl.add_row(); cell=row.cells[0]
    for i in range(1,len(row.cells)): cell=cell.merge(row.cells[i])
    _bg(cell,DVB); _bdr(cell,bot=DBR); _mar(cell,100,100,140,140)
    para=cell.paragraphs[0]; para.clear()
    para.paragraph_format.space_before=Pt(4); para.paragraph_format.space_after=Pt(4)
    if text:
        for i,line in enumerate(text.split("\n")):
            if i==0: _run(para,line,size=10,color=DTX)
            else:
                np=cell.add_paragraph(); np.paragraph_format.space_before=Pt(2); np.paragraph_format.space_after=Pt(2)
                _run(np,line,size=10,color=DTX)
    used=max(1,len(text.split("\n")) if text else 0)
    for _ in range(max(0,min_lines-used)):
        p=cell.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)

def build_doc(data,path):
    doc=Document(); sec=doc.sections[0]
    sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.left_margin=Cm(1.5); sec.right_margin=Cm(1.5)
    sec.top_margin=Cm(1.2);  sec.bottom_margin=Cm(1.2)
    hp=doc.sections[0].header.paragraphs[0]; hp.clear()
    hp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    _run(hp,"FUNCTIONAL SAFETY INCIDENT LOG  —  ISO 26262 Powertrain",bold=True,color=DN,size=8)
    _run(hp,"          CONFIDENTIAL — INTERNAL USE ONLY",italic=True,color=DSC,size=8)
    hp.paragraph_format.space_after=Pt(6)
    tp=doc.add_paragraph(); tp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before=Pt(0); tp.paragraph_format.space_after=Pt(4)
    _run(tp,"FS INCIDENT LOG",bold=True,color=DN,size=20)
    sp=doc.add_paragraph(); sp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    sp.paragraph_format.space_before=Pt(0); sp.paragraph_format.space_after=Pt(14)
    _run(sp,"ISO 26262  —  Powertrain  —  v3.0",italic=True,color=DB,size=10)
    CW=int(18*567); C1=int(3.5*567); C2=CW-C1
    tbl=doc.add_table(rows=0,cols=2); tbl.style="Table Grid"; tbl.autofit=False
    tg=tbl._tbl.find(qn("w:tblGrid"))
    if tg is not None:
        for gc in tg.findall(qn("w:gridCol")): tg.remove(gc)
        for w in [C1,C2]:
            gc=OxmlElement("w:gridCol"); gc.set(qn("w:w"),str(w)); tg.append(gc)
    tpr=tbl._tbl.find(qn("w:tblPr"))
    if tpr is None: tpr=OxmlElement("w:tblPr"); tbl._tbl.insert(0,tpr)
    tw=OxmlElement("w:tblW"); tw.set(qn("w:w"),str(CW)); tw.set(qn("w:type"),"dxa"); tpr.insert(0,tw)
    def S(t): _mrow(tbl,DN,f"  {t}",bold=True,color="FFFFFF",size=10.5)
    def U(t): _mrow(tbl,DB,f"  {t}",bold=True,color="FFFFFF",size=9.5)
    # 1 — IDENTIFICATION
    S("1.  IDENTIFICATION")
    r=tbl.add_row(); ca=r.cells[0].merge(r.cells[1])
    _bg(ca,DVB); _bdr(ca,top=DBR,bot=DBR,left=DBR,right=DBR); _mar(ca,0,0,0,0)
    inn=OxmlElement("w:tbl"); ip=OxmlElement("w:tblPr"); iw=OxmlElement("w:tblW")
    iw.set(qn("w:w"),str(CW)); iw.set(qn("w:type"),"dxa"); ip.append(iw)
    ib=OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e=OxmlElement(f"w:{edge}"); e.set(qn("w:val"),"single"); e.set(qn("w:sz"),"4"); e.set(qn("w:color"),DBR); ib.append(e)
    ip.append(ib); inn.append(ip)
    ig=OxmlElement("w:tblGrid"); qw=CW//4
    for _ in range(4):
        gc=OxmlElement("w:gridCol"); gc.set(qn("w:w"),str(qw)); ig.append(gc)
    inn.append(ig); ir=OxmlElement("w:tr")
    for lt,val in [("TICKET / ID",data["ticket_id"]),("DATE",data["date"]),("ASIL",data["asil"]),("PROJECT / PLATFORM",data["project"])]:
        tc=OxmlElement("w:tc"); tp2=OxmlElement("w:tcPr"); tw2=OxmlElement("w:tcW")
        tw2.set(qn("w:w"),str(qw)); tw2.set(qn("w:type"),"dxa"); tp2.append(tw2)
        sh=OxmlElement("w:shd"); sh.set(qn("w:val"),"clear"); sh.set(qn("w:color"),"auto"); sh.set(qn("w:fill"),DVB); tp2.append(sh)
        m2=OxmlElement("w:tcMar")
        for edge in ["top","bottom","left","right"]:
            me=OxmlElement(f"w:{edge}"); me.set(qn("w:w"),"100"); me.set(qn("w:type"),"dxa"); m2.append(me)
        tp2.append(m2); tc.append(tp2)
        lp=OxmlElement("w:p"); lr=OxmlElement("w:r"); lrp=OxmlElement("w:rPr")
        lb=OxmlElement("w:b"); lc=OxmlElement("w:color"); lc.set(qn("w:val"),DB)
        ls=OxmlElement("w:sz"); ls.set(qn("w:val"),"17")
        lrp.append(lb); lrp.append(lc); lrp.append(ls); lr.append(lrp)
        ltt=OxmlElement("w:t"); ltt.text=lt; lr.append(ltt); lp.append(lr); tc.append(lp)
        vp=OxmlElement("w:p"); vr=OxmlElement("w:r"); vrp=OxmlElement("w:rPr")
        vc=OxmlElement("w:color"); vc.set(qn("w:val"),DTX if val else DPH)
        vs=OxmlElement("w:sz"); vs.set(qn("w:val"),"20")
        if not val: vi=OxmlElement("w:i"); vrp.append(vi)
        vrp.append(vc); vrp.append(vs); vr.append(vrp)
        vt=OxmlElement("w:t"); vt.text=val or "—"; vr.append(vt); vp.append(vr); tc.append(vp); ir.append(tc)
    inn.append(ir); ca._tc.append(inn)
    _lv(tbl,"ECU / SYSTEM",data["ecu"]); _lv(tbl,"FUNCTION / TOPIC",data["function"])
    _lv(tbl,"SAFETY GOAL REF.",data["safety_goal"]); _lv(tbl,"SW — APPEARED IN",data["sw_appeared"])
    _lv(tbl,"SW — FIXED IN",data["sw_fixed"]); _lv(tbl,"REPORTED BY",data["reported_by"])
    _lv(tbl,"CHANGE REQUEST (CR)",data["cr_id"])
    # 2 — CONTEXT
    S("2.  CONTEXT")
    U("Powertrain Type"); _cb(tbl,["ICE","MHEV","HEV","PHEV","BEV","FCEV"],data["powertrain"])
    U("Operating Mode");  _cb(tbl,["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"],data["op_mode"])
    _lv(tbl,"SPEED RANGE",data["speed_range"]); _lv(tbl,"THERMAL / SOC",data["thermal_soc"])
    rr=tbl.add_row(); rc0,rc1=rr.cells[0],rr.cells[1]
    _bg(rc0,DLB); _bdr(rc0,top=DBR,bot=DBR,left=DBR,right=DBR)
    _bg(rc1,DVB); _bdr(rc1,bot=DBR); _mar(rc0,80,80,130,130); _mar(rc1,60,60,130,130)
    _cp(rc0,"REPRODUCIBILITY",bold=True,color=DB,size=9.5)
    pr=rc1.paragraphs[0]; pr.clear(); pr.paragraph_format.space_before=Pt(3); pr.paragraph_format.space_after=Pt(3)
    for j,opt in enumerate(["Always","Intermittent","Rare"]):
        if j>0: _run(pr,"     ",size=10,color=DTX)
        sel=data["reproducibility"]==opt
        _run(pr,f"{'☑' if sel else '☐'}  {opt}",size=10,color=DB if sel else DSC,bold=sel)
    # 3 — INCIDENT
    S("3.  INCIDENT")
    _lv(tbl,"OBSERVED BEHAVIOR",""); _ta(tbl,data["observed"])
    _lv(tbl,"TRIGGER CONDITION",""); _ta(tbl,data["trigger"])
    _lv(tbl,"SIDE EFFECT","");       _ta(tbl,data["side_effect"])
    rcv=tbl.add_row(); cc0,cc1=rcv.cells[0],rcv.cells[1]
    for cx,lbl_t,items,sel_ in [
        (cc0,"CUSTOMER VISIBLE?",["Yes","No"],[data["customer_visible"]] if data["customer_visible"] else []),
        (cc1,"SEVERITY",["Safety-relevant","Functional","Comfort"],[data["severity"]] if data["severity"] else []),
    ]:
        _bg(cx,DVB); _bdr(cx,top=DBR,bot=DBR,left=DBR,right=DBR); _mar(cx,80,80,130,130)
        _cp(cx,lbl_t,bold=True,color=DB,size=9.5)
        p2=cx.add_paragraph(); p2.paragraph_format.space_before=Pt(3); p2.paragraph_format.space_after=Pt(3)
        for j,opt in enumerate(items):
            if j>0: _run(p2,"     ",size=10,color=DTX)
            sel=opt in sel_; _run(p2,f"{'☑' if sel else '☐'}  {opt}",size=10,color=DB if sel else DSC,bold=sel)
    # 4 — SIGNAL CONFLICT
    S("4.  SIGNAL CONFLICT")
    sr2=tbl.add_row(); ss0,ss1=sr2.cells[0],sr2.cells[1]
    _bg(ss0,DLB); _bdr(ss0,top=DBR,bot=DBR,left=DBR,right=DBR); _mar(ss0,80,80,130,130)
    _cp(ss0,"CONFLICT TABLE",bold=True,color=DB,size=9.5)
    _bg(ss1,DVB); _bdr(ss1,top=DBR,bot=DBR,left=DBR,right=DBR); _mar(ss1,80,80,100,100)
    sc=["Role","Signal / Function","Expected","Observed","Note"]
    sw=[int(C2*.14),int(C2*.24),int(C2*.18),int(C2*.18),C2-int(C2*.14)-int(C2*.24)-int(C2*.18)*2]
    st=OxmlElement("w:tbl"); stp=OxmlElement("w:tblPr")
    stw=OxmlElement("w:tblW"); stw.set(qn("w:w"),str(C2)); stw.set(qn("w:type"),"dxa"); stp.append(stw)
    stb=OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e=OxmlElement(f"w:{edge}"); e.set(qn("w:val"),"single"); e.set(qn("w:sz"),"3"); e.set(qn("w:color"),DBR); stb.append(e)
    stp.append(stb); st.append(stp); sg=OxmlElement("w:tblGrid")
    for w in sw: gc=OxmlElement("w:gridCol"); gc.set(qn("w:w"),str(w)); sg.append(gc)
    st.append(sg)
    def _str(values,is_hdr=False,rfill=DVB):
        tr=OxmlElement("w:tr")
        for val,w in zip(values,sw):
            tc=OxmlElement("w:tc"); tp3=OxmlElement("w:tcPr"); tw3=OxmlElement("w:tcW")
            tw3.set(qn("w:w"),str(w)); tw3.set(qn("w:type"),"dxa"); tp3.append(tw3)
            sh3=OxmlElement("w:shd"); sh3.set(qn("w:val"),"clear"); sh3.set(qn("w:color"),"auto")
            sh3.set(qn("w:fill"),DB if is_hdr else rfill); tp3.append(sh3)
            m3=OxmlElement("w:tcMar")
            for edge in ["top","bottom","left","right"]:
                me=OxmlElement(f"w:{edge}"); me.set(qn("w:w"),"80"); me.set(qn("w:type"),"dxa"); m3.append(me)
            tp3.append(m3); tc.append(tp3)
            wp3=OxmlElement("w:p"); wr3=OxmlElement("w:r"); wrp3=OxmlElement("w:rPr")
            if is_hdr: wb=OxmlElement("w:b"); wrp3.append(wb)
            wc3=OxmlElement("w:color"); wc3.set(qn("w:val"),"FFFFFF" if is_hdr else DTX); wrp3.append(wc3)
            ws3=OxmlElement("w:sz"); ws3.set(qn("w:val"),"17"); wrp3.append(ws3); wr3.append(wrp3)
            wt3=OxmlElement("w:t"); wt3.text=val or ""; wr3.append(wt3); wp3.append(wr3); tc.append(wp3); tr.append(tc)
        return tr
    st.append(_str(sc,is_hdr=True))
    rfills=[DLB,DVB,"EBF5EE",DVB,"FDF8EC",DVB]; rlbls=["Driver Intent","","Physical State","","System State",""]
    sigs=data["signals"]
    for ri in range(6):
        if ri<len(sigs): s=sigs[ri]; vals=[s.get("role",""),s.get("signal",""),s.get("expected",""),s.get("observed",""),s.get("note","")]
        else: vals=[rlbls[ri],"","","",""]
        st.append(_str(vals,rfill=rfills[ri]))
    for s in sigs[6:]: st.append(_str([s.get("role",""),s.get("signal",""),s.get("expected",""),s.get("observed",""),s.get("note","")]))
    ss1._tc.append(st)
    # 5 — SAFETY RESPONSE
    S("5.  SAFETY RESPONSE")
    _lv(tbl,"DETECTION FUNCTION",""); _ta(tbl,data["detection"])
    _lv(tbl,"SAFE STATE ENTERED","");  _ta(tbl,data["safe_state"])
    rf=tbl.add_row(); rf0,rf1=rf.cells[0],rf.cells[1]
    _bg(rf0,DLB); _bdr(rf0,top=DBR,bot=DBR,left=DBR,right=DBR)
    _bg(rf1,DVB); _bdr(rf1,bot=DBR); _mar(rf0,80,80,130,130); _mar(rf1,60,60,130,130)
    _cp(rf0,"FTTI RESPECTED?",bold=True,color=DB,size=9.5)
    pf=rf1.paragraphs[0]; pf.clear(); pf.paragraph_format.space_before=Pt(3); pf.paragraph_format.space_after=Pt(3)
    for j,opt in enumerate(["Yes","No — explain in analysis","Not assessed"]):
        if j>0: _run(pf,"          ",size=10,color=DTX)
        sel=data["ftti"]==opt; _run(pf,f"{'☑' if sel else '☐'}  {opt}",size=10,color=DB if sel else DSC,bold=sel)
    # 6 — ROOT CAUSE & FIX
    S("6.  ROOT CAUSE & FIX")
    U("Root Cause Layer"); _cb(tbl,["SW Logic","Signal Definition","Calibration","Interface / Communication","Timing / Scheduling","Architecture"],data["rc_layer"])
    U("Root Cause Type");  _cb(tbl,["Incorrect rule","Missing condition","Wrong threshold","State machine error","Signal misinterpretation","Missing debounce","Race condition","Incomplete spec"],data["rc_type"])
    _lv(tbl,"ROOT CAUSE STATEMENT",""); _ta(tbl,data["rc_statement"],min_lines=4)
    U("Resolution Category"); _cb(tbl,["Logic refinement","State reclassification","Timing / debounce","Architectural change","Calibration correction","Signal definition fix","Spec update only","No fix — accepted risk"],data["resolution"])
    _lv(tbl,"FIX APPLIED",""); _ta(tbl,data["fix_applied"])
    # 7 — ENGINEER SPACE
    S("7.  ENGINEER SPACE")
    nr=tbl.add_row(); nc=nr.cells[0].merge(nr.cells[1])
    _bg(nc,DVB); _bdr(nc,top=DBR,bot=DBR,left=DBR,right=DBR); _mar(nc,120,120,160,160)
    nc.paragraphs[0].clear(); nc.paragraphs[0].paragraph_format.space_before=Pt(4); nc.paragraphs[0].paragraph_format.space_after=Pt(4)
    if data["engineer_notes"]:
        first=True
        for line in data["engineer_notes"].split("\n"):
            if first: _run(nc.paragraphs[0],line,size=10,color=DTX); first=False
            else:
                np2=nc.add_paragraph(); np2.paragraph_format.space_before=Pt(2); np2.paragraph_format.space_after=Pt(2); _run(np2,line,size=10,color=DTX)
        for _ in range(8):
            p=nc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
    else:
        _run(nc.paragraphs[0],"Write freely — narrative, insights, hypotheses, observations, anything relevant.",italic=True,color=DPH,size=9.5)
        for _ in range(12):
            p=nc.add_paragraph(); p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
    # 8 — STATUS
    S("8.  STATUS")
    for sl,sv,sc2 in [
        ("SAFETY MECHANISM CHANGED?",data["sm_changed"] or "—","006400" if data["sm_changed"]=="No — false trigger removed only" else ("8B0000" if data["sm_changed"]=="Yes — safety behavior modified" else DTX)),
        ("V&V STATUS",data["vv_status"] or "—",DTX),
        ("INCIDENT STATUS",data["incident_status"] or "—",DGR if data["incident_status"]=="Closed" else (DOR if data["incident_status"]=="Open" else DTX)),
    ]:
        rr2=tbl.add_row(); ra,rb=rr2.cells[0],rr2.cells[1]
        _bg(ra,DLB); _bdr(ra,top=DBR,bot=DBR,left=DBR,right=DBR)
        _bg(rb,DSF); _bdr(rb,top=DBR,bot=DBR,left=DBR,right=DBR)
        _mar(ra,80,80,130,130); _mar(rb,80,80,130,130)
        _cp(ra,sl,bold=True,color=DB,size=9.5); _cp(rb,sv,bold=(bool(sv) and sv!="—"),color=sc2,size=10)
    # 9 — TECHNICAL SUMMARY
    smr=tbl.add_row(); smc=smr.cells[0].merge(smr.cells[1])
    _bg(smc,DN); _bdr(smc,top=DN,bot=DN,left=DN,right=DN); _mar(smc,120,120,160,160)
    smc.paragraphs[0].clear(); lp2=smc.paragraphs[0]
    lp2.paragraph_format.space_before=Pt(5); lp2.paragraph_format.space_after=Pt(4)
    _run(lp2,"TECHNICAL SUMMARY",bold=True,color="FFFFFF",size=9.5)
    sp2=smc.add_paragraph(); sp2.paragraph_format.space_before=Pt(4); sp2.paragraph_format.space_after=Pt(6)
    if data["summary"]: _run(sp2,data["summary"],color="D0E8FF",size=10.5)
    else: _run(sp2,"[ Root cause ]  —  [ Safety reaction ]  —  [ Functional effect ]",italic=True,color=DSC,size=10)
    doc.save(path)


if __name__ == "__main__":
    App().mainloop()
