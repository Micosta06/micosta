#!/usr/bin/env python3
"""
FS Incident Log — Interactive CLI Tool
ISO 26262 Powertrain | Fill field by field, generate a complete .docx at the end
"""

import sys, os, json, re
from pathlib import Path
from datetime import date
from copy import deepcopy

# ── DEPENDENCY CHECK ─────────────────────────────────────────────────────────
def check_deps():
    missing = []
    try: import rich
    except ImportError: missing.append("rich")
    try: import docx
    except ImportError: missing.append("python-docx")
    if missing:
        print(f"\n  Missing: {', '.join(missing)}")
        print(f"  Run: pip install {' '.join(missing)}\n")
        sys.exit(1)
check_deps()

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich.table import Table as RTable
from rich.text import Text
from rich.rule import Rule
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.columns import Columns
from rich.padding import Padding

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Emu, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── CONSOLE ──────────────────────────────────────────────────────────────────
con = Console()

# iOS Chill Corporate Pro — CLI palette
NAVY  = "#1D2D44"
BLUE  = "#0A84FF"
HINT  = "#8E8E93"
GOOD  = "#34C759"
WARN  = "#FF9500"

# ── BLANK DATA STRUCTURE ─────────────────────────────────────────────────────
def blank_data():
    return {
        "ticket_id":        "",
        "date":             str(date.today()),
        "asil":             "",
        "project":          "",
        "ecu":              "",
        "function":         "",
        "safety_goal":      "",
        "sw_appeared":      "",
        "sw_fixed":         "",
        "reported_by":      "",
        "cr_id":            "",
        "powertrain":       [],
        "op_mode":          [],
        "speed_range":      "",
        "thermal_soc":      "",
        "reproducibility":  "",
        "observed":         "",
        "trigger":          "",
        "side_effect":      "",
        "customer_visible": "",
        "severity":         "",
        "signals":          [],
        "detection":        "",
        "safe_state":       "",
        "ftti":             "",
        "rc_layer":         [],
        "rc_type":          [],
        "rc_statement":     "",
        "resolution":       [],
        "fix_applied":      "",
        "engineer_notes":   "",
        "attachments":      [],
        "sm_changed":       "",
        "vv_status":        "",
        "incident_status":  "",
        "summary":          "",
    }

# ── CLI HELPERS ───────────────────────────────────────────────────────────────
def sec_header(title, num, total):
    con.print()
    con.rule(
        f"[bold white on #1D2D44]  {title}  [/]  [dim]{num}/{total}[/]",
        style=f"bold {BLUE}"
    )
    con.print()

def field_prompt(label, hint="", default=""):
    if hint:
        con.print(f"  [bold {BLUE}]{label}[/]")
        con.print(f"  [dim]{hint}[/]")
    else:
        con.print(f"  [bold {BLUE}]{label}[/]")
    val = Prompt.ask(f"  [dim]>[/]", default=default, console=con)
    con.print()
    return val.strip()

def multiline_prompt(label, hint=""):
    con.print(f"  [bold {BLUE}]{label}[/]")
    if hint:
        con.print(f"  [dim]{hint}[/]")
    con.print(f"  [dim](type your text, press Enter twice when done)[/]")
    lines = []
    while True:
        try:
            line = con.input("  > ")
        except (EOFError, KeyboardInterrupt):
            break
        if line == "" and lines and lines[-1] == "":
            break
        lines.append(line)
    text = "\n".join(lines).strip()
    con.print()
    return text

def checkbox_prompt(label, options, hint="", multi=True):
    con.print(f"  [bold {BLUE}]{label}[/]")
    if hint:
        con.print(f"  [dim]{hint}[/]")
    con.print()
    for i, opt in enumerate(options, 1):
        con.print(f"    [dim]{i}.[/] {opt}")
    con.print()
    if multi:
        raw = Prompt.ask(
            "  [dim]Enter numbers (e.g. 1,3) or leave blank to skip >[/]",
            default="", console=con
        )
    else:
        raw = Prompt.ask(
            "  [dim]Enter number or leave blank to skip >[/]",
            default="", console=con
        )
    con.print()
    selected = []
    for part in raw.replace(" ", "").split(","):
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(options):
                selected.append(options[idx])
    return selected

def single_choice(label, options, hint=""):
    result = checkbox_prompt(label, options, hint, multi=False)
    return result[0] if result else ""

def file_prompt(label):
    con.print(f"  [bold {BLUE}]{label}[/]")
    con.print(f"  [dim]Add image/file paths one at a time. Leave blank when done.[/]")
    paths = []
    while True:
        raw = Prompt.ask(
            "  [dim]File path (or blank to finish) >[/]",
            default="", console=con
        )
        if not raw.strip():
            break
        p = Path(raw.strip())
        if p.exists():
            paths.append(str(p.resolve()))
            con.print(f"    [{GOOD}]✓ Added:[/] {p.name}")
        else:
            con.print(f"    [red]✗ Not found:[/] {raw}")
    con.print()
    return paths

def signal_prompt():
    con.print(f"  [bold {BLUE}]SIGNAL CONFLICT TABLE[/]")
    con.print(f"  [dim]Add signals involved in the conflict. Leave Role blank to stop.[/]")
    con.print()
    signals = []
    roles = ["Driver Intent", "Physical State", "System State", "Other"]
    row = 1
    while True:
        con.print(f"  [bold]Row {row}[/]")
        role_raw = Prompt.ask("  Role (or blank to finish)", default="", console=con)
        if not role_raw.strip():
            break
        if role_raw.strip().isdigit():
            idx = int(role_raw.strip()) - 1
            role = roles[idx] if 0 <= idx < len(roles) else role_raw.strip()
        else:
            role = role_raw.strip()
        signal   = Prompt.ask("  Signal / Function name", default="", console=con)
        expected = Prompt.ask("  Expected value / state",  default="", console=con)
        observed = Prompt.ask("  Observed value / state",  default="", console=con)
        note     = Prompt.ask("  Interpretation / note",   default="", console=con)
        signals.append({
            "role": role, "signal": signal,
            "expected": expected, "observed": observed, "note": note
        })
        con.print()
        row += 1
    return signals

# ── DATA COLLECTION ───────────────────────────────────────────────────────────
def collect(save_path):
    data = blank_data()
    TOTAL = 9

    # ── 1. IDENTIFICATION ─────────────────────────────────────────────────────
    sec_header("IDENTIFICATION", 1, TOTAL)
    data["ticket_id"]    = field_prompt("Ticket / ID",          "Primary reference (Jira, StarTeam, customer ticket)")
    data["date"]         = field_prompt("Date",                 "YYYY-MM-DD", default=str(date.today()))
    data["asil"]         = single_choice("ASIL Level",          ["ASIL A","ASIL B","ASIL C","ASIL D","QM"])
    data["project"]      = field_prompt("Project / Platform",   "e.g. MEB, MLB evo, PPE, C3")
    data["ecu"]          = field_prompt("ECU / System",         "e.g. HCP1, GE3, DCDC")
    data["function"]     = field_prompt("Function / Topic",     "e.g. MonMTqO, DrvDmdasil")
    data["safety_goal"]  = field_prompt("Safety Goal Ref.",     "SG-xx  |  FSC-xx  |  SwSR-xx")
    data["sw_appeared"]  = field_prompt("SW Version — Appeared in", "e.g. HCP1_SW_23.40.1")
    data["sw_fixed"]     = field_prompt("SW Version — Fixed in",    "e.g. HCP1_SW_23.44.0  |  Pending")
    data["reported_by"]  = field_prompt("Reported By",          "Name, role, date")
    data["cr_id"]        = field_prompt("Change Request (CR)",  "StarTeam / Jira CR ID")
    autosave(data, save_path)

    # ── 2. CONTEXT ────────────────────────────────────────────────────────────
    sec_header("CONTEXT", 2, TOTAL)
    data["powertrain"]   = checkbox_prompt("Powertrain Type",
        ["ICE","MHEV","HEV","PHEV","BEV","FCEV"])
    data["op_mode"]      = checkbox_prompt("Operating Mode",
        ["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"])
    data["speed_range"]  = field_prompt("Speed Range",          "e.g. standstill / <30 km/h / highway")
    data["thermal_soc"]  = field_prompt("Thermal / SOC State",  "e.g. cold start, warm, low SOC")
    data["reproducibility"] = single_choice("Reproducibility",  ["Always","Intermittent","Rare"])
    autosave(data, save_path)

    # ── 3. INCIDENT ───────────────────────────────────────────────────────────
    sec_header("INCIDENT", 3, TOTAL)
    data["observed"]     = multiline_prompt("Observed Behavior",
        "What did the system do that was unexpected? Facts only, no analysis.")
    data["trigger"]      = multiline_prompt("Trigger Condition",
        "When / under what driver action or state transition does it occur?")
    data["side_effect"]  = multiline_prompt("Side Effect",
        "Which legitimate functionality was lost or degraded?")
    data["customer_visible"] = single_choice("Customer Visible?", ["Yes","No"])
    data["severity"]     = single_choice("Severity",
        ["Safety-relevant","Functional","Comfort"])
    autosave(data, save_path)

    # ── 4. SIGNAL CONFLICT ────────────────────────────────────────────────────
    sec_header("SIGNAL CONFLICT", 4, TOTAL)
    data["signals"] = signal_prompt()
    autosave(data, save_path)

    # ── 5. SAFETY RESPONSE ────────────────────────────────────────────────────
    sec_header("SAFETY RESPONSE", 5, TOTAL)
    data["detection"]    = multiline_prompt("Detection Function",
        "Function name + detection method (e.g. MonMTqO — plausibility check)")
    data["safe_state"]   = multiline_prompt("Safe State Entered",
        "Describe the reaction and safe state (e.g. torque = 0 Nm, ICE inhibit)")
    data["ftti"]         = single_choice("FTTI Respected?",
        ["Yes","No — explain in analysis","Not assessed"])
    autosave(data, save_path)

    # ── 6. ROOT CAUSE & FIX ───────────────────────────────────────────────────
    sec_header("ROOT CAUSE & FIX", 6, TOTAL)
    data["rc_layer"]     = checkbox_prompt("Root Cause Layer",
        ["SW Logic","Signal Definition","Calibration",
         "Interface / Communication","Timing / Scheduling","Architecture"])
    data["rc_type"]      = checkbox_prompt("Root Cause Type",
        ["Incorrect rule","Missing condition","Wrong threshold","State machine error",
         "Signal misinterpretation","Missing debounce","Race condition","Incomplete spec"])
    data["rc_statement"] = multiline_prompt("Root Cause Statement",
        "One sentence — what was conceptually wrong?")
    data["resolution"]   = checkbox_prompt("Resolution Category",
        ["Logic refinement","State reclassification","Timing / debounce",
         "Architectural change","Calibration correction","Signal definition fix",
         "Spec update only","No fix — accepted risk"])
    data["fix_applied"]  = multiline_prompt("Fix Applied",
        "Brief description of the specific correction")
    autosave(data, save_path)

    # ── 7. ENGINEER SPACE ────────────────────────────────────────────────────
    sec_header("ENGINEER SPACE", 7, TOTAL)
    con.print(f"  [dim]Write freely — your narrative, insights, observations, anything.[/]")
    con.print()
    data["engineer_notes"] = multiline_prompt("Notes & Analysis", "")
    data["attachments"]    = file_prompt("Attach Files")
    autosave(data, save_path)

    # ── 8. STATUS ────────────────────────────────────────────────────────────
    sec_header("STATUS", 8, TOTAL)
    data["sm_changed"]      = single_choice("Safety Mechanism Changed?",
        ["No — false trigger removed only","Yes — safety behavior modified"])
    data["vv_status"]       = single_choice("V&V Status",
        ["Pending","In Progress","Complete","Waived"])
    data["incident_status"] = single_choice("Incident Status",
        ["Open","In Progress","Closed","Monitoring"])
    autosave(data, save_path)

    # ── 9. TECHNICAL SUMMARY ─────────────────────────────────────────────────
    sec_header("TECHNICAL SUMMARY", 9, TOTAL)
    data["summary"] = multiline_prompt("One-Sentence Summary",
        "[ Root cause ] — [ Safety reaction ] — [ Functional effect ]")
    autosave(data, save_path)

    return data

def autosave(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# ── DOCX HELPERS ─────────────────────────────────────────────────────────────
# iOS Chill Corporate Pro — document palette
D_NAVY      = "1D2D44"   # section header bg (deep navy)
D_BLUE      = "0A84FF"   # sub-header bg / label text (iOS system blue)
D_LABEL_BG  = "F0F4FF"   # label cell bg (airy blue-white)
D_VAL_BG    = "FFFFFF"   # value cell bg
D_BORDER    = "E2E8F0"   # separator (iOS separator gray)
D_TEXT      = "1C1C1E"   # primary text (iOS label)
D_SECONDARY = "6E6E73"   # secondary text (iOS secondary label)
D_PLACEHOLDER = "AEAEB2" # placeholder / empty hint (iOS tertiary)
D_GREEN     = "34C759"   # positive status (iOS green)
D_ORANGE    = "FF9500"   # warning status (iOS orange)
D_SURFACE   = "F5F5F7"   # light surface (Apple off-white)

def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_bg(cell, color):
    color = color.lstrip("#")
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for s in tcPr.findall(qn("w:shd")):
        tcPr.remove(s)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  color.upper())
    tcPr.append(shd)

def set_borders(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for b in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(b)
    tcBorders = OxmlElement("w:tcBorders")
    for edge, color in [("top", top), ("bottom", bottom), ("start", left), ("end", right)]:
        el = OxmlElement(f"w:{edge}")
        if color:
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    "4")
            el.set(qn("w:color"), color.lstrip("#").upper())
        else:
            el.set(qn("w:val"), "nil")
        tcBorders.append(el)
    shd_el = tcPr.find(qn("w:shd"))
    if shd_el is not None:
        shd_el.addprevious(tcBorders)
    else:
        tcPr.append(tcBorders)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for m in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(m)
    mar = OxmlElement("w:tcMar")
    for edge, val in [("top", top), ("bottom", bottom), ("start", left), ("end", right)]:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)

def set_col_width(cell, width_twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for w in tcPr.findall(qn("w:tcW")):
        tcPr.remove(w)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(width_twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.insert(0, tcW)

def add_run(para, text, bold=False, italic=False, color=D_TEXT,
            size=10, font="Arial"):
    run = para.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.name = font
    run.font.size = Pt(size)
    if color:
        r, g, b = hex_to_rgb(color)
        run.font.color.rgb = RGBColor(r, g, b)
    return run

def cell_para(cell, text, bold=False, italic=False, color=D_TEXT,
              size=10, align=WD_ALIGN_PARAGRAPH.LEFT, spacing_before=60,
              spacing_after=60, font="Arial", clear=True):
    if clear:
        cell.paragraphs[0].clear()
        para = cell.paragraphs[0]
    else:
        para = cell.add_paragraph()
    para.alignment = align
    para.paragraph_format.space_before = Pt(spacing_before / 20)
    para.paragraph_format.space_after  = Pt(spacing_after  / 20)
    if text:
        add_run(para, text, bold=bold, italic=italic, color=color, size=size, font=font)
    return para

def merge_row(table, fill, text, bold=True, color="FFFFFF", size=10.5):
    row  = table.add_row()
    cell = row.cells[0]
    for i in range(1, len(row.cells)):
        cell = cell.merge(row.cells[i])
    set_bg(cell, fill)
    set_borders(cell, top=fill, bottom=fill, left=fill, right=fill)
    set_cell_margins(cell, 70, 70, 140, 140)
    cell_para(cell, text, bold=bold, color=color, size=size)
    return row

def label_value_row(table, label, value, lb_fill=D_LABEL_BG, val_fill=D_VAL_BG,
                    val_italic=False, val_color=D_TEXT):
    row = table.add_row()
    c0, c1 = row.cells[0], row.cells[1]
    for i in range(2, len(row.cells)):
        c1 = c1.merge(row.cells[i])
    set_bg(c0, lb_fill);  set_borders(c0, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_bg(c1, val_fill); set_borders(c1, top=None, bottom=D_BORDER, left=None, right=None)
    set_cell_margins(c0, 80, 80, 130, 130)
    set_cell_margins(c1, 80, 80, 130, 130)
    cell_para(c0, label, bold=True, color=D_BLUE, size=9.5)
    cell_para(c1, value, italic=val_italic, color=val_color, size=10)
    return row

def checkbox_row(table, items, selected, fill=D_VAL_BG):
    row  = table.add_row()
    cell = row.cells[0]
    for i in range(1, len(row.cells)):
        cell = cell.merge(row.cells[i])
    set_bg(cell, fill)
    set_borders(cell, top=None, bottom=D_BORDER, left=None, right=None)
    set_cell_margins(cell, 60, 60, 130, 130)
    para = cell.paragraphs[0]
    para.clear()
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after  = Pt(3)
    for j, item in enumerate(items):
        if j > 0:
            add_run(para, "          ", size=10, color=D_TEXT)
        tick = "☑" if item in selected else "☐"
        clr  = D_BLUE if item in selected else D_SECONDARY
        add_run(para, f"{tick}  {item}", size=10, color=clr, bold=(item in selected))
    return row

def text_area_row(table, text, fill=D_VAL_BG, min_lines=6):
    row  = table.add_row()
    cell = row.cells[0]
    for i in range(1, len(row.cells)):
        cell = cell.merge(row.cells[i])
    set_bg(cell, fill)
    set_borders(cell, top=None, bottom=D_BORDER, left=None, right=None)
    set_cell_margins(cell, 100, 100, 140, 140)
    para = cell.paragraphs[0]
    para.clear()
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after  = Pt(4)
    if text:
        for i, line in enumerate(text.split("\n")):
            if i == 0:
                add_run(para, line, size=10, color=D_TEXT)
            else:
                new_para = cell.add_paragraph()
                new_para.paragraph_format.space_before = Pt(2)
                new_para.paragraph_format.space_after  = Pt(2)
                add_run(new_para, line, size=10, color=D_TEXT)
    else:
        add_run(para, "", size=10, color=D_TEXT)
    lines_used = max(1, len(text.split("\n")) if text else 0)
    for _ in range(max(0, min_lines - lines_used)):
        pad = cell.add_paragraph()
        pad.paragraph_format.space_before = Pt(2)
        pad.paragraph_format.space_after  = Pt(2)
        add_run(pad, "", size=10)
    return row

# ── DOCUMENT BUILDER ─────────────────────────────────────────────────────────
def build_doc(data, output_path):
    doc = Document()

    section = doc.sections[0]
    section.page_width    = Cm(21)
    section.page_height   = Cm(29.7)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)
    section.top_margin    = Cm(1.2)
    section.bottom_margin = Cm(1.2)

    # Header
    hdr      = doc.sections[0].header
    hdr_para = hdr.paragraphs[0]
    hdr_para.clear()
    hdr_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(hdr_para, "FUNCTIONAL SAFETY INCIDENT LOG  —  ISO 26262 Powertrain",
            bold=True, color=D_NAVY, size=8)
    add_run(hdr_para, "          CONFIDENTIAL — INTERNAL USE ONLY",
            italic=True, color=D_SECONDARY, size=8)
    hdr_para.paragraph_format.space_after = Pt(6)

    # ── TITLE BLOCK ──────────────────────────────────────────────────────────
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_before = Pt(0)
    title_para.paragraph_format.space_after  = Pt(4)
    add_run(title_para, "FS INCIDENT LOG", bold=True, color=D_NAVY, size=20)

    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_para.paragraph_format.space_before = Pt(0)
    subtitle_para.paragraph_format.space_after  = Pt(14)
    add_run(subtitle_para, "ISO 26262  —  Powertrain  —  v3.0",
            italic=True, color=D_BLUE, size=10)

    # Content width in twips (A4 21cm − 3cm margins = 18cm)
    CW = int(18 * 567)   # ~10206 twips
    C1 = int(3.5 * 567)  # label col ~1985 twips
    C2 = CW - C1

    # ── MAIN TABLE ───────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style  = "Table Grid"
    tbl.autofit = False

    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        for gc in tblGrid.findall(qn("w:gridCol")):
            tblGrid.remove(gc)
        for w in [C1, C2]:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(w))
            tblGrid.append(gc)

    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl._tbl.insert(0, tblPr)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(CW))
    tblW.set(qn("w:type"), "dxa")
    tblPr.insert(0, tblW)

    def sec(heading):
        merge_row(tbl, D_NAVY, f"  {heading}", bold=True, color="FFFFFF", size=10.5)

    def sub(heading):
        merge_row(tbl, D_BLUE, f"  {heading}", bold=True, color="FFFFFF", size=9.5)

    def lv(label, value):
        label_value_row(tbl, label, value,
                        val_italic=not bool(value),
                        val_color=D_PLACEHOLDER if not value else D_TEXT)

    def cb(items, selected):
        checkbox_row(tbl, items, selected)

    def ta(text, lines=5):
        text_area_row(tbl, text, min_lines=lines)

    # ── 1. IDENTIFICATION ─────────────────────────────────────────────────────
    sec("1.  IDENTIFICATION")

    # Quad header row: ticket | date | asil | project (inner 4-col table)
    r = tbl.add_row()
    cell_all = r.cells[0].merge(r.cells[1])
    set_bg(cell_all, D_VAL_BG)
    set_borders(cell_all, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_cell_margins(cell_all, 0, 0, 0, 0)

    inner = OxmlElement("w:tbl")
    iTblPr  = OxmlElement("w:tblPr")
    iTblW   = OxmlElement("w:tblW")
    iTblW.set(qn("w:w"),    str(CW))
    iTblW.set(qn("w:type"), "dxa")
    iTblPr.append(iTblW)
    iTblBdr = OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"),   "single")
        e.set(qn("w:sz"),    "4")
        e.set(qn("w:color"), D_BORDER)
        iTblBdr.append(e)
    iTblPr.append(iTblBdr)
    inner.append(iTblPr)

    iTblGrid = OxmlElement("w:tblGrid")
    qw = CW // 4
    for _ in range(4):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(qw))
        iTblGrid.append(gc)
    inner.append(iTblGrid)

    iRow = OxmlElement("w:tr")
    quad_fields = [
        ("TICKET / ID",        data["ticket_id"]),
        ("DATE",               data["date"]),
        ("ASIL",               data["asil"]),
        ("PROJECT / PLATFORM", data["project"]),
    ]
    for label, val in quad_fields:
        tc = OxmlElement("w:tc")
        tcPr2 = OxmlElement("w:tcPr")
        tcW2  = OxmlElement("w:tcW")
        tcW2.set(qn("w:w"),    str(qw))
        tcW2.set(qn("w:type"), "dxa")
        tcPr2.append(tcW2)
        shd2 = OxmlElement("w:shd")
        shd2.set(qn("w:val"),   "clear")
        shd2.set(qn("w:color"), "auto")
        shd2.set(qn("w:fill"),  D_VAL_BG)
        tcPr2.append(shd2)
        mar2 = OxmlElement("w:tcMar")
        for edge in ["top","bottom","left","right"]:
            me = OxmlElement(f"w:{edge}")
            me.set(qn("w:w"),    "100")
            me.set(qn("w:type"), "dxa")
            mar2.append(me)
        tcPr2.append(mar2)
        tc.append(tcPr2)

        lp = OxmlElement("w:p")
        lr = OxmlElement("w:r")
        lrPr = OxmlElement("w:rPr")
        lbold  = OxmlElement("w:b")
        lcolor = OxmlElement("w:color")
        lcolor.set(qn("w:val"), D_BLUE)
        lsize  = OxmlElement("w:sz")
        lsize.set(qn("w:val"), "17")
        lrPr.append(lbold); lrPr.append(lcolor); lrPr.append(lsize)
        lr.append(lrPr)
        lt = OxmlElement("w:t"); lt.text = label
        lr.append(lt); lp.append(lr)
        tc.append(lp)

        vp = OxmlElement("w:p")
        vr = OxmlElement("w:r")
        vrPr = OxmlElement("w:rPr")
        vcolor = OxmlElement("w:color")
        vcolor.set(qn("w:val"), D_TEXT if val else D_PLACEHOLDER)
        vsize = OxmlElement("w:sz")
        vsize.set(qn("w:val"), "20")
        if not val:
            vi = OxmlElement("w:i"); vrPr.append(vi)
        vrPr.append(vcolor); vrPr.append(vsize)
        vr.append(vrPr)
        vt = OxmlElement("w:t"); vt.text = val if val else "—"
        vr.append(vt); vp.append(vr)
        tc.append(vp)
        iRow.append(tc)

    inner.append(iRow)
    cell_all._tc.append(inner)

    lv("ECU / SYSTEM",        data["ecu"])
    lv("FUNCTION / TOPIC",    data["function"])
    lv("SAFETY GOAL REF.",    data["safety_goal"])
    lv("SW — APPEARED IN",    data["sw_appeared"])
    lv("SW — FIXED IN",       data["sw_fixed"])
    lv("REPORTED BY",         data["reported_by"])
    lv("CHANGE REQUEST (CR)", data["cr_id"])

    # ── 2. CONTEXT ────────────────────────────────────────────────────────────
    sec("2.  CONTEXT")
    sub("Powertrain Type")
    cb(["ICE","MHEV","HEV","PHEV","BEV","FCEV"], data["powertrain"])
    sub("Operating Mode")
    cb(["Normal","Degraded / Limp-Home","Post-Fault","READY","Cranking","Charging"], data["op_mode"])
    lv("SPEED RANGE",   data["speed_range"])
    lv("THERMAL / SOC", data["thermal_soc"])

    r_rep = tbl.add_row()
    c0, c1 = r_rep.cells[0], r_rep.cells[1]
    set_bg(c0, D_LABEL_BG); set_borders(c0, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_bg(c1, D_VAL_BG);   set_borders(c1, top=None, bottom=D_BORDER, left=None, right=None)
    set_cell_margins(c0, 80, 80, 130, 130)
    set_cell_margins(c1, 60, 60, 130, 130)
    cell_para(c0, "REPRODUCIBILITY", bold=True, color=D_BLUE, size=9.5)
    options_rep = ["Always", "Intermittent", "Rare"]
    para_rep = c1.paragraphs[0]; para_rep.clear()
    para_rep.paragraph_format.space_before = Pt(3)
    para_rep.paragraph_format.space_after  = Pt(3)
    for j, opt in enumerate(options_rep):
        if j > 0:
            add_run(para_rep, "          ", size=10, color=D_TEXT)
        sel = data["reproducibility"] == opt
        add_run(para_rep, f"{'☑' if sel else '☐'}  {opt}",
                size=10, color=D_BLUE if sel else D_SECONDARY, bold=sel)

    # ── 3. INCIDENT ───────────────────────────────────────────────────────────
    sec("3.  INCIDENT")
    lv("OBSERVED BEHAVIOR", ""); ta(data["observed"])
    lv("TRIGGER CONDITION", ""); ta(data["trigger"])
    lv("SIDE EFFECT",       ""); ta(data["side_effect"])

    r_cv = tbl.add_row()
    c0, c1 = r_cv.cells[0], r_cv.cells[1]
    for c_, label_, items_, sel_ in [
        (c0, "CUSTOMER VISIBLE?", ["Yes","No"],
         [data["customer_visible"]] if data["customer_visible"] else []),
        (c1, "SEVERITY", ["Safety-relevant","Functional","Comfort"],
         [data["severity"]] if data["severity"] else []),
    ]:
        set_bg(c_, D_VAL_BG)
        set_borders(c_, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
        set_cell_margins(c_, 80, 80, 130, 130)
        cell_para(c_, label_, bold=True, color=D_BLUE, size=9.5)
        p2 = c_.add_paragraph()
        p2.paragraph_format.space_before = Pt(3)
        p2.paragraph_format.space_after  = Pt(3)
        for j, opt in enumerate(items_):
            if j > 0:
                add_run(p2, "     ", size=10)
            sel = opt in sel_
            add_run(p2, f"{'☑' if sel else '☐'}  {opt}",
                    size=10, color=D_BLUE if sel else D_SECONDARY, bold=sel)

    # ── 4. SIGNAL CONFLICT ────────────────────────────────────────────────────
    sec("4.  SIGNAL CONFLICT")
    sig_row = tbl.add_row()
    sc0, sc1 = sig_row.cells[0], sig_row.cells[1]
    set_bg(sc0, D_LABEL_BG)
    set_borders(sc0, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_cell_margins(sc0, 80, 80, 130, 130)
    cell_para(sc0, "CONFLICT TABLE", bold=True, color=D_BLUE, size=9.5)
    set_bg(sc1, D_VAL_BG)
    set_borders(sc1, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_cell_margins(sc1, 80, 80, 100, 100)

    sig_cols   = ["Role","Signal / Function","Expected","Observed","Note"]
    sig_widths = [
        int(C2 * 0.14), int(C2 * 0.24), int(C2 * 0.18), int(C2 * 0.18),
        C2 - int(C2*0.14) - int(C2*0.24) - int(C2*0.18)*2
    ]
    sig_tbl_xml = OxmlElement("w:tbl")
    stPr = OxmlElement("w:tblPr")
    stW  = OxmlElement("w:tblW")
    stW.set(qn("w:w"),    str(C2))
    stW.set(qn("w:type"), "dxa")
    stPr.append(stW)
    stBdr = OxmlElement("w:tblBorders")
    for edge in ["top","bottom","left","right","insideH","insideV"]:
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"),   "single")
        e.set(qn("w:sz"),    "3")
        e.set(qn("w:color"), D_BORDER)
        stBdr.append(e)
    stPr.append(stBdr)
    sig_tbl_xml.append(stPr)
    stGrid = OxmlElement("w:tblGrid")
    for w in sig_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        stGrid.append(gc)
    sig_tbl_xml.append(stGrid)

    def sig_tbl_row(values, is_header=False, row_fill=D_VAL_BG):
        tr = OxmlElement("w:tr")
        for val, w in zip(values, sig_widths):
            tc = OxmlElement("w:tc")
            tcPr3 = OxmlElement("w:tcPr")
            tw3   = OxmlElement("w:tcW")
            tw3.set(qn("w:w"),    str(w))
            tw3.set(qn("w:type"), "dxa")
            tcPr3.append(tw3)
            shd3 = OxmlElement("w:shd")
            shd3.set(qn("w:val"),   "clear")
            shd3.set(qn("w:color"), "auto")
            shd3.set(qn("w:fill"),  D_BLUE if is_header else row_fill)
            tcPr3.append(shd3)
            mar3 = OxmlElement("w:tcMar")
            for edge in ["top","bottom","left","right"]:
                me3 = OxmlElement(f"w:{edge}")
                me3.set(qn("w:w"),    "80")
                me3.set(qn("w:type"), "dxa")
                mar3.append(me3)
            tcPr3.append(mar3)
            tc.append(tcPr3)
            wp3  = OxmlElement("w:p")
            wr3  = OxmlElement("w:r")
            wrPr3 = OxmlElement("w:rPr")
            if is_header:
                wb = OxmlElement("w:b"); wrPr3.append(wb)
            wc3 = OxmlElement("w:color")
            wc3.set(qn("w:val"), "FFFFFF" if is_header else D_TEXT)
            wrPr3.append(wc3)
            ws3 = OxmlElement("w:sz")
            ws3.set(qn("w:val"), "17")
            wrPr3.append(ws3)
            wr3.append(wrPr3)
            wt3 = OxmlElement("w:t"); wt3.text = val or ""
            wr3.append(wt3); wp3.append(wr3); tc.append(wp3)
            tr.append(tc)
        return tr

    sig_tbl_xml.append(sig_tbl_row(sig_cols, is_header=True))
    row_fills   = [D_LABEL_BG, D_VAL_BG, "EBF5EE", D_VAL_BG, "FDF8EC", D_VAL_BG]
    role_labels = ["Driver Intent", "", "Physical State", "", "System State", ""]
    signals = data["signals"]
    for ri in range(6):
        if ri < len(signals):
            s = signals[ri]
            vals = [s.get("role",""), s.get("signal",""), s.get("expected",""),
                    s.get("observed",""), s.get("note","")]
        else:
            vals = [role_labels[ri], "", "", "", ""]
        sig_tbl_xml.append(sig_tbl_row(vals, row_fill=row_fills[ri]))
    for s in signals[6:]:
        vals = [s.get("role",""), s.get("signal",""), s.get("expected",""),
                s.get("observed",""), s.get("note","")]
        sig_tbl_xml.append(sig_tbl_row(vals))
    sc1._tc.append(sig_tbl_xml)

    # ── 5. SAFETY RESPONSE ────────────────────────────────────────────────────
    sec("5.  SAFETY RESPONSE")
    lv("DETECTION FUNCTION", ""); ta(data["detection"])
    lv("SAFE STATE ENTERED",  ""); ta(data["safe_state"])

    r_ftti = tbl.add_row()
    cf0, cf1 = r_ftti.cells[0], r_ftti.cells[1]
    set_bg(cf0, D_LABEL_BG); set_borders(cf0, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_bg(cf1, D_VAL_BG);   set_borders(cf1, top=None, bottom=D_BORDER, left=None, right=None)
    set_cell_margins(cf0, 80, 80, 130, 130)
    set_cell_margins(cf1, 60, 60, 130, 130)
    cell_para(cf0, "FTTI RESPECTED?", bold=True, color=D_BLUE, size=9.5)
    pf = cf1.paragraphs[0]; pf.clear()
    pf.paragraph_format.space_before = Pt(3)
    pf.paragraph_format.space_after  = Pt(3)
    ftti_opts = ["Yes", "No — explain in analysis", "Not assessed"]
    for j, opt in enumerate(ftti_opts):
        if j > 0:
            add_run(pf, "          ", size=10)
        sel = data["ftti"] == opt
        add_run(pf, f"{'☑' if sel else '☐'}  {opt}",
                size=10, color=D_BLUE if sel else D_SECONDARY, bold=sel)

    # ── 6. ROOT CAUSE & FIX ───────────────────────────────────────────────────
    sec("6.  ROOT CAUSE & FIX")
    sub("Root Cause Layer")
    cb(["SW Logic","Signal Definition","Calibration",
        "Interface / Communication","Timing / Scheduling","Architecture"],
       data["rc_layer"])
    sub("Root Cause Type")
    cb(["Incorrect rule","Missing condition","Wrong threshold","State machine error",
        "Signal misinterpretation","Missing debounce","Race condition","Incomplete spec"],
       data["rc_type"])
    lv("ROOT CAUSE STATEMENT", ""); ta(data["rc_statement"], lines=4)
    sub("Resolution Category")
    cb(["Logic refinement","State reclassification","Timing / debounce","Architectural change",
        "Calibration correction","Signal definition fix","Spec update only","No fix — accepted risk"],
       data["resolution"])
    lv("FIX APPLIED", ""); ta(data["fix_applied"])

    # ── 7. ENGINEER SPACE ─────────────────────────────────────────────────────
    sec("7.  ENGINEER SPACE")
    notes_row = tbl.add_row()
    nc = notes_row.cells[0].merge(notes_row.cells[1])
    set_bg(nc, D_VAL_BG)
    set_borders(nc, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
    set_cell_margins(nc, 120, 120, 160, 160)
    nc.paragraphs[0].clear()
    nc.paragraphs[0].paragraph_format.space_before = Pt(4)
    nc.paragraphs[0].paragraph_format.space_after  = Pt(4)
    if data["engineer_notes"]:
        first = True
        for line in data["engineer_notes"].split("\n"):
            if first:
                add_run(nc.paragraphs[0], line, size=10, color=D_TEXT)
                first = False
            else:
                np2 = nc.add_paragraph()
                np2.paragraph_format.space_before = Pt(2)
                np2.paragraph_format.space_after  = Pt(2)
                add_run(np2, line, size=10, color=D_TEXT)
        for _ in range(8):
            pad = nc.add_paragraph()
            pad.paragraph_format.space_before = Pt(2)
            pad.paragraph_format.space_after  = Pt(2)
    else:
        add_run(nc.paragraphs[0],
                "Write freely — narrative, insights, hypotheses, observations, anything relevant.",
                italic=True, color=D_PLACEHOLDER, size=9.5)
        for _ in range(12):
            pad = nc.add_paragraph()
            pad.paragraph_format.space_before = Pt(2)
            pad.paragraph_format.space_after  = Pt(2)

    if data["attachments"]:
        att_label = tbl.add_row()
        al_cell = att_label.cells[0].merge(att_label.cells[1])
        set_bg(al_cell, D_LABEL_BG)
        set_borders(al_cell, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
        set_cell_margins(al_cell, 60, 60, 130, 130)
        cell_para(al_cell, f"  ATTACHMENTS  ({len(data['attachments'])} file(s))",
                  bold=True, color=D_BLUE, size=9.5)
        for fpath in data["attachments"]:
            att_row = tbl.add_row()
            ac = att_row.cells[0].merge(att_row.cells[1])
            set_bg(ac, D_VAL_BG)
            set_borders(ac, top=None, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
            set_cell_margins(ac, 80, 80, 140, 140)
            ac.paragraphs[0].clear()
            ext = Path(fpath).suffix.lower()
            img_exts = {".png",".jpg",".jpeg",".gif",".bmp",".tiff",".webp"}
            if ext in img_exts and Path(fpath).exists():
                try:
                    fname_para = ac.paragraphs[0]
                    fname_para.paragraph_format.space_before = Pt(4)
                    fname_para.paragraph_format.space_after  = Pt(4)
                    add_run(fname_para, f"  {Path(fpath).name}",
                            bold=True, color=D_BLUE, size=9)
                    img_para = ac.add_paragraph()
                    img_para.paragraph_format.space_before = Pt(4)
                    img_para.paragraph_format.space_after  = Pt(8)
                    img_para.add_run().add_picture(fpath, width=Inches(5.5))
                except Exception as e:
                    add_run(ac.paragraphs[0],
                            f"  [Image: {Path(fpath).name}] (could not embed: {e})",
                            italic=True, color=D_SECONDARY, size=9)
            else:
                add_run(ac.paragraphs[0],
                        f"  Attachment: {Path(fpath).name}",
                        bold=True, color=D_BLUE, size=9.5)
                if Path(fpath).exists():
                    np3 = ac.add_paragraph()
                    np3.paragraph_format.space_before = Pt(2)
                    np3.paragraph_format.space_after  = Pt(4)
                    add_run(np3, f"  {fpath}", italic=True, color=D_SECONDARY, size=8.5)

    # ── 8. STATUS ─────────────────────────────────────────────────────────────
    sec("8.  STATUS")
    status_fields = [
        ("SAFETY MECHANISM CHANGED?",
         data["sm_changed"] or "—",
         D_TEXT if not data["sm_changed"] else
         ("006400" if data["sm_changed"] == "No — false trigger removed only" else "8B0000")),
        ("V&V STATUS",
         data["vv_status"] or "—",
         D_TEXT),
        ("INCIDENT STATUS",
         data["incident_status"] or "—",
         D_GREEN   if data["incident_status"] == "Closed"      else
         D_ORANGE  if data["incident_status"] == "Open"        else
         D_TEXT),
    ]
    for s_label, s_val, s_col in status_fields:
        sr = tbl.add_row()
        sc_a, sc_b = sr.cells[0], sr.cells[1]
        set_bg(sc_a, D_LABEL_BG); set_borders(sc_a, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
        set_bg(sc_b, D_SURFACE);  set_borders(sc_b, top=D_BORDER, bottom=D_BORDER, left=D_BORDER, right=D_BORDER)
        set_cell_margins(sc_a, 80, 80, 130, 130)
        set_cell_margins(sc_b, 80, 80, 130, 130)
        cell_para(sc_a, s_label, bold=True, color=D_BLUE, size=9.5)
        cell_para(sc_b, s_val,   bold=(bool(s_val) and s_val != "—"), color=s_col, size=10)

    # ── 9. TECHNICAL SUMMARY ──────────────────────────────────────────────────
    sum_row  = tbl.add_row()
    sum_cell = sum_row.cells[0].merge(sum_row.cells[1])
    set_bg(sum_cell, D_NAVY)
    set_borders(sum_cell, top=D_NAVY, bottom=D_NAVY, left=D_NAVY, right=D_NAVY)
    set_cell_margins(sum_cell, 120, 120, 160, 160)
    sum_cell.paragraphs[0].clear()
    lp2 = sum_cell.paragraphs[0]
    lp2.paragraph_format.space_before = Pt(5)
    lp2.paragraph_format.space_after  = Pt(4)
    add_run(lp2, "TECHNICAL SUMMARY", bold=True, color="FFFFFF", size=9.5)
    sp2 = sum_cell.add_paragraph()
    sp2.paragraph_format.space_before = Pt(4)
    sp2.paragraph_format.space_after  = Pt(6)
    if data["summary"]:
        add_run(sp2, data["summary"], color="D0E8FF", size=10.5)
    else:
        add_run(sp2,
                "[ Root cause ]  —  [ Safety reaction ]  —  [ Functional effect ]",
                italic=True, color=D_SECONDARY, size=10)

    doc.save(output_path)

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    import argparse
    parser = argparse.ArgumentParser(description="FS Incident Log — Interactive Filler")
    parser.add_argument("--output", "-o", default="", help="Output .docx path")
    parser.add_argument("--resume", "-r", default="", help="Resume from saved JSON")
    args = parser.parse_args()

    con.print()
    con.print(Panel(
        f"[bold white]FS INCIDENT LOG[/]\n[dim]ISO 26262 Powertrain — Interactive Filler[/]\n\n"
        f"Fill each field one at a time.\n"
        f"Press [bold]Enter[/] to skip any field.\n"
        f"Your progress is saved automatically after each section.",
        title=f"[bold {BLUE}]Welcome[/]",
        border_style=NAVY,
        padding=(1, 4),
    ))
    con.print()

    if not args.output:
        default_name = f"FS_Incident_{date.today().strftime('%Y%m%d')}.docx"
        out = Prompt.ask("  Output file", default=default_name, console=con)
    else:
        out = args.output
    if not out.endswith(".docx"):
        out += ".docx"

    save_path = out.replace(".docx", "_data.json")

    if args.resume and Path(args.resume).exists():
        with open(args.resume) as f:
            data = json.load(f)
        con.print(f"\n  [{GOOD}]Resuming from[/] {args.resume}\n")
    else:
        data = collect(save_path)

    con.print()
    con.rule(f"[bold white]Generating document[/]", style=BLUE)
    con.print()

    with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                  console=con) as progress:
        task = progress.add_task("Building document...", total=None)
        build_doc(data, out)
        progress.update(task, description="Done.")

    con.print()
    con.print(Panel(
        f"[bold {GOOD}]✓ Document saved:[/]\n[white]{Path(out).resolve()}[/]\n\n"
        f"[dim]Data saved to:[/] {save_path}",
        title=f"[bold {GOOD}]Complete[/]",
        border_style=GOOD,
        padding=(1, 4),
    ))
    con.print()

if __name__ == "__main__":
    main()
