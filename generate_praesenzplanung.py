"""
GF-Präsenzplanung – Automatische Generierung
Ausführen: python3 generate_praesenzplanung.py
Ergebnis:  GF_Praesenzplanung_<Jahr>_H2.xlsx im selben Ordner
"""

import json, requests
from datetime import date, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
DUMP_PATH = HERE / "absences_raw_dump.json"

# ── Config ────────────────────────────────────────────────────────────────────

PEOPLE = {
    "Roksana Lichte":          "C",
    "Ben Min-Woo Illigens":    "D",
    "Sebastian Woinar":        "E",
    "Pablo Ezequiel Guerrero": "F",
}
COL_IDX = {"C": 3, "D": 4, "E": 5, "F": 6}

TYPE_MAP = {
    "Office Berlin":      ("Berlin",                  1),
    "Office Potsdam":     ("Potsdam",                 1),
    "Home Office":        ("Home Office",              1),
    "Work Travel":        ("Abwesend (Dienstreise)",   2),
    "Remote Work Abroad": ("Abwesend (Dienstreise)",   2),
    "Education Day":      ("Abwesend (Dienstreise)",   2),
    "Compensation Time":  ("Abwesend (Urlaub)",        2),
    "Paid vacation":      ("Abwesend (Urlaub)",        3),
    "Special Leave":      ("Abwesend (Urlaub)",        3),
    "Sick Leave":         ("Abwesend (Krank)",         3),
}

VACATION_TYPES = {"Paid vacation", "Compensation Time", "Special Leave"}

def fetch_holidays(year: int) -> dict:
    """
    Fetch public holidays for Berlin (BE) and Brandenburg (BB)
    from feiertage-api.de. Returns {date: label}.
    """
    holidays = {}
    be_dates, bb_dates = {}, {}

    for land, store in (("BE", be_dates), ("BB", bb_dates)):
        try:
            r = requests.get(
                f"https://feiertage-api.de/api/",
                params={"jahr": year, "nur_land": land},
                timeout=10,
            )
            r.raise_for_status()
            for name, info in r.json().items():
                d = date.fromisoformat(info["datum"])
                store[d] = name
        except Exception as e:
            print(f"  ⚠ Feiertage-API ({land}) nicht erreichbar: {e}")

    all_dates = set(be_dates) | set(bb_dates)
    for d in all_dates:
        in_be = d in be_dates
        in_bb = d in bb_dates
        name  = be_dates.get(d) or bb_dates.get(d)
        if in_be and in_bb:
            label = f"{name} (BE + BB)"
        elif in_be:
            label = f"{name} (nur BE)"
        else:
            label = f"{name} (nur BB)"
        holidays[d] = label

    print(f"✓ Feiertage {year}: {len(holidays)} Tage "
          f"(BE: {len(be_dates)}, BB: {len(bb_dates)})")
    return holidays

MONTHS_DE = {1:"Januar",2:"Februar",3:"März",4:"April",5:"Mai",6:"Juni",
             7:"Juli",8:"August",9:"September",10:"Oktober",11:"November",12:"Dezember"}

# ── Styles ────────────────────────────────────────────────────────────────────

HDR_BG = "1F3864"; HDR_FG = "FFFFFF"
GRP_BG = "EBF5FB"; HOL_BG = "F2F2F2"
OK_BG  = "C6EFCE"; OK_FG  = "276221"
WRN_BG = "FFEB9C"; WRN_FG = "9C5700"
ERR_BG = "FFC7CE"; ERR_FG = "9C0006"
VAC_BG = "FCE4D6"; VAC_FG = "833C00"

def fill(h):   return PatternFill("solid", start_color=h, end_color=h)
thin = Side(style="thin", color="BFBFBF")
def bdr():     return Border(left=thin, right=thin, top=thin, bottom=thin)
def hdr_font(bold=False, sz=10): return Font(name="Arial", color=HDR_FG, bold=bold, size=sz)
def std_font(col="000000", bold=False, italic=False, sz=10):
    return Font(name="Arial", color=col, bold=bold, italic=italic, size=sz)

DAYS_DE = ["Mo","Di","Mi","Do","Fr"]
OPTIONS = ('"Berlin,Potsdam,Home Office,'
           'Abwesend (Urlaub),Abwesend (Dienstreise),Abwesend (Krank)"')
COLS = [
    ("Datum", 14), ("Wochentag", 13), ("Roksana", 14), ("Ben", 14),
    ("Sebastian", 14), ("Pablo\n(Fallback)", 14),
    ("Berlin\ngedeckt?", 18), ("Hinweise", 35),
]

# ── Step 1: Load credentials ──────────────────────────────────────────────────

def load_env():
    env = {}
    for line in (HERE / ".env").read_text().splitlines():
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env

# ── Step 2: Fetch Personio data ───────────────────────────────────────────────

def _auth() -> tuple[dict, str]:
    """Returns (headers, token) after authenticating."""
    env  = load_env()
    resp = requests.post(
        "https://api.personio.de/v1/auth",
        json={"client_id": env["PERSONIO_CLIENT_ID"],
              "client_secret": env["PERSONIO_CLIENT_SECRET"]},
        timeout=15,
    )
    resp.raise_for_status()
    token = resp.json()["data"]["token"]
    return {"Authorization": f"Bearer {token}"}, token


def _get_employee_ids(headers: dict) -> dict[str, int]:
    """Fetch all employees and return {full_name: id} for our PEOPLE."""
    ids   = {}
    limit, page = 200, 1
    while True:
        r = requests.get(
            "https://api.personio.de/v1/company/employees",
            headers=headers,
            params={"limit": limit, "offset": (page - 1) * limit},
            timeout=15,
        )
        r.raise_for_status()
        body = r.json()
        for e in body["data"]:
            a     = e["attributes"]
            first = a.get("first_name", {}).get("value", "")
            last  = a.get("last_name",  {}).get("value", "")
            name  = f"{first} {last}".strip()
            if name in PEOPLE:
                ids[name] = a["id"]["value"]
        if page >= body["metadata"].get("total_pages", 1):
            break
        page += 1
    return ids


def fetch_absences(start: date, end: date) -> list:
    headers, _ = _auth()
    print("✓ Personio: authentifiziert")

    all_raw = []
    limit, page = 200, 1
    total_pages = None
    while True:
        r = requests.get(
            "https://api.personio.de/v1/company/time-offs",
            headers=headers,
            params={
                "start_date": start.isoformat(),
                "end_date":   end.isoformat(),
                "limit":      limit,
                "page":       page,
            },
            timeout=15,
        )
        r.raise_for_status()
        body        = r.json()
        batch       = body["data"]
        total_pages = body.get("metadata", {}).get("total_pages", 1)
        all_raw.extend(batch)
        print(f"  Seite {page}/{total_pages}: {len(batch)} Einträge (gesamt: {len(all_raw)})")
        if page >= total_pages:
            break
        page += 1

    # Save raw dump for inspection (wird am Ende des Laufs wieder gelöscht)
    DUMP_PATH.write_text(json.dumps(all_raw, indent=2, ensure_ascii=False))
    print(f"✓ Raw-Dump gespeichert: {DUMP_PATH.name}")

    parsed = []
    for a in all_raw:
        if not a["attributes"].get("employee"):
            continue
        try:
            emp   = a["attributes"]["employee"]["attributes"]
            first = emp.get("first_name", {}).get("value", "")
            last  = emp.get("last_name",  {}).get("value", "")
            parsed.append({
                "employee_name": f"{first} {last}".strip(),
                "type":          a["attributes"]["time_off_type"]["attributes"]["name"],
                "start_date":    a["attributes"]["start_date"][:10],
                "end_date":      a["attributes"]["end_date"][:10],
                "status":        a["attributes"]["status"],
            })
        except (KeyError, TypeError):
            continue

    print(f"✓ Personio: {len(parsed)} Einträge gesamt ({start} – {end})")

    # Per-person summary
    from collections import Counter
    counts = Counter(p["employee_name"] for p in parsed if p["employee_name"] in PEOPLE)
    for name in PEOPLE:
        print(f"  {name}: {counts.get(name, 0)} Einträge")

    return parsed

# ── Step 3: Build presence lookup ─────────────────────────────────────────────

def build_presence(absences: list) -> dict:
    presence = {n: {} for n in PEOPLE}
    for ab in absences:
        name = ab["employee_name"]
        if name not in PEOPLE:
            continue
        label, prio = TYPE_MAP.get(ab["type"], ("Home Office", 0))
        d   = date.fromisoformat(ab["start_date"])
        end = date.fromisoformat(ab["end_date"])
        while d <= end:
            if d.weekday() < 5:
                ex = presence[name].get(d)
                if ex is None or prio > ex[1]:
                    presence[name][d] = (label, prio)
            d += timedelta(days=1)
    return presence

# ── Step 4: Build Excel ───────────────────────────────────────────────────────

def make_month_sheet(wb, month_num, year, presence, holidays):
    name = MONTHS_DE[month_num]
    ws   = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    ws["A1"] = f"GF-Präsenzplanung – {name} {year}"
    ws["A1"].font = Font(name="Arial", color=HDR_FG, bold=True, size=13)
    ws["A1"].fill = fill(HDR_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:H2")
    ws["A2"] = ("Werte: Berlin | Potsdam | Home Office | Abwesend (Urlaub) | "
                "Abwesend (Dienstreise) | Abwesend (Krank)")
    ws["A2"].font = std_font("595959", italic=True, sz=9)
    ws["A2"].fill = fill("F2F2F2")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 30
    for ci, (label, width) in enumerate(COLS, start=1):
        c = ws.cell(3, ci, label)
        c.font = hdr_font(bold=True); c.fill = fill(HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr()
        ws.column_dimensions["ABCDEFGH"[ci-1]].width = width

    dv = DataValidation(type="list", formula1=OPTIONS, allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)

    row = 4
    d   = date(year, month_num, 1)
    while d.month == month_num:
        if d.weekday() < 5:
            is_hol   = d in holidays
            row_fill = fill(HOL_BG) if is_hol else (fill(GRP_BG) if row % 2 == 0 else fill("FFFFFF"))

            c = ws.cell(row, 1, d); c.number_format = "DD.MM.YYYY"
            c.font = std_font(); c.fill = row_fill; c.border = bdr()
            c.alignment = Alignment(horizontal="center")

            c = ws.cell(row, 2, DAYS_DE[d.weekday()])
            c.font = std_font(); c.fill = row_fill; c.border = bdr()
            c.alignment = Alignment(horizontal="center")

            for col_letter, person in zip("CDEF", PEOPLE.keys()):
                entry = presence[person].get(d)
                val   = entry[0] if entry else ""
                c     = ws.cell(row, COL_IDX[col_letter], val)
                c.fill = row_fill; c.border = bdr()
                c.font = std_font()
                c.alignment = Alignment(horizontal="center")
                dv.add(c)

            r = row
            formula = (f'=IF(OR(C{r}="Berlin",D{r}="Berlin",E{r}="Berlin"),'
                       f'"✓ GF anwesend",'
                       f'IF(F{r}="Berlin","⚠ Fallback (Pablo)","✗ Nicht gedeckt"))')
            c = ws.cell(row, 7, formula)
            c.font = std_font(); c.fill = row_fill; c.border = bdr()
            c.alignment = Alignment(horizontal="center")

            hint = holidays.get(d, "")
            c = ws.cell(row, 8, hint)
            c.font = std_font("595959", italic=bool(hint))
            c.fill = row_fill; c.border = bdr()
            c.alignment = Alignment(horizontal="left", wrap_text=True)

            row += 1
        d += timedelta(days=1)

    last    = row - 1
    g_range = f"G4:G{last}"
    ws.conditional_formatting.add(g_range, FormulaRule(
        formula=['G4="✓ GF anwesend"'],
        fill=fill(OK_BG), font=Font(name="Arial", color=OK_FG, bold=True)))
    ws.conditional_formatting.add(g_range, FormulaRule(
        formula=['G4="⚠ Fallback (Pablo)"'],
        fill=fill(WRN_BG), font=Font(name="Arial", color=WRN_FG, bold=True)))
    ws.conditional_formatting.add(g_range, FormulaRule(
        formula=['G4="✗ Nicht gedeckt"'],
        fill=fill(ERR_BG), font=Font(name="Arial", color=ERR_FG, bold=True)))
    for col in "CDEF":
        ws.conditional_formatting.add(f"{col}4:{col}{last}", FormulaRule(
            formula=[f'{col}4="Abwesend (Urlaub)"'],
            fill=fill(VAC_BG), font=Font(name="Arial", color=VAC_FG, bold=True)))

    ws.freeze_panes = "A4"
    return last


def make_vacation_sheet(wb, absences, holidays):
    vs = wb.create_sheet("Urlaubsübersicht")
    vs.sheet_view.showGridLines = False

    vs.merge_cells("A1:F1")
    vs["A1"] = "Urlaubsübersicht – GF & Fallback"
    vs["A1"].font = Font(name="Arial", color=HDR_FG, bold=True, size=13)
    vs["A1"].fill = fill(HDR_BG)
    vs["A1"].alignment = Alignment(horizontal="center", vertical="center")
    vs.row_dimensions[1].height = 28

    for ci, (h, w) in enumerate(zip(
        ["Name", "Typ", "Von", "Bis", "Werktage", "Monat"],
        [24, 24, 14, 14, 12, 14]
    ), start=1):
        c = vs.cell(2, ci, h)
        c.font = hdr_font(bold=True); c.fill = fill(HDR_BG); c.border = bdr()
        c.alignment = Alignment(horizontal="center")
        vs.column_dimensions["ABCDEF"[ci-1]].width = w

    entries = []
    for ab in absences:
        if ab["employee_name"] not in PEOPLE or ab["type"] not in VACATION_TYPES:
            continue
        s  = date.fromisoformat(ab["start_date"])
        e  = date.fromisoformat(ab["end_date"])
        wd = sum(1 for i in range((e - s).days + 1)
                 if (s + timedelta(i)).weekday() < 5
                 and (s + timedelta(i)) not in holidays)
        entries.append((ab["employee_name"], ab["type"], s, e, wd,
                        MONTHS_DE.get(s.month, ""), ab["status"]))
    entries.sort(key=lambda x: (x[2], x[0]))

    for i, (name, typ, s, e, wd, mon, status) in enumerate(entries):
        rf = fill(GRP_BG) if i % 2 == 0 else fill("FFFFFF")
        for ci, val in enumerate([name, typ, s, e, wd, mon], start=1):
            c = vs.cell(3 + i, ci, val)
            c.fill = rf; c.border = bdr()
            c.font = std_font("9C5700" if status == "requested" else "000000",
                              italic=(status == "requested"))
            c.alignment = Alignment(horizontal="left" if ci <= 2 else "center")
            if ci in (3, 4): c.number_format = "DD.MM.YYYY"
    vs.freeze_panes = "A3"


def make_summary_sheet(wb, months, last_rows):
    ss = wb.create_sheet("Übersicht")
    ss.sheet_view.showGridLines = False

    ss.merge_cells("A1:F1")
    ss["A1"] = "Monatsübersicht – GF Berliner Büropräsenz"
    ss["A1"].font = Font(name="Arial", color=HDR_FG, bold=True, size=13)
    ss["A1"].fill = fill(HDR_BG)
    ss["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ss.row_dimensions[1].height = 28

    for ci, (h, w) in enumerate(zip(
        ["Monat", "Werktage", "✓ GF anwesend", "⚠ Fallback", "✗ Nicht gedeckt", "Abdeckung %"],
        [14, 13, 16, 14, 18, 14]
    ), start=1):
        c = ss.cell(2, ci, h)
        c.font = hdr_font(bold=True); c.fill = fill(HDR_BG); c.border = bdr()
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ss.column_dimensions["ABCDEF"[ci-1]].width = w
    ss.row_dimensions[2].height = 30

    for ri, m in enumerate(months, start=3):
        mname = MONTHS_DE[m]
        last  = last_rows[m]
        rf    = fill(GRP_BG) if ri % 2 == 0 else fill("FFFFFF")
        formulas = [
            mname,
            f"=COUNTA('{mname}'!A4:A{last})",
            f"=COUNTIF('{mname}'!G4:G{last},\"✓ GF anwesend\")",
            f"=COUNTIF('{mname}'!G4:G{last},\"⚠ Fallback (Pablo)\")",
            f"=COUNTIF('{mname}'!G4:G{last},\"✗ Nicht gedeckt\")",
            f"=IFERROR((C{ri}+D{ri})/B{ri},0)",
        ]
        for ci, val in enumerate(formulas, start=1):
            c = ss.cell(ri, ci, val)
            c.fill = rf; c.border = bdr(); c.font = std_font(bold=(ci==1))
            c.alignment = Alignment(horizontal="center")
            if ci == 6: c.number_format = "0%"
    ss.freeze_panes = "A3"


def build_excel(absences: list, year: int, month: int) -> Path:
    holidays = fetch_holidays(year)
    presence = build_presence(absences)

    wb = Workbook()
    wb.remove(wb.active)

    last_row = make_month_sheet(wb, month, year, presence, holidays)
    print(f"  ✓ {MONTHS_DE[month]}")

    make_vacation_sheet(wb, absences, holidays)
    make_summary_sheet(wb, [month], {month: last_row})

    month_str = f"{year}-{month:02d}"
    out = HERE / f"GF_Praesenzplanung_{month_str}.xlsx"
    wb.save(out)
    return out


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, calendar

    today = date.today()

    # Allow optional CLI override: python3 generate_praesenzplanung.py [YYYY-MM]
    if len(sys.argv) > 1:
        year, month = map(int, sys.argv[1].split("-"))
    else:
        # Default: next month
        if today.month == 12:
            year, month = today.year + 1, 1
        else:
            year, month = today.year, today.month + 1

    print(f"\nGeneriere GF-Präsenzplanung {MONTHS_DE[month]} {year}\n")

    month_start = date(year, month, 1)
    month_end   = date(year, month, calendar.monthrange(year, month)[1])

    absences = fetch_absences(start=month_start, end=month_end)

    try:
        print("\nBaue Excel ...")
        out = build_excel(absences, year, month)
    finally:
        # Raw-Dump enthält Personendaten – nach dem Lauf entfernen
        if DUMP_PATH.exists():
            DUMP_PATH.unlink()
            print(f"✓ Raw-Dump gelöscht: {DUMP_PATH.name}")

    print(f"\n✓ Fertig: {out.name}")
