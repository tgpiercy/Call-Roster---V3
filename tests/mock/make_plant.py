"""Generate the mock plant (P9): tests/mock/*.csv - about 70 made-up people over the 8 units, four crews and DAYS, with
the opening figures, the historic OT Log rows of periods 21-22 (Oct 4-31, 2026), the bookings, switches, postings,
planned moves and restrictions the alpha plan of Nov 1-14 exercises, and the mock Kronos data behind Worked (H).

    py -3.14 tests/mock/make_plant.py

Every name, phone, email and Payroll # is made up (names from two word lists, phones 905-555-01nn, emails
first.last@example.test, Payroll # 9 + 7 digits). The historic rows are produced by the engine in plant.py with the
same rules the roster applies, so IntegrityCheck finds them consistent. Deterministic: the same files every run.
tests/mock/README.md describes every column.
"""
from __future__ import annotations

import sys
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import rules  # noqa: E402
from plant import (LOG_HEADERS, UNITS, YEAR_FROM, Person, Plant, period_start, write_csv)  # noqa: E402

FIRST = ["Avery", "Blair", "Casey", "Dana", "Ellis", "Finley", "Harper", "Jordan", "Kendall", "Logan", "Morgan", "Quinn",
         "Reese", "Riley", "Rowan", "Sawyer", "Skyler", "Taylor", "Devon", "Emery", "Hayden", "Jesse", "Lane", "Marlow",
         "Oakley", "Parker", "Remy", "Shay", "Tatum", "Wren", "Arden", "Bryn", "Corin", "Darcy", "Eden", "Flynn", "Greer",
         "Hollis", "Indy", "Jules", "Kit", "Linden", "Merritt", "Nico", "Odell", "Perry", "Rafe", "Sidney", "Tobin",
         "Vale", "Wynn", "Ash", "Blake", "Cary", "Dale", "Evan", "Frankie", "Gale", "Haven", "Ira", "Jamie", "Kai",
         "Lee", "Max", "Noel", "Pat", "Robin", "Sam", "Teal", "Val"]
LAST = ["Ashgrove", "Bellweather", "Cardwell", "Dunmore", "Everhart", "Fairbanks", "Greenlaw", "Hollins", "Ironside",
        "Juniper", "Kestrel", "Lockridge", "Marsden", "Northway", "Oakhurst", "Pemberton", "Quarrie", "Redfern",
        "Stonebridge", "Thornbury", "Underhill", "Vantage", "Wexford", "Yardley", "Ashby", "Brightwater", "Coldstream",
        "Dovetail", "Ellery", "Fenwick", "Glassford", "Hartwell", "Islington", "Kingsley", "Larkspur", "Moorcroft",
        "Nettleton", "Orchard", "Pennyworth", "Ravenscroft", "Saltmarsh", "Tillingham", "Vauxhall", "Whitlock",
        "Amberley", "Birchwood", "Crossley", "Dalrymple", "Eastwick", "Fairweather", "Goldfinch", "Hawthorne",
        "Ingleby", "Jarrow", "Kelsall", "Lindhurst", "Mallory", "Norwood", "Overton", "Prentice", "Rushmore",
        "Silverton", "Tredwell", "Upton", "Wainwright", "Yeoman", "Ashworth", "Bramble", "Copperfield", "Driscoll"]

# (crew, unit, [grade codes]) - the plant's shape: four rotating crews of 14 and 13 people on DAYS
SHAPE = [
    ("A", "Phosphine", ["P1", "P1", "P2", "P3"]), ("A", "Derivatives", ["D1", "D1", "D1.5", "D2", "D3"]),
    ("A", "Cypure", ["G1", "G2"]), ("A", "Material Handler", ["M1", "M2"]), ("A", "QC Derivatives", ["Q1"]),
    ("B", "Phosphine", ["P1", "P1.5", "P2", "P2"]), ("B", "Derivatives", ["D1", "D1", "D2", "D2", "D3"]),
    ("B", "Cypure", ["G1", "G3"]), ("B", "Material Handler", ["M1", "M3"]), ("B", "QC Derivatives", ["Q2"]),
    ("C", "Phosphine", ["P1", "P1", "P2", "P3"]), ("C", "Derivatives", ["D1", "D1.5", "D2", "D2", "D3"]),
    ("C", "Cypure", ["G1", "G2"]), ("C", "Material Handler", ["M1", "M2"]), ("C", "QC Derivatives", ["Q1"]),
    ("D", "Phosphine", ["P1", "P1.5", "P2", "P2"]), ("D", "Derivatives", ["D1", "D1", "D2", "D2.5", "D3"]),
    ("D", "Cypure", ["G1", "G2"]), ("D", "Material Handler", ["M1", "M2"]), ("D", "QC Derivatives", ["Q3"]),
    ("DAYS", "QC Gases", ["QG1", "QG2", "QG3"]), ("DAYS", "Utilities", ["U1", "U2", "U3"]), ("DAYS", "SBP", ["S1", "S2", "S3"]),
    ("DAYS", "QC Derivatives", ["Q1", "Q2"]), ("DAYS", "Phosphine", ["P1"]), ("DAYS", "Derivatives", ["D1"]),
]

# seniority dates, in order of the people as SHAPE lists them (two exact-G ties by seniority, one by name - see TARGETS)
SENIORITY = [
    "2008-04-14", "2015-09-21", "2019-02-04", "2023-06-12", "2006-11-06", "2011-05-02", "2017-03-13", "2020-10-05", "2024-01-15",
    "2012-08-20", "2021-04-19", "2009-07-27", "2022-02-14", "2013-10-07",
    "2007-03-19", "2016-01-11", "2018-06-25", "2020-03-02", "2011-05-02", "2010-09-13", "2019-11-18", "2022-08-29", "2024-05-06",
    "2014-02-17", "2023-09-11", "2005-06-20", "2021-01-25", "2012-04-30",
    "2014-03-03", "2017-07-10", "2019-09-16", "2024-03-25", "2009-01-19", "2015-05-18", "2018-10-29", "2021-08-16", "2023-11-27",
    "2013-06-03", "2020-07-20", "2008-12-08", "2022-11-21", "2016-04-11",
    "2016-08-15", "2018-02-26", "2020-05-11", "2023-03-06", "2007-10-15", "2012-12-10", "2017-09-04", "2021-06-14", "2024-08-19",
    "2015-01-12", "2022-05-23", "2010-02-22", "2019-04-08", "2014-09-29",
    "2011-11-14", "2018-08-13", "2023-01-09", "2009-05-25", "2016-10-17", "2022-09-26", "2010-06-28", "2017-12-04", "2021-03-15",
    "2013-03-25", "2019-07-29", "2012-02-06", "2015-11-23",
]

# Opening Total Hours G per person (same order), hit through the OB- row: G = H + OB + historic charges.
# Ties (v29.1, review 2.6): a target is H + I at Day 0 (provisional OT sits on top of it). Index 29 (C Phosphine P1,
# Wren Fenwick) 173.5 carries the 12 h he accepted on the Oct 30 night (S-0005) as provisional OT (185.5 on screen at
# Day 0); D1.1's Oct 31 report credits that night (12 h CALLIN OT = +21 Worked, -12 provisional), so from Day 1 he reads
# 194.5 = index 43 (D Phosphine P1, Merritt Vauxhall): a REAL tie on the D1.5 list, broken by seniority 2016-08-15
# before 2017-07-10. (Before, the Kronos row was Merritt's by mistake and the Tatum / Wren tie at 171.5 existed only
# because Wren's provisional never dropped; a tie between the acceptor and the person left a message is impossible once
# Kronos moves - the acceptor ranks higher and gains more - and lowering Wren would swap the Oct 30 session's call
# order, so Merritt's target moved instead.) Index 28 (Tatum Ellery, 171.5) stays rank 1 of the Oct 30 session.
# index 5 (A Derivatives D1) and 18 (B Derivatives D1) both 203 with the SAME seniority 2011-05-02: the name decides.
TARGETS = [
    224, 161, 298.25, 350, 189, 203, 243.25, 310.5, 402, 176.5, 288, 154, 231, 267,
    205.25, 142, 266, 319, 203, 181.75, 258.5, 301, 388, 196, 275.5, 149, 240, 212,
    171.5, 173.5, 272, 338.5, 177, 199.25, 251, 294.75, 365, 163, 283, 158.25, 221, 247.5,
    194.5, 193, 262.25, 327, 169, 207.5, 236, 289, 379.5, 182, 296, 147, 229.75, 254,
    133, 188, 241.5, 126, 174, 219.5, 139, 195.5, 228, 151, 237, 143.5, 206,
]

# The mock Kronos year-to-date (Jan 11 - Oct 31) per person: pay code hours. H = the weighted total, rounded whole.
# (OVERTIME x 1.75, CALLIN OT x 1.75, MIN 4 x 1.5, LIEU OT x 1.5, OT - DT Split x 2, LIEU DT x 2, LIEU ST x 1)
def kronos_codes(i: int) -> dict:
    base = [("OVERTIME", 6 + (i * 7) % 30), ("CALLIN OT", (i * 3) % 10), ("MIN 4", 4 * ((i * 5) % 3))]
    if i % 3 == 0:
        base.append(("LIEU OT", 8))
    if i % 4 == 1:
        base.append(("OT - DT Split", 4))
    if i % 7 == 2:
        base.append(("LIEU DT", 4))
    if i % 5 == 3:
        base.append(("LIEU ST", 6))
    out = {c: h for c, h in base if h > 0}
    if i == 31:
        out["LIEU TKN"] = 8          # not a roster code: the import names it, IntegrityCheck's Kronos section lists it
    return out


LOGINS = ["mihaychukc", "sheltond", "campeauc", "piercyg"]
PERIOD_21 = date(2026, 10, 4)
PERIOD_22 = date(2026, 10, 18)
DAY0 = date(2026, 10, 31)


def build_people() -> list[Person]:
    people = []
    i = 0
    for crew, unit, codes in SHAPE:
        for code in codes:
            nm = f"{FIRST[i]} {LAST[i]}"
            p = Person(payroll=f"9{100001 + i:07d}", home=unit, name=nm, crew=crew, phone=f"905-555-01{i:02d}",
                       worked=0.0, seniority=date.fromisoformat(SENIORITY[i]),
                       email=f"{FIRST[i].lower()}.{LAST[i].lower()}@example.test", codes={u: "" for u in UNITS})
            p.codes[unit] = code
            people.append(p)
            i += 1
    # the two cross-coded people: B Cypure's G1 also M2 in Material Handler; the DAYS Derivatives D1 also Q2 in QC Derivatives
    by_name = {p.name: p for p in people}
    cyp_b = [p for p in people if p.crew == "B" and p.home == "Cypure"][0]
    cyp_b.codes["Material Handler"] = "M2"
    der_days = [p for p in people if p.crew == "DAYS" and p.home == "Derivatives"][0]
    der_days.codes["QC Derivatives"] = "Q2"
    return people


def who(people: list[Person], crew: str, unit: str, k: int = 0) -> Person:
    """The k-th person of a crew in a unit (home unit), in SHAPE order."""
    return [p for p in people if p.crew == crew and p.home == unit][k]


def main() -> int:
    people = build_people()
    plant = Plant()
    for p in people:
        plant.people[p.name] = p
    # Worked H from the mock cumulative
    totals = {}
    for i, p in enumerate(people):
        spec = kronos_codes(i)
        totals[p.name] = spec
    plant.import_cumulative(DAY0, totals)
    plant.hours_history = []                      # the Day-0 import writes the Oct 31 snapshot itself
    # ---- crew history: one dated row per person; one crew change on Oct 20 (CrewOn matters for Oct 4-19 history)
    crew_rows = []
    mover = who(people, "C", "Derivatives", 2)    # C Derivatives D2: was on crew A until Oct 19
    for p in people:
        if p is mover:
            crew_rows.append([p.name, "A", "2025-12-28", "piercyg", "2025-12-28 08:00:00"])
            crew_rows.append([p.name, "C", "2026-10-20", "piercyg", "2026-10-19 15:10:00"])
            plant.crew_hist[p.name] = [("A", YEAR_FROM), ("C", date(2026, 10, 20))]
        else:
            crew_rows.append([p.name, p.crew, "2025-12-28", "piercyg", "2025-12-28 08:00:00"])
            plant.crew_hist[p.name] = [(p.crew, YEAR_FROM)]
    # ---- the named people of the plan (comments say the role; the plan and expected.py name them by these roles)
    roles = {
        "sick_d1": who(people, "A", "Phosphine", 1),            # A Phosphine P1: sick on the Oct 31 night (3rd occurrence)
        "vacation": who(people, "D", "Derivatives", 3),         # D Derivatives D2.5: vacation Nov 2-5, open lines Nov 2-3 days
        "leave": who(people, "B", "Derivatives", 1),            # B Derivatives D1: STD leave since Oct 5, returns Mon Nov 9
        "personal": who(people, "B", "Phosphine", 2),           # B Phosphine P2: a personal day Tue Nov 10 (B days)
        "payback_owed": who(people, "A", "Cypure", 0),          # A Cypure G1 (off Sat Sep 12): worked that day shift for ...
        "payback_owes": who(people, "C", "Cypure", 0),          # C Cypure G1 (on days Sep 12): ... who gave it up and still owes the payback
        "restricted": who(people, "A", "Derivatives", 0),       # A Derivatives D1: OT restricted Nov 9-13
        "planned": who(people, "D", "Cypure", 1),               # D Cypure G2: planned move to Material Handler on Sun Nov 8
        "crew_change": mover,
    }
    # ---- absences (the loader writes them as the Absences table rows)
    abs_rows = []       # id, emp, unit, type, start, end, hours, times, note, home, group, shift, status, entered by, entered at
    cov_rows = []       # coverage lines: date, shift, unit, left, right, source, kind, status, note
    sick = roles["sick_d1"]
    # two past occurrences ON crew A's scheduled shifts: the first A day shift from Aug 15, the first A night from Sep 20
    s1 = first_shift("A", date(2026, 8, 15), "Days")
    s2 = first_shift("A", date(2026, 9, 20), "Nights")
    abs_rows.append(["SC-0001", sick.name, "Phosphine", "SIC", s1.isoformat(), s1.isoformat(), "", "", "", "Phosphine", "", "Days", "", "sheltond", f"{s1.isoformat()} 06:20:00"])
    abs_rows.append(["SC-0002", sick.name, "Phosphine", "SIC", s2.isoformat(), s2.isoformat(), "", "", "", "Phosphine", "", "Nights", "", "campeauc", f"{s2.isoformat()} 17:45:00"])
    v = roles["vacation"]
    abs_rows.append(["AB-0001", v.name, "Derivatives", "VAC", "2026-11-02", "2026-11-05", "", "", "", "Derivatives", "", "", "", "piercyg", "2026-09-30 09:12:00"])
    cov_rows.append(["CV-0001", "2026-11-02", "Days", "Derivatives", f"{v.name} (Vac)", "", "AB-0001", "VAC", "", ""])
    cov_rows.append(["CV-0002", "2026-11-03", "Days", "Derivatives", f"{v.name} (Vac)", "", "AB-0001", "VAC", "", ""])
    lv = roles["leave"]
    abs_rows.append(["AB-0002", lv.name, "Derivatives", "STD", "2026-10-05", "", "", "", "", "Derivatives", "", "", "", "piercyg", "2026-10-05 08:30:00"])
    per = roles["personal"]
    abs_rows.append(["AB-0003", per.name, "Phosphine", "PER", "2026-11-10", "2026-11-10", "", "", "", "Phosphine", "", "", "", "piercyg", "2026-10-28 14:05:00"])
    cov_rows.append(["CV-0003", "2026-11-10", "Days", "Phosphine", f"{per.name} (Personal)", "", "AB-0003", "PER", "", ""])
    old_vac = who(people, "C", "Phosphine", 3)
    abs_rows.append(["AB-0004", old_vac.name, "Phosphine", "VAC", "2026-09-14", "2026-09-17", "", "", "", "Phosphine", "", "", "", "piercyg", "2026-08-20 10:00:00"])
    for a in abs_rows:
        plant.absences.append({"id": a[0], "emp": a[1], "unit": a[2], "type": a[3], "start": date.fromisoformat(a[4]),
                               "end": date.fromisoformat(a[5]) if a[5] else None, "hours": a[6], "shift": a[11], "status": a[12], "home": a[9]})
    # ---- switches: one payback owed since Sep 12 (50 days by Nov 1: over PAYBACK_WARN_DAYS 42)
    sw_rows = []        # id, date, shift, unit, off, off crew, works, works crew, pay for, paid by, note, status, entered by, entered at
    sw_rows.append(["SS-0001", "2026-09-12", "Days", "Cypure", roles["payback_owes"].name, "C", roles["payback_owed"].name, "A",
                    "", "", "payback to be arranged", "Active", "mihaychukc", "2026-09-10 13:40:00"])
    plant.switches.append({"id": "SS-0001", "date": date(2026, 9, 12), "shift": "Days", "unit": "Cypure", "off": roles["payback_owes"].name,
                           "works": roles["payback_owed"].name, "payfor": "", "paidby": "", "status": "Active"})
    cov_rows.append(["CV-0004", "2026-09-12", "Days", "Cypure", f"{roles['payback_owes'].name} S/S", f"{roles['payback_owed'].name} S/S", "SS-0001", "SS", "", ""])
    # ---- restriction: Nov 9-13
    rs_rows = [["RS-0001", roles["restricted"].name, "2026-11-09", "2026-11-13", "medical - no overtime this week", "Active", "piercyg", "2026-10-29 16:00:00", ""]]
    plant.restrictions.append({"id": "RS-0001", "emp": roles["restricted"].name, "start": date(2026, 11, 9), "end": date(2026, 11, 13),
                               "reason": "medical - no overtime this week", "status": "Active"})
    # ---- planned move: D Cypure G2 -> Material Handler M2 on Sun Nov 8 (crew stays D)
    pm_rows = [["PM-0001", roles["planned"].name, "Material Handler", "M2", "Cypure", "", "D", "2026-11-08", "Planned", "piercyg", "2026-10-27 11:30:00", "", ""]]
    plant.planned.append({"id": "PM-0001", "emp": roles["planned"].name, "unit": "Material Handler", "code": "M2", "clear": "Cypure",
                          "crew": "", "start": date(2026, 11, 8), "status": "Planned"})

    # ---- the OT Log: OB- rows first (two passes: the historic call orders rest on H + OB, then OB is re-solved)
    hist_spec = [
        # (kind, ref, date, unit, shift, hours, login, time, outcomes by rank or award spec)
        ("S", "2026-10-06", "Phosphine", "Nights", 12, "sheltond", "17:35", [("Refused", 0, ""), ("Unavailable", 0, "Out of town"), ("Refused", 0, ""), ("Accepted", 0, "")]),
        ("P", "2026-10-10", "Derivatives", "Days", 12, "campeauc", "09:10", {"posted": "2026-09-24", "closes": "2026-09-26", "awarded_on": "2026-09-28",
                                                                              "reason": "Vacation Coverage", "award_rank": 3, "signed_ranks": [3, 5]}),
        ("S", "2026-10-13", "Cypure", "Days", 12, "mihaychukc", "05:50", [("Left Message", 0, ""), ("Accepted", 0, "")]),
        ("S", "2026-10-20", "Derivatives", "Nights", 12, "sheltond", "16:20", [("Refused", 0, ""), ("Refused", 0, ""), ("Exception", 0, rules_exception()), ("Accepted", 0, "")]),
        ("P", "2026-10-24", "Material Handler", "Days", 12, "piercyg", "08:05", {"posted": "2026-10-08", "closes": "2026-10-10", "awarded_on": "2026-10-16",
                                                                                   "reason": "Training Coverage", "no_takers": True}),
        ("S", "2026-10-24", "Utilities", "Days", 12, "campeauc", "05:30", [("Refused", 0, ""), ("Accepted", 0, "")]),
        ("S", "2026-10-30", "Phosphine", "Nights", 12, "mihaychukc", "17:10", [("Left Message", 0, ""), ("Accepted", 0, "")]),
        ("S", "2026-10-31", "Derivatives", "Days", 12, "sheltond", "05:40", [("Refused", 0, ""), ("Left Message", 0, ""), ("Accepted", 0, "")]),
    ]
    for pass_no in (1, 2):
        hist_charges = {p.name: 0.0 for p in people} if pass_no == 1 else {
            p.name: sum(float(r["Hours Charged"] or 0) for r in plant.log if r["Employee"] == p.name and r["Ref ID"] != f"OB-{people.index(p) + 1:04d}")
            for p in people}
        plant.log = []
        plant.entry_next = 1
        plant.postings, plant.posting_lists = {}, {}
        plant.coverage = []
        ob_rows = []
        for i, p in enumerate(people):
            ob = round(TARGETS[i] - plant.worked[p.name] - hist_charges[p.name], 2)
            if ob < 0:
                raise SystemExit(f"{p.name}: the target {TARGETS[i]} is under H {plant.worked[p.name]} + historic charges {hist_charges[p.name]}")
            r = plant.blank_row()
            r.update({"Ref ID": f"OB-{i + 1:04d}", "Type": "Opening Balance", "Shift Date": date(2026, 1, 4), "Unit": p.home, "Shift": "",
                      "Crew": "", "Reason": "Opening balance", "Hours": 0, "Call Order": "", "Employee": p.name,
                      "Dept Code": p.codes[p.home], "Total Hrs Then": 0, "Outcome": "Adjustment",
                      "Note": "Opening balance: charges carried into 2026 before the roster went live", "Hours Charged": ob,
                      "Period Start": YEAR_FROM, "Logged By": "piercyg", "Timestamp": datetime(2026, 1, 4, 9, 0, 0), "Status": "", "Void Note": ""})
            ob_rows.append(r)
        plant.append_rows(ob_rows)
        run_history(plant, people, hist_spec)
    # the historic coverage lines for every Accepted / Awarded row
    n_cv = len(cov_rows)
    for r in plant.log:
        if r["Outcome"] in ("Accepted", "Awarded") and r["Status"] != "VOID":
            n_cv += 1
            cov_rows.append([f"CV-{n_cv:04d}", r["Shift Date"].isoformat(), r["Shift"], r["Unit"],
                             f"Call-out {r['Ref ID']}" if r["Type"] == "Callout" else r["Reason"], r["Employee"], r["Ref ID"],
                             "SUP" if r["Type"] == "Callout" else REASON_TYPE.get(r["Reason"], "SUP"), "", ""])
    # ---- the three open postings (their lists frozen with the totals of their posting day: today's figures, as the
    #      loader's verify sees them - the plan reads the list as loaded)
    make_open_postings(plant, people, roles)

    # the calendar lines the postings wrote: POSTED on P-0003's shift line (the vacation's Nov 3 line), need lines for
    # P-0004 and P-0005, and P-0002's NO TAKERS line (CreatePostingFrom -> CalendarPosted; the award / no-takers outcome)
    for c in cov_rows:
        if c[6] == "AB-0001" and c[1] == "2026-11-03":
            c[5], c[6], c[8] = "POSTED P-0003", "P-0003", "AB-0001"      # the line keeps its absence (CV_ABS): SlotsFor joins them by it
    for pid, p in plant.postings.items():
        if p["status"] == "Open" and pid != "P-0003":
            n_cv += 1
            cov_rows.append([f"CV-{n_cv:04d}", p["date"].isoformat(), p["shift"], p["unit"], p["reason"], f"POSTED {pid}", pid,
                             REASON_TYPE.get(p["reason"], "SUP"), "", ""])
        if p["status"] == "No takers":
            n_cv += 1
            cov_rows.append([f"CV-{n_cv:04d}", p["date"].isoformat(), p["shift"], p["unit"], p["reason"], "NO TAKERS", pid,
                             REASON_TYPE.get(p["reason"], "SUP"), "", ""])
    # ---- write everything
    out = HERE
    write_csv(out / "people.csv", ["Payroll #", "Home Unit", "Name", "Crew", "Phone", "Worked H", "Seniority", "Email"] + UNITS,
              [[p.payroll, p.home, p.name, p.crew, p.phone, f"{plant.worked[p.name]:g}", p.seniority.isoformat(), p.email] + [p.codes[u] for u in UNITS]
               for p in people])
    write_csv(out / "crew_history.csv", ["Employee", "Crew", "Effective From", "Logged By", "Timestamp"], crew_rows)
    write_csv(out / "kronos_totals.csv", ["Payroll #", "Employee", "Pay Code", "Hours"],
              [[p.payroll, p.name, c, h] for p in people for c, h in kronos_codes(people.index(p)).items()])
    # Hours History: the two snapshots before Day 0 (through Oct 3 and Oct 17), H a little under the opening figure
    hh = []
    for i, p in enumerate(people):
        h31 = plant.worked[p.name]
        h17 = max(0.0, h31 - (7 + (i * 3) % 19))
        h3 = max(0.0, h17 - (5 + (i * 5) % 23))
        hh.append(["KR-0021", "2026-10-04 07:40:00", "2026-10-04", p.name, f"{h3:g}", "piercyg", "2026-10-03", 2026])
        hh.append(["KR-0022", "2026-10-18 07:35:00", "2026-10-18", p.name, f"{h17:g}", "piercyg", "2026-10-17", 2026])
    write_csv(out / "hours_history.csv", ["Import ID", "Imported At", "Kronos File Date", "Employee", "Worked Hours", "Imported By", "Data Through", "Year"], hh)
    # as the loader leaves the file: H = the Oct 18 snapshot, KronosDataThrough = Oct 17 (the Day-0 import moves both)
    h17 = {r[3]: float(r[4]) for r in hh if r[6] == "2026-10-17"}
    plant.worked = dict(h17)
    plant.kronos_through = date(2026, 10, 17)
    el = []
    for p in people:
        g, h, i_, c = plant.ghic(p.name)
        psum = sum(float(r["Hours Charged"] or 0) for r in plant.log if r["Employee"] == p.name and r["Status"] != "VOID")
        el.append([p.name, f"{g:.2f}", f"{i_:.2f}", f"{psum:.2f}", f"{h:g}", f"{c:g}"])
    el.append(["(log total Hours Charged)", f"{sum(float(r['Hours Charged'] or 0) for r in plant.log if r['Status'] != 'VOID'):.2f}", "", "", "", ""])
    write_csv(out / "expected_load.csv", ["Employee", "G Total", "I Refused", "Period cells sum", "H Worked", "C Provisional"], el)
    write_csv(out / "absences.csv", ["Absence ID", "Employee", "Unit", "Type", "Start", "End", "Hours / Shift", "Times", "Note", "Status",
                                     "Logged By", "Timestamp", "Void Note", "Home Unit", "Leave Group", "Shift"],
              [[a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], "Active", a[13], a[14], "", a[9], a[10], a[11]] for a in abs_rows])
    write_csv(out / "coverage.csv", ["Line ID", "Date", "Shift", "Unit", "Absence ID", "Off / Need", "Type", "Covered By", "Hours", "Source",
                                     "Status", "Logged By", "Timestamp", "Note"],
              [[c[0], c[1], c[2], c[3], c[8] or (c[6] if c[6].startswith("AB-") else ""), c[4], c[7], c[5], "", c[6], "Active", "piercyg", c[1] + " 08:00:00", ""] for c in cov_rows])
    write_csv(out / "switches.csv", ["SS ID", "Shift Date", "Shift", "Unit", "Gives Up", "Off Crew", "Works", "Works Crew", "Payback For",
                                     "Paid Back By", "Note", "Status", "Logged By", "Timestamp", "Void Note"], [r + [""] for r in sw_rows])
    write_csv(out / "restrictions.csv", ["Restriction ID", "Employee", "From", "To", "Reason", "Status", "Entered By", "Entered At", "Void Note"], rs_rows)
    write_csv(out / "planned_moves.csv", ["ID", "Employee", "New Unit", "Grade Code", "Codes To Clear", "New Crew", "Old Crew", "Effective From",
                                          "Status", "Entered By", "Entered On", "Closed On", "Note"], pm_rows)
    log_rows = []
    for r in plant.log:
        row = []
        for h in LOG_HEADERS:
            v = r[h]
            if isinstance(v, datetime):
                v = v.strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(v, date):
                v = v.isoformat()
            elif isinstance(v, float):
                v = f"{v:g}"
            row.append(v)
        log_rows.append(row)
    write_csv(out / "log_opening.csv", LOG_HEADERS, log_rows)
    post_rows, list_rows = [], []
    for pid, p in plant.postings.items():
        post_rows.append([pid, p["status"], p["posted"].isoformat(), p["closes"].isoformat(), p["date"].isoformat(), p["unit"], p["shift"],
                          p["crew"], f"{p['hours']:g}", p["positions"], p["reason"], p["award_to"], p.get("award_on", ""),
                          p.get("award_by", ""), p["posted_by"]])
        for row in plant.posting_lists[pid]:
            list_rows.append([pid, row["rank"], row["name"], row["crew"], row["code"], f"{row['total']:g}",
                              "YES" if row["signed"] else "", row["result"], f"{row['charged']:g}" if row.get("charged") else "",
                              f"{row['signup_hrs']:g}" if row["signup_hrs"] else ""])
    write_csv(out / "postings.csv", ["Posting ID", "Status", "Posted On", "Closes On", "Shift Date", "Unit", "Shift", "Crew", "Hours",
                                     "Positions", "Reason", "Awarded To", "Awarded On", "Awarded By", "Posted By"], post_rows)
    write_csv(out / "posting_list.csv", ["Posting ID", "Rank", "Employee", "Shift", "Dept Code", "Total Hrs (posted)", "Signed Up", "Result",
                                         "Hours Charged", "Signup Hours"], list_rows)
    # only the log-based prefixes read a mark (NextRefID / NextEntryID / NextRequestId); AB, SC, CV, SS, PM, RS, KR count from their tables
    marks = {"OB": len(people), "S": 0, "P": 0, "ADJ": 0, "RQ": 0, "LV": 0, "UM": 0, "NH": 0, "MAN": 0}
    for r in plant.log:
        pre, _, num = r["Ref ID"].partition("-")
        marks[pre] = max(marks.get(pre, 0), int(num))
    for pid in plant.postings:
        marks["P"] = max(marks["P"], int(pid[2:]))
    write_csv(out / "idmarks.csv", ["Prefix", "Number"], [[k, v] for k, v in marks.items()])
    # the loader's verify: G / H / I / C per person after the load, and the log's charged total
    write_csv(out / "roles.csv", ["Role", "Employee"], [[k, v.name] for k, v in roles.items()])
    # the one-day Kronos reports: Report = day (the morning after) or re (the Nov 3 day re-pulled on Nov 6)
    by_name = {p.name: p for p in people}
    kd = []
    for rep, d, nm, code, h in KRONOS_DAYS:
        kd.append([rep, d, by_name[nm].payroll, nm, code, h])
    write_csv(out / "kronos_days.csv", ["Report", "Date", "Payroll #", "Employee", "Pay Code", "Hours"], kd)
    print(f"people {len(people)}, log rows {len(plant.log)}, postings {len(plant.postings)}, coverage {len(cov_rows)}")
    return 0


# (report, day, person, pay code, hours): the OT the plan's accepted call-outs and awards produce, plus fillers. A NIGHT
# shift's hours are dated the day the night ENDS (Kronos's day; the roster's coverage rule: a night on D is covered when
# D + 1 <= Data Through): Wren's Oct 30 night (S-0005, v29.1 review 2.6: the log says Wren accepted it, not Merritt) -> Oct 31,
# Merritt's Oct 31 night (6 h) -> Nov 1, Odell's Nov 4 night -> Nov 5,
# Finley's Nov 11 night -> Nov 12; Teal's Oct 31 DAY -> Oct 31.
KRONOS_DAYS = [
    ("day", "2026-10-31", "Wren Fenwick", "CALLIN OT", 12), ("day", "2026-10-31", "Teal Copperfield", "OVERTIME", 12),
    ("day", "2026-10-31", "Rowan Oakhurst", "OVERTIME", 4), ("day", "2026-10-31", "Hollis Orchard", "MIN 4", 4),
    ("day", "2026-11-01", "Merritt Vauxhall", "CALLIN OT", 6), ("day", "2026-11-01", "Avery Ashgrove", "OVERTIME", 2),
    ("day", "2026-11-01", "Casey Cardwell", "MIN 4", 4), ("day", "2026-11-01", "Jules Ravenscroft", "LIEU OT", 4),
    ("day", "2026-11-02", "Hayden Underhill", "OVERTIME", 12), ("day", "2026-11-02", "Kit Saltmarsh", "OVERTIME", 3),
    ("day", "2026-11-03", "Harper Greenlaw", "OVERTIME", 12), ("day", "2026-11-03", "Cary Jarrow", "MIN 4", 4),
    ("re", "2026-11-03", "Harper Greenlaw", "OVERTIME", 10), ("re", "2026-11-03", "Cary Jarrow", "MIN 4", 4),
    ("day", "2026-11-04", "Nico Whitlock", "OVERTIME", 12), ("day", "2026-11-04", "Parker Brightwater", "OVERTIME", 2),
    ("day", "2026-11-05", "Odell Amberley", "OVERTIME", 12), ("day", "2026-11-05", "Dale Kelsall", "OVERTIME", 2),
    ("day", "2026-11-05", "Gale Norwood", "MIN 4", 4),
    ("day", "2026-11-06", "Tobin Eastwick", "MIN 4", 4), ("day", "2026-11-06", "Max Upton", "OVERTIME", 3),
    ("day", "2026-11-07", "Hayden Underhill", "OVERTIME", 12), ("day", "2026-11-07", "Ash Hawthorne", "LIEU ST", 4),
    ("day", "2026-11-08", "Sidney Dalrymple", "LIEU ST", 6), ("day", "2026-11-08", "Frankie Mallory", "OVERTIME", 2),
    ("day", "2026-11-09", "Hollis Orchard", "CALLIN OT", 12), ("day", "2026-11-09", "Wynn Goldfinch", "OVERTIME", 2),
    ("day", "2026-11-10", "Jamie Rushmore", "MIN 4", 4), ("day", "2026-11-10", "Reese Marsden", "OVERTIME", 3),
    ("day", "2026-11-11", "Sam Bramble", "STATWK 2.0", 12),
    ("day", "2026-11-12", "Finley Fairbanks", "OVERTIME", 12), ("day", "2026-11-12", "Perry Birchwood", "MIN 4", 4),
    ("day", "2026-11-12", "Noel Wainwright", "OVERTIME", 2),
    ("day", "2026-11-13", "Greer Nettleton", "OVERTIME", 3), ("day", "2026-11-13", "Quinn Lockridge", "MIN 4", 4),
    ("cum", "2026-11-14", "Taylor Redfern", "OVERTIME", 2),
]


REASON_TYPE = {"Vacation Coverage": "VAC", "Leave Coverage": "LV", "Training Coverage": "TRN", "PHA": "SUP", "Unit Operator Support": "SUP", "Other": "SUP"}


def first_shift(crew: str, start: date, shift: str) -> date:
    d = start
    while rules.crew_shift_on(crew, d) != shift:
        d += timedelta(days=1)
    return d


def rules_exception() -> str:
    return "Plant upset - called in for the unit, not an OT refusal"


def run_history(plant: Plant, people: list[Person], spec: list) -> None:
    """The historic sessions and postings of periods 21-22, through the engine (call order as the roster would have had it)."""
    n_s = n_p = 0
    for kind, d_txt, unit, shift, hours, login, hm, detail in spec:
        d = date.fromisoformat(d_txt)
        when = datetime.combine(d, datetime.strptime(hm, "%H:%M").time())
        if kind == "S":
            lst, removed = plant.callout_list(unit, d, shift)
            if len(lst) < len(detail):
                raise SystemExit(f"history {unit} {d} {shift}: only {len(lst)} listed, {len(detail)} outcomes wanted; off: {removed}")
            outcomes = [(lst[k].name, oc, taken, reason) for k, (oc, taken, reason) in enumerate(detail)]
            plant.callout(unit, d, shift, hours, outcomes, when, login)
            n_s += 1
        else:
            n_p += 1
            pid = f"P-{n_p:04d}"
            posted = date.fromisoformat(detail["posted"])
            plant.create_posting(pid, unit, d, shift, hours, 1, date.fromisoformat(detail["closes"]), detail["reason"], posted, login)
            aw_when = datetime.combine(date.fromisoformat(detail["awarded_on"]), datetime.strptime(hm, "%H:%M").time())
            if detail.get("no_takers"):
                plant.no_takers(pid, aw_when, login)
            else:
                lst = plant.posting_lists[pid]
                for rk in detail["signed_ranks"]:
                    plant.sign_up(pid, lst[rk - 1]["name"])
                res = plant.award(pid, {lst[detail["award_rank"] - 1]["name"]: 0.0}, aw_when, login)
                if res.get("problem"):
                    raise SystemExit(f"history {pid}: award problem {res['problem']}")
            plant.postings[pid]["award_on"] = aw_when.date().isoformat()
            plant.postings[pid]["award_by"] = login


def make_open_postings(plant: Plant, people: list[Person], roles: dict) -> None:
    """P-0003: Derivatives Tue Nov 3 Days (the vacation), closes Mon Nov 2 - two sign-up files already in Signups\\.
    P-0004: Material Handler Thu Nov 5 Days, closes Sun Nov 1 - no sign-ups (Ray's exception: closing soon, then no takers).
    P-0005: Phosphine Fri Nov 20 Days (period 24), closes Fri Nov 6."""
    n = max(int(k[2:]) for k in plant.postings)
    plant.create_posting(f"P-{n + 1:04d}", "Derivatives", date(2026, 11, 3), "Days", 12, 1, date(2026, 11, 2), "Vacation Coverage",
                         date(2026, 10, 20), "campeauc")
    plant.create_posting(f"P-{n + 2:04d}", "Material Handler", date(2026, 11, 5), "Days", 12, 1, date(2026, 11, 1), "Training Coverage",
                         date(2026, 10, 24), "sheltond")
    plant.create_posting(f"P-{n + 3:04d}", "Phosphine", date(2026, 11, 20), "Days", 12, 1, date(2026, 11, 6), "Unit Operator Support",
                         date(2026, 10, 28), "piercyg")


if __name__ == "__main__":
    sys.exit(main())
