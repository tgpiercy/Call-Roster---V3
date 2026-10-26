"""Python mirror of the Call Out Roster's pure rules.

Each function here mirrors a VBA procedure named in its docstring. The
known answers in cases.csv are shared with the workbook's SelfTest
(modSelfTest.bas), so the two cannot drift apart. Nothing here reads a
workbook: the rotation, holidays, DST, call order, charging, the rotation
part of eligibility and the provisional-OT date logic are all arithmetic.
"""
from __future__ import annotations

import re

import math
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

# ---------------------------------------------------------------- settings
CHARGE_MULTIPLIER = 1.75        # modOvertimeEntry.CHARGE_MULTIPLIER
CHARGE_MAX_GRADE = 2            # modOvertimeEntry.CHARGE_MAX_GRADE
PEER_MAX_GRADE = 2              # modLeave.PEER_MAX_GRADE
DST_USE_ACTUAL_HOURS = True     # modOvertimeEntry.DST_USE_ACTUAL_HOURS
USUAL_SHIFT_HOURS = 12.0
# v27 (briefs/v27-kronos.md): the Kronos pay codes that count toward Worked (H) and their straight-time multipliers
# (modOvertimeEntry.PREMIUM_CODES / PREMIUM_MULTIPLIERS, parallel lists; test_drift compares them)
PREMIUM_CODES = "OVERTIME,OT - DT Split,LIEU OT,LIEU DT,LIEU OT - DT Split,LIEU ST,CALLIN OT,MIN 4,STATWK 2.0"
PREMIUM_MULTIPLIERS = "1.75,2,1.5,2,1.5,1,1.75,1.5,2"
KRONOS_YEAR_SLACK = 14          # ImportKronosData.KRONOS_YEAR_SLACK: a cumulative report starts within this of the year's first day

# ---------------------------------------------------------------- rotation
# modShiftCycle: 28-day cycle, day 1 = Sun Dec 28, 2025 (B days / A nights)
ROT_ANCHOR = date(2025, 12, 28)
ROT_DAYS = "BDDAACCCBBDDAAACCBBDDDAACCBB"
ROT_NIGHTS = "ACCBBDDDAACCBBBDDAACCCBBDDAA"
ROT_LEN = 28


def rot_idx(d: date) -> int:
    """modShiftCycle.RotIdx: 1-based position in the cycle."""
    return (d - ROT_ANCHOR).days % ROT_LEN + 1


def days_crew(d: date) -> str:
    return ROT_DAYS[rot_idx(d) - 1]


def nights_crew(d: date) -> str:
    return ROT_NIGHTS[rot_idx(d) - 1]


def crews_off(d: date) -> str:
    return "/".join(c for c in "ABCD" if c not in (days_crew(d), nights_crew(d)))


def norm_code(code: str) -> str:
    return code.strip().upper().replace(" ", "")


def crew_shift_on(code: str, d: date, observed: dict | None = None) -> str:
    """modShiftCycle.CrewShiftOn: 'Days', 'Nights', '' (off) or '?' (unknown).

    A/B and C/D work their pair's day shifts only; DAYS works Mon-Fri days
    except plant holidays (the observed date when one is entered).
    """
    s = norm_code(code)
    if len(s) == 1 and s in "ABCD":
        if s == days_crew(d):
            return "Days"
        if s == nights_crew(d):
            return "Nights"
        return ""
    if len(s) == 3 and s[1] == "/" and s[0] in "ABCD" and s[2] in "ABCD":
        return "Days" if days_crew(d) in s else ""
    if s == "DAYS":
        if d.weekday() <= 4 and not days_holiday_off(d, observed):
            return "Days"
        return ""
    return "?"


def available_crew(d: date, is_night: bool) -> str:
    """modShiftCycle.AvailableCrew: the off crews minus the rest-rule crew."""
    blocked = days_crew(d + timedelta(days=1)) if is_night else nights_crew(d - timedelta(days=1))
    return "/".join(c for c in "ABCD" if c not in (days_crew(d), nights_crew(d)) and c != blocked)


# ---------------------------------------------------------------- holidays
HOLIDAY_NAMES = [
    "New Year's Day", "Family Day", "Good Friday", "Victoria Day", "Canada Day",
    "Civic Holiday", "Labour Day", "Thanksgiving", "Remembrance Day",
    "Christmas Day", "Boxing Day",
]


def nth_monday(y: int, m: int, n: int) -> date:
    d = date(y, m, 1)
    d += timedelta(days=(7 - d.weekday()) % 7)      # first Monday
    return d + timedelta(days=7 * (n - 1))


def easter_sunday(y: int) -> date:
    """modHolidays.EasterSunday (anonymous Gregorian algorithm)."""
    a, b, c = y % 19, y // 100, y % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    month = (h + L - 7 * m + 114) // 31
    day = (h + L - 7 * m + 114) % 31 + 1
    return date(y, month, day)


def holiday_date(name: str, y: int) -> date:
    """modHolidays.HolidayDate."""
    if name == "New Year's Day":
        return date(y, 1, 1)
    if name == "Family Day":
        return nth_monday(y, 2, 3)
    if name == "Good Friday":
        return easter_sunday(y) - timedelta(days=2)
    if name == "Victoria Day":
        d = date(y, 5, 24)
        return d - timedelta(days=d.weekday())        # Monday on/before May 24
    if name == "Canada Day":
        return date(y, 7, 1)
    if name == "Civic Holiday":
        return nth_monday(y, 8, 1)
    if name == "Labour Day":
        return nth_monday(y, 9, 1)
    if name == "Thanksgiving":
        return nth_monday(y, 10, 2)
    if name == "Remembrance Day":
        return date(y, 11, 11)
    if name == "Christmas Day":
        return date(y, 12, 25)
    if name == "Boxing Day":
        return date(y, 12, 26)
    raise ValueError(name)


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def days_holiday_off(d: date, observed: dict | None = None) -> bool:
    """modHolidays.DaysHolidayOff: are DAYS workers off on d?

    observed maps (year, holiday name) -> the weekday the plant gives off.
    With an observed date, that date is the day off (whatever weekday the
    holiday falls on); without one, a weekday holiday is off on the day
    and a weekend holiday gives no day off until it is entered.
    """
    observed = observed or {}
    for yy in (d.year - 1, d.year, d.year + 1):
        for nm in HOLIDAY_NAMES:
            hd = holiday_date(nm, yy)
            obs = observed.get((yy, nm))
            if obs is not None:
                if obs == d:
                    return True
            elif not is_weekend(hd) and hd == d:
                return True
    return False


# ---------------------------------------------------------------- DST
def nth_sunday(y: int, m: int, n: int) -> date:
    d = date(y, m, 1)
    d += timedelta(days=(6 - d.weekday()) % 7)      # first Sunday
    return d + timedelta(days=7 * (n - 1))


def dst_fall_back(y: int) -> date:
    """modShiftCycle.DstFallBack: the Saturday whose night is 13 h."""
    return nth_sunday(y, 11, 1) - timedelta(days=1)


def dst_spring_forward(y: int) -> date:
    """modShiftCycle.DstSpringForward: the Saturday whose night is 11 h."""
    return nth_sunday(y, 3, 2) - timedelta(days=1)


def dst_night_hours(d: date) -> float:
    """modShiftCycle.DstNightHours: 13, 11 or 12."""
    if d == dst_fall_back(d.year):
        return 13.0
    if d == dst_spring_forward(d.year):
        return 11.0
    return 12.0


def default_shift_hours(d: date, is_night: bool, usual: float = USUAL_SHIFT_HOURS) -> float:
    """modShiftCycle.DefaultShiftHours."""
    if is_night and DST_USE_ACTUAL_HOURS and dst_night_hours(d) != 12:
        return dst_night_hours(d)
    return usual


def default_shift_date(now: datetime) -> date:
    """modShiftCycle.DefaultShiftDate (v26.5.5, B1): before 07:00 the night under way is dated yesterday."""
    today = now.date()
    return today - timedelta(days=1) if now.hour < 7 else today


# ---------------------------------------------------------------- call order
@dataclass
class Emp:
    name: str
    grade: float
    total_hrs: float
    seniority: date | None     # earliest first; None sorts last


def seniority_key(s: date | None) -> float:
    """modOvertimeEntry.SeniorityKey: the date serial, or 1E15 when blank."""
    if s is None:
        return 1e15
    return float((s - date(1899, 12, 30)).days)


def call_order(people: list[Emp]) -> list[Emp]:
    """modOvertimeEntry.SortEmpRecs / RecBefore:
    grade (low first) > total hours (low first) > seniority (earliest) > name."""
    return sorted(people, key=lambda e: (e.grade, round(e.total_hrs, 2), seniority_key(e.seniority), e.name.casefold()))


def list_insert_index(frozen: list[Emp], new: Emp) -> int:
    """modOvertimeEntry.ListInsertIndex: where a newcomer goes in a list that is
    already in order and is not re-sorted (a posting's frozen list): before the
    first row they come before in call order. 1 .. n + 1."""
    def key(e: Emp):
        return (e.grade, e.total_hrs, seniority_key(e.seniority), e.name.casefold())
    for i, e in enumerate(frozen, start=1):
        if key(new) < key(e):
            return i
    return len(frozen) + 1


# ---------------------------------------------------------------- charging
def callout_charge(outcome: str, grade: float, hours: float) -> float:
    """modOvertimeEntry.CalloutCharge."""
    if grade > CHARGE_MAX_GRADE:
        return 0.0
    if outcome in ("Refused", "Left Message"):
        return hours * CHARGE_MULTIPLIER
    return 0.0


@dataclass
class Listed:
    rank: int
    grade: float
    signed_up: bool = False
    awarded: bool = False
    charged: float = field(default=0.0)
    result: str = ""
    exempt: bool = False       # v24: moved unit or left before the shift - never charged on this posting
    signup_hrs: float = 0.0    # v24: 0 = signed up for the whole shift; otherwise the part they'll take


def compute_award(hours: float, lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeAward: anyone ahead of the LAST awarded person in
    call order who did not sign up is charged (grade cap applies)."""
    max_rank = max((p.rank for p in lst if p.awarded), default=0)
    n_awarded = sum(1 for p in lst if p.awarded)
    for p in lst:
        p.charged = 0.0
        if p.awarded:
            p.result = "Awarded"
        elif p.signed_up:
            p.result = "Signed up"
        elif p.exempt:
            p.result = "Exempt"
        elif n_awarded > 0 and p.rank < max_rank and p.grade <= CHARGE_MAX_GRADE:
            p.charged = hours * CHARGE_MULTIPLIER
            p.result = "Charged"
        else:
            p.result = ""
    return lst


def is_partial_signup(post_hours: float, signup_hrs: float) -> bool:
    """modPostings.IsPartialSignup."""
    return 0 < signup_hrs < post_hours


def awarded_hours(post_hours: float, signup_hrs: float) -> float:
    """modPostings.AwardedHours."""
    return signup_hrs if is_partial_signup(post_hours, signup_hrs) else post_hours


def award_order_problem(hours: float, positions: int, lst: list[Listed]) -> str:
    """modPostings.AwardOrderProblem: '' / 'full-first' / 'filled' / 'over-hours'."""
    n_full_aw = n_part_aw = n_full_waiting = 0
    part_hrs = 0.0
    for p in lst:
        if p.awarded:
            if is_partial_signup(hours, p.signup_hrs):
                n_part_aw += 1
                part_hrs += p.signup_hrs
            else:
                n_full_aw += 1
        elif p.signed_up and not is_partial_signup(hours, p.signup_hrs):
            n_full_waiting += 1
    if n_part_aw == 0:
        return ""
    pos = max(positions, 1)
    if n_full_aw >= pos:
        return "filled"
    if n_full_waiting > 0:
        return "full-first"
    if part_hrs > (pos - n_full_aw) * hours + 0.0001:
        return "over-hours"
    return ""


def award_coverage(hours: float, positions: int, lst: list[Listed]) -> tuple[float, float]:
    """modPostings.AwardCoverage: (hours covered, hours left uncovered)."""
    covered = sum(awarded_hours(hours, p.signup_hrs) for p in lst if p.awarded)
    return covered, max(0.0, max(positions, 1) * hours - covered)


def compute_no_takers(hours: float, lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeNoTakers: everyone eligible is charged."""
    for p in lst:
        p.awarded = False
        p.charged = 0.0
        p.result = "Exempt" if p.exempt else ""
        if not p.exempt and p.grade <= CHARGE_MAX_GRADE:
            p.charged = hours * CHARGE_MULTIPLIER
            p.result = "Charged"
    return lst


def compute_outside(hours: float, lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeOutside: in-department people who didn't sign up
    are charged as for no takers; anyone who signed up isn't."""
    for p in lst:
        p.awarded = False
        p.charged = 0.0
        p.result = "Exempt" if p.exempt else ""
        if p.signed_up:
            p.result = "Signed up"
        elif not p.exempt and p.grade <= CHARGE_MAX_GRADE:
            p.charged = hours * CHARGE_MULTIPLIER
            p.result = "Charged"
    return lst


def compute_cancel(lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeCancel: nobody is awarded, nobody is charged."""
    for p in lst:
        p.awarded = False
        p.charged = 0.0
        p.result = ""
    return lst


# ---------------------------------------------------------------- eligibility
def is_night_shift(shift_lbl: str) -> bool:
    return shift_lbl.strip().upper().startswith("N")


def rotation_free_reason(code: str, d: date, shift_lbl: str, observed: dict | None = None) -> str:
    """modCalendar.RotationFreeReason: 'scheduled', 'rest' or ''."""
    return rotation_free_reason3(code, code, code, d, shift_lbl, observed)


def rotation_free_reason3(crew_prev: str, crew_on_d: str, crew_next: str, d: date, shift_lbl: str,
                          observed: dict | None = None) -> str:
    """modCalendar.RotationFreeReason3 (v24): the same rule when the crew on
    d - 1, d and d + 1 can differ (crew history)."""
    night = is_night_shift(shift_lbl)
    on_d = crew_shift_on(crew_on_d, d, observed)
    if on_d in ("Days", "Nights"):
        return "scheduled"
    if on_d != "?":
        if not night and crew_shift_on(crew_prev, d - timedelta(days=1), observed) == "Nights":
            return "rest"
        if night and crew_shift_on(crew_next, d + timedelta(days=1), observed) == "Days":
            return "rest"
    return ""


# ---------------------------------------------------------------- crew history (v24)
def norm_crew(s: str) -> str:
    """modPeople.NormCrew."""
    s = s.strip().replace(" ", "").upper().replace("&", "/")
    return "DAYS" if s == "DAY" else s


def crew_on(hist: list[tuple[str, date]], d: date, fallback: str = "") -> str:
    """modPeople.CrewOnCore for ONE person's rows (crew, effective from), in
    table order: the row with the latest date on or before d (of two rows
    with the same date, the later row); before the first row, the first
    row's crew; no rows, the fallback."""
    le = None
    first = None
    for crew, frm in hist:
        if frm <= d and (le is None or frm >= le[0]):
            le = (frm, norm_crew(crew))
        if first is None or frm < first[0]:
            first = (frm, norm_crew(crew))
    if le is not None:
        return le[1]
    if first is not None:
        return first[1]
    return fallback


LEAVE_CODES = ("LOA", "STD", "PAR", "WSIB", "LTD", "OLV")       # modCalendar.IsLeaveCode


def type_tag(code: str) -> str:
    """modCalendar.TypeTag: what the calendar shows (never STD / PAR / WSIB / LTD outside the Absences sheet)."""
    c = code.upper()
    if c in LEAVE_CODES:
        return "Leave"
    return {"VAC": "Vac", "TRN": "Trn", "SIC": "Sick", "PER": "Personal",
            "SS": "S/S", "PHA": "PHA", "SUP": "Support"}.get(c, "Off")


def public_type(code: str) -> str:
    """modPublish.PublicType: leave of any kind goes out as LV, a sick call or personal day as OFF (v24.3.1)."""
    c = code.strip(" ").upper()
    if c in LEAVE_CODES:
        return "LV"
    if c in ("SIC", "PER"):
        return "OFF"
    return c


def hours_text(x: float) -> str:
    """modPostings.HoursText: CStr(Round(x, 2)) - 8 -> '8', 22.75 -> '22.75', never a trailing point."""
    s = "%.2f" % round(x, 2)
    s = s.rstrip("0").rstrip(".")
    return "0" if s in ("", "-0") else s


def dept_code_value(code: str) -> float:
    """modOvertimeEntry.DeptCodeValue: the numeric part of a grade code (D1 -> 1, P1.5 -> 1.5), 999 when none."""
    num = ""
    found = False
    for ch in code:
        if ch.isdigit() and ch in "0123456789" or ch == ".":
            num += ch
            found = True
        elif found:
            break
    if num == "":
        return 999
    try:
        return float(num)
    except ValueError:                                      # '.' or '1.5.2': not numeric
        return 999


def totals_plain(packed: str) -> str:
    """modPeople.TotalsPlain: '12|21|21' -> 'Total 12, Refused 21, periods 21'."""
    p = packed.split("|")
    if len(p) < 3:
        return packed

    def plain(s: str) -> str:
        try:
            out = hours_text(float(s))
        except ValueError:
            out = s
        return out[:-1] if out.endswith(".") or out.endswith(",") else out

    return f"Total {plain(p[0])}, Refused {plain(p[1])}, periods {plain(p[2])}"


def totals_difference(before: dict[str, str], after: dict[str, str], skip_name: str = "") -> str:
    """modPeople.TotalsDifference: every name whose packed totals differ, or that is in one map only; '' = nothing moved."""
    s = ""
    n = 0
    skip = skip_name.lower()
    for k, v in before.items():
        if k.lower() == skip:
            continue
        if k not in after:
            n += 1
            if n <= 10:
                s += f"  {k}: no longer on the roster\n"
        elif v != after[k]:
            n += 1
            if n <= 10:
                s += f"  {k}: was {totals_plain(v)}, now {totals_plain(after[k])}\n"
    for k in after:
        if k.lower() == skip:
            continue
        if k not in before:
            n += 1
            if n <= 10:
                s += f"  {k}: appeared on the roster\n"
    if n > 10:
        s += f"  ...and {n - 10} more\n"
    return s


def in_bar_list(list_txt: str, item: str) -> bool:
    """modPeople.InBarList: is item one of the '|'-separated entries (trimmed, any case; a blank item never is)."""
    item = item.strip(" ")
    if item == "":
        return False
    return any(part.strip(" ").lower() == item.lower() for part in list_txt.split("|"))


def trim_bar_list(list_txt: str) -> str:
    """modPeople.TrimBarList: 'Phosphine | Cypure |' -> 'Phosphine|Cypure'."""
    return "|".join(p.strip(" ") for p in list_txt.split("|") if p.strip(" ") != "")


def plan_id_in(note: str) -> str:
    """modPostings.PlanIdIn: 'Exempt - moves Feb 1 (PM-0007)' -> 'PM-0007' ('' = no planned-move ID in the note)."""
    a = note.upper().find("(PM-")
    if a < 0:
        return ""
    b = note.find(")", a)
    return note[a + 1:b] if b > a else ""


def name_key(s: str) -> str:
    """ImportKronosData.NameKey: the Kronos / roster name as a match key (no spaces or points, lower case)."""
    return s.strip(" ").replace(" ", "").replace(".", "").lower()


def match_import_name(txt: str, roster: list[str]) -> str:
    """modLeave.MatchImportName: 'Paul M' against roster names -> the one match, '?' for several, '' for none."""
    parts = txt.strip(" ").split(" ")
    first = parts[0].strip(" ").upper()
    ini = parts[1].strip(" ")[:1].upper() if len(parts) >= 2 else ""
    if first == "":
        return ""
    hits, hit = 0, ""
    for nm in roster:
        p = nm.find(",")
        if p >= 0:
            r_last = nm[:p].strip(" ").upper()
            r_first = nm[p + 1:].strip(" ").upper()
        else:
            p = nm.rfind(" ") + 1                                 # InStrRev: 0 when no space
            r_first = nm[:p].strip(" ").upper() if p > 0 else ""
            r_last = nm[p:].strip(" ").upper()
        if r_first[:len(first)] == first and (ini == "" or r_last[:1] == ini):
            hits += 1
            hit = nm
    return hit if hits == 1 else ("?" if hits > 1 else "")


def attend_public_status(entitled: bool, paid_left: int, occurrences: int, wait_until: date | None) -> str:
    """modPublish.AttendPublicStatus: the four things the members' file may say about paid sick time."""
    if not entitled:
        return "No paid sick until " + (wait_until.strftime("%b ") + str(wait_until.day) + wait_until.strftime(", %Y") if wait_until else "Dec 30, 1899")
    if paid_left <= 0:
        return "Paid sick days used up"
    if occurrences >= PAID_SICK_OCC:
        return "Paid sick only while the current occurrence continues"
    return "Paid sick available"


def not_free_reason(code: str, d: date, shift_lbl: str, *, booked_off: str = "",
                    ss_working: bool = False, ss_gave_up: bool = False,
                    ss_rest: bool = False, ss_gave_up_adjoining: bool = False) -> str:
    """modCalendar.NotFreeReason with the table lookups replaced by flags:
    booked_off = absence type code ('VAC', 'LOA' ...) active on d;
    ss_working = working a switch on d; ss_gave_up = gave up a shift on d;
    ss_rest = working a switch on the adjoining shift (nights before a day
    OT, days after a night OT); ss_gave_up_adjoining (v25, D-C3) = gave up
    that adjoining shift, which relieves the rotation's rest rule."""
    night = is_night_shift(shift_lbl)
    if booked_off:
        return type_tag(booked_off)
    why = rotation_free_reason(code, d, shift_lbl)
    if why == "rest" and ss_gave_up_adjoining:
        why = ""
    if why:
        return why
    if ss_working:
        return "S/S"
    if ss_gave_up:
        return "S/S off"
    if ss_rest:
        return "rest"
    return ""


# ---------------------------------------------------------------- v22: sick calls, hours as of a date, leave
def sick_blocks(sick_date: date, sick_night: bool, d: date, night: bool) -> bool:
    """modLeave.SickBlocks: the sick shift and any shift starting within 24 h
    after it ends are blocked. Days 07:00-19:00, nights 19:00-07:00 (+1)."""
    sick_start = sick_date.toordinal() + (19 / 24 if sick_night else 7 / 24)
    sick_end = sick_start + 0.5
    cand_start = d.toordinal() + (19 / 24 if night else 7 / 24)
    return sick_start - 1e-4 <= cand_start < sick_end + 1 - 1e-4


def pick_snapshot(hh: list[tuple[date, float]], d: date, year_from: date):
    """modLeave.PickSnapshot for one employee: (through, hours, estimated) or None.
    hh = [(data_through, worked_hours), ...]. Snapshots before the year start
    are last year's figures: ignored, and if nothing else exists the worked
    figure is 0 at the reset (through None, not estimated)."""
    in_year = [(t, h) for t, h in hh if t >= year_from]
    le = [(t, h) for t, h in in_year if t <= d]
    if le:
        t, h = max(le, key=lambda x: x[0])
        return t, h, False
    if len(in_year) < len(hh) and d < year_from + timedelta(days=14):
        return None, 0.0, False      # first pay period, before any import of the new year: nothing worked yet
    if in_year:
        t, h = min(in_year, key=lambda x: x[0])
        return t, h, True
    if hh:
        return None, 0.0, False
    return None


def hours_as_of(hh: list[tuple[date, float]], lg: list[dict], d: date, year_from: date, year_to: date):
    """modLeave.HoursAsOfCore for one employee. lg rows: dict(date, outcome, status,
    hours, charged, night). Returns (total, through, estimated)."""
    if d < year_from:
        return 0.0, None, False
    snap = pick_snapshot(hh, d, year_from)
    if snap is None:
        through, worked, est = None, 0.0, True
    else:
        through, worked, est = snap
    charged = sum(r["charged"] for r in lg
                  if r["status"].upper() != "VOID" and year_from <= r["date"] <= year_to and r["date"] <= d)
    prov = 0.0
    for r in lg:
        if r["status"].upper() == "VOID" or r["outcome"] not in ("Accepted", "Awarded"):
            continue
        if not (year_from <= r["date"] <= year_to and r["date"] <= d):
            continue
        covered = through is not None and provisional_covered(r["date"], r["night"], through)
        if not covered:
            prov += r["hours"]
    return worked + charged + prov, through, est


def peer_eligible(grade: float, is_returner: bool, leave_overlaps: bool, has_history_before: bool) -> bool:
    """modLeave.PeerEligible."""
    return grade <= PEER_MAX_GRADE and not is_returner and not leave_overlaps and has_history_before


def leave_average(deltas: list[float]) -> float:
    """modLeave.LeaveAverage: mean rounded to 0.1 half away from zero (0 with no peers)."""
    import math
    if not deltas:
        return 0.0
    x = sum(deltas) / len(deltas)
    return math.copysign(math.floor(abs(x) * 10 + 0.5 + 1e-6) / 10, x)


def merge_runs(dates: list[date]) -> list[tuple[date, date]]:
    """modLeave.MergeRuns: sorted unique dates -> consecutive runs."""
    runs: list[list[date]] = []
    for d in sorted(set(dates)):
        if runs and d == runs[-1][1] + timedelta(days=1):
            runs[-1][1] = d
        else:
            runs.append([d, d])
    return [(a, b) for a, b in runs]


# ---------------------------------------------------------------- provisional OT
def refused_counts(period_start: date, through: date) -> bool:
    """modOvertimeEntry.RefusedCounts (v26.5, Graeme Oct 19): a charge counts in Refused and Total Hours once the pay
    period it carries (its Period Start) has started, i.e. period_start <= the current period's start."""
    return period_start <= through


def provisional_covered(shift_date: date, is_night: bool, data_through: date) -> bool:
    """modOvertimeEntry.ProvisionalCovered: the shift has fully ended by the
    end of the data-through date (day: D <= through; night: D + 1 <= through)."""
    ends_on = shift_date + timedelta(days=1) if is_night else shift_date
    return ends_on <= data_through


def provisional_counts(outcome: str, status: str, shift_date: date, is_night: bool,
                       data_through: date, year_from: date, year_to: date) -> bool:
    """modOvertimeEntry.ProvisionalCounts: mirrors the column-C SUMIFS."""
    if status.upper() == "VOID":
        return False
    if outcome not in ("Accepted", "Awarded"):
        return False
    if shift_date < year_from or shift_date > year_to:
        return False
    return not provisional_covered(shift_date, is_night, data_through)


# ---------------------------------------------------------------------------
# v22.1 attendance: sick and personal day counters (modAttendance.AttendanceCore)
# ---------------------------------------------------------------------------
PAID_SICK_DAYS = 5
PAID_SICK_OCC = 3
PERSONAL_DAYS = 5
COMBINED_LIMIT = 5


def add_year(d: date) -> date:
    """VBA DateAdd('yyyy', 1, d): Feb 29 -> Feb 28."""
    try:
        return d.replace(year=d.year + 1)
    except ValueError:
        return d.replace(year=d.year + 1, day=28)


def _ev_pos(d: date, night: bool, kind: int) -> int:
    return d.toordinal() * 4 + (2 if night else 0) + (kind - 1)


def sched_between(crew: str, d1: date, n1: bool, d2: date, n2: bool, observed: dict | None = None) -> bool:
    """True when a scheduled shift of the crew lies strictly between two shifts."""
    p1, p2 = _ev_pos(d1, n1, 1), _ev_pos(d2, n2, 1)
    if p2 <= p1:
        return False
    if (d2 - d1).days > 60:
        return True
    x = d1
    while x <= d2:
        s = crew_shift_on(crew, x, observed)
        if s == "?":
            s = "Days"
        if s:
            p = _ev_pos(x, s == "Nights", 1)
            if p1 < p < p2:
                return True
        x += timedelta(days=1)
    return False


def attendance(crew: str, sick: list[tuple[date, bool]], per: list[date], as_of: date,
               observed: dict | None = None) -> dict:
    """modAttendance.AttendanceCore. sick = [(date, night), ...]; per = one date per
    personal day (scheduled shift). Returns the status fields as a dict."""
    events = [(d, n, 1) for d, n in sick if d <= as_of] + [(d, False, 2) for d in per if d <= as_of]
    events.sort(key=lambda e: _ev_pos(*e))
    y0 = min([as_of.year] + [e[0].year for e in events])
    s = dict(next_jan1=date(y0 + 1, 1, 1), window=date(y0, 1, 1),
             waiting=False, wait_until=None, crossed=None, reset_on=None,
             used=0, paid=0, occ=0, occ1=None, occ2=None, occ_txt=[], per=0,
             have_last=False, last_d=None, last_n=False, occs=[])
    out = dict(last_sick=None, last_paid=False, last_reason="")

    def reset():
        s.update(used=0, paid=0, occ=0, occ1=None, occ2=None, occ_txt=[], per=0, have_last=False, occs=[])

    def advance_to(d: date):
        while True:
            nxt, kind = None, 0
            if s["next_jan1"] <= d:
                nxt, kind = s["next_jan1"], 1
            if s["waiting"] and s["wait_until"] <= d and (kind == 0 or s["wait_until"] < nxt):
                nxt, kind = s["wait_until"], 2
            if kind == 0:
                return
            reset()
            s["window"] = nxt
            if kind == 1:
                s["next_jan1"] = date(nxt.year + 1, 1, 1)
                s["reset_on"] = None
            else:
                s["waiting"] = False
                s["reset_on"] = s["wait_until"]
                s["wait_until"] = None
                s["crossed"] = None

    def start_wait(d: date):
        s["waiting"] = True
        s["crossed"] = d
        if s["occ"] >= 2:
            s["wait_until"] = add_year(s["occ2"])
        elif s["occ"] >= 1:
            s["wait_until"] = add_year(s["occ1"])
        else:
            s["wait_until"] = add_year(d)

    for d, night, kind in events:
        if d < s["window"]:
            continue
        advance_to(d)
        if kind == 1:
            s["used"] += 1
            same = s["have_last"] and not sched_between(crew, s["last_d"], s["last_n"], d, night, observed)
            if not same:
                s["occ"] += 1
                if s["occ"] == 1:
                    s["occ1"] = d
                if s["occ"] == 2:
                    s["occ2"] = d
                s["occ_txt"].append(d)
                s["occs"].append([d, d, 0, 0])              # v24.1: first day, last day, days, paid days
            s["occs"][-1][1] = d
            s["occs"][-1][2] += 1
            s.update(have_last=True, last_d=d, last_n=night)
            if s["waiting"]:
                paid, why = False, "unpaid - not entitled until %s" % s["wait_until"]
            elif s["occ"] > PAID_SICK_OCC:
                paid, why = False, "unpaid - occurrence %d" % s["occ"]
            elif s["used"] > PAID_SICK_DAYS:
                paid, why = False, "unpaid - sick day %d" % s["used"]
            else:
                paid, why = True, "paid sick day %d of %d, occurrence %d of %d" % (s["used"], PAID_SICK_DAYS, s["occ"], PAID_SICK_OCC)
            if not s["waiting"] and s["used"] + s["per"] > COMBINED_LIMIT:
                start_wait(d)
                paid, why = False, "unpaid - sick + personal day %d passes %d" % (s["used"] + s["per"], COMBINED_LIMIT)
            if paid:
                s["paid"] += 1
                s["occs"][-1][3] += 1
            out.update(last_sick=d, last_paid=paid, last_reason=why)
        else:
            s["per"] += 1
            if not s["waiting"] and s["used"] + s["per"] > COMBINED_LIMIT:
                start_wait(d)
    advance_to(as_of)
    paid_left = max(0, PAID_SICK_DAYS - s["used"])      # days still payable inside the current occurrences
    # v24.3.1: 0 once a 4th occurrence has started, or once the 3rd has ended (a scheduled shift
    # after its last sick shift and before the night of as_of: a day shift on as_of counts, the
    # night on as_of does not)
    if s["occ"] > PAID_SICK_OCC:
        paid_left = 0
    elif s["occ"] == PAID_SICK_OCC and s["have_last"] and \
            sched_between(crew, s["last_d"], s["last_n"], as_of, True, observed):
        paid_left = 0
    if s["waiting"]:
        paid_left = 0
    out.update(window_start=s["window"], sick_used=s["used"], paid_used=s["paid"], paid_left=paid_left,
               occurrences=s["occ"], occ_dates=list(s["occ_txt"]), personal_used=s["per"],
               personal_left=max(0, PERSONAL_DAYS - s["per"]), combined=s["used"] + s["per"],
               entitled=not s["waiting"], wait_until=s["wait_until"], crossed_on=s["crossed"],
               reset_on=s["reset_on"], occs=[tuple(o) for o in s["occs"]])
    return out



# ---------------------------------------------------------------------------
# v23 year end: the roster year, ID marks, integrity helpers, archive split
# ---------------------------------------------------------------------------
def first_period_start(anchor: date, y: int) -> date:
    """modOvertimeEntry.FirstPeriodStart: the period containing Jan 1 of y, on the 14-day grid through anchor."""
    return anchor + timedelta(days=14 * ((date(y, 1, 1) - anchor).days // 14))


def roster_year_of(first_start: date) -> int:
    return (first_start + timedelta(days=13)).year


def periods_in_year(anchor: date, y: int) -> int:
    return (first_period_start(anchor, y + 1) - first_period_start(anchor, y)).days // 14


def next_number(max_in_log: int, mark: int) -> int:
    """modOvertimeEntry.NextNumber: IDs carry on from the higher of the log and the mark."""
    return max(max_in_log, mark) + 1


def period_cell_code(has_formula: bool, formula_text: str, v) -> str:
    """modIntegrity.PeriodCellCode. v: None = empty, ('err',) = error value, str, or number."""
    if has_formula:
        return "" if "tblotlog" in formula_text.lower() else "notlog"
    if isinstance(v, tuple):
        return "error"
    if v is None:
        return "empty"
    if isinstance(v, str):
        return "textformula" if v.startswith("=") else "text"
    return "typed"


def duplicates_in(keys: list[str]) -> str:
    """modIntegrity.DuplicatesIn: 'S-0004 x2; P-0010 x3' in first-seen order, '' when none."""
    seen: dict[str, list] = {}
    for k in keys:
        k = k.strip()
        if not k:
            continue
        low = k.lower()
        if low in seen:
            seen[low][1] += 1
        else:
            seen[low] = [k, 1]
    return "; ".join(f"{k} x{n}" for k, n in seen.values() if n > 1)


def goes_to_archive(period_start, cut: date) -> bool:
    """modRollover.GoesToArchive. period_start: None (empty), str (text), date, or an Excel serial number."""
    if period_start is None or isinstance(period_start, str):
        return False
    if isinstance(period_start, date):
        return period_start < cut
    serial = int(period_start // 1)
    if serial <= 0:
        return False
    return date(1899, 12, 30) + timedelta(days=serial) < cut


def is_dry_run_name(file_name: str) -> bool:
    u = file_name.upper()
    return any(t in u for t in ("COPY", "DRY RUN", "DRYRUN", "DRY-RUN"))


BACKUP_BASENAME = "Call Out Roster"


def backup_file_date(file_name: str) -> date | None:
    """modSystem.BackupFileDate: the date in 'Call Out Roster yyyy-mm-dd.xlsm', else None."""
    pat_len = len(BACKUP_BASENAME) + 1 + 10 + 5
    if len(file_name) != pat_len or not file_name.startswith(BACKUP_BASENAME + " ") or not file_name.endswith(".xlsm"):
        return None
    s = file_name[len(BACKUP_BASENAME) + 1:len(BACKUP_BASENAME) + 11]
    digits = "0123456789"                                   # VBA's "#" matches 0-9 only, not every Unicode digit
    if not (all(c in digits for c in s[0:4] + s[5:7] + s[8:10]) and s[4] == "-" and s[7] == "-"):
        return None
    return date_serial(int(s[0:4]), int(s[5:7]), int(s[8:10]))


def date_serial(y: int, m: int, d: int) -> date:
    """VBA DateSerial: a month or day out of range rolls over (month 13 = January next year, day 0 = the day before the 1st)."""
    first = date(y + (m - 1) // 12, (m - 1) % 12 + 1, 1)
    return first + timedelta(days=d - 1)


def pre_change_copy_date(file_name: str, roster_name: str) -> date | None:
    """modSystem.PreChangeCopyDate (v26): 'Backup before <what> yyyy-mm-dd hhmm[ss] - <roster>' -> its date, else None."""
    head = "Backup before "
    tail = " - " + roster_name
    if not roster_name or len(file_name) <= len(head) + len(tail):
        return None
    if file_name[:len(head)].upper() != head.upper() or file_name[-len(tail):].upper() != tail.upper():
        return None
    core = file_name[:len(file_name) - len(tail)]
    p = core.rfind(" ")
    if p < 0:
        return None
    tm = core[p + 1:]
    digits = "0123456789"
    if len(tm) not in (4, 6) or not all(c in digits for c in tm):
        return None
    core = core[:p]
    p = core.rfind(" ")
    if p < len(head):                                       # VBA: p <= Len(HEAD) with p 1-based
        return None
    ds = core[p + 1:]
    if len(ds) != 10 or ds[4] != "-" or ds[7] != "-" or not all(c in digits for c in ds[0:4] + ds[5:7] + ds[8:10]):
        return None
    return date_serial(int(ds[0:4]), int(ds[5:7]), int(ds[8:10]))


BACKUP_KEEP_DAYS = 30            # modOvertimeEntry (v27 item 12 adds the two below)
BACKUP_HOURLY_KEEP_DAYS = 7
BACKUP_PERIOD_KEEP_DAYS = 400


def backup_keep(kind: str, file_date: date, today: date, period_start: bool) -> bool:
    """modSystem.BackupKeepCore: a daily backup kept BACKUP_KEEP_DAYS (BACKUP_PERIOD_KEEP_DAYS when it is the first daily
    backup of its pay period - the caller decides); an hourly backup BACKUP_HOURLY_KEEP_DAYS. Plain dates on both sides
    (VBA takes Int of both; pass a date, not a datetime)."""
    age = (today - file_date).days
    if kind.lower() == "hourly":
        return age <= BACKUP_HOURLY_KEEP_DAYS
    if period_start:
        return age <= BACKUP_PERIOD_KEEP_DAYS
    return age <= BACKUP_KEEP_DAYS


def hourly_backup_date(file_name: str) -> datetime | None:
    """modSystem.HourlyBackupDate: '<basename> backup yyyy-mm-dd hh.xlsm' -> that day and hour, else None."""
    head = BACKUP_BASENAME + " backup "
    if len(file_name) != len(head) + 13 + 5 or not file_name.startswith(head) or not file_name.endswith(".xlsm"):
        return None
    s = file_name[len(head):len(head) + 13]
    digits = "0123456789"
    if not (all(c in digits for c in s[0:4] + s[5:7] + s[8:10] + s[11:13]) and s[4] == "-" and s[7] == "-" and s[10] == " "):
        return None
    y, mo, dy, hh = int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13])
    if not (1 <= mo <= 12 and 1 <= dy <= 31 and hh <= 23):
        return None
    try:
        return datetime(y, mo, dy, hh)
    except ValueError:
        return None


def is_backup_copy_name(file_name: str) -> bool:
    """modSystem.IsBackupCopyName (v24.3.1): a daily backup, the FINAL, a 'Backup before ...' copy."""
    if backup_file_date(file_name) is not None:
        return True
    u = file_name.upper()
    return "FINAL" in u or "BACKUP" in u


def snapshot_through_for_year(through: date, snap_yr: int, cur_yr: int, yr_from: date) -> date:
    """modLeave.SnapshotThroughForYear."""
    if snap_yr <= 0 or snap_yr >= cur_yr:
        return through
    return yr_from - timedelta(days=1) if through >= yr_from else through



# ---------------------------------------------------------------- people changes (v24)
def new_hire_average(peers: list[tuple[float, bool, float]]) -> float:
    """modPeople.NewHireAverage. peers = (grade, on leave on the start date,
    Total Hours on the start date): the average over grade 2 and below, not
    on leave, rounded like the leave top-up. 0 with no peers."""
    return leave_average([tot for grade, on_leave, tot in peers if peer_eligible(grade, False, on_leave, True)])


def highest_peer_total(peers: list[tuple[float, bool, float]]):
    """modPeople.HighestPeerTotal. peers = (grade, is the mover, Total Hours):
    the highest among grade 2 and below who aren't the mover; None if nobody."""
    vals = [tot for grade, mover, tot in peers if grade <= PEER_MAX_GRADE and not mover]
    return max(vals) if vals else None


def unit_move_top_up(own: float, highest_peer) -> float:
    """modPeople.UnitMoveTopUp: up to the target, never lowered (2 decimals)."""
    if highest_peer is None or highest_peer <= own:
        return 0.0
    return math.floor((highest_peer - own) * 100 + 0.5 + 0.000001) / 100


def planned_code(plans: list[tuple[str, str, str, str, date, str]], emp: str, unit: str, d: date,
                 roster_code: str) -> str:
    """modPeople.PlannedCodeCore. plans = (employee, new unit, grade code,
    'unit|unit' to clear, effective from, status), in table order. Only
    Planned rows for emp dated on or before d count (of two, the later date;
    of two with the same date, the later row). The new unit gets the planned
    code, a cleared unit gets '', any other unit keeps the roster's code."""
    best = None
    for who, new_unit, code, clear, frm, status in plans:
        if status.strip().lower() != "planned" or who.strip().lower() != emp.strip().lower():
            continue
        if frm <= d and (best is None or frm >= best[3]):
            best = (new_unit, code, clear, frm)
    if best is None:
        return roster_code
    new_unit, code, clear, _ = best
    if new_unit.strip().lower() == unit.strip().lower():
        return code.strip()
    if unit.strip().lower() in [c.strip().lower() for c in clear.split("|") if c.strip()]:
        return ""
    return roster_code


def planned_list_hours(own: float, peers: list[tuple[float, bool, float]]) -> float:
    """modPeople.PlannedListHours: the Total a planned mover is listed at in
    the new unit before the top-up. peers = (grade, is a mover, Total Hours)
    for everyone holding a code in the unit on the date: the highest among
    grade 2 and below who aren't movers, never under their own."""
    top = highest_peer_total(peers)
    if top is None or top <= own:
        return own
    return top


def rename_in_list(txt: str, old: str, new: str) -> str:
    """modPeople.RenameInList: one full name inside 'A; B (out of dept)'."""
    if not txt.strip():
        return txt
    out = []
    for part in txt.split(";"):
        p = part.strip()
        tail = ""
        if p.lower().endswith(" (out of dept)"):
            tail = p[-14:]
            p = p[:-14].strip()
        if p.casefold() == old.casefold():
            p = new
        out.append(p + tail)
    return "; ".join(out)


def replace_name_token(txt: str, old: str, new: str) -> str:
    """modPeople.ReplaceNameToken: a short calendar name replaced only where
    it stands as a whole name."""
    if not old:
        return txt
    start = 0
    low_old = old.lower()
    while True:
        p = txt.lower().find(low_old, start)
        if p < 0:
            break
        before = txt[p - 1] if p > 0 else ""
        after = txt[p + len(old)] if p + len(old) < len(txt) else ""
        if before in ("", " ", "/", ";") and after in ("", " ", "/", ";", "("):
            txt = txt[:p] + new + txt[p + len(old):]
            start = p + len(new)
        else:
            start = p + 1
    return txt


# ---------------------------------------------------------------- refused hours by kind (v24.1)
def is_plain_prefix(s: str) -> bool:
    """modOvertimeEntry.IsPlainPrefix (v26, was modRollover): letters A-Z only, 1 to 8 of them."""
    return 1 <= len(s) <= 8 and all("A" <= ch <= "Z" for ch in s)


def ref_prefix_max_text(ref_list: str) -> str:
    """modOvertimeEntry.RefPrefixMaxText (v26, I8): Ref IDs joined by | -> 'LV=3|P=17|S=431', '' when none parses."""
    d: dict[str, int] = {}
    for raw in ref_list.split("|"):
        s = raw.strip().upper()
        p = s.rfind("-")
        if p < 1 or p >= len(s) - 1:
            continue
        tail = s[p + 1:]
        if not tail.isdigit():                              # VBA IsNumeric is wider, but the IDs are digits
            continue
        pre = s[:p]
        if not is_plain_prefix(pre):
            continue
        k = int(tail)
        if k > d.get(pre, -1):
            d[pre] = k
    return "|".join(f"{k}={d[k]}" for k in sorted(d))


def ordinal_text(n: int) -> str:
    """modReports.OrdinalText (v26): 1st, 2nd, 3rd, 11th-13th, 21st ..."""
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def charge_why(typ: str, outcome: str, call_order: int, hours: float, charged: float, code: str) -> str:
    """modReports.ChargeWhy (v26): 'why was I charged' from an OT Log row's own fields."""
    t = typ.strip().lower()
    o = outcome.strip().lower()
    oc = outcome.strip()
    grade_txt = "no code" if code.strip() == "" else f"grade {code.strip()}"
    above_txt = "no code" if code.strip() == "" else f"{grade_txt} is above {hours_text(CHARGE_MAX_GRADE)}"
    mult = hours_text(CHARGE_MULTIPLIER)
    if t == "callout":
        pos = f"{ordinal_text(call_order)} on the call list, " if call_order > 0 else ""
        if o == "exception":
            return f"Not charged: {pos}plant upset (exception)"
        if charged != 0:
            return f"Charged {hours_text(charged)} h: {pos}{oc}, {hours_text(hours)} h offered x {mult} ({grade_txt})"
        if o == "accepted":
            return f"Not charged: {pos}accepted {hours_text(hours)} h"
        if o in ("refused", "left message") and dept_code_value(code) > CHARGE_MAX_GRADE:
            return f"Not charged: {pos}{oc}, {above_txt}"
        return f"Not charged: {pos}{oc or 'no outcome'}"
    if t == "posting":
        pos = f"#{call_order} in call order, " if call_order > 0 else ""
        if charged != 0:
            return f"Charged {hours_text(charged)} h: {pos}didn't sign up, posting hours {hours_text(hours)} x {mult} ({grade_txt})"
        if o == "awarded":
            return f"Not charged: {pos}awarded {hours_text(hours)} h"
        if o == "signed up":
            return f"Not charged: {pos}signed up, not awarded"
        if o == "refused" and dept_code_value(code) > CHARGE_MAX_GRADE:
            return f"Not charged: {pos}didn't sign up, {above_txt}"
        return f"Not charged: {pos}{oc or 'no outcome'}"
    head = typ.strip() or "Entry"
    sign = "+" if charged >= 0 else ""
    return f"{head}: {sign}{hours_text(charged)} h" + (f" ({oc})" if oc else "")


def staff_short(scheduled: int, off: int, covered: int, minimum: int, part: bool = False) -> str:
    """modMinStaff.StaffShort (v26): '' or 'SHORT staffed/minimum'; staffed = scheduled - off + covered; no minimum = never
    short. v27.5 (item 3): ' (part)' added when a partial cover's fraction was dropped from covered."""
    if minimum <= 0:
        return ""
    staffed = scheduled - off + covered
    return f"SHORT {staffed}/{minimum}" + (" (part)" if part else "") if staffed < minimum else ""


FATIGUE_MAX_HOURS = 16.0   # modOvertimeEntry (v26.5.5): the plant's maximum consecutive hours
FATIGUE_MIN_REST = 11.0    # ... and the rest before the next shift


def fatigue_max_hours(worked_before: float, hours_to_next: float) -> float:
    """modCalendar.FatigueMaxHours (v26.5.5, item 9): the most OT that fits, given the consecutive hours of work
    ending when the OT starts and the hours from the OT's start to the next work's start (99 = none)."""
    m = FATIGUE_MAX_HOURS - worked_before
    m = min(m, hours_to_next - FATIGUE_MIN_REST)
    return max(m, 0.0)


def fatigue_problem(worked_before: float, hours_offered: float, hours_to_next: float) -> str:
    """modCalendar.FatigueProblem: '' or '16 h rule' when the hours offered do not fit."""
    return "16 h rule" if hours_offered > fatigue_max_hours(worked_before, hours_to_next) + 0.001 else ""


def fatigue_block_core(h_prev_prev: float, h_prev: float, h_next1: float, h_next2: float) -> tuple[float, float]:
    """modCalendar.FatigueBlockCore: the consecutive block ending when the offer starts (the touching shift plus the
    one before it whenever both were worked) and the hours to the next work (12 h steps; 99 = none)."""
    worked_before = h_prev + h_prev_prev if h_prev > 0 else 0.0
    hours_to_next = USUAL_SHIFT_HOURS if h_next1 > 0 else 2 * USUAL_SHIFT_HOURS if h_next2 > 0 else 99.0
    return worked_before, hours_to_next


# ---- v27.1: the shift-switch hard stops as pure words (modShiftSwitch SwitchStopsCore; suite switchstop) ----
SF_KEYS = ["bkoff", "gaveoff", "workoff", "bkworks", "workworks", "fat",
           "p_bkworks", "p_gaveworks", "p_workworks", "p_bkoff", "p_workoff", "p_fat",
           "s_bkoff", "s_bkworks", "s_bkoff2", "s_bkworks1", "s_offon", "s_workson",
           "s_fatworks", "s_fatoff"]                   # = the SF_* positions 0..19 (v27.5: the two same-day 16 h facts)
SF_BOOKED = {"bkoff", "bkworks", "p_bkworks", "p_bkoff", "s_bkoff", "s_bkworks", "s_bkoff2", "s_bkworks1"}
SF_ON = {"s_offon", "s_workson"}


def switch_facts(spec: str) -> list[str]:
    """modSelfTest.SwitchFacts: 'key=value;...' -> the 20 facts (a booked fact '+type', an 'on' fact '+')."""
    f = [""] * len(SF_KEYS)
    for part in spec.split(";") if spec else []:
        k, _, v = part.partition("=")
        i = SF_KEYS.index(k)                        # an unknown key fails the case
        f[i] = "+" + v if k in SF_BOOKED else "+" if k in SF_ON else v
    return f


def _num(x: float) -> str:
    return str(int(x)) if x == int(x) else str(x)


def eligibility_text(off: str, works: str, d_txt: str, bk_off: str, gave_off: str, work_off: str,
                     bk_works: str, work_works: str) -> str:
    """modShiftSwitch.EligibilityText: booked off / already on a switch, for the one giving up and the one working."""
    s = ""
    if bk_off:
        s += f"- {off} is booked off ({bk_off[1:]}) on {d_txt}\n"
    if gave_off:
        s += f"- {off} already switched off that day ({gave_off})\n"
    if work_off:
        s += f"- {off} is already working a switch that day ({work_off})\n"
    if bk_works:
        s += f"- {works} is booked off ({bk_works[1:]}) on {d_txt}\n"
    if work_works:
        s += f"- {works} is already working a switch that day ({work_works})\n"
    return s


def fatigue_stop_text(works: str, d_txt: str, shift: str, reason: str) -> str:
    """modShiftSwitch.FatigueStopText: the 16 h stop from NotFreeReason's answer ('' = fine)."""
    if shift == "":
        return ""
    if reason == "16 h rule":                   # v27.5 (item 1): a scheduled shift beside it counts too
        return (f"- {works} can't work {d_txt} {shift.lower()} under the 16 h rule (with the shifts beside it - "
                f"scheduled / OT / a switch - it would run past {_num(FATIGUE_MAX_HOURS)} h or leave under "
                f"{_num(FATIGUE_MIN_REST)} h rest)\n")
    if reason == "on OT":
        return f"- {works} already has OT accepted on {d_txt} {shift.lower()}\n"
    return ""


def switch_stops_core(off: str, works: str, d_txt: str, shift: str, two_way: bool, same_day: bool,
                      d2_txt: str, sh2: str, f: list[str]) -> str:
    """modShiftSwitch.SwitchStopsCore: one-way = eligibility + the 16 h stop for the one working it; two-way on another
    day = the same for the payback, roles swapped; a same-day sister swap = bookings and switches that day, then (v27.5)
    only "on OT" from the two 16 h facts (the swap of both shifts is exempt from the 16 h arithmetic)."""
    stops = eligibility_text(off, works, d_txt, *f[0:5]) + fatigue_stop_text(works, d_txt, shift, f[5])
    if two_way:
        if same_day:
            stops = ""
            if f[12]:
                stops += f"- {off} is booked off ({f[12][1:]})\n"
            if f[13]:
                stops += f"- {works} is booked off ({f[13][1:]})\n"
            if f[14]:
                stops += f"- {off} can't work the {sh2.lower()} shift ({f[14][1:]})\n"
            if f[15]:
                stops += f"- {works} can't work the {shift.lower()} shift ({f[15][1:]})\n"
            if f[16]:
                stops += f"- {off} is already on a switch that day\n"
            if f[17]:
                stops += f"- {works} is already on a switch that day\n"
            # v27.5 (Graeme Oct 25 after review): the swap of both shifts is exempt from the 16 h arithmetic; only "on OT" stops
            if f[18] == "on OT":
                stops += fatigue_stop_text(works, d_txt, shift, f[18])
            if f[19] == "on OT":
                stops += fatigue_stop_text(off, d_txt, sh2, f[19])
        else:
            stops += eligibility_text(works, off, d2_txt, *f[6:11]) + fatigue_stop_text(off, d2_txt, sh2, f[11])
    return stops


def fatigue_shift_reason_fixture(code: str, d: date, shift_lbl: str, pending_lbl: str) -> str:
    """modCalendar.FatigueShiftReason on the rotation alone, as SelfTest's CaseFatShift runs it (v27.5, item 1 review): no OT,
    no bookings, no saved switch; pending_lbl = the shift on d this person gives up in the switch being saved ('' = none;
    SetPendingGiveUp -> WorkedHoursOn 0). '16 h rule' / '' (fits)."""
    if shift_lbl == "":
        return ""
    one = timedelta(days=1)

    def worked(dd: date, is_night: bool) -> float:
        lbl = "Nights" if is_night else "Days"
        if dd == d and lbl == pending_lbl:
            return 0.0
        return default_shift_hours(dd, is_night) if crew_shift_on(code, dd) == lbl else 0.0

    night = is_night_shift(shift_lbl)
    if night:                                     # FatigueInputs: a night offer - that date's day shift, then the night before it
        hp = worked(d, False)
        hpp = worked(d - one, True) if hp > 0 else 0.0
        hn1 = worked(d + one, False)
        hn2 = worked(d + one, True) if hn1 == 0 else 0.0
    else:                                         # a day offer - the night before, then that night's day shift
        hp = worked(d - one, True)
        hpp = worked(d - one, False) if hp > 0 else 0.0
        hn1 = worked(d, True)
        hn2 = worked(d + one, False) if hn1 == 0 else 0.0
    wb, htn = fatigue_block_core(hpp, hp, hn1, hn2)
    return "16 h rule" if fatigue_max_hours(wb, htn) < default_shift_hours(d, night) - 0.001 else ""


def capped_posting_hours(post_hours: float, max_hours: float, max_set: bool) -> float:
    """modPostings.CappedPostingHours: the hours a listed person is charged on - the posting's, or their 16 h rule
    maximum when the award pop-up worked one out and it is lower (0 = they could take none: not charged)."""
    if max_set and max_hours < post_hours:
        return max_hours if max_hours > 0 else 0.0
    return post_hours


def practice_copy_date(file_name: str) -> datetime | None:
    """modSystem.PracticeCopyDate (v26.5.5, 5.4): '<roster> COPY yyyy-mm-dd hhmm.xlsm' -> that time, else None."""
    p = file_name.lower().find(" copy ")
    if p < 0:
        return None
    s = file_name[p + 6:]
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{4}\.xlsm", s):
        return None
    try:
        return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]), int(s[13:15]))
    except ValueError:
        return None


def next_step_text(n_accepted: int, partial_hours: float, n_left_msg: int, is_admin: bool, admin_name: str) -> str:
    """modOvertimeEntry.NextStepText (v26.5.5, B13): the last line of a submitted call-out session."""
    if n_accepted == 0:
        return ("The line stays OPEN on the calendar. Next: "
                + ("call back later (Fill a shift now shows ? beside them), " if n_left_msg > 0 else "")
                + "post the shift, or "
                + ("award out of department." if is_admin else f"ask {admin_name} to award out of department."))
    if partial_hours > 0:
        return f"The shift is closed at {hours_text(partial_hours)} h; the rest is dropped, nobody below is charged."
    return "The line is covered on the calendar."


def owed_status_text(days: int, warn_days: int) -> str:
    """modShiftSwitch.OwedStatusText (v26): the Needs Attention status for a payback owed `days` days."""
    if days <= warn_days:
        return "Payback owed"
    return f"Payback owed {days} {'day' if days == 1 else 'days'}"


def refused_bucket(typ: str) -> str:
    """modReports.RefusedBucket."""
    t = typ.strip().lower()
    return t if t in ("callout", "posting") else "other"


def refused_split(rows: list[tuple[str, float, bool, bool]]) -> tuple[float, float, float]:
    """modReports.SplitTotals for ONE person. rows = (type, hours charged, void,
    period inside the roster year). Returns (call-out, posting, other)."""
    out = {"callout": 0.0, "posting": 0.0, "other": 0.0}
    for typ, hrs, void, in_year in rows:
        if void or not in_year or hrs == 0:
            continue
        out[refused_bucket(typ)] += hrs
    return out["callout"], out["posting"], out["other"]


# ---------------------------------------------------------------- part of a call-out accepted (v24.1)
def callout_offers(hours: float, calls: list[tuple[str, float]]) -> list[float]:
    """modOvertimeEntry.CalloutOffers. calls = (outcome, hours taken; 0 = all on
    offer) in call order. A PARTIAL acceptance lowers the offer for everyone
    below; a whole-shift acceptance changes nothing. v25 (D-E4): a partial
    acceptance CLOSES the shift - nobody below is offered the rest."""
    remain = hours
    out = []
    for outcome, taken in calls:
        out.append(remain)
        if outcome == "Accepted" and 0 < taken < remain:
            remain = 0
    return out


def accepted_hours(offer: float, taken: float) -> float:
    """modOvertimeEntry.AcceptedHours."""
    return taken if 0 < taken < offer else offer


# ---------------------------------------------------------------- electronic sign-ups (v24.1)
SIGNUP_STAMP_SLACK_MIN = 10          # modSignups.SIGNUP_STAMP_SLACK_MIN (v27, item 14)


def signup_suspect(stamp: datetime | None, file_time: datetime | None, clock_change_between: bool = False) -> bool:
    """modSignups.SignupSuspect: the file was written more than the slack after the stamp it claims, in whole
    seconds; a clock change between the two makes a difference within the slack of an hour not suspect."""
    if stamp is None or file_time is None:
        return False
    secs = round((file_time - stamp).total_seconds())
    if clock_change_between and abs(secs - 3600) <= SIGNUP_STAMP_SLACK_MIN * 60:
        return False
    return secs > SIGNUP_STAMP_SLACK_MIN * 60


def clock_change_between(d1: datetime, d2: datetime) -> bool:
    """modSignups.ClockChangeBetween: a clock change (2 AM the morning after a change night) lies between the two times."""
    a, b = (d1, d2) if d1 <= d2 else (d2, d1)
    if (b - a).days > 366:
        b = a + timedelta(days=366)
    k = a.date() - timedelta(days=1)
    while k <= b.date():
        if dst_night_hours(k) != 12:
            m = datetime(k.year, k.month, k.day) + timedelta(days=1, hours=2)
            if a < m <= b:
                return True
        k += timedelta(days=1)
    return False


def reduce_signups(rows: list[tuple], closes_at: datetime | None):
    """modSignups.ReduceSignups. rows = (employee, action, hours, stamp[, login[, file_time]]) in any
    order. Each person's action with the latest stamp wins (same stamp: the
    later row). Returns [(employee, hours, late, suspect, login_mix)] for people whose newest action
    is a sign-up, in the order each name first appeared. v27 (item 14): suspect = the winning file's time
    on disk is more than the slack after its stamp; login_mix = another file for that name came from a
    different (non-empty) login."""
    newest: dict[str, int] = {}
    for i, row in enumerate(rows):
        key = row[0].strip().casefold()
        if not key:
            continue
        if key not in newest:
            newest[key] = i
        elif row[3] >= rows[newest[key]][3]:
            newest[key] = i
    out = []
    for key, i in newest.items():
        row = rows[i]
        emp, action, hours, stamp = row[0], row[1], row[2], row[3]
        login = row[4].strip() if len(row) > 4 else ""
        file_time = row[5] if len(row) > 5 else None
        if action.strip().lower() != "signup":
            continue
        mix = bool(login) and any(
            j != i and r[0].strip().casefold() == key and len(r) > 4 and r[4].strip()
            and r[4].strip().casefold() != login.casefold()
            for j, r in enumerate(rows))
        out.append((emp.strip(), hours, closes_at is not None and stamp > closes_at,
                    signup_suspect(stamp, file_time), mix))
    return out


def withdrawn_signups(rows: list[tuple[str, str, float, datetime, str]]):
    """modSignups.WithdrawnSignups (v26, M3): rows = (employee, action, hours, stamp, login). The people whose
    newest action is a withdraw, in first-appearance order: [(employee, login, stamp, mismatch)], mismatch = a
    sign-up at or before the withdraw came from a different login."""
    newest: dict[str, int] = {}
    for i, row in enumerate(rows):
        key = row[0].strip().casefold()
        if not key:
            continue
        if key not in newest:
            newest[key] = i
        elif row[3] >= rows[newest[key]][3]:
            newest[key] = i
    out = []
    for key, i in newest.items():
        emp, action, _hours, stamp, login = rows[i]
        if action.strip().lower() != "withdraw":
            continue
        wd_login = login.strip()
        mism = any(r[0].strip().casefold() == key and j != i and r[1].strip().lower() == "signup"
                   and r[3] <= stamp and r[4].strip().casefold() != wd_login.casefold()
                   for j, r in enumerate(rows))
        out.append((emp.strip(), wd_login, stamp, mism))
    return out


# ---------------------------------------------------------------------------
# v24.2: the publish's privacy check (modPublish) and the Payroll # rules
# (modOvertimeEntry)
# ---------------------------------------------------------------------------
def _single_run(s: str) -> int:
    best = run = 0
    for ch in s:
        if ch.isdigit() and ch.isascii():
            run += 1
            best = max(best, run)
        else:
            run = 0
    return best


def _joined_digits(s: str) -> int:
    """modPublish.JoinedDigits: digits counted across phone separators, yyyy-mm-dd blanked first."""
    chars = list(s)
    i = 0
    while i <= len(chars) - 10:
        w = chars[i:i + 10]
        if all((c.isdigit() and c.isascii()) if k not in (4, 7) else c == "-" for k, c in enumerate(w)):
            chars[i:i + 10] = [" "] * 10
            i += 10
        else:
            i += 1
    best = run = 0
    for ch in chars:
        if ch.isdigit() and ch.isascii():
            run += 1
            best = max(best, run)
        elif ch in " -().":
            pass
        else:
            run = 0
    return best


def _dashed_seven(s: str) -> bool:
    """modPublish.DashedSeven (v26, M4): '###-####' with no digit directly before or after it."""
    digits = "0123456789"
    for p in range(0, len(s) - 7):
        seg = s[p:p + 8]
        if seg[3] == "-" and all(c in digits for c in seg[0:3] + seg[4:8]):
            before = s[p - 1] if p > 0 else ""
            after = s[p + 8] if p + 8 < len(s) else ""
            if (before == "" or before not in digits) and (after == "" or after not in digits):
                return True
    return False


def longest_digit_run(s: str) -> int:
    joined = _joined_digits(s)
    run = joined if joined >= 10 else _single_run(s)
    if run < 7 and _dashed_seven(s):                        # v26 (M4): "555-0142" counts as 7
        run = 7
    return run


def private_value_kind(x) -> str:
    """modPublish.PrivateValueKind: '' / email / phone / number."""
    if isinstance(x, str):
        if "@" in x:
            return "email"
        if longest_digit_run(x) >= 7:
            return "phone"
        return ""
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return "number" if abs(float(x)) >= 1_000_000 else ""
    return ""                                   # dates, blanks, errors


def attend_count_ok(x) -> bool:
    """A sick-time count may be a number or blank - never text and never a date."""
    return not isinstance(x, (str, date))


_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def attend_status_ok(s: str) -> bool:
    if s in ("Paid sick available", "Paid sick days used up", "Paid sick only while the current occurrence continues"):
        return True
    head = "No paid sick until "
    if not s.startswith(head):
        return False
    rest = s[len(head):]                        # Like "[A-Z][a-z][a-z] *, ####"
    if len(rest) < 4 + 2 + 4 or not (rest[0].isupper() and rest[1:3].islower() and rest[3] == " "):
        return False
    return rest[-6:-4] == ", " and rest[-4:].isdigit()


LEAVE_WORDS = ("STD", "LOA", "WSIB", "LTD", "OLV", "PAR", "PARENTAL", "DISABILITY", "SICK", "PERSONAL")


def leave_words_in(text: str) -> str:
    """modPublish.LeaveWordsIn (v24.3.1): the first whole word (letters only, upper-cased) that
    would tell members the kind of leave, or 'LEAVE OF ABSENCE'; '' when clean."""
    toks = [t.upper() for t in re.findall(r"[A-Za-z]+", text)]
    for i, t in enumerate(toks):
        if t in LEAVE_WORDS:
            return t
        if t == "LEAVE" and toks[i + 1:i + 3] == ["OF", "ABSENCE"]:
            return "LEAVE OF ABSENCE"
    return ""


def payroll_text(x) -> str:
    """modOvertimeEntry.PayrollText."""
    if x is None or isinstance(x, tuple):
        return ""
    if isinstance(x, str):
        return x.strip()
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return str(int(round(float(x)))) if float(x) > 0 else ""
    return ""


def is_payroll_header(s: str) -> bool:
    t = "".join(ch for ch in s.lower() if ch.isascii() and (ch.isalnum() or ch == "#"))
    return t in ("payroll#", "payroll", "payrollno", "payrollnumber", "payrollnum", "payrollid")


def is_header_match(cell_text: str, want: str) -> bool:
    """modOvertimeEntry.IsHeaderMatch (v24.3.1): letters and digits only, case and spacing ignored."""
    def key(s: str) -> str:
        return "".join(ch for ch in s.lower() if ch.isascii() and ch.isalnum())
    return key(cell_text) != "" and key(cell_text) == key(want)


def payroll_decision(roster_num: str, num_in_kronos: bool, names_agree: bool) -> str:
    if roster_num == "" or not num_in_kronos:
        return "name"
    return "number" if names_agree else "mismatch"


# ---------------------------------------------------------------- typed dates (v25, D-C7)
_MONTH_INDEX = {m.lower(): i + 1 for i, m in enumerate(_MONTHS)}


def _date_parts(text: str):
    """modPeople.TypedDateParts: (year or None, month, day) or None. Month-name forms, yyyy-mm-dd,
    m/d[/y] (month / day, as the boxes' hints say); a time after the date is dropped."""
    toks: list[str] = []
    clock_seen = False
    for w in text.replace(",", " ").replace(".", " ").split():
        lw = w.lower()
        if ":" in w:
            clock_seen = True
            continue
        if lw[:1].isdigit() and lw.endswith(("am", "pm")) and lw[:-2].isdigit():
            continue
        if lw in ("am", "pm"):
            if toks and toks[-1].isdigit() and not clock_seen:
                toks.pop()
            continue
        toks.append(w)
    if len(toks) == 1:
        tok = toks[0]
        parts = tok.split("-")
        if len(parts) == 3 and len(parts[0]) == 4 and all(p.isdigit() for p in parts) and 1 <= len(parts[1]) <= 2 and 1 <= len(parts[2]) <= 2:
            return int(parts[0]), int(parts[1]), int(parts[2])
        parts = tok.split("/")
        if len(parts) in (2, 3) and all(p.isdigit() for p in parts) and len(parts[0]) <= 2 and len(parts[1]) <= 2 \
                and (len(parts) == 2 or len(parts[2]) in (1, 2, 4)):
            year = None
            if len(parts) == 3:
                year = int(parts[2])
                year = year + 2000 if year < 100 else year
            return year, int(parts[0]), int(parts[1])
        return None
    if len(toks) not in (2, 3):
        return None

    def month_of(w: str):
        lw = w.lower()
        return _MONTH_INDEX.get(lw[:3]) if len(lw) >= 3 and lw.isalpha() else None

    def digits(w: str) -> bool:
        return w.isdigit() and len(w) <= 4

    if month_of(toks[0]):
        if not digits(toks[1]):
            return None
        month, day = month_of(toks[0]), int(toks[1])
    elif month_of(toks[1]) and digits(toks[0]):
        month, day = month_of(toks[1]), int(toks[0])
    else:
        return None
    year = None
    if len(toks) == 3:
        if not digits(toks[2]):
            return None
        year = int(toks[2])
        year = year + 2000 if year < 100 else year
    return year, month, day


def _valid_ymd(y: int, m: int, d: int) -> bool:
    if y < 100 or y > 9999 or not 1 <= m <= 12 or not 1 <= d <= 31:
        return False
    try:
        date(y, m, d)
        return True
    except ValueError:
        return False


def typed_date(text: str, today: date) -> str:
    """modPeople.TypedDateCore: 'yyyy-mm-dd', or the hint word: blank / not a date / check the year.
    A yearless date lands in the year that puts it inside today-90 .. today+275 (v25 review)."""
    t = text.strip()
    if not t:
        return "blank"
    parts = _date_parts(t)
    if parts is None:
        return "not a date"
    year, month, day = parts
    if year is None:
        year = today.year
        if not 1 <= month <= 12 or not 1 <= day <= 31:
            return "not a date"
        d = date(year, month, min(day, 28))                   # the year is chosen first ("Feb 29" before a leap February)
        if d < today - timedelta(days=90):
            year += 1
        elif d > today + timedelta(days=275):
            year -= 1
    if not _valid_ymd(year, month, day):
        return "not a date"
    d = date(year, month, day)
    if abs((d - today).days) > 400:
        return "check the year"
    return d.isoformat()


# ---------------------------------------------------------------- time off on scheduled shifts (v25, from frmTimeOff)
def trim_kind(ds: date, de: date, d_first: date, d_last: date) -> str:
    """modCalendar.TrimKind: which end(s) of a booking are days off: '' / start / end / both."""
    a, b = d_first > ds, d_last < de
    return "both" if a and b else "start" if a else "end" if b else ""


# ---------------------------------------------------------------- the Award pop-up's hours box (v25, from frmPostingEntry)
def partial_hours_value(text: str, post_hours: float) -> str:
    """modPostings.PartialHoursValue as the SelfTest shows it: 'whole' (blank, or the
    posting's hours or more), the hours as text, or 'bad' (not a number, 0 or less)."""
    t = text.strip()
    if not t:
        return "whole"
    try:
        h = float(t)
    except ValueError:
        return "bad"
    if h <= 0:
        return "bad"
    if h >= post_hours:
        return "whole"
    return f"{h:g}"


# ---------------------------------------------------------------- postings: the one closing test (v25, P6)
POSTING_CLOSE_HOUR = 8          # modPostings.POSTING_CLOSE_TIME "8:00 AM"


def posting_closed(closes_on: date, as_of: date | datetime) -> bool:
    """modPostings.PostingClosed: closed from POSTING_CLOSE_TIME on the closing date (v25 review: the one
    moment the printed sheet and the members' file use). A date as-of is that day's midnight."""
    at = as_of if isinstance(as_of, datetime) else datetime(as_of.year, as_of.month, as_of.day)
    return at >= datetime(closes_on.year, closes_on.month, closes_on.day, POSTING_CLOSE_HOUR)


# ---------------------------------------------------------------- the Kronos report (v27, ImportKronosData)
def premium_multiplier(code: str, codes: str = PREMIUM_CODES, mults: str = PREMIUM_MULTIPLIERS) -> float:
    """ImportKronosData.PremiumMultiplier: the code without its 'CA ' prefix, any case, against the parallel
    lists; -1 = not a roster code."""
    c = code.strip()
    if c[:3].upper() == "CA ":
        c = c[3:].strip()
    if not c:
        return -1
    cs = [x.strip() for x in codes.split(",")]
    ms = [x.strip() for x in mults.split(",")]
    for i, x in enumerate(cs):
        if x.lower() == c.lower():
            return _val(ms[i]) if i < len(ms) else -1
    return -1


def _val(s: str) -> float:
    """VBA's Val: the leading number of a string, 0 when there is none."""
    import re as _re
    m = _re.match(r"\s*[-+]?\d*\.?\d+", s)
    return float(m.group(0)) if m else 0.0


def round_half_up(x: float) -> float:
    """ImportKronosData.RoundHalfUp: Excel's ROUND(x, 0), half away from zero (VBA's Round is half-to-even)."""
    import math
    sign = -1 if x < 0 else 1
    out = float(math.floor(abs(x) + 0.5000001)) * sign
    return 0.0 if out == 0 else out          # never -0.0 (VBA prints 0)


def weighted_total(spec: str, codes: str = PREMIUM_CODES, mults: str = PREMIUM_MULTIPLIERS) -> str:
    """ImportKronosData.WeightedTotal: 'code=hours;code=hours' -> the roster codes' hours x their multiplier,
    summed and rounded whole, then '|' + the codes not counted (when any)."""
    tot = 0.0
    other: list[str] = []
    for part in spec.split(";"):
        if not part.strip() or "=" not in part:
            continue
        code, hrs = part.split("=", 1)
        m = premium_multiplier(code.strip(), codes, mults)
        if m < 0:
            other.append(code.strip())
        else:
            tot += _val(hrs) * m
    out = f"{round_half_up(tot):g}"
    return out + ("|" + ";".join(other) if other else "")


def time_period(text: str) -> str:
    """ImportKronosData.TimePeriodCore: the two m/d/yyyy dates in the report's Time Period line ->
    'yyyy-mm-dd|yyyy-mm-dd', or 'no dates' / 'end before start'."""
    import re as _re
    found: list[date] = []
    # as the VBA: every run of digits and slashes is a token; a token is a date only when it is exactly m/d/yyyy
    for tok in _re.findall(r"[\d/]+", text):
        parts = tok.split("/")
        if len(parts) != 3 or not all(parts):
            continue
        m, d_, y = (int(x) for x in parts)
        if not (1900 <= y <= 2999 and 1 <= m <= 12 and 1 <= d_ <= 31):
            continue
        try:
            found.append(date(y, m, d_))
        except ValueError:
            continue
    if len(found) < 2:
        return "no dates"
    a, b = found[0], found[1]
    if b < a:
        return "end before start"
    return f"{a.isoformat()}|{b.isoformat()}"


def report_kind(from_d: date, to_d: date, year_from: date, year_to: date) -> str:
    """ImportKronosData.ReportKindCore: 'day' (one day: a snapshot), 'cumulative' (starts within KRONOS_YEAR_SLACK of
    the year's first day), 'range' (several days, not the year to date: refused), 'outside' (not this year)."""
    if from_d is None or to_d is None:
        return "outside"
    if from_d < year_from or to_d < year_from or from_d > year_to:
        return "outside"
    if from_d == to_d:
        return "day"
    return "cumulative" if from_d <= year_from + timedelta(days=KRONOS_YEAR_SLACK) else "range"


# ---------------------------------------------------------------- the weekly cap (v27 item 21, modOvertimeEntry)
def week_cap_over(hours_before: float, offer: float, cap: float) -> float:
    """modOvertimeEntry.WeekCapOver: the hours past the weekly cap (0 = the offer fits)."""
    if cap <= 0:
        return 0.0                              # 0 = the warning is off (the default since Oct 22)
    o = hours_before + offer - cap
    return o if o > 0.001 else 0.0


# ---------------------------------------------------------------- the receipts (v27, modReceipt)
_REF = re.compile(r"(?<![A-Za-z0-9])[A-Z]{1,4}-[0-9]+(?![A-Za-z0-9])")


def first_ref(txt: str) -> str:
    """modReceipt.FirstRefIn: the first Ref ID standing on its own (1-4 capitals, a dash, digits), else ''."""
    m = _REF.search(txt)
    return m.group(0) if m else ""


# ---------------------------------------------------------------- the Stats report (v27, modStats)
def stats_quarter(d: date, jan1: date) -> str:
    """modStats.StatsQuarter: the calendar quarter, Q1..Q4; the roster year's late-December start (before jan1) is Q1."""
    if d < jan1:
        return "Q1"
    return f"Q{(d.month - 1) // 3 + 1}"


def stats_bucket(typ: str, outcome: str, status: str) -> str:
    """modStats.StatsBucket: what an OT Log row counts as."""
    if status.strip().upper() == "VOID":
        return "void"
    t, o = typ.strip().lower(), outcome.strip().lower()
    if t == "callout":
        return {"accepted": "accepted", "refused": "refused", "left message": "message", "unavailable": "unavailable",
                "exception": "exception"}.get(o, "other")
    if t == "posting":
        return {"awarded": "awarded", "refused": "p-refused", "signed up": "signedup"}.get(o, "other")
    if t == "adjustment":
        return "adjustment"
    return "other"


def time_of_day_band(hr: int) -> str:
    """modStats.TimeOfDayBand: the band a session was logged in."""
    if 0 <= hr <= 5:
        return "night 0-6"
    if 6 <= hr <= 11:
        return "morning 6-12"
    if 12 <= hr <= 17:
        return "afternoon 12-18"
    return "evening 18-24"


def spread(spec: str) -> str:
    """modStats.SpreadCore: 'h;h;h' -> 'min|max|mean|sd|gap' to one decimal (population SD); '-|-|-|-|-' for none."""
    import math
    vals = [float(x) for x in spec.split(";") if x.strip()]
    if not vals:
        return "-|-|-|-|-"
    mean = sum(vals) / len(vals)
    sd = math.sqrt(sum((x - mean) ** 2 for x in vals) / len(vals))

    def one(x: float) -> str:
        # VBA Format$(x, "0.0") rounds half away from zero (22.25 -> 22.3); Python's format is half-to-even on the binary value
        from decimal import Decimal, ROUND_HALF_UP
        return str(Decimal(repr(x)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))

    return "|".join(one(x) for x in (min(vals), max(vals), mean, sd, max(vals) - min(vals)))


# ---------------------------------------------------------------- the Reconcile sheet (v27, modReconcile)
def reconcile_diff(paper: str, roster: str) -> str:
    """modReconcile.ReconcileDiffCore: names top to bottom, ';' between -> 'positions that differ|paper names
    the roster does not list|roster names not on the paper', '-' for an empty part; names compare trimmed,
    any case; a blank paper line is an empty position."""
    p = [x.strip() for x in paper.split(";")] if paper.strip() else []
    r = [x.strip() for x in roster.split(";")] if roster.strip() else []
    n = max(len(p), len(r))

    def at(lst: list[str], i: int) -> str:
        return lst[i] if i < len(lst) else ""

    diff = [str(i + 1) for i in range(n) if at(p, i).lower() != at(r, i).lower()]
    rl = [x.lower() for x in r]
    pl = [x.lower() for x in p]
    missing = [x for x in p if x and x.lower() not in rl]
    extra = [x for x in r if x and x.lower() not in pl]
    return "|".join(";".join(part) if part else "-" for part in (diff, missing, extra))


def view_mode_for(login: str, admin_list: str, desk_list: str, sup_list: str = "") -> str:
    """v25.3 modDesk.ViewModeFor: 'desk' = on ADMIN_USERS and DESK_USERS, 'admin' = ADMIN_USERS only,
    'supervisor' (v25.5) = on SUPERVISOR_USERS and not an admin, 'user' = none of them (a desk-only login
    has no admin rights). Lists comma separated, any case, spaces ignored."""
    u = login.strip().lower()
    if not u:
        return "user"

    def in_list(lst: str) -> bool:
        return u in [x for x in lst.replace(" ", "").lower().split(",") if x]

    a, d, s = in_list(admin_list), in_list(desk_list), in_list(sup_list)
    return "desk" if a and d else "admin" if a else "supervisor" if s else "user"


def data_shape_problem(kind: str, v: str) -> str:
    """modPeople.DataShapeProblem (v27 item 25): '' when a roster detail looks right, else what is wrong."""
    v = v.strip()
    k = kind.lower()
    if k == "phone":
        if v == "":
            return ""
        low = v.lower()
        digits = sum(1 for c in low if c in "0123456789")
        bad = any(c in "abcdefghijklmnopqrstuvwxyz" and c not in "ext" for c in low)
        return "phone does not look like a phone number" if bad or digits < 7 or digits > 15 else ""
    if k == "email":
        if v == "":
            return ""
        at = v.find("@") + 1                          # 1-based like InStr; 0 = none
        if at < 2:
            bad = True
        else:
            bad = ("@" in v[at:] or v.find(".", at + 1) == -1 or " " in v or v.endswith("."))
        return "email does not look like an address" if bad else ""
    if k == "code":
        return f"code {v} has no grade number" if v != "" and dept_code_value(v) == 999 else ""
    if k == "crew":
        return "crew blank" if v == "" else ""
    return ""


# ---------------------------------------------------------------- v27 (item 30, C8): constants the drift test ties to the VBA
EXEMPT_TAG = "Exempt - "              # modPostings.EXEMPT_TAG: starts a list row's never-charged note
PAYBACK_WARN_DAYS = 42                # modOvertimeEntry.PAYBACK_WARN_DAYS (the Settings sheet may override it)
SIGNUP_STALE_DAYS = 7                 # modOvertimeEntry.SIGNUP_STALE_DAYS
POSTING_CLOSE_DAYS = 14               # modPostings.POSTING_CLOSE_DAYS


def order_tie(grade_above: float, hours_above: float, grade: float, hours: float) -> bool:
    """modOvertimeEntry.OrderTie (v27 item 27): tied with the row above on grade and Total Hours."""
    return abs(grade_above - grade) < 0.0001 and round(hours_above, 2) == round(hours, 2)


def public_not_free_word(why: str) -> str:
    """modPublish.PublicNotFreeWord (v27 Members M4 item 37): a not-free reason as members may read it."""
    w = why.strip().lower()
    if w == "":
        return ""
    return {"scheduled": "scheduled", "rest": "rest rule", "s/s": "switch", "s/s off": "switch",
            "16 h rule": "16 h rule", "on ot": "on OT that shift"}.get(w, "off")


# ---- v27.1: the Kronos ledger rule's pure cores (ImportKronosData.bas; suites kledger / kedit / kday / kname) ----
EPOCH = date(1899, 12, 30)          # VBA / Excel day 0


def serial(d: date) -> int:
    """A whole day's VBA serial (CLng(Int(CDbl(d))))."""
    if isinstance(d, datetime):
        d = d.date()
    return (d - EPOCH).days


def _ser(d: date | None) -> float:
    """A date as the number VBA compares (0 = none)."""
    if d is None:
        return 0.0
    if isinstance(d, datetime):
        return (d - datetime(1899, 12, 30)).total_seconds() / 86400.0
    return float(serial(d))


def ledger_cum_of(cum: date | None, year_from: date) -> date | None:
    """ImportKronosData.LedgerCumOf (Private): a Totals snapshot from before the roster year is none."""
    if cum is None or _ser(cum) < _ser(year_from):
        return None
    return cum


def ledger_run(cum: date | None, year_from: date, day_serials: str) -> date | None:
    """ImportKronosData.LedgerRunCore: the Totals' Through, then each next day the ledger holds along the unbroken run;
    only days after the Through and in the roster year count. None = no snapshot this year."""
    cum = ledger_cum_of(cum, year_from)
    out = cum
    days: set[int] = set()
    for part in day_serials.split(";"):
        if part.strip():
            k = round(_val(part))            # CLng(Val(...)): half to even, as Python's round
            if k > _ser(cum) and k >= _ser(year_from):
                days.add(k)
    if cum is not None:
        d = cum + timedelta(days=1)
        while serial(d) in days:
            out = d
            d = d + timedelta(days=1)
    return out


def ledger_day_in(d: date, cum: date | None, run_end: date | None) -> bool:
    """ImportKronosData.LedgerDayIn: after the Totals' Through and inside the run."""
    return _ser(d) > _ser(cum) and _ser(d) <= _ser(run_end)


def ledger_person(snap_spec: str, day_spec: str) -> float:
    """ImportKronosData.LedgerPersonCore: the snapshot's 'code=h;...' then the run's, weighted and rounded once
    (WeightedTotal), the codes not counted dropped."""
    s = snap_spec
    if day_spec != "":
        s = s + ("" if s == "" else ";") + day_spec
    w = weighted_total(s)
    if "|" in w:
        w = w[:w.index("|")]
    return _val(w)


def day_reimported_after(reimported_at: datetime | None, when) -> bool:
    """ImportKronosData.DayReimportedAfter (Private): the day re-imported at or after the edit's When."""
    if reimported_at is None:
        return False
    if not isinstance(when, (date, datetime)):
        return False                                # IsDate False: never "after"
    return _ser(reimported_at) >= _ser(when)


def edit_counts(d: date, year_from: date, ledger_from: date | None, reimported_at: datetime | None, when) -> bool:
    """ImportKronosData.EditCountsCore: a timecard edit dated d counts in Worked Periods when it is this year, from the
    first one-day report on, and the day was not re-imported at or after the edit was logged."""
    return (_ser(d) >= _ser(year_from) and _ser(ledger_from) > 0 and _ser(d) >= _ser(ledger_from)
            and not day_reimported_after(reimported_at, when))


def gap_days_text(old_through: date | None, to_d: date, held: set[int]) -> str:
    """ImportKronosData.GapDaysText (v27.5, item 2): the days in (oldThrough, toD] the ledger does not hold, as
    'yyyy-mm-dd; yyyy-mm-dd' ('' = none; no snapshot yet = none). held = whole-day serials the ledger holds."""
    if _ser(old_through) <= 0:
        return ""
    out = []
    x = old_through + timedelta(days=1)
    while x <= to_d:
        if int(_ser(x)) not in held:
            out.append(x.strftime("%Y-%m-%d"))
        x += timedelta(days=1)
    return "; ".join(out)


def gap_note_days(note: str) -> list[int]:
    """ImportKronosData.GapNoteDays: every 'yyyy-mm-dd' (digits in every place, month 1-12, day 1-31) in a gap row's
    Note, as whole-day serials, left to right (the VBA moves on by one character after a window that is not a date; the
    two agree on every Note the import writes)."""
    out: list[int] = []
    for m in re.finditer(r"\d{4}-\d{2}-\d{2}", note):
        s = m.group(0)
        mo, dd = int(s[5:7]), int(s[8:10])
        if 1 <= mo <= 12 and 1 <= dd <= 31:
            out.append(int(_ser(date(int(s[0:4]), 1, 1)) + (date(int(s[0:4]), mo, 1) - date(int(s[0:4]), 1, 1)).days + dd - 1))
    return out


def gap_counts(d: date, year_from: date, ledger_from: date | None, gap_days: list[int], held: set[int]) -> bool:
    """ImportKronosData.GapCountsCore: a gap row dated d counts in Worked Periods when it is this year, from the first
    one-day report on, and the person has no ledger row on any day it names (held = the person's days)."""
    if not (_ser(d) >= _ser(year_from) and _ser(ledger_from) > 0 and _ser(d) >= _ser(ledger_from)):
        return False
    return not any(g in held for g in gap_days)


def ledger_day_drop(row_date, d: date, emp: str, on_roster: bool) -> bool:
    """ImportKronosData.LedgerDayDrop: a one-day report for d replaces a ledger row of that day that is a roster
    person's or the quiet-day marker (blank Employee)."""
    if not isinstance(row_date, (date, datetime)):
        return False
    if int(_ser(row_date) // 1) == int(_ser(d) // 1):
        return emp == "" or on_roster
    return False


def roster_name(by_num: str, by_name: str) -> str:
    """ImportKronosData.RosterNameCore: clash / number / name / none."""
    if by_num != "" and by_name != "" and by_num.lower() != by_name.lower():
        return "clash"
    if by_num != "":
        return "number"
    if by_name != "":
        return "name"
    return "none"


# ---- v27.1: the staffing counts as a pure core (modMinStaff.StaffCountsCore; suite staffcount) ----
_NUMERIC = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?$")
_VAL = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")


def vb_trim(s: str) -> str:
    """Trim$: spaces only."""
    return s.strip(" ")


def vb_isnumeric(s: str) -> bool:
    """IsNumeric for the text CoveredShare gives it: a run with no spaces between the last space and the final h.
    Plain decimal numbers only (hex, currency and thousands forms are out of scope; the cases avoid them)."""
    return bool(_NUMERIC.match(s))


def vb_val(s: str) -> float:
    """Val on a text with no spaces: the leading number, 0 when none."""
    m = _VAL.match(s)
    return float(m.group(0)) if m else 0.0


def covered_share(rt: str, shift_hours: float = USUAL_SHIFT_HOURS) -> float:
    """modMinStaff.CoveredShare (v27.5, item 3; CoveredCount before): one person per part (" / " between the names of one
    award, "; " between two covers of one line); a part ending in "Nh" with 0 < N < the shift's hours counts N / hours of a
    person; a blank part and a NO TAKERS / POSTED / NCN part add nobody; no rounding (the core rounds once)."""
    total = 0.0
    if shift_hours <= 0:
        shift_hours = USUAL_SHIFT_HOURS
    for raw in rt.replace("; ", " / ").split(" / "):
        part = vb_trim(raw)
        up = part.upper()
        if part == "" or up.startswith("NO TAKERS") or up.startswith("POSTED") or up == "NCN":
            continue
        hrs = 0.0
        if part[-1:].lower() == "h":
            p = part.rfind(" ") + 1                         # InStrRev, 1-based (0 = no space)
            if p > 0 and vb_isnumeric(part[p:len(part) - 1]):   # Mid$(part, p + 1, Len(part) - p - 1)
                hrs = vb_val(part[p:])                          # Val(Mid$(part, p + 1))
        if 0 < hrs < shift_hours - 0.001:
            total += hrs / shift_hours
        else:
            total += 1
    return total


def staff_counts_core(people: list[str], right_texts: list[str],
                      shift_hours: float = USUAL_SHIFT_HOURS) -> tuple[int, int, int, bool]:
    """modMinStaff.StaffCountsCore: scheduled, off, covered (every line's share added once, rounded DOWN) and part
    (a fraction was dropped) - v27.5, item 3; shift_hours = the shift's length (a clock-change night's 11 / 13)."""
    scheduled = off = covered = 0
    total = 0.0
    for tok in people:
        if "S" in tok:
            scheduled += 1
            if "A" in tok or "W" in tok:
                off += 1
    for raw in right_texts:
        rt = vb_trim(raw)
        if rt != "":
            total += covered_share(rt, shift_hours)
    covered = int(math.floor(total + 0.000001))              # CLng(Int(total + 0.000001))
    return scheduled, off, covered, (total - covered > 0.001)
