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
    return sorted(people, key=lambda e: (e.grade, e.total_hrs, seniority_key(e.seniority), e.name.casefold()))


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


def type_tag(code: str) -> str:
    """modCalendar.TypeTag (never STD / PAR outside the Absences sheet)."""
    c = code.upper()
    return {"VAC": "Vac", "TRN": "Trn", "LOA": "Leave", "STD": "Leave", "PAR": "Leave",
            "SS": "S/S", "PHA": "PHA", "SUP": "Support"}.get(c, "Off")


def not_free_reason(code: str, d: date, shift_lbl: str, *, booked_off: str = "",
                    ss_working: bool = False, ss_gave_up: bool = False,
                    ss_rest: bool = False) -> str:
    """modCalendar.NotFreeReason with the table lookups replaced by flags:
    booked_off = absence type code ('VAC', 'LOA' ...) active on d;
    ss_working = working a switch on d; ss_gave_up = gave up a shift on d;
    ss_rest = working a switch on the adjoining shift (nights before a day
    OT, days after a night OT)."""
    night = is_night_shift(shift_lbl)
    if booked_off:
        return type_tag(booked_off)
    why = rotation_free_reason(code, d, shift_lbl)
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
    return grade <= CHARGE_MAX_GRADE and not is_returner and not leave_overlaps and has_history_before


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
    if not (s[0:4].isdigit() and s[4] == "-" and s[5:7].isdigit() and s[7] == "-" and s[8:10].isdigit()):
        return None
    try:
        return date(int(s[0:4]), int(s[5:7]), int(s[8:10]))
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
    below; a whole-shift acceptance changes nothing."""
    remain = hours
    out = []
    for outcome, taken in calls:
        out.append(remain)
        if outcome == "Accepted" and 0 < taken < remain:
            remain -= taken
    return out


def accepted_hours(offer: float, taken: float) -> float:
    """modOvertimeEntry.AcceptedHours."""
    return taken if 0 < taken < offer else offer


# ---------------------------------------------------------------- electronic sign-ups (v24.1)
def reduce_signups(rows: list[tuple[str, str, float, datetime]], closes_at: datetime | None):
    """modSignups.ReduceSignups. rows = (employee, action, hours, stamp) in any
    order. Each person's action with the latest stamp wins (same stamp: the
    later row). Returns [(employee, hours, late)] for people whose newest action
    is a sign-up, in the order each name first appeared."""
    newest: dict[str, int] = {}
    names: dict[str, str] = {}
    for i, (emp, _action, _hours, stamp) in enumerate(rows):
        key = emp.strip().casefold()
        if not key:
            continue
        if key not in newest:
            newest[key] = i
            names[key] = emp.strip()
        elif stamp >= rows[newest[key]][3]:
            newest[key] = i
    out = []
    for key, i in newest.items():
        emp, action, hours, stamp = rows[i]
        if action.strip().lower() == "signup":
            out.append((emp.strip(), hours, closes_at is not None and stamp > closes_at))
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


def longest_digit_run(s: str) -> int:
    joined = _joined_digits(s)
    return joined if joined >= 10 else _single_run(s)


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

