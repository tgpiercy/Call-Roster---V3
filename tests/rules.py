"""Python mirror of the Call Out Roster's pure rules.

Each function here mirrors a VBA procedure named in its docstring. The
known answers in cases.csv are shared with the workbook's SelfTest
(modSelfTest.bas), so the two cannot drift apart. Nothing here reads a
workbook: the rotation, holidays, DST, call order, charging, the rotation
part of eligibility and the provisional-OT date logic are all arithmetic.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

# ---------------------------------------------------------------- settings
CHARGE_MULTIPLIER = 1.75        # modOvertimeEntry.CHARGE_MULTIPLIER
CHARGE_MAX_GRADE = 2            # modOvertimeEntry.CHARGE_MAX_GRADE
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
        elif n_awarded > 0 and p.rank < max_rank and p.grade <= CHARGE_MAX_GRADE:
            p.charged = hours * CHARGE_MULTIPLIER
            p.result = "Charged"
        else:
            p.result = ""
    return lst


def compute_no_takers(hours: float, lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeNoTakers: everyone eligible is charged."""
    for p in lst:
        p.awarded = False
        p.charged = 0.0
        p.result = ""
        if p.grade <= CHARGE_MAX_GRADE:
            p.charged = hours * CHARGE_MULTIPLIER
            p.result = "Charged"
    return lst


def compute_outside(hours: float, lst: list[Listed]) -> list[Listed]:
    """modPostings.ComputeOutside: in-department people who didn't sign up
    are charged as for no takers; anyone who signed up isn't."""
    for p in lst:
        p.awarded = False
        p.charged = 0.0
        p.result = ""
        if p.signed_up:
            p.result = "Signed up"
        elif p.grade <= CHARGE_MAX_GRADE:
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
    night = is_night_shift(shift_lbl)
    on_d = crew_shift_on(code, d, observed)
    if on_d in ("Days", "Nights"):
        return "scheduled"
    if on_d != "?":
        if not night and crew_shift_on(code, d - timedelta(days=1), observed) == "Nights":
            return "rest"
        if night and crew_shift_on(code, d + timedelta(days=1), observed) == "Days":
            return "rest"
    return ""


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
