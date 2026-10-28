"""The mock plant's state engine (P9, the alpha of Nov 1-14, 2026).

One in-memory model of what the roster holds - people, crew history, absences, switches, the OT Log, postings,
restrictions, planned moves, the Kronos snapshot and ledger - and the readers the plan's EXPECT figures rest on.
Every rule here is tests/rules.py's (the Python mirror of the pure VBA rules, shared with SelfTest); this file only
walks the tables the way the sheet-facing VBA does (NotFreeReason, FilterNotFree, LoadEligible, FatigueInputs,
SubmitCalloutSession, WriteOutcome, the ledger). Nothing in it is a rule of its own: where a figure needs arithmetic
the mirror does not have, the mirror gets the function first (tests/rules.py + cases.csv), never this file.

Used by tests/mock/make_plant.py (to generate the historic rows consistently) and tests/mock/expected.py (to replay
docs/alpha-test-plan.md's steps and write expected.csv).
"""
from __future__ import annotations

import csv
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import rules  # noqa: E402

# ---------------------------------------------------------------- the roster's shape (modOvertimeEntry)
UNITS = ["Phosphine", "Derivatives", "Cypure", "Material Handler", "QC Gases", "QC Derivatives", "Utilities", "SBP"]
UNIT_COL = {"QC Derivatives": 12, "Derivatives": 13, "Cypure": 14, "Phosphine": 15,
            "Material Handler": 16, "QC Gases": 17, "Utilities": 18, "SBP": 19}
YEAR_FROM = date(2025, 12, 28)        # the roster's period-1 start (T2)
YEAR_TO = date(2026, 12, 26)
LOG_HEADERS = ["Ref ID", "Type", "Shift Date", "Unit", "Shift", "Crew", "Reason", "Hours", "Call Order", "Employee",
               "Dept Code", "Total Hrs Then", "Outcome", "Note", "Hours Charged", "Period Start", "Entry ID",
               "Logged By", "Timestamp", "Status", "Void Note", "Payroll #"]
CREW_ALL = "All not scheduled"
USUAL = rules.USUAL_SHIFT_HOURS


def period_start(d: date) -> date:
    """modOvertimeEntry.PeriodStartFor: the Sunday that starts d's biweekly pay period."""
    k = (d - YEAR_FROM).days // 14
    return YEAR_FROM + timedelta(days=14 * k)


def period_number(d: date) -> int:
    return (d - YEAR_FROM).days // 14 + 1


def iso(d: date | None) -> str:
    return d.isoformat() if d else ""


def hours_text(x: float) -> str:
    """modPostings.HoursText = CStr(Round(x, 2)): two decimals at most, no trailing zeros, never a significant-digit cut."""
    s = f"{round(x, 2):.2f}".rstrip("0").rstrip(".")
    return "0" if s in ("", "-0") else s


def shift_matches(emp_shift: str, sel: str, include_days: bool = True) -> bool:
    """modOvertimeEntry.ShiftMatches."""
    es = rules.norm_crew(emp_shift)
    ss = rules.norm_crew(sel)
    if ss == "ALL":
        return True
    if es == "" or ss == "":
        return False
    if es == "DAYS":
        return ss == "DAYS" or include_days
    if ss == "DAYS":
        return False
    return any(e and e == s for e in es.split("/") for s in ss.split("/"))


def crew_filter(sel: str) -> str:
    return "ALL" if sel.strip() == "" or sel.strip().lower() == CREW_ALL.lower() else sel


# ---------------------------------------------------------------- records
@dataclass
class Person:
    payroll: str
    home: str
    name: str
    crew: str
    phone: str
    worked: float            # Worked (H) opening, before any import this engine applies
    seniority: date | None
    email: str
    codes: dict              # unit -> grade code ("" = none)

    def grade_in(self, unit: str) -> float:
        return rules.dept_code_value(self.codes.get(unit, "") or "")


@dataclass
class Rec:
    """One listed person (modOvertimeEntry.EmpRec)."""
    name: str
    grade: float
    total: float
    seniority: date | None
    code: str
    crew: str
    max_hours: float = rules.FATIGUE_MAX_HOURS
    est: bool = False
    roster_total: float = 0.0


@dataclass
class Plant:
    people: dict = field(default_factory=dict)          # name -> Person (roster order kept by insertion)
    crew_hist: dict = field(default_factory=dict)       # name -> [(crew, from)] in table order
    absences: list = field(default_factory=list)        # dicts: id, emp, unit, type, start, end, hours, shift, status, note, home
    coverage: list = field(default_factory=list)        # dicts: the calendar lines (display only here)
    switches: list = field(default_factory=list)        # dicts: id, date, shift, unit, off, works, payfor, paidby, status
    restrictions: list = field(default_factory=list)    # dicts: id, emp, start, end, reason, status
    planned: list = field(default_factory=list)         # dicts: id, emp, unit, code, clear, crew, start, status
    log: list = field(default_factory=list)             # dicts keyed by LOG_HEADERS, in ENTRY order (oldest first)
    postings: dict = field(default_factory=dict)        # id -> dict (status, posted, closes, date, unit, shift, crew, hours, positions, reason, award_to)
    posting_lists: dict = field(default_factory=dict)   # id -> [dict(rank, name, crew, code, total, signed, awarded, result, signup_hrs)]
    hours_history: list = field(default_factory=list)   # (name, data_through, worked) - the KR- snapshots
    idmarks: dict = field(default_factory=dict)         # prefix -> highest number ever used
    worked: dict = field(default_factory=dict)          # name -> H as the roster shows it now
    kronos_through: date | None = None                  # KronosDataThrough (None = nothing imported)
    kronos_cum_through: date | None = None              # the last cumulative's Through
    kronos_snapshot: dict = field(default_factory=dict) # name -> {code: hours} of the last cumulative
    kronos_ledger: dict = field(default_factory=dict)   # (name, date) -> {code: hours}
    ledger_days: set = field(default_factory=set)       # every day any one-day report was held for
    refused_through: date = date(2026, 11, 1)           # RefusedThrough: the current period's start
    today: date = date(2026, 11, 1)
    entry_next: int = 1                                 # the next Entry ID (#n)

    # ------------------------------------------------------------ people and crews
    def crew_on(self, name: str, d: date) -> str:
        p = self.people[name]
        return rules.crew_on(self.crew_hist.get(name, []), d, rules.norm_crew(p.crew))

    def unit_code_on(self, name: str, unit: str, d: date) -> str:
        """modPlannedMoves.UnitCodeOn: the roster's code unless a Planned move dated on or before d says otherwise."""
        code = self.people[name].codes.get(unit, "") or ""
        if not self.planned:
            return code
        plans = [(m["emp"], m["unit"], m["code"], m.get("clear", ""), m["start"], m["status"]) for m in self.planned]
        return rules.planned_code(plans, name, unit, d, code)

    def moved_in(self, name: str, unit: str, d: date) -> bool:
        return self.unit_code_on(name, unit, d) != "" and (self.people[name].codes.get(unit, "") or "") == ""

    # ------------------------------------------------------------ absences and switches
    def absent_for_shift(self, name: str, d: date, shift: str) -> str:
        """modCalendar.AbsentForShift: the type tag ('' = not booked off for that shift)."""
        night = rules.is_night_shift(shift)
        for a in self.absences:
            if a["emp"].casefold() != name.casefold() or a["status"].upper() == "VOID":
                continue
            code = a["type"].upper()
            if code == "SIC" and a.get("shift"):
                if rules.sick_blocks(a["start"], rules.is_night_shift(a["shift"]), d, night):
                    return rules.type_tag(code)
            elif a["start"] <= d and (a["end"] is None or d <= a["end"]):
                return rules.type_tag(code)
        return ""

    def absent_on(self, name: str, d: date) -> str:
        for a in self.absences:
            if a["emp"].casefold() == name.casefold() and a["status"].upper() != "VOID":
                if a["start"] <= d and (a["end"] is None or d <= a["end"]):
                    return rules.type_tag(a["type"].upper())
        return ""

    def _switch(self, col: str, name: str, d: date, shift: str = "") -> str:
        for s in self.switches:
            if s["status"].upper() == "VOID":
                continue
            if s[col].casefold() == name.casefold() and s["date"] == d and (shift == "" or s["shift"].casefold() == shift.casefold()):
                return s["id"]
        return ""

    def switch_working_on(self, name: str, d: date, shift: str = "") -> str:
        return self._switch("works", name, d, shift)

    def switch_off_on(self, name: str, d: date, shift: str = "") -> str:
        return self._switch("off", name, d, shift)

    def restricted_on(self, name: str, d: date) -> bool:
        for r in self.restrictions:
            if r["emp"].casefold() == name.casefold() and r["status"].upper() == "ACTIVE":
                if r["start"] <= d and (r["end"] is None or d <= r["end"]):
                    return True
        return False

    # ------------------------------------------------------------ the log and the totals
    def live_rows(self, name: str | None = None):
        for r in self.log:
            if r["Status"].upper() == "VOID":
                continue
            if name is None or r["Employee"].casefold() == name.casefold():
                yield r

    def ot_worked_on(self, name: str, d: date, night: bool) -> float:
        """modCalendar.OtWorkedOn: OT accepted / awarded on one shift."""
        tot = 0.0
        for r in self.live_rows(name):
            if r["Outcome"] in ("Accepted", "Awarded") and r["Shift Date"] == d and rules.is_night_shift(r["Shift"]) == night:
                tot += float(r["Hours"] or 0)
        return tot

    def refused(self, name: str) -> float:
        """Roster column I: charges whose pay period has started (RefusedFormula)."""
        return sum(float(r["Hours Charged"] or 0) for r in self.live_rows(name)
                   if r["Period Start"] and rules.refused_counts(r["Period Start"], self.refused_through))

    def provisional(self, name: str) -> float:
        """Roster column C: accepted / awarded OT not yet covered by Kronos (ProvisionalCounts)."""
        tot = 0.0
        through = self.kronos_through or date(1899, 12, 30)
        for r in self.live_rows(name):
            if rules.provisional_counts(r["Outcome"], r["Status"], r["Shift Date"], rules.is_night_shift(r["Shift"]),
                                        through, YEAR_FROM, YEAR_TO):
                tot += float(r["Hours"] or 0)
        return tot

    def total(self, name: str) -> float:
        """Roster column G = H + I + C."""
        return round(self.worked.get(name, 0.0) + self.refused(name) + self.provisional(name), 2)

    def ghic(self, name: str) -> tuple[float, float, float, float]:
        h = self.worked.get(name, 0.0)
        i = self.refused(name)
        c = self.provisional(name)
        return round(h + i + c, 2), h, i, c

    # ------------------------------------------------------------ the 16 h rule (modCalendar)
    def worked_hours_on(self, name: str, d: date, night: bool) -> float:
        """modCalendar.WorkedHoursOn."""
        lbl = "Nights" if night else "Days"
        if self.absent_for_shift(name, d, lbl):
            return 0.0
        ot = self.ot_worked_on(name, d, night)
        if ot > 0:
            return ot
        full = rules.default_shift_hours(d, night)
        if self.switch_working_on(name, d, lbl):
            return full
        if self.switch_off_on(name, d, lbl):
            return 0.0
        if rules.crew_shift_on(self.crew_on(name, d), d) == lbl:
            return full
        return 0.0

    def fatigue_inputs(self, name: str, d: date, night: bool) -> tuple[float, float]:
        """modCalendar.FatigueInputs -> (worked before, hours to next)."""
        h_pp = h_p = h_n1 = h_n2 = 0.0
        if night:
            h_p = self.worked_hours_on(name, d, False)
            if h_p > 0:
                h_pp = self.worked_hours_on(name, d - timedelta(days=1), True)
            h_n1 = self.worked_hours_on(name, d + timedelta(days=1), False)
            if h_n1 == 0:
                h_n2 = self.worked_hours_on(name, d + timedelta(days=1), True)
        else:
            h_p = self.worked_hours_on(name, d - timedelta(days=1), True)
            if h_p > 0:
                h_pp = self.worked_hours_on(name, d - timedelta(days=1), False)
            h_n1 = self.worked_hours_on(name, d, True)
            if h_n1 == 0:
                h_n2 = self.worked_hours_on(name, d + timedelta(days=1), False)
        return rules.fatigue_block_core(h_pp, h_p, h_n1, h_n2)

    def fatigue_max_for(self, name: str, d: date, shift: str) -> float:
        wb, htn = self.fatigue_inputs(name, d, rules.is_night_shift(shift))
        return rules.fatigue_max_hours(wb, htn)

    def fatigue_shift_reason(self, name: str, d: date, shift: str) -> str:
        """modCalendar.FatigueShiftReason: the 16 h rule for a WHOLE shift worked in a switch."""
        night = rules.is_night_shift(shift)
        if self.ot_worked_on(name, d, night) > 0:
            return "on OT"
        if self.fatigue_max_for(name, d, shift) < rules.default_shift_hours(d, night) - 0.001:
            return "16 h rule"
        return ""

    # ------------------------------------------------------------ who is free (modCalendar.NotFreeReason)
    def not_free_reason(self, name: str, d: date, shift: str, hours_offered: float = 0.0) -> tuple[str, float]:
        night = rules.is_night_shift(shift)
        t = self.absent_for_shift(name, d, shift)
        if t:
            return t, rules.FATIGUE_MAX_HOURS
        why = rules.rotation_free_reason3(self.crew_on(name, d - timedelta(days=1)), self.crew_on(name, d),
                                          self.crew_on(name, d + timedelta(days=1)), d, shift)
        if why == "rest":
            if not night and self.switch_off_on(name, d - timedelta(days=1), "Nights"):
                why = ""
            if night and self.switch_off_on(name, d + timedelta(days=1), "Days"):
                why = ""
        if why:
            return why, rules.FATIGUE_MAX_HOURS
        if self.switch_working_on(name, d):
            return "S/S", rules.FATIGUE_MAX_HOURS
        if self.switch_off_on(name, d):
            return "S/S off", rules.FATIGUE_MAX_HOURS
        if not night and self.switch_working_on(name, d - timedelta(days=1), "Nights"):
            return "rest", rules.FATIGUE_MAX_HOURS
        if night and self.switch_working_on(name, d + timedelta(days=1), "Days"):
            return "rest", rules.FATIGUE_MAX_HOURS
        if self.ot_worked_on(name, d, night) > 0:
            return "on OT", rules.FATIGUE_MAX_HOURS
        mx = self.fatigue_max_for(name, d, shift)
        if mx < (hours_offered if hours_offered > 0 else 0.5) - 0.001:
            return "16 h rule", mx
        return "", mx

    # ------------------------------------------------------------ the list (LoadEligible + FilterNotFree)
    def load_eligible(self, unit: str, crew_sel: str, d: date, max_grade: float = 999) -> list[Rec]:
        sel = crew_filter(crew_sel)
        recs: list[Rec] = []
        peers: list[tuple[float, bool, float]] = []
        movers: list[Rec] = []
        for p in self.people.values():
            code = self.unit_code_on(p.name, unit, d)
            if not code:
                continue
            mv = self.moved_in(p.name, unit, d)
            peers.append((rules.dept_code_value(code), mv, self.total(p.name)))
            crew = self.crew_on(p.name, d)
            if not shift_matches(crew, sel, True):
                continue
            g = rules.dept_code_value(code)
            if g > max_grade:
                continue
            r = Rec(p.name, g, self.total(p.name), p.seniority, code, crew, roster_total=self.total(p.name))
            recs.append(r)
            if mv:
                movers.append(r)
        for r in movers:
            lh = rules.planned_list_hours(r.total, peers)
            if lh > r.total:
                r.total, r.est = lh, True
        recs.sort(key=lambda r: (r.grade, round(r.total, 2), rules.seniority_key(r.seniority), r.name.casefold()))
        return recs

    def filter_not_free(self, recs: list[Rec], d: date, shift: str) -> tuple[list[Rec], dict]:
        kept, removed = [], {}
        for r in recs:
            if self.restricted_on(r.name, d):
                removed[r.name] = "restricted"
                continue
            why, mx = self.not_free_reason(r.name, d, shift, 0)
            if why:
                removed[r.name] = why
            else:
                r.max_hours = mx
                kept.append(r)
        return kept, removed

    def callout_list(self, unit: str, d: date, shift: str, crew_sel: str = CREW_ALL) -> tuple[list[Rec], dict]:
        """The call-out pop-up's list in call order, and who was left off and why."""
        return self.filter_not_free(self.load_eligible(unit, crew_sel, d), d, shift)

    # ------------------------------------------------------------ writing the log
    def next_ref(self, prefix: str) -> str:
        mx = 0
        for r in self.log:
            s = r["Ref ID"].upper()
            if s.startswith(prefix.upper() + "-"):
                try:
                    mx = max(mx, int(s[len(prefix) + 1:]))
                except ValueError:
                    pass
        n = rules.next_number(mx, int(self.idmarks.get(prefix, 0)))
        return f"{prefix.upper()}-{n:04d}"

    def blank_row(self) -> dict:
        return {h: "" for h in LOG_HEADERS}

    def append_rows(self, rows: list[dict]) -> None:
        """modOvertimeEntry.AppendLogRows: Entry IDs filled here, the Payroll # from the roster."""
        for r in rows:
            r["Entry ID"] = self.entry_next
            self.entry_next += 1
            if not r.get("Payroll #") and r["Employee"] in self.people:
                r["Payroll #"] = self.people[r["Employee"]].payroll
            for h in LOG_HEADERS:
                r.setdefault(h, "")
            self.log.append(r)

    def callout(self, unit: str, d: date, shift: str, hours: float, outcomes: list[tuple], when: datetime,
                user: str, crew_sel: str = CREW_ALL) -> dict:
        """modOvertimeEntry.SubmitCalloutSession. outcomes = [(name, outcome, taken, reason)] - every named person must
        be on the list; the rows are written in CALL ORDER. Returns dict(ref, rows, charged, offers, lst, removed)."""
        lst, removed = self.callout_list(unit, d, shift, crew_sel)
        by = {o[0]: o for o in outcomes}
        for nm in by:
            if nm not in [r.name for r in lst]:
                raise ValueError(f"{nm} is not on the {unit} {d} {shift} list (left off: {removed.get(nm, 'not coded / not matched')})")
        calls = [(by[r.name][1] if r.name in by else "", by[r.name][2] if r.name in by else 0.0) for r in lst]
        offers0 = rules.callout_offers(hours, calls)
        offers, capped = [], []
        for r, o in zip(lst, offers0):
            if 0 < r.max_hours < o - 0.001:
                offers.append(r.max_hours)
                capped.append(True)
            else:
                offers.append(o)
                capped.append(False)
        ref = self.next_ref("S")
        ps = period_start(d)
        rows, charged_total = [], 0.0
        for i, r in enumerate(lst):
            if r.name not in by:
                continue
            _, outcome, taken, reason = by[r.name]
            charged = rules.callout_charge(outcome, r.grade, offers[i])
            charged_total += charged
            row = self.blank_row()
            row.update({"Ref ID": ref, "Type": "Callout", "Shift Date": d, "Unit": unit, "Shift": shift,
                        "Crew": crew_sel, "Hours": offers[i], "Call Order": i + 1, "Employee": r.name,
                        "Dept Code": r.code, "Total Hrs Then": r.roster_total if r.est else r.total,
                        "Outcome": outcome, "Note": reason, "Hours Charged": charged, "Period Start": ps,
                        "Logged By": user, "Timestamp": when, "Status": "", "Void Note": ""})
            if outcome == "Accepted":
                row["Hours"] = rules.accepted_hours(offers[i], taken)
            if outcome == "Accepted" and 0 < taken < offers0[i]:
                row["Note"] = (f"Took {hours_text(taken)} of {hours_text(offers0[i])} h (the rest dropped, uncharged"
                               + (f"; the 16 h rule allowed {hours_text(offers[i])} h" if capped[i] else "") + ")"
                               + (f"  {reason}" if reason else ""))
            elif offers[i] == 0 and reason == "":
                row["Note"] = "Not offered: a partial acceptance closed the shift"
            elif capped[i] and (reason == "" or outcome == "Accepted"):
                row["Note"] = (f"Offered {hours_text(offers[i])} of {hours_text(hours)} h (the 16 h rule: OT accepted, "
                               f"an award or a switch touches this shift)" + (f"  {reason}" if reason else ""))
            elif offers[i] < hours and (reason == "" or outcome == "Accepted"):
                row["Note"] = f"Offered {hours_text(offers[i])} of {hours_text(hours)} h (part already taken)" + (f"  {reason}" if reason else "")
            rows.append(row)
        self.append_rows(rows)
        return {"ref": ref, "rows": rows, "charged": charged_total, "offers": offers, "lst": lst, "removed": removed,
                "capped": capped}

    # ------------------------------------------------------------ postings (modPostings)
    def create_posting(self, pid: str, unit: str, d: date, shift: str, hours: float, positions: int, closes: date,
                       reason: str, posted: date, user: str, crew_sel: str = CREW_ALL) -> list[dict]:
        recs = self.load_eligible(unit, crew_filter(crew_sel), d)
        kept, removed = self.filter_not_free(recs, d, shift)
        if not kept:
            raise ValueError(f"nobody free for posting {pid} {unit} {d} {shift}")
        self.postings[pid] = {"id": pid, "status": "Open", "posted": posted, "closes": closes, "date": d, "unit": unit,
                              "shift": shift, "crew": crew_sel, "hours": hours, "positions": positions, "reason": reason,
                              "posted_by": user, "award_to": "", "removed": removed}
        self.posting_lists[pid] = [{"rank": i + 1, "name": r.name, "crew": r.crew, "code": r.code, "total": r.total,
                                    "grade": r.grade, "signed": False, "awarded": False, "result": "", "signup_hrs": 0.0,
                                    "max_hours": r.max_hours, "esign": ""} for i, r in enumerate(kept)]
        return self.posting_lists[pid]

    def _listed(self, pid: str) -> list[rules.Listed]:
        out = []
        for row in self.posting_lists[pid]:
            out.append(rules.Listed(row["rank"], row["grade"], row["signed"], row["awarded"], 0.0, "",
                                    row["result"].startswith("Exempt"), row["signup_hrs"]))
        return out

    def _posting_rows(self, pid: str, lst: list[rules.Listed], status: str, when: datetime, user: str,
                      max_rank: int = 0, out_name: str = "", out_code: str = "", out_hrs: float = 0.0,
                      out_take: float = 0.0) -> list[dict]:
        p = self.postings[pid]
        ps = period_start(p["date"])
        rows = []
        for row, l in zip(self.posting_lists[pid], lst):
            if not (l.awarded or l.signed_up or l.charged > 0):
                continue
            r = self.blank_row()
            r.update({"Ref ID": pid, "Type": "Posting", "Shift Date": p["date"], "Unit": p["unit"], "Shift": p["shift"],
                      "Crew": p["crew"], "Reason": p["reason"], "Hours": p["hours"], "Call Order": row["rank"],
                      "Employee": row["name"], "Dept Code": row["code"], "Total Hrs Then": row["total"],
                      "Hours Charged": l.charged, "Period Start": ps, "Logged By": user, "Timestamp": when,
                      "Status": "", "Void Note": ""})
            if l.awarded:
                r["Outcome"] = "Awarded"
                r["Hours"] = rules.awarded_hours(p["hours"], row["signup_hrs"])
                if rules.is_partial_signup(p["hours"], row["signup_hrs"]):
                    r["Note"] = f"Partial award: {hours_text(row['signup_hrs'])} of {hours_text(p['hours'])} h"
            elif l.charged > 0:
                r["Outcome"] = "Refused"
                r["Hours"] = rules.capped_posting_hours(p["hours"], row["max_hours"], True)
                if status == "No takers":
                    r["Note"] = "Posting - no takers"
                elif out_name:
                    r["Note"] = "Posting - didn't sign up (awarded out of department)"
                else:
                    r["Note"] = f"Posting - didn't sign up (#{row['rank']}, awarded #{max_rank})"
            else:
                r["Outcome"] = "Signed Up"
                if rules.is_partial_signup(p["hours"], row["signup_hrs"]):
                    r["Note"] = f"Signed up for {hours_text(row['signup_hrs'])} of {hours_text(p['hours'])} h"
            if row["esign"] and l.signed_up:
                r["Note"] = (r["Note"] + f"  [electronic sign-up {row['esign']}]").strip()
            rows.append(r)
        if out_name:
            r = self.blank_row()
            r.update({"Ref ID": pid, "Type": "Posting", "Shift Date": p["date"], "Unit": p["unit"], "Shift": p["shift"],
                      "Crew": p["crew"], "Reason": p["reason"], "Hours": rules.awarded_hours(p["hours"], out_take),
                      "Employee": out_name, "Dept Code": out_code, "Total Hrs Then": out_hrs, "Outcome": "Awarded",
                      "Note": "Out of department" + (f" - partial award: {hours_text(out_take)} of {hours_text(p['hours'])} h"
                                                     if rules.is_partial_signup(p["hours"], out_take) else ""),
                      "Hours Charged": 0.0, "Period Start": ps, "Logged By": user, "Timestamp": when, "Status": "", "Void Note": ""})
            rows.append(r)
        self.append_rows(rows)
        return rows

    def sign_up(self, pid: str, name: str, hours: float = 0.0, esign: str = "") -> None:
        for row in self.posting_lists[pid]:
            if row["name"].casefold() == name.casefold():
                row["signed"], row["signup_hrs"], row["esign"] = True, hours, esign
                return
        raise ValueError(f"{name} is not on posting {pid}'s list")

    def award(self, pid: str, awarded: dict, when: datetime, user: str) -> dict:
        """Submit Award: awarded = {name: signup hours (0 = whole)}. The caller ticks the sign-ups first (sign_up)."""
        p = self.postings[pid]
        for row in self.posting_lists[pid]:
            row["awarded"] = row["name"] in awarded
            if row["awarded"]:
                row["signed"] = True
                row["signup_hrs"] = awarded[row["name"]]
        lst = self._listed(pid)
        problem = rules.award_order_problem(p["hours"], p["positions"], lst)
        if problem:
            return {"problem": problem}
        lst = rules.compute_award(p["hours"], lst)
        max_rank = max(l.rank for l in lst if l.awarded)
        rows = self._posting_rows(pid, lst, "Awarded", when, user, max_rank)
        for row, l in zip(self.posting_lists[pid], lst):
            row["result"] = l.result
        p["status"] = "Awarded"
        p["award_to"] = "; ".join(n for n in awarded)
        covered, uncovered = rules.award_coverage(p["hours"], p["positions"], lst)
        return {"rows": rows, "charged": sum(l.charged for l in lst), "n_charged": sum(1 for l in lst if l.charged > 0),
                "max_rank": max_rank, "covered": covered, "uncovered": uncovered, "problem": ""}

    def no_takers(self, pid: str, when: datetime, user: str) -> dict:
        p = self.postings[pid]
        lst = rules.compute_no_takers(p["hours"], self._listed(pid))
        rows = self._posting_rows(pid, lst, "No takers", when, user)
        for row, l in zip(self.posting_lists[pid], lst):
            row["result"] = l.result
        p["status"] = "No takers"
        return {"rows": rows, "charged": sum(l.charged for l in lst), "n_charged": sum(1 for l in lst if l.charged > 0)}

    def award_outside(self, pid: str, out_name: str, when: datetime, user: str, out_take: float = 0.0) -> dict:
        p = self.postings[pid]
        lst = rules.compute_outside(p["hours"], self._listed(pid))
        crew = self.crew_on(out_name, p["date"]) if out_name in self.people else ""
        code = f"OD {crew}" if out_name in self.people else "OD"
        hrs = self.total(out_name) if out_name in self.people else 0.0
        rows = self._posting_rows(pid, lst, "Awarded", when, user, 0, out_name, code, hrs, out_take)
        for row, l in zip(self.posting_lists[pid], lst):
            row["result"] = l.result
        p["status"] = "Awarded"
        p["award_to"] = f"{out_name} (out of dept)"
        return {"rows": rows, "charged": sum(l.charged for l in lst), "n_charged": sum(1 for l in lst if l.charged > 0)}

    def cancel_posting(self, pid: str) -> None:
        for row in self.posting_lists[pid]:
            row["awarded"], row["result"] = False, ""
        self.postings[pid]["status"] = "Cancelled"

    def void_ref(self, ref: str, note: str, outcome: str = "") -> int:
        n = 0
        for r in self.log:
            if r["Ref ID"].upper() == ref.upper() and r["Status"].upper() != "VOID" and (outcome == "" or r["Outcome"] == outcome):
                r["Status"], r["Void Note"] = "VOID", note
                n += 1
        return n

    # ------------------------------------------------------------ Kronos (ImportKronosData, the v27 ledger)
    def _weighted(self, spec: dict) -> float:
        s = ";".join(f"{c}={h}" for c, h in spec.items())
        return rules.ledger_person(s, "")

    def recompute_worked(self) -> None:
        """H per person = the snapshot's codes + the ledger days along the unbroken run after it, weighted once."""
        cum = self.kronos_cum_through
        run_end = rules.ledger_run(cum, YEAR_FROM, ";".join(str(rules.serial(d)) for d in sorted(self.ledger_days)))
        self.kronos_through = run_end
        for nm in self.people:
            snap = self.kronos_snapshot.get(nm, {})
            day_parts = []
            for (who, d), spec in self.kronos_ledger.items():
                if who == nm and rules.ledger_day_in(d, cum, run_end):
                    day_parts.append(";".join(f"{c}={h}" for c, h in spec.items()))
            if not snap and not day_parts:
                self.worked[nm] = 0.0
                continue
            snap_spec = ";".join(f"{c}={h}" for c, h in snap.items())
            self.worked[nm] = rules.ledger_person(snap_spec, ";".join(day_parts))

    def import_cumulative(self, through: date, totals: dict) -> None:
        """totals = {name: {code: hours}} - the report's Totals block for roster people."""
        self.kronos_cum_through = through
        self.kronos_snapshot = {nm: dict(spec) for nm, spec in totals.items()}
        self.recompute_worked()
        for nm in self.people:
            self.hours_history.append((nm, through, self.worked[nm]))

    def import_day(self, d: date, rows: dict) -> None:
        """rows = {name: {code: hours}} - a one-day report's Totals block (a quiet day = {})."""
        for key in [k for k in self.kronos_ledger if k[1] == d]:
            del self.kronos_ledger[key]
        for nm, spec in rows.items():
            self.kronos_ledger[(nm, d)] = dict(spec)
        self.ledger_days.add(d)
        self.recompute_worked()
        for nm in self.people:
            self.hours_history.append((nm, self.kronos_through, self.worked[nm]))

    # ------------------------------------------------------------ hours as of a date (modLeave)
    def hours_as_of(self, name: str, d: date):
        hh = [(t, h) for (nm, t, h) in self.hours_history if nm == name]
        lg = [{"date": r["Shift Date"], "outcome": r["Outcome"], "status": r["Status"], "hours": float(r["Hours"] or 0),
               "charged": float(r["Hours Charged"] or 0), "night": rules.is_night_shift(r["Shift"])}
              for r in self.log if r["Employee"].casefold() == name.casefold()]
        return rules.hours_as_of(hh, lg, d, YEAR_FROM, YEAR_TO)


# ---------------------------------------------------------------- CSV helpers
def read_csv(path: Path) -> list[dict]:
    with path.open(newline="", encoding="ascii") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, headers: list[str], rows: list[list]) -> None:
    with path.open("w", newline="", encoding="ascii") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(headers)
        for r in rows:
            w.writerow(["" if v is None else v for v in r])


def pdate(s: str) -> date | None:
    s = (s or "").strip()
    return date.fromisoformat(s[:10]) if s else None


def pstamp(s: str) -> datetime | None:
    s = (s or "").strip()
    return datetime.fromisoformat(s) if s else None
