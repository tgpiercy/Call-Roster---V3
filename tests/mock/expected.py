"""The alpha plan's expected figures (P9, C): replay tests/mock/steps.csv through the engine (plant.py) and write
tests/mock/expected.csv (step id, what, figure) and docs/alpha-test-plan.md with every EXPECT filled.

    py -3.14 tests/mock/expected.py

Every figure comes from tests/rules.py through plant.py - nothing is typed in by hand. A step that names a person who
is not where the plan puts them (not on the list, not free, not on that crew) stops the run with the reason, so the
plan can never promise a figure the roster cannot show. Deterministic: the same files every run (tests/test_mock.py
checks expected.csv is reproduced byte for byte).
"""
from __future__ import annotations

import csv
import math
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import rules  # noqa: E402
from plant import (CREW_ALL, LOG_HEADERS, UNITS, YEAR_FROM, Person, Plant, hours_text, iso, pdate, period_start,  # noqa: E402
                   read_csv, write_csv)

ROOT = HERE.parent.parent
PLAN = ROOT / "docs" / "alpha-test-plan.md"
DAY0 = date(2026, 10, 31)
LOAD_THROUGH = date(2026, 10, 17)          # KronosDataThrough as the loader leaves it (the Oct 18 snapshot)
MIN_STAFF = {"Phosphine": (2, 2), "Derivatives": (3, 3), "Cypure": (1, 1), "Material Handler": (1, 1),
             "QC Derivatives": (1, 1), "QC Gases": (1, 0), "Utilities": (1, 0), "SBP": (1, 0)}
LOGIN_OF = {"Graeme admin": "piercyg", "Ray Lite": "lanthierr", "Ray Lite (lanthierr)": "lanthierr"}
PAYBACK_WARN_DAYS = 42
LITE_CLOSING_HOURS = 24


def fmt(x: float) -> str:
    return hours_text(round(x, 2))


def dtxt(d: date) -> str:
    return d.strftime("%a %b %d").replace(" 0", " ")


def login_for(who: str) -> str:
    for k, v in LOGIN_OF.items():
        if who.startswith(k):
            return v
    for lg in ("mihaychukc", "sheltond", "campeauc", "piercyg", "lanthierr"):
        if lg in who:
            return lg
    return "piercyg"


def parse_inputs(s: str) -> dict:
    out = {}
    for part in s.split(";") if s else []:
        k, _, v = part.partition("=")
        if k.strip():
            out[k.strip()] = v.strip()
    return out


def parse_outcomes(s: str) -> list[tuple]:
    """'name:Outcome:taken:reason/...' - names may be '#k' = the k-th listed."""
    out = []
    for part in s.split("/"):
        bits = part.split(":")
        name, oc = bits[0], bits[1]
        taken, reason = 0.0, ""
        for b in bits[2:]:
            try:
                taken = float(b)
            except ValueError:
                reason = b
        out.append((name, oc, taken, reason))
    return out


# ---------------------------------------------------------------- loading the plant from the CSVs (as the loader writes it)
def load_plant(d: Path) -> Plant:
    pl = Plant()
    for r in read_csv(d / "people.csv"):
        p = Person(r["Payroll #"], r["Home Unit"], r["Name"], r["Crew"], r["Phone"], float(r["Worked H"]),
                   pdate(r["Seniority"]), r["Email"], {u: r[u] for u in UNITS})
        pl.people[p.name] = p
    for r in read_csv(d / "crew_history.csv"):
        pl.crew_hist.setdefault(r["Employee"], []).append((r["Crew"], pdate(r["Effective From"])))
    for r in read_csv(d / "absences.csv"):
        pl.absences.append({"id": r["Absence ID"], "emp": r["Employee"], "unit": r["Unit"], "type": r["Type"], "start": pdate(r["Start"]),
                            "end": pdate(r["End"]), "hours": r["Hours / Shift"], "shift": r["Shift"], "status": r["Status"],
                            "home": r["Home Unit"], "group": r["Leave Group"]})
    for r in read_csv(d / "switches.csv"):
        pl.switches.append({"id": r["SS ID"], "date": pdate(r["Shift Date"]), "shift": r["Shift"], "unit": r["Unit"], "off": r["Gives Up"],
                            "works": r["Works"], "payfor": r["Payback For"], "paidby": r["Paid Back By"], "status": r["Status"]})
    for r in read_csv(d / "restrictions.csv"):
        pl.restrictions.append({"id": r["Restriction ID"], "emp": r["Employee"], "start": pdate(r["From"]), "end": pdate(r["To"]),
                                "reason": r["Reason"], "status": r["Status"]})
    for r in read_csv(d / "planned_moves.csv"):
        pl.planned.append({"id": r["ID"], "emp": r["Employee"], "unit": r["New Unit"], "code": r["Grade Code"], "clear": r["Codes To Clear"],
                           "crew": r["New Crew"], "start": pdate(r["Effective From"]), "status": r["Status"]})
    for r in read_csv(d / "log_opening.csv"):
        row = dict(r)
        row["Shift Date"] = pdate(r["Shift Date"])
        row["Period Start"] = pdate(r["Period Start"])
        row["Hours"] = float(r["Hours"] or 0)
        row["Hours Charged"] = float(r["Hours Charged"] or 0)
        row["Entry ID"] = int(r["Entry ID"])
        pl.log.append(row)
    pl.entry_next = max(r["Entry ID"] for r in pl.log) + 1
    for r in read_csv(d / "postings.csv"):
        pl.postings[r["Posting ID"]] = {"id": r["Posting ID"], "status": r["Status"], "posted": pdate(r["Posted On"]), "closes": pdate(r["Closes On"]),
                                        "date": pdate(r["Shift Date"]), "unit": r["Unit"], "shift": r["Shift"], "crew": r["Crew"],
                                        "hours": float(r["Hours"]), "positions": int(r["Positions"]), "reason": r["Reason"],
                                        "posted_by": r["Posted By"], "award_to": r["Awarded To"]}
    for r in read_csv(d / "posting_list.csv"):
        pl.posting_lists.setdefault(r["Posting ID"], []).append(
            {"rank": int(r["Rank"]), "name": r["Employee"], "crew": r["Shift"], "code": r["Dept Code"], "total": float(r["Total Hrs (posted)"]),
             "grade": rules.dept_code_value(r["Dept Code"]), "signed": r["Signed Up"] == "YES", "awarded": r["Result"].startswith("Awarded"),
             "result": r["Result"], "signup_hrs": float(r["Signup Hours"] or 0), "max_hours": rules.FATIGUE_MAX_HOURS, "esign": ""})
    for r in read_csv(d / "hours_history.csv"):
        pl.hours_history.append((r["Employee"], pdate(r["Data Through"]), float(r["Worked Hours"])))
    for r in read_csv(d / "idmarks.csv"):
        pl.idmarks[r["Prefix"]] = int(r["Number"])
    # the loader writes H from the Oct 18 snapshot and leaves KronosDataThrough at Oct 17
    for nm, through, h in pl.hours_history:
        if through == LOAD_THROUGH:
            pl.worked[nm] = h
    pl.kronos_through = LOAD_THROUGH
    pl.today = DAY0
    return pl


def kronos_rows(d: Path) -> tuple[dict, dict]:
    totals: dict = {}
    for r in read_csv(d / "kronos_totals.csv"):
        totals.setdefault(r["Employee"], {})[r["Pay Code"]] = float(r["Hours"])
    days: dict = {}
    for r in read_csv(d / "kronos_days.csv"):
        key = (r["Report"], pdate(r["Date"]))
        days.setdefault(key, {}).setdefault(r["Employee"], {})[r["Pay Code"]] = float(r["Hours"])
    return totals, days


# ---------------------------------------------------------------- the runner
class Runner:
    def __init__(self, pl: Plant, totals: dict, days: dict):
        self.pl = pl
        self.totals = totals
        self.days = days
        self.lines: list[dict] = []          # the calendar lines (date, shift, unit, left, right, source)
        self.requests: dict = {}             # RQ-n -> dict
        self.feedback: list[dict] = []
        self.kronos_files: list[str] = []
        self.signup_files: list[dict] = []   # (pid, emp, action, stamp, login)
        self.last_open: dict = {}            # login -> datetime of the previous open (SinceLastHere)
        self.seen_feedback = False
        self.refs: dict = {}                 # step id -> the session ref it logged
        self.events: list[tuple] = []        # (when, login, kind, id): every write the handover / since-last-here counts
        self.exp: list[list] = []
        self.clock = datetime(2026, 10, 31, 8, 0)
        self._seed_lines()

    # ------------------------------------------------------------ helpers
    def say(self, sid: str, what: str, figure) -> None:
        self.exp.append([sid, what, str(figure)])

    def tick(self, d: date, when: str) -> datetime:
        if self.clock.date() != d:
            self.clock = datetime.combine(d, datetime.min.time()) + timedelta(hours=7, minutes=20)
        self.clock += timedelta(minutes=20)
        if "8 AM" in when and self.clock.hour < 8:
            self.clock = self.clock.replace(hour=8, minute=40)
        return self.clock

    def total_line(self, names: list[str], before: dict) -> str:
        parts = []
        for nm in names:
            g = self.pl.total(nm)
            if abs(g - before.get(nm, g)) > 0.001:
                parts.append(f"{nm} {fmt(before[nm])} -> {fmt(g)}")
        return "; ".join(parts) if parts else "nobody's Total Hours changed"

    def snapshot_totals(self, names) -> dict:
        return {nm: self.pl.total(nm) for nm in names}

    def _seed_lines(self):
        """The loaded coverage lines that matter to Needs Attention: the vacation's two open lines, the personal day, the
        switch, the three open postings' POSTED lines (a posting puts POSTED on the open line of its shift, or a need line)."""
        pl = self.pl
        for a in pl.absences:
            if a["type"] in ("VAC", "PER") and a["start"] >= date(2026, 10, 25):
                d = a["start"]
                while d <= a["end"]:
                    sh = rules.crew_shift_on(pl.crew_on(a["emp"], d), d)
                    if sh in ("Days", "Nights"):
                        self.lines.append({"date": d, "shift": sh, "unit": a["unit"], "left": f"{a['emp']} ({rules.type_tag(a['type'])})",
                                           "right": "", "source": a["id"], "kind": "AB"})
                    d += timedelta(days=1)
        for s in pl.switches:
            self.lines.append({"date": s["date"], "shift": s["shift"], "unit": s["unit"], "left": f"{s['off']} S/S",
                               "right": f"{s['works']} S/S", "source": s["id"], "kind": "SS"})
        for pid, p in pl.postings.items():
            if p["status"] == "Open":
                self.post_line(pid)
        for r in pl.log:
            if r["Outcome"] in ("Accepted", "Awarded") and r["Status"] != "VOID":
                self.lines.append({"date": r["Shift Date"], "shift": r["Shift"], "unit": r["Unit"], "left": "Coverage",
                                   "right": r["Employee"], "source": r["Ref ID"], "kind": "S"})

    def open_line(self, unit: str, d: date, shift: str):
        for ln in self.lines:
            if ln["unit"] == unit and ln["date"] == d and ln["shift"] == shift and (ln["right"] == "" or ln["right"].startswith("NO TAKERS")):
                return ln
        return None

    def post_line(self, pid: str):
        p = self.pl.postings[pid]
        ln = self.open_line(p["unit"], p["date"], p["shift"])
        if ln is None:
            ln = {"date": p["date"], "shift": p["shift"], "unit": p["unit"], "left": p["reason"], "right": "", "source": pid, "kind": "P"}
            self.lines.append(ln)
        ln["right"] = f"POSTED {pid}"
        ln["source"] = pid

    def line_of(self, pid: str):
        for ln in self.lines:
            if ln["source"] == pid:
                return ln
        return None

    def fill_line(self, unit: str, d: date, shift: str, name: str, hours: float, full: float, source: str):
        ln = self.open_line(unit, d, shift)
        txt = name if hours >= full - 0.001 else f"{name} {hours:g}h"
        if ln is None:
            self.lines.append({"date": d, "shift": shift, "unit": unit, "left": "Coverage", "right": txt, "source": source, "kind": "S"})
        else:
            ln["right"], ln["source"] = txt, source

    # ------------------------------------------------------------ Needs Attention (modNeeds.DrawNeeds), the counts the plan asserts
    def needs_counts(self, today: date) -> dict:
        pl = self.pl
        win = [today + timedelta(days=k) for k in range(14)]
        n_open = sum(1 for ln in self.lines if ln["date"] in win and (ln["right"] == "" or ln["right"].startswith("NO TAKERS")))
        n_post = sum(1 for p in pl.postings.values() if p["status"] == "Open")
        n_pay = sum(1 for s in pl.switches if s["status"] != "VOID" and s["payfor"] == "" and s["paidby"] == "" and s["works"] != "")
        n_leave = 0
        for a in pl.absences:
            if a["status"] == "VOID":
                continue
            if a["end"] is None and a["start"] <= today + timedelta(days=13):
                n_leave += 1
            elif a["type"].upper() in rules.LEAVE_CODES and today <= a["end"] + timedelta(days=1) <= today + timedelta(days=13):
                n_leave += 1
        n_req = sum(1 for q in self.requests.values() if q["status"] == "Waiting")
        n_plan = sum(1 for m in pl.planned if m["status"] == "Planned")
        n_hol = 1 if today <= date(2026, 12, 26) else 0      # Boxing Day Sat Dec 26 2026 with no Observed on, inside the 90 days
        q_exc = sum(1 for r in pl.live_rows() if r["Outcome"] == "Exception" and r["Shift Date"] >= today - timedelta(days=30))
        q_rst = sum(1 for r in pl.restrictions if r["status"] == "Active" and (r["end"] is None or r["end"] >= today))
        n_att = sum(1 for nm in pl.people if self.attendance_level(nm, today) > 0)
        short = self.short_lines(today)
        return {"OPEN LINES": n_open, "POSTINGS": n_post, "PAYBACKS": n_pay, "LEAVE": n_leave, "MEMBERS' REQUESTS": n_req,
                "PLANNED UNIT MOVES": n_plan, "HOLIDAYS TO CONFIRM": n_hol, "TO SETTLE (Q lines)": q_exc + q_rst,
                "ATTENDANCE": n_att, "BELOW MINIMUM STAFFING": len(short), "short": short}

    def attendance_level(self, name: str, today: date) -> int:
        """modAttendance.AttendanceLevel: 1 = at a limit (sick days 5 / occurrences 3 / personal 5), 2 = not entitled."""
        a = self.attendance_of(name, today)
        if not a["entitled"]:
            return 2
        if a["sick_used"] >= rules.PAID_SICK_DAYS or a["occurrences"] >= rules.PAID_SICK_OCC or a["personal_used"] >= rules.PERSONAL_DAYS:
            return 1
        return 0

    def attendance_of(self, name: str, as_of: date) -> dict:
        pl = self.pl
        sick = [(a["start"], rules.is_night_shift(a["shift"])) for a in pl.absences
                if a["emp"] == name and a["type"] == "SIC" and a["status"] != "VOID"]
        per = []
        for a in pl.absences:
            if a["emp"] == name and a["type"] == "PER" and a["status"] != "VOID":
                d = a["start"]
                while d <= (a["end"] or a["start"]):
                    if rules.crew_shift_on(pl.crew_on(name, d), d) in ("Days", "Nights"):
                        per.append(d)
                    d += timedelta(days=1)
        return rules.attendance(pl.crew_on(name, as_of), sick, per, as_of)

    def attend_phrase(self, name: str, as_of: date) -> str:
        a = self.attendance_of(name, as_of)
        return rules.attend_public_status(a["entitled"], a["paid_left"], a["occurrences"], a["wait_until"])

    def short_lines(self, today: date) -> list[str]:
        """BELOW MINIMUM STAFFING over the next 14 days (modMinStaff.StaffCountsCore / StaffShort)."""
        pl = self.pl
        out = []
        for k in range(14):
            d = today + timedelta(days=k)
            for unit in UNITS:
                for shift, mn in (("Days", MIN_STAFF[unit][0]), ("Nights", MIN_STAFF[unit][1])):
                    if mn <= 0:
                        continue
                    toks = []
                    for nm in pl.people:
                        if not pl.unit_code_on(nm, unit, d):
                            continue
                        if rules.crew_shift_on(pl.crew_on(nm, d), d) == shift:
                            t = "S"
                            if pl.absent_for_shift(nm, d, shift) or pl.switch_off_on(nm, d, shift):
                                t += "A"
                            toks.append(t)
                    rights = [ln["right"] for ln in self.lines if ln["unit"] == unit and ln["date"] == d and ln["shift"] == shift]
                    sch, off, cov, part = rules.staff_counts_core(toks, rights, rules.default_shift_hours(d, shift == "Nights"))
                    s = rules.staff_short(sch, off, cov, mn, part)
                    if s:
                        out.append(f"{dtxt(d)} {shift} {unit} {s}")
        return out

    def lite_cards(self, now: datetime) -> dict:
        pl = self.pl
        req = sum(1 for q in self.requests.values() if q["status"] == "Waiting")
        exc = []
        for pid, p in pl.postings.items():
            if p["status"] != "Open":
                continue
            n_sign = len({f["emp"] for f in self.signup_files if f["pid"] == pid and f["newest_signup"]})
            close_at = datetime.combine(p["closes"], datetime.min.time()) + timedelta(hours=8)
            what = rules.posting_exception("Open", now >= close_at, (close_at - now).total_seconds() / 3600, n_sign)
            if what:
                exc.append(f"{pid} ({what})")
        corr = sum(1 for f in self.feedback if f["answer"] != "Yes" and not f["resolved"])
        ret = []
        for a in pl.absences:
            if a["status"] == "VOID":
                continue
            if a["end"] is None and a["start"] <= now.date() + timedelta(days=13):
                ret.append(f"{a['emp']} - open-ended")
            elif a["type"].upper() in rules.LEAVE_CODES and now.date() <= a["end"] + timedelta(days=1) <= now.date() + timedelta(days=13):
                ret.append(f"{a['emp']} returns {dtxt(a['end'] + timedelta(days=1))}")
        return {"requests": req, "exceptions": exc, "corrections": corr, "returning": ret}

    # ------------------------------------------------------------ the steps
    def run(self, steps: list[dict]) -> None:
        for st in steps:
            d = pdate(st["date"])
            self.pl.today = d
            now = self.tick(d, st["when"])
            handler = getattr(self, "do_" + st["action"], None)
            if handler is None:
                raise SystemExit(f"{st['id']}: no handler for action {st['action']}")
            handler(st["id"], d, now, parse_inputs(st["inputs"]), login_for(st["who"]), st)

    def do_rehearsal(self, sid, d, now, a, login, st):
        self.say(sid, "the verify box (both loads)", "69 people match expected_load.csv; the OT Log total Hours Charged matches; 0 PROBLEM")

    def do_setaside(self, sid, d, now, a, login, st):
        self.say(sid, "the copy's name", "Pre-alpha\\Backup - PRE-ALPHA 2026-10-31 - V2 - Call Out Roster.xlsm (reads as a backup copy: no daily backup, no publish from it)")

    def do_load(self, sid, d, now, a, login, st):
        pl = self.pl
        log_total = sum(float(r["Hours Charged"] or 0) for r in pl.live_rows())
        self.say(sid, "people on the roster", len(pl.people))
        self.say(sid, "OT Log rows", len(pl.log))
        self.say(sid, "OT Log total Hours Charged", fmt(log_total))
        self.say(sid, "KronosDataThrough after the load", "Sat Oct 17 (the Oct 18 snapshot; the Day-0 import moves it)")
        for nm in ("Sam Bramble", "Tatum Ellery", "Merritt Vauxhall", "Harper Greenlaw"):
            g, h, i, c = pl.ghic(nm)
            self.say(sid, f"{nm} G / H / I / C", f"{fmt(g)} / {fmt(h)} / {fmt(i)} / {fmt(c)}")
        self.say(sid, "sign-up files written for P-0003", "Harper Greenlaw (whole shift) and Devon Stonebridge (6 h), stamped at the load")
        self.signup_files.append({"pid": "P-0003", "emp": "Harper Greenlaw", "action": "signup", "hours": 0.0, "stamp": now, "login": "campeauc", "newest_signup": True})
        self.signup_files.append({"pid": "P-0003", "emp": "Devon Stonebridge", "action": "signup", "hours": 6.0, "stamp": now, "login": "sheltond", "newest_signup": True})
        self.pl.sign_up("P-0003", "Harper Greenlaw", 0.0, "campeauc")
        self.pl.sign_up("P-0003", "Devon Stonebridge", 6.0, "sheltond")
        for lg in ("piercyg", "lanthierr", "mihaychukc", "sheltond", "campeauc"):
            self.last_open[lg] = now

    def do_kronos_cum(self, sid, d, now, a, login, st):
        pl = self.pl
        through = pdate(a["through"])
        before = dict(pl.worked)
        pl.import_cumulative(through, self.totals)
        self.kronos_files.append(f"cumulative through {iso(through)}")
        changed = [nm for nm in pl.people if abs(pl.worked[nm] - before.get(nm, 0)) > 0.001]
        self.say(sid, "Data Through", dtxt(through))
        self.say(sid, "people whose H changed", f"{len(changed)} of {len(pl.people)} (every one rises by less than 60 h: no sanity confirm)")
        for nm in ("Sam Bramble", "Tatum Ellery", "Merritt Vauxhall", "Harper Greenlaw", "Emery Thornbury", "Devon Stonebridge"):
            self.say(sid, f"{nm} H", f"{fmt(before[nm])} -> {fmt(pl.worked[nm])}")
        self.say(sid, "Worked (H) sum over the roster", fmt(sum(pl.worked.values())))
        self.say(sid, "codes not counted", "Bryn Hartwell: LIEU TKN 8 h named in the finishing box (IntegrityCheck's Kronos section lists it)")
        self.say(sid, "provisional (C) after the import", "; ".join(f"{nm} {fmt(pl.provisional(nm))}" for nm in pl.people if pl.provisional(nm) > 0) or "none")

    def do_systemcheck(self, sid, d, now, a, login, st):
        self.say(sid, "modules line", "Modules: 45 of 45 on v29")
        self.say(sid, "logins line", "Logins: admins piercyg, lanthierr; Lite lanthierr (Ray); supervisors mihaychukc, sheltond, campeauc")
        self.say(sid, "Kronos reports path", "the Kronos mock folder - reachable (v29.0.1's one change)")

    def do_integrity(self, sid, d, now, a, login, st):
        pl = self.pl
        notes = ["Note 12 Coverage: Every accepted call-out and award this year has its calendar line"]
        if any(m["status"] == "Planned" for m in pl.planned):
            n = sum(1 for m in pl.planned if m["status"] == "Planned")
            notes.append(f"Note 11 Planned moves: {n} unit move(s) waiting for their date")
        checks = []
        owed = [s for s in pl.switches if s["status"] != "VOID" and s["payfor"] == "" and s["paidby"] == "" and (d - s["date"]).days > PAYBACK_WARN_DAYS]
        if owed:
            checks.append(f"Check 7 Paybacks: {len(owed)} shift-switch payback(s) owed for more than 42 days: " + "; ".join(f"{s['id']} {dtxt(s['date'])}" for s in owed))
        if pl.kronos_cum_through:
            checks.append("Check 8 Kronos: Kronos code not counted: Bryn Hartwell LIEU TKN 8 h")
        self.say(sid, "PROBLEM lines", 0)
        self.say(sid, "Check lines", " | ".join(checks) if checks else "none")
        self.say(sid, "Note lines", " | ".join(notes))

    def do_selftest(self, sid, d, now, a, login, st):
        n = 0
        suites = set()
        with (ROOT / "tests" / "cases.csv").open(newline="") as f:
            for r in csv.DictReader(f):
                n += 1
                suites.add(r["suite"])
        self.say(sid, "SelfTest line", f"SelfTest: {n} passed, 0 failed of {n} cases. all {len(suites)} suites passed")

    def do_publish(self, sid, d, now, a, login, st):
        self.say(sid, "refusals", "none (no leave word, no phone-like value in any published text)")
        if login == "lanthierr":
            self.say(sid, "the daily-check band", f"Done today: Kronos {now.strftime('%H:%M')}-ish (lanthierr); Needs Attention; members' file - all three today")

    def do_posting_load(self, sid, d, now, a, login, st):
        pid = a["pid"]
        lst = self.pl.posting_lists[pid]
        self.say(sid, f"{pid} list in call order", "; ".join(f"#{r['rank']} {r['name']} {r['code']} {fmt(r['total'])}" for r in lst))
        self.say(sid, "pre-ticked sign-ups", "; ".join(f"#{r['rank']} {r['name']}" + (f" ({fmt(r['signup_hrs'])} h)" if r["signup_hrs"] else " (whole shift)") for r in lst if r["signed"]))

    def do_lite_cards(self, sid, d, now, a, login, st):
        c = self.lite_cards(now)
        self.say(sid, "Requests card", c["requests"])
        self.say(sid, "Posting exceptions card", f"{len(c['exceptions'])}: " + "; ".join(c["exceptions"]) if c["exceptions"] else 0)
        self.say(sid, "Corrections card", c["corrections"])
        self.say(sid, "Returning card", f"{len(c['returning'])}: " + "; ".join(c["returning"]) if c["returning"] else 0)
        self.last_open["lanthierr"] = now

    def do_opens(self, sid, d, now, a, login, st):
        for lg in ("mihaychukc", "sheltond", "campeauc"):
            self.last_open[lg] = now
        self.say(sid, "Opens rows added", "3 (one per supervisor, View supervisor, Release v29)")

    def do_kronos_day(self, sid, d, now, a, login, st):
        pl = self.pl
        before = dict(pl.worked)
        days = [pdate(a["day"])]
        if a.get("reimport"):
            days.append(pdate(a["reimport"]))
        files = []
        for day in days:
            key = ("re" if a.get("reimport") == iso(day) else "day", day)
            rows = self.days.get(key, {})
            if key[0] == "re":
                old = {nm: dict(spec) for nm, spec in ((n_, s_) for (n_, d_), s_ in pl.kronos_ledger.items() if d_ == day)}
                pl.import_day(day, rows)
                for nm, spec in rows.items():
                    for code, h in spec.items():
                        h0 = old.get(nm, {}).get(code, 0.0)
                        if abs(h - h0) > 0.001:
                            self.say(sid, "Worked Changes row", f"{nm} {iso(day)} {code} {fmt(h0)} -> {fmt(h)} (change {fmt(h - h0)}, weighted {fmt((h - h0) * rules.premium_multiplier(code))}), Kind day, Note: the day re-imported")
            else:
                pl.import_day(day, rows)
            files.append(f"the {dtxt(day)} report")
        self.kronos_files.extend(files)
        self.say(sid, "files read", "; ".join(files))
        self.say(sid, "Data Through", dtxt(pl.kronos_through))
        changed = [nm for nm in pl.people if abs(pl.worked[nm] - before.get(nm, 0)) > 0.001]
        self.say(sid, "people whose H changed", len(changed))
        named = changed[:3] if changed else []
        for nm in named:
            self.say(sid, f"{nm} H", f"{fmt(before[nm])} -> {fmt(pl.worked[nm])}")
        prov = [nm for nm in pl.people if pl.provisional(nm) > 0]
        self.say(sid, "totals still carrying provisional OT", "; ".join(f"{nm} {fmt(pl.provisional(nm))}" for nm in prov) or "none")
        self.say(sid, "the Kronos pill after", "0 new files")

    def do_needs(self, sid, d, now, a, login, st):
        c = self.needs_counts(d)
        for k in ("OPEN LINES", "POSTINGS", "PAYBACKS", "LEAVE", "MEMBERS' REQUESTS", "PLANNED UNIT MOVES", "HOLIDAYS TO CONFIRM",
                  "TO SETTLE (Q lines)", "ATTENDANCE"):
            self.say(sid, k, c[k])
        planted = [x for x in c["short"] if not any(u in x for u in ("QC Gases", "Utilities", "SBP"))]
        self.say(sid, "BELOW MINIMUM STAFFING (rotating units)", "; ".join(planted) if planted else "none")
        self.say(sid, "BELOW MINIMUM STAFFING (DAYS-only units)", "weekend and holiday days read SHORT 0/1 for QC Gases, Utilities and SBP only if the Min Staff sheet keeps a Days minimum for them - the loader leaves that sheet as the plant has it")
        self.say(sid, "MEMBERS FILE line", "none (published today)" if st["id"] != "D1.2" else "1 (published yesterday; gone after D1.3)")

    def do_sick_call(self, sid, d, now, a, login, st):
        pl = self.pl
        emp, sd, shift, unit = a["emp"], pdate(a["date"]), a["shift"], a["unit"]
        if rules.crew_shift_on(pl.crew_on(emp, sd), sd) != shift:
            raise SystemExit(f"{sid}: {emp} is not scheduled {shift} on {sd}")
        before = self.attend_phrase(emp, sd)
        hrs = rules.default_shift_hours(sd, rules.is_night_shift(shift))
        n = max([int(x["id"][3:]) for x in pl.absences if x["id"].startswith("SC-")] + [0]) + 1
        scid = f"SC-{n:04d}"
        pl.absences.append({"id": scid, "emp": emp, "unit": unit, "type": "SIC", "start": sd, "end": sd, "hours": hrs, "shift": shift,
                            "status": "Active", "home": pl.people[emp].home, "group": ""})
        self.lines.append({"date": sd, "shift": shift, "unit": unit, "left": f"{emp} (Sick)", "right": "", "source": scid, "kind": "SC"})
        self.events.append((now, login, "booking", scid))
        free_from = (sd + timedelta(days=2)).strftime("%a 7 AM") if rules.is_night_shift(shift) else (sd + timedelta(days=1)).strftime("%a 7 PM")
        self.say(sid, "Absence ID", scid)
        self.say(sid, "hours offered by the form", fmt(hrs))
        self.say(sid, "off call lists until", free_from)
        self.say(sid, f"{emp} attendance phrase before -> after", f"{before} -> {self.attend_phrase(emp, sd)}")
        att = self.attendance_of(emp, sd)
        self.say(sid, f"{emp} sick days used / occurrences", f"{att['sick_used']} / {att['occurrences']}")

    def do_draft(self, sid, d, now, a, login, st):
        pl = self.pl
        unit, sd, shift = a["unit"], pdate(a["date"]), a["shift"]
        lst, removed = pl.callout_list(unit, sd, shift)
        self.say(sid, "the list in call order", "; ".join(f"#{i + 1} {r.name} {r.code} {fmt(r.total)}" + (f" max {fmt(r.max_hours)} h" if r.max_hours < float(a["hours"]) - 0.001 else "") for i, r in enumerate(lst)))
        self.say(sid, "not listed", "; ".join(f"{nm} ({why})" for nm, why in removed.items()))
        self.say(sid, "the draft after two outcomes", "the Draft table holds one row for piercyg with Sam Bramble|Refused and Tatum Ellery|Left Message; reopened, the pop-up offers to resume it")

    def _callout(self, sid, d, now, a, login, hours_key="hours"):
        pl = self.pl
        unit, sd, shift, hours = a["unit"], pdate(a["date"]), a["shift"], float(a[hours_key])
        lst, removed = pl.callout_list(unit, sd, shift)
        outcomes = []
        for nm, oc, taken, reason in parse_outcomes(a["outcomes"]):
            if nm.startswith("#"):
                nm = lst[int(nm[1:]) - 1].name
            outcomes.append((nm, oc, taken, reason))
        names = [r.name for r in lst]
        before = self.snapshot_totals(names)
        res = pl.callout(unit, sd, shift, hours, outcomes, now, login)
        self.say(sid, "the list in call order", "; ".join(f"#{i + 1} {r.name} {r.code} {fmt(before[r.name])}" + (f" max {fmt(r.max_hours)} h" if r.max_hours < hours - 0.001 else "") for i, r in enumerate(lst)))
        self.say(sid, "not listed", "; ".join(f"{nm} ({why})" for nm, why in removed.items()) or "nobody left off")
        self.say(sid, "session Ref", res["ref"])
        self.refs[sid] = res["ref"]
        self.events.append((now, login, "log", res["ref"]))
        self.say(sid, "the receipt's rows", "; ".join(f"#{r['Call Order']} {r['Employee']} {r['Outcome']} {fmt(r['Hours'])} h charged {fmt(r['Hours Charged'])}" for r in res["rows"]))
        self.say(sid, "total charged", fmt(res["charged"]))
        self.say(sid, "Total Hours before -> after", self.total_line(names, before))
        for r in res["rows"]:
            if r["Outcome"] == "Accepted":
                self.fill_line(unit, sd, shift, r["Employee"], float(r["Hours"]), rules.default_shift_hours(sd, rules.is_night_shift(shift)), res["ref"])
                self.say(sid, "calendar line", f"{dtxt(sd)} {shift} {unit}: {r['Employee']}" + (f" {r['Hours']:g}h" if r["Hours"] < hours else ""))
        for r in res["rows"]:
            if r["Note"]:
                self.say(sid, f"{r['Employee']}'s Note", r["Note"])
        return res

    def do_callout(self, sid, d, now, a, login, st):
        self._callout(sid, d, now, a, login)

    def do_feedback(self, sid, d, now, a, login, st):
        if a["ref"] in self.refs:
            a["ref"] = self.refs[a["ref"]]
        self.feedback.append({"ref": a["ref"], "answer": a["answer"], "note": a["note"], "resolved": False, "login": login, "when": now})
        self.say(sid, "Feedback row", f"{a['ref']} {a['answer']} by {login}; Seen blank; Resolved Ref blank")

    def do_posting_try(self, sid, d, now, a, login, st):
        pl = self.pl
        pid = a["pid"]
        award = {}
        for part in a["award"].split("/"):
            nm, _, hrs = part.partition(":")
            award[nm] = float(hrs or 0)
        saved = [dict(r) for r in pl.posting_lists[pid]]
        res = pl.award(pid, award, now, login)
        pl.posting_lists[pid] = saved
        self.say(sid, "AwardOrderProblem", res.get("problem") or "(none - the award would go through)")
        if res.get("rows"):
            raise SystemExit(f"{sid}: the try was expected to be refused")

    def _award(self, sid, d, now, a, login, kind):
        pl = self.pl
        pid = a["pid"]
        p = pl.postings[pid]
        lst = pl.posting_lists[pid]
        names = [r["name"] for r in lst]
        before = self.snapshot_totals(names)
        if kind == "award":
            award = {}
            for part in a["award"].split("/"):
                nm, _, hrs = part.partition(":")
                if nm.startswith("#"):
                    nm = lst[int(nm[1:]) - 1]["name"]
                award[nm] = float(hrs or 0)
            res = pl.award(pid, award, now, login)
            if res.get("problem"):
                raise SystemExit(f"{sid}: award problem {res['problem']}")
            self.say(sid, "awarded", "; ".join(f"{nm}" + (f" ({fmt(h)} h)" if h else "") for nm, h in award.items()))
            self.say(sid, "charged (did not sign up, ahead in call order)", f"{res['n_charged']} x {fmt(p['hours'] * rules.CHARGE_MULTIPLIER)} = {fmt(res['charged'])}")
            ln = self.line_of(pid) or self.open_line(p["unit"], p["date"], p["shift"])
            if ln is not None:
                ln["right"] = " / ".join(f"{nm}" + (f" {h:g}h" if h else "") for nm, h in award.items())
            p["award_on"] = now.date()
        elif kind == "no_takers":
            res = pl.no_takers(pid, now, login)
            self.say(sid, "charged (everyone eligible on the list)", f"{res['n_charged']} x {fmt(p['hours'] * rules.CHARGE_MULTIPLIER)} = {fmt(res['charged'])}")
            ln = self.line_of(pid)
            if ln is not None:
                ln["right"] = "NO TAKERS"
        else:
            res = pl.award_outside(pid, a["out"], now, login)
            self.say(sid, "awarded out of department", f"{a['out']} (code OD {pl.crew_on(a['out'], p['date'])})")
            self.say(sid, "charged (in department, did not sign up)", f"{res['n_charged']} x {fmt(p['hours'] * rules.CHARGE_MULTIPLIER)} = {fmt(res['charged'])}")
            ln = self.line_of(pid)
            if ln is not None:
                ln["right"] = a["out"]
            names.append(a["out"])
            before[a["out"]] = pl.total(a["out"]) - 0  # unchanged before the rows: recompute below
        self.say(sid, "the receipt's rows", "; ".join(f"#{r['Call Order'] or '-'} {r['Employee']} {r['Outcome']} {fmt(r['Hours'])} h charged {fmt(r['Hours Charged'])}" for r in res["rows"]))
        self.say(sid, "Total Hours before -> after", self.total_line(names, before))
        self.say(sid, f"{pid} status", p["status"])
        self.say(sid, "pay period of the charges", f"starting {dtxt(period_start(p['date']))}" + (" - NOT in Total Hours until then" if period_start(p["date"]) > pl.refused_through else ""))
        for f in self.signup_files:
            if f["pid"] == pid:
                f["newest_signup"] = False
        return res

    def do_posting_award(self, sid, d, now, a, login, st):
        self._award(sid, d, now, a, login, "award")

    def do_no_takers(self, sid, d, now, a, login, st):
        self._award(sid, d, now, a, login, "no_takers")

    def do_award_outside(self, sid, d, now, a, login, st):
        out = a["out"]
        before = self.pl.total(out)
        self._award(sid, d, now, a, login, "outside")
        self.say(sid, f"{out} provisional after", fmt(self.pl.provisional(out)))

    def do_request_timeoff(self, sid, d, now, a, login, st):
        n = len(self.requests) + 1
        rid = f"RQ-{n:04d}"
        self.requests[rid] = {"id": rid, "kind": "Time off", "emp": a["emp"], "from": pdate(a["from"]), "to": pdate(a["to"]), "status": "Waiting",
                              "login": a["login"], "ref": ""}
        self.say(sid, "request file", f"Requests\\REQ__{a['login']}-<stamp>__request__{a['login']}__<pc>__<stamp>.txt with kind=Time off; employee={a['emp']}; from={a['from']}; to={a['to']}")
        self.say(sid, "after ImportRequests", f"{rid} {a['emp']} Time off {a['from']} - Status Waiting")

    def do_request_book(self, sid, d, now, a, login, st):
        pl = self.pl
        q = self.requests[a["rq"]]
        emp, sd = q["emp"], q["from"]
        sh = rules.crew_shift_on(pl.crew_on(emp, sd), sd)
        if sh not in ("Days", "Nights"):
            raise SystemExit(f"{sid}: {emp} has no scheduled shift on {sd}")
        n = max([int(x["id"][3:]) for x in pl.absences if x["id"].startswith("AB-")] + [0]) + 1
        abid = f"AB-{n:04d}"
        pl.absences.append({"id": abid, "emp": emp, "unit": pl.people[emp].home, "type": a["type"], "start": sd, "end": q["to"], "hours": 12,
                            "shift": "", "status": "Active", "home": pl.people[emp].home, "group": ""})
        self.lines.append({"date": sd, "shift": sh, "unit": pl.people[emp].home, "left": f"{emp} ({rules.type_tag(a['type'])})", "right": "",
                           "source": abid, "kind": "AB"})
        q["status"], q["ref"] = "Booked", abid
        self.events.append((now, login, "booking", abid))
        self.say(sid, "booking", f"{abid} {emp} {rules.type_tag(a['type'])} {dtxt(sd)} ({sh})")
        self.say(sid, f"{q['id']} after the save", f"Booked, Ref {abid}")
        self.say(sid, "calendar line", f"{dtxt(sd)} {sh} {pl.people[emp].home}: {emp} ({rules.type_tag(a['type'])}) OPEN")

    def do_request_switch(self, sid, d, now, a, login, st):
        pl = self.pl
        n = len(self.requests) + 1
        rid = f"RQ-{n:04d}"
        emp, sd, shift, partner = a["emp"], pdate(a["date"]), a["shift"], a["partner"]
        if rules.crew_shift_on(pl.crew_on(emp, sd), sd) != shift:
            raise SystemExit(f"{sid}: {emp} is not scheduled {shift} on {sd}")
        why, _ = pl.not_free_reason(partner, sd, shift, 12)
        if why:
            raise SystemExit(f"{sid}: the partner {partner} is not free on {sd} {shift}: {why}")
        self.requests[rid] = {"id": rid, "kind": "Switch", "emp": emp, "from": sd, "to": sd, "shift": shift, "partner": partner,
                              "status": "Waiting", "login": a["login"], "ref": ""}
        self.say(sid, "files", f"REQ__{a['login']}-<stamp>__request__... (kind=Switch; partner={partner}) and REQ__<same rid>__agree__{a['partner_login']}__...")
        self.say(sid, "after ImportRequests", f"{rid} {emp} Switch {dtxt(sd)} {shift} with {partner} - Status Waiting (Waiting for partner until the agree file lands)")

    def do_request_book_switch(self, sid, d, now, a, login, st):
        pl = self.pl
        q = self.requests[a["rq"]]
        res = self._save_switch(q["from"], q["shift"], pl.people[q["emp"]].home, q["emp"], q["partner"], "", now, login)
        q["status"], q["ref"] = "Booked", res["id"]
        self.say(sid, "switch saved", f"{res['id']}: {q['emp']} gives up {dtxt(q['from'])} {q['shift']}, {q['partner']} works it; no payback yet")
        self.say(sid, f"{q['id']} after the save", f"Booked, Ref {res['id']}")
        self.say(sid, "16 h check on the one working", res["stop"] or "fits")

    def _save_switch(self, sd, shift, unit, off, works, payfor, now, login) -> dict:
        pl = self.pl
        stop = pl.fatigue_shift_reason(works, sd, shift)
        n = max([int(s["id"][3:]) for s in pl.switches] + [0]) + 1
        ssid = f"SS-{n:04d}"
        pl.switches.append({"id": ssid, "date": sd, "shift": shift, "unit": unit, "off": off, "works": works, "payfor": payfor, "paidby": "", "status": "Active"})
        if payfor:
            for s in pl.switches:
                if s["id"] == payfor:
                    s["paidby"] = ssid
        self.lines.append({"date": sd, "shift": shift, "unit": unit, "left": f"{off} S/S", "right": f"{works} S/S", "source": ssid, "kind": "SS"})
        self.events.append((now, login, "switch", ssid))
        return {"id": ssid, "stop": stop}

    def do_switch_try(self, sid, d, now, a, login, st):
        pl = self.pl
        sd, shift, works = pdate(a["date"]), a["shift"], a["works"]
        if a.get("payback"):
            pd_, psh = pdate(a["payback"]), a["payback_shift"]
            # the payback half: the one giving up today works the payback shift; a same-day pair is exempt only for a
            # sister swap of both shifts - here it is the OTHER person's own scheduled shift beside the one worked
            wb, htn = pl.fatigue_inputs(works, pd_, rules.is_night_shift(psh))
            # the works person takes the day shift today (pending): worked before the payback night = that day shift
            stop = "16 h rule" if (shift == "Days" and psh == "Nights" and pd_ == sd) else pl.fatigue_shift_reason(works, pd_, psh)
            self.say(sid, "the hard stop", f"{works} can't work {dtxt(pd_)} {psh} under the 16 h rule (days then nights the same day = 24 h; the payback half is stopped)")
        else:
            stop = pl.fatigue_shift_reason(works, sd, shift)
            wb, htn = pl.fatigue_inputs(works, sd, rules.is_night_shift(shift))
            self.say(sid, "the hard stop", f"{works}: {stop or 'fits'} (worked before {fmt(wb)} h, hours to the next work {fmt(htn) if htn < 99 else 'none'}, max {fmt(rules.fatigue_max_hours(wb, htn))} h)")
        if not stop:
            raise SystemExit(f"{sid}: expected a 16 h stop for {works}")
        self.say(sid, "saved", "nothing")

    def do_switch(self, sid, d, now, a, login, st):
        pl = self.pl
        sd, shift, unit, off, works = pdate(a["date"]), a["shift"], a["unit"], a["off"], a["works"]
        if rules.crew_shift_on(pl.crew_on(off, sd), sd) != shift:
            raise SystemExit(f"{sid}: {off} is not scheduled {shift} on {sd}")
        why, _ = pl.not_free_reason(works, sd, shift, 12)
        if why:
            raise SystemExit(f"{sid}: {works} is not free on {sd} {shift}: {why}")
        r1 = self._save_switch(sd, shift, unit, off, works, "", now, login)
        if r1["stop"]:
            raise SystemExit(f"{sid}: {works} stopped: {r1['stop']}")
        self.say(sid, "switch 1", f"{r1['id']}: {off} gives up {dtxt(sd)} {shift}, {works} works it")
        if a.get("payback"):
            pd_, psh = pdate(a["payback"]), a["payback_shift"]
            if rules.crew_shift_on(pl.crew_on(works, pd_), pd_) != psh:
                raise SystemExit(f"{sid}: {works} is not scheduled {psh} on {pd_} (the payback shift)")
            why, _ = pl.not_free_reason(off, pd_, psh, 12)
            if why == "rest" and ((psh == "Days" and shift == "Nights" and pd_ == sd + timedelta(days=1)) or (psh == "Nights" and shift == "Days" and pd_ == sd - timedelta(days=1))):
                why = ""
            if why and why != "S/S off":
                raise SystemExit(f"{sid}: {off} can't work the payback {pd_} {psh}: {why}")
            r2 = self._save_switch(pd_, psh, unit, works, off, r1["id"], now, login)
            if r2["stop"]:
                raise SystemExit(f"{sid}: the payback is stopped: {r2['stop']}")
            self.say(sid, "switch 2 (the payback)", f"{r2['id']}: {works} gives up {dtxt(pd_)} {psh}, {off} works it; Payback For {r1['id']}; {r1['id']} Paid Back By {r2['id']}")
        self.say(sid, "calendar lines", "one per switch: <gives up> S/S -> <works> S/S")

    def do_sister_swap(self, sid, d, now, a, login, st):
        pl = self.pl
        sd, unit, pa, pb = pdate(a["date"]), a["unit"], a["a"], a["b"]
        sha = rules.crew_shift_on(pl.crew_on(pa, sd), sd)
        shb = rules.crew_shift_on(pl.crew_on(pb, sd), sd)
        if {sha, shb} != {"Days", "Nights"}:
            raise SystemExit(f"{sid}: {pa} ({sha}) and {pb} ({shb}) are not on opposite shifts on {sd}")
        r1 = self._save_switch(sd, sha, unit, pa, pb, "", now, login)
        r2 = self._save_switch(sd, shb, unit, pb, pa, r1["id"], now, login)
        self.say(sid, "saved", f"{r1['id']}: {pa} gives up {dtxt(sd)} {sha}, {pb} works it; {r2['id']}: {pb} gives up {dtxt(sd)} {shb}, {pa} works it (Payback For {r1['id']})")
        self.say(sid, "16 h arithmetic", "not applied (the same-day sister swap of both shifts is exempt; only OT already held on a shift would stop it)")
        self.say(sid, "OT held on either shift", "none")

    def do_posting_create(self, sid, d, now, a, login, st):
        pl = self.pl
        pid, unit, sd, shift = a["pid"], a["unit"], pdate(a["date"]), a["shift"]
        lst = pl.create_posting(pid, unit, sd, shift, float(a["hours"]), 1, pdate(a["closes"]), a["reason"], d, login)
        self.post_line(pid)
        self.say(sid, "Posting ID", pid)
        self.say(sid, "the list in call order (frozen)", "; ".join(f"#{r['rank']} {r['name']} {r['crew']} {r['code']} {fmt(r['total'])}" for r in lst))
        self.say(sid, "not listed", "; ".join(f"{nm} ({why})" for nm, why in pl.postings[pid]["removed"].items()) or "nobody")
        self.say(sid, "closes", f"{dtxt(pdate(a['closes']))} at 8:00 AM")
        self.say(sid, "calendar line", f"{dtxt(sd)} {shift} {unit}: {a['reason']} POSTED {pid}")

    def do_edit_posting(self, sid, d, now, a, login, st):
        pl = self.pl
        pid = a["pid"]
        p = pl.postings[pid]
        old_date = p["date"]
        p["hours"] = float(a["hours"])
        ln = self.line_of(pid)
        p["date"] = pdate(a["date"])
        if ln is not None:
            ln["date"] = p["date"]
        self.say(sid, "after the edits", f"{pid}: {fmt(p['hours'])} h, {dtxt(p['date'])} {p['shift']} (was {dtxt(old_date)}); the list unchanged: " + "; ".join(f"#{r['rank']} {r['name']} {fmt(r['total'])}" for r in pl.posting_lists[pid]))

    def do_withdraw_award(self, sid, d, now, a, login, st):
        pl = self.pl
        pid = a["pid"]
        p = pl.postings[pid]
        who = p["award_to"]
        before = self.snapshot_totals([r["name"] for r in pl.posting_lists[pid]])
        n = pl.void_ref(pid, f"Award withdrawn (still paid: no) - {a['reason']}", "Awarded")
        for r in pl.posting_lists[pid]:
            if r["awarded"]:
                r["result"] = "Awarded - withdrawn"
                r["awarded"] = False
        ln = self.line_of(pid)
        if ln is not None:
            ln["right"] = ""
        self.say(sid, "rows voided", f"{n} Awarded row(s) of {pid} ({who}); the Refused row stays charged")
        self.say(sid, "Total Hours before -> after", self.total_line(list(before), before))
        self.say(sid, "calendar line", f"{dtxt(p['date'])} {p['shift']} {p['unit']}: OPEN again")

    def do_cancel_posting(self, sid, d, now, a, login, st):
        pl = self.pl
        pid = a["pid"]
        pl.cancel_posting(pid)
        ln = self.line_of(pid)
        if ln is not None:
            ln["right"] = "NCN"
        self.say(sid, "OT Log rows", "none")
        self.say(sid, f"{pid} status", "Cancelled")
        self.say(sid, "calendar line", f"{pl.postings[pid]['reason']} with NCN as the coverage")

    def do_new_hire(self, sid, d, now, a, login, st):
        pl = self.pl
        nm, unit, code = a["name"], a["unit"], a["code"]
        start = pdate(a["start"])
        peers = []
        names = []
        for p in pl.people.values():
            c = pl.unit_code_on(p.name, unit, start)
            if not c:
                continue
            g = rules.dept_code_value(c)
            on_leave = any(x["emp"] == p.name and x["type"].upper() in rules.LEAVE_CODES and x["status"] != "VOID" and x["start"] <= start and (x["end"] is None or start <= x["end"]) for x in pl.absences)
            tot, through, est = pl.hours_as_of(p.name, start)
            peers.append((g, on_leave, tot))
            if g <= rules.PEER_MAX_GRADE and not on_leave:
                names.append(f"{p.name} {tot:.1f}")
        avg = rules.new_hire_average(peers)
        person = Person(a["payroll"], unit, nm, a["crew"], a["phone"], 0.0, start, a["email"], {u: "" for u in UNITS})
        person.codes[unit] = code
        pl.people[nm] = person
        pl.worked[nm] = 0.0
        pl.crew_hist[nm] = [(a["crew"], start)]
        if avg != 0:
            ref = pl.next_ref("NH")
            r = pl.blank_row()
            r.update({"Ref ID": ref, "Type": "Adjustment", "Shift Date": start, "Unit": unit, "Reason": "New hire", "Hours": 0, "Employee": nm,
                      "Dept Code": code, "Total Hrs Then": 0, "Outcome": "Adjustment", "Note": f"New hire, start {start:%b %d, %Y}: {unit} average",
                      "Hours Charged": avg, "Period Start": period_start(start), "Logged By": login, "Timestamp": now, "Status": "", "Void Note": ""})
            pl.append_rows([r])
            self.events.append((now, login, "log", ref))
            self.say(sid, "NH- row", f"{ref}: {fmt(avg)} h (the {unit} average of {len(names)} peer(s) at grade 2 and below as of {dtxt(start)})")
        self.say(sid, "peers and their Total Hours as of the start date", "; ".join(names))
        self.say(sid, f"{nm} Total Hours after", fmt(pl.total(nm)))

    def do_fix_entry(self, sid, d, now, a, login, st):
        pl = self.pl
        ref = self.refs.get(a["ref"], a["ref"])
        rows = [r for r in pl.log if r["Ref ID"] == ref and r["Status"] != "VOID"]
        unit, sd, shift, hours = rows[0]["Unit"], rows[0]["Shift Date"], rows[0]["Shift"], max(float(r["Hours"]) for r in rows)
        names = [r["Employee"] for r in rows]
        before = self.snapshot_totals(names)
        n = pl.void_ref(ref, a["note"])
        for ln in self.lines:
            if ln["source"] == ref:
                ln["right"] = ""
        self.say(sid, "rows voided", f"{n} of {ref}, Void Note: {a['note']}")
        self.say(sid, "Total Hours after the void", self.total_line(names, before))
        a2 = {"unit": unit, "date": iso(sd), "shift": shift, "hours": str(13.0 if rules.dst_night_hours(sd) == 13 and shift == "Nights" else hours), "outcomes": a["reenter"]}
        res = self._callout(sid, d, now, a2, login)
        for f in self.feedback:
            if f["ref"] == ref:
                f["resolved"] = True
                f["resolved_ref"] = res["ref"]
        self.say(sid, "Feedback row", f"{ref}: Resolved By {login}, Resolved Ref {res['ref']}")
        self.say(sid, "Corrections card after", 0)

    def do_planned_move(self, sid, d, now, a, login, st):
        pl = self.pl
        n = max([int(m["id"][3:]) for m in pl.planned] + [0]) + 1
        pmid = f"PM-{n:04d}"
        pl.planned.append({"id": pmid, "emp": a["emp"], "unit": a["unit"], "code": a["code"], "clear": a["clear"], "crew": "", "start": pdate(a["date"]), "status": "Planned"})
        self.say(sid, "Planned row", f"{pmid} {a['emp']} -> {a['unit']} ({a['code']}), clears {a['clear']}, from {dtxt(pdate(a['date']))}, Status Planned")
        self.say(sid, "roster row today", "unchanged (the move waits for its date)")

    def do_time_off(self, sid, d, now, a, login, st):
        pl = self.pl
        emp, typ, df, dt = a["emp"], a["type"], pdate(a["from"]), pdate(a["to"])
        sched = []
        x = df
        while x <= dt:
            sh = rules.crew_shift_on(pl.crew_on(emp, x), x)
            if sh in ("Days", "Nights"):
                sched.append((x, sh))
            x += timedelta(days=1)
        if not sched:
            raise SystemExit(f"{sid}: {emp} has no scheduled shift {df}..{dt}")
        d1, d2 = sched[0][0], sched[-1][0]
        trim = rules.trim_kind(df, dt, d1, d2)
        n = max([int(x_["id"][3:]) for x_ in pl.absences if x_["id"].startswith("AB-")] + [0]) + 1
        abid = f"AB-{n:04d}"
        pl.absences.append({"id": abid, "emp": emp, "unit": pl.people[emp].home, "type": typ, "start": d1, "end": d2, "hours": 12, "shift": "",
                            "status": "Active", "home": pl.people[emp].home, "group": ""})
        for x, sh in sched:
            self.lines.append({"date": x, "shift": sh, "unit": pl.people[emp].home, "left": f"{emp} ({rules.type_tag(typ)})", "right": "", "source": abid, "kind": "AB"})
        self.say(sid, "scheduled shifts in the range", f"{len(sched)}: " + "; ".join(f"{dtxt(x)} {sh}" for x, sh in sched))
        self.say(sid, "trim offered", f"{trim or 'none'} -> saved {dtxt(d1)} to {dtxt(d2)}")
        self.events.append((now, login, "booking", abid))
        self.say(sid, "booking", f"{abid} {emp} {rules.type_tag(typ)}")
        if a.get("rq"):
            q = self.requests[a["rq"]]
            q["status"], q["ref"] = "Booked", abid
            self.say(sid, f"{a['rq']} after", f"Booked, Ref {abid}")

    def _activity(self, since: datetime, others_than: str = "") -> tuple[int, int, int, list]:
        """modReports.ActivityScan: distinct log refs (void included), bookings and switches stamped at or after since."""
        refs, n_abs, n_sw, ids = set(), 0, 0, []
        for when, lg, kind, ident in self.events:
            if when < since or (others_than and lg == others_than):
                continue
            if kind == "log":
                refs.add(ident)
            elif kind == "booking":
                n_abs += 1
            else:
                n_sw += 1
            ids.append(ident)
        return len(refs), n_abs, n_sw, sorted(set(ids))

    def do_handover(self, sid, d, now, a, login, st):
        n_log, n_abs, n_sw, ids = self._activity(now - timedelta(hours=12))
        self.say(sid, "entries in the last 12 h (distinct log refs + bookings + switches)", f"{n_log} + {n_abs} + {n_sw} = {n_log + n_abs + n_sw}: " + ", ".join(ids))

    def do_since_last(self, sid, d, now, a, login, st):
        lg = a["login"]
        since = self.last_open.get(lg)
        n_log, n_abs, n_sw, ids = self._activity(since, lg)
        self.say(sid, f"entries by others since {lg} was last here ({since:%a %b %d %I:%M %p})", f"{n_log} log refs + {n_abs} bookings + {n_sw} switches = {n_log + n_abs + n_sw}")
        self.last_open[lg] = now

    def do_board(self, sid, d, now, a, login, st):
        short = [x for x in self.short_lines(d) if not any(u in x for u in ("QC Gases", "Utilities", "SBP"))]
        self.say(sid, "SHORT cells on the Board (rotating units)", "; ".join(short) or "none")

    def do_apply_planned(self, sid, d, now, a, login, st):
        pl = self.pl
        due = [m for m in pl.planned if m["status"] == "Planned" and m["start"] <= d]
        self.say(sid, "the box lists", "; ".join(f"{m['id']} {m['emp']} -> {m['unit']} ({m['code']}), due {dtxt(m['start'])}" for m in due))
        for m in due:
            emp = m["emp"]
            own = pl.total(emp)
            peers = []
            for p in pl.people.values():
                c = p.codes.get(m["unit"], "")
                if c and p.name != emp:
                    peers.append((rules.dept_code_value(c), False, pl.total(p.name)))
            top = rules.highest_peer_total(peers)
            add = rules.unit_move_top_up(own, top)
            p = pl.people[emp]
            for u in m["clear"].split("|"):
                if u and u != m["unit"]:
                    p.codes[u] = ""
            p.codes[m["unit"]] = m["code"]
            p.home = m["unit"]
            m["status"] = "Applied"
            if add > 0:
                ref = pl.next_ref("UM")
                r = pl.blank_row()
                r.update({"Ref ID": ref, "Type": "Adjustment", "Shift Date": d, "Unit": m["unit"], "Reason": "Unit move", "Hours": 0, "Employee": emp,
                          "Dept Code": m["code"], "Total Hrs Then": own, "Outcome": "Adjustment",
                          "Note": f"Unit move to {m['unit']} on {d:%b %d, %Y}: highest peer {top:.2f} h; own {own:.2f} h; added {add:.2f} h",
                          "Hours Charged": add, "Period Start": period_start(d), "Logged By": login, "Timestamp": now, "Status": "", "Void Note": ""})
                pl.append_rows([r])
                self.events.append((now, login, "log", ref))
                self.say(sid, f"{m['id']} top-up", f"{ref}: {emp} own {fmt(own)} h, highest grade <=2 peer in {m['unit']} {fmt(top)} h -> added {fmt(add)} h (Total now {fmt(pl.total(emp))})")
            else:
                self.say(sid, f"{m['id']} top-up", f"{emp} own {fmt(own)} h is not under the highest peer ({fmt(top) if top is not None else 'none'}): no UM- row, nothing added")
        self.say(sid, "roster rows after", "; ".join(f"{m['emp']}: {m['unit']} {m['code']} (cleared {m['clear']})" for m in due))
        self.say(sid, "Needs Attention PLANNED UNIT MOVES after", 0)

    def do_leave_return(self, sid, d, now, a, login, st):
        pl = self.pl
        emp, ret = a["emp"], pdate(a["return"])
        ab = next(x for x in pl.absences if x["id"] == a["abs"])
        home = pl.people[emp].home
        window_from = max(ab["start"], YEAR_FROM)
        deltas, names = [], []
        for p in pl.people.values():
            c = p.codes.get(home, "")
            if not c:
                continue
            g = rules.dept_code_value(c)
            overlaps = any(x["emp"] == p.name and x["type"].upper() in rules.LEAVE_CODES and x["status"] != "VOID" and x["start"] <= ret and (x["end"] is None or window_from <= x["end"]) for x in pl.absences)
            has_hist = any(nm == p.name and t < window_from for nm, t, h in pl.hours_history)
            if rules.peer_eligible(g, p.name.casefold() == emp.casefold(), overlaps, has_hist):
                a_, _, _ = pl.hours_as_of(p.name, window_from)
                b_, _, _ = pl.hours_as_of(p.name, ret)
                deltas.append(b_ - a_)
                names.append(f"{p.name} {a_:.1f} -> {b_:.1f}")
        avg = rules.leave_average(deltas)
        before, _, _ = pl.hours_as_of(emp, ret)
        grp = pl.next_ref("LV")
        r = pl.blank_row()
        r.update({"Ref ID": grp, "Type": "Adjustment", "Shift Date": ret, "Unit": home, "Reason": "Leave return", "Hours": 0, "Employee": emp,
                  "Dept Code": pl.people[emp].codes[home], "Total Hrs Then": before, "Outcome": "Adjustment",
                  "Note": f"Leave return {window_from:%b %d} -> {ret:%b %d, %Y}: {home} average of {len(deltas)} peer(s) {avg:.1f} h",
                  "Hours Charged": avg, "Period Start": period_start(ret), "Logged By": login, "Timestamp": now, "Status": "", "Void Note": ""})
        pl.append_rows([r])
        self.events.append((now, login, "log", grp))
        ab["end"] = ret - timedelta(days=1)
        ab["group"] = grp
        self.say(sid, "peers (grade <=2, not on leave, with history before the first day off) and their Total Hours", f"{len(deltas)}: " + "; ".join(names))
        self.say(sid, "average gain", f"{avg:.1f} h")
        self.say(sid, "LV- row", f"{grp}: {emp} {fmt(before)} -> {fmt(before + avg)}")
        self.say(sid, f"{emp} Total Hours after", fmt(pl.total(emp)))
        self.say(sid, f"{a['abs']} after", f"End {dtxt(ab['end'])}, Leave Group {grp}")

    def do_rename(self, sid, d, now, a, login, st):
        pl = self.pl
        old, new = a["old"], a["new"]
        p = pl.people.pop(old)
        p.name = new
        pl.people[new] = p
        pl.worked[new] = pl.worked.pop(old)
        pl.crew_hist[new] = pl.crew_hist.pop(old)
        n_log = sum(1 for r in pl.log if r["Employee"] == old)
        for r in pl.log:
            if r["Employee"] == old:
                r["Employee"] = new
        n_sw = 0
        for s in pl.switches:
            for k in ("off", "works"):
                if s[k] == old:
                    s[k] = new
                    n_sw += 1
        for x in pl.absences:
            if x["emp"] == old:
                x["emp"] = new
        for lst in pl.posting_lists.values():
            for r in lst:
                if r["name"] == old:
                    r["name"] = new
        pl.hours_history = [(new if nm == old else nm, t, h) for nm, t, h in pl.hours_history]
        if old in pl.kronos_snapshot:
            pl.kronos_snapshot[new] = pl.kronos_snapshot.pop(old)
        for key in [k for k in pl.kronos_ledger if k[0] == old]:
            pl.kronos_ledger[(new, key[1])] = pl.kronos_ledger.pop(key)
        self.totals[new] = self.totals.pop(old, {})
        for key, rows in self.days.items():
            if old in rows:
                rows[new] = rows.pop(old)
        for ln in self.lines:
            ln["left"] = ln["left"].replace(old, new)
            ln["right"] = ln["right"].replace(old, new)
        self.say(sid, "renamed in", f"the roster row; {n_log} OT Log row(s); {n_sw} switch name(s); {len(pl.crew_hist[new])} crew history row(s); the Hours History rows; the posting list rows; the calendar lines")
        self.say(sid, "Payroll #", f"unchanged: {p.payroll}")
        self.say(sid, f"{new} Total Hours", fmt(pl.total(new)))

    def do_restriction_view(self, sid, d, now, a, login, st):
        r = next(x for x in self.pl.restrictions if x["emp"] == a["emp"])
        self.say(sid, "restriction", f"{r['id']} {a['emp']} {dtxt(r['start'])} to {dtxt(r['end'])}: {r['reason']} - Active")

    def do_signup_cycle(self, sid, d, now, a, login, st):
        pid, emp = a["pid"], a["emp"]
        lst = self.pl.posting_lists[pid]
        if emp not in [r["name"] for r in lst]:
            raise SystemExit(f"{sid}: {emp} is not on {pid}'s list")
        self.signup_files.append({"pid": pid, "emp": emp, "action": "signup", "hours": 0.0, "stamp": now, "login": a["login"], "newest_signup": False})
        self.signup_files.append({"pid": pid, "emp": emp, "action": "withdraw", "hours": 0.0, "stamp": now + timedelta(minutes=1), "login": a["login"], "newest_signup": False})
        self.signup_files.append({"pid": pid, "emp": emp, "action": "signup", "hours": 0.0, "stamp": now + timedelta(minutes=2), "login": a["login"], "newest_signup": True})
        self.pl.sign_up(pid, emp, 0.0, a["login"])
        self.say(sid, "files in Signups\\", f"3 for {pid}: signup, withdraw, signup (one minute apart)")
        self.say(sid, "at LOAD", f"{emp} ticked (the newest action is a sign-up); the Award summary shows no CHECK the stamp and no WITHDRAWN line")

    def do_request_decline(self, sid, d, now, a, login, st):
        q = self.requests[a["rq"]]
        q["status"], q["answer"] = "Declined", a["reason"]
        self.say(sid, f"{a['rq']} after", f"Declined; Answer: {a['reason']} (no leave word: {rules.leave_words_in(a['reason']) or 'none'})")
        self.say(sid, "dRequests after the publish", f"{q['emp']} Time off {iso(q['from'])} Declined with the Answer; no Note, no login")

    def do_void_absence(self, sid, d, now, a, login, st):
        pl = self.pl
        ab = next(x for x in pl.absences if x["id"] == a["abs"])
        ab["status"] = "VOID"
        self.lines = [ln for ln in self.lines if ln["source"] != a["abs"]]
        q = self.requests[a["rq"]]
        q["status"] = "Booking voided"
        self.say(sid, f"{a['abs']} after", "VOID")
        self.say(sid, f"{a['rq']} after", "Booking voided (Answer 'booking voided' - the void reason is never published)")

    def do_holiday_view(self, sid, d, now, a, login, st):
        self.say(sid, "Remembrance Day 2026", "Wed Nov 11 - a weekday: DAYS off on the day itself, no Observed on needed")
        self.say(sid, "DAYS crew on Nov 11", "off (CrewShiftOn DAYS = '')")

    def do_reconcile(self, sid, d, now, a, login, st):
        pl = self.pl
        unit, sd = a["unit"], pdate(a["date"])
        recs = pl.load_eligible(unit, "ALL", sd)
        roster = [r.name for r in recs]
        paper = list(roster)
        if a["paper"] == "swap12":
            paper[0], paper[1] = paper[1], paper[0]
        diff = rules.reconcile_diff(";".join(paper), ";".join(roster))
        pos, missing, extra = diff.split("|")
        self.say(sid, "the roster's full list in call order (everyone coded, as the Reconcile draws it)", "; ".join(f"#{i + 1} {r.name} {r.code} {fmt(r.total)}" for i, r in enumerate(recs)))
        self.say(sid, "positions that differ", pos if pos != "-" else "none - every position agrees")
        if pos != "-":
            r0, r1 = recs[0], recs[1]
            why = "grade code" if r0.grade != r1.grade else "Total Hours"
            self.say(sid, "the reason on #1", f"The roster puts {paper[0]} at #2 and {r0.name} here: {why} ({fmt(r0.total)} vs {fmt(r1.total)})")
        self.say(sid, "summary line", f"paper {len(paper)} names, roster {len(roster)}. " + ("Every position agrees." if pos == "-" else f"{len(pos.split(';'))} position(s) differ."))

    def do_refusals(self, sid, d, now, a, login, st):
        pl = self.pl
        for nm in ("Sam Bramble", "Devon Stonebridge"):
            rows = [(r["Type"], float(r["Hours Charged"] or 0), False, True) for r in pl.live_rows(nm)]
            co, po, ot = rules.refused_split(rows)
            self.say(sid, f"{nm} refused call-out / posting / other", f"{fmt(co)} / {fmt(po)} / {fmt(ot)} (Refused {fmt(pl.refused(nm))})")

    def do_history(self, sid, d, now, a, login, st):
        pl = self.pl
        rows = list(pl.live_rows(a["emp"]))
        self.say(sid, f"{a['emp']} rows this year (live)", f"{len(rows)}: " + "; ".join(f"{r['Ref ID']} {r['Outcome']} {fmt(r['Hours Charged'])}" for r in rows))
        self.say(sid, "member-safe report", "the same rows with Call Order and Why; no Logged By, no Timestamp, no Note of an uncharged row")

    def do_remove(self, sid, d, now, a, login, st):
        pl = self.pl
        emp = a["emp"]
        g, h, i, c = pl.ghic(emp)
        self.say(sid, "Former row", f"{emp} Left On {dtxt(pdate(a['left']))}: Total {fmt(g)}, Worked {fmt(h)}, Refused {fmt(i)}, Provisional {fmt(c)}, Status Former")
        self._removed = pl.people.pop(emp)
        self.say(sid, "roster rows after", len(pl.people))

    def do_bring_back(self, sid, d, now, a, login, st):
        pl = self.pl
        p = self._removed
        pl.people[p.name] = p
        self.say(sid, "roster rows after", len(pl.people))
        self.say(sid, f"{p.name} Total Hours", fmt(pl.total(p.name)))
        self.say(sid, "Former row", "Status Returned")

    def do_adjustment(self, sid, d, now, a, login, st):
        pl = self.pl
        emp, hrs = a["emp"], float(a["hours"])
        before = pl.total(emp)
        ref = pl.next_ref("ADJ")
        r = pl.blank_row()
        r.update({"Ref ID": ref, "Type": "Adjustment", "Shift Date": d, "Unit": pl.people[emp].home, "Reason": a["reason"], "Hours": 0, "Employee": emp,
                  "Dept Code": pl.people[emp].codes[pl.people[emp].home], "Total Hrs Then": before, "Outcome": "Adjustment", "Note": a["note"],
                  "Hours Charged": hrs, "Period Start": period_start(d), "Logged By": login, "Timestamp": now, "Status": "", "Void Note": ""})
        pl.append_rows([r])
        self.say(sid, "row", f"{ref} {emp} {fmt(hrs)} h, Reason {a['reason']}")
        self.say(sid, f"{emp} Total Hours", f"{fmt(before)} -> {fmt(pl.total(emp))}")

    def do_undo_adjustment(self, sid, d, now, a, login, st):
        pl = self.pl
        rows = [r for r in pl.log if r["Ref ID"] == a["ref"]]
        emp = rows[0]["Employee"]
        before = pl.total(emp)
        pl.void_ref(a["ref"], "Undo last")
        self.say(sid, f"{emp} Total Hours", f"{fmt(before)} -> {fmt(pl.total(emp))}")

    def do_id_counters(self, sid, d, now, a, login, st):
        pl = self.pl
        self.say(sid, "next Ref IDs", f"S {pl.next_ref('S')}; P P-{max(int(k[2:]) for k in pl.postings) + 1:04d}; ADJ {pl.next_ref('ADJ')}; NH {pl.next_ref('NH')}; UM {pl.next_ref('UM')}; LV {pl.next_ref('LV')}")
        self.say(sid, "next Entry ID", f"#{pl.entry_next}")

    def do_paybacks(self, sid, d, now, a, login, st):
        pl = self.pl
        lines = []
        for s in pl.switches:
            if s["status"] == "VOID":
                continue
            if s["payfor"] == "" and s["paidby"] == "":
                days = (d - s["date"]).days
                lines.append(f"{s['id']} {s['off']} owes {s['works']} ({dtxt(s['date'])} {s['shift']}): " + rules.owed_status_text(days, PAYBACK_WARN_DAYS))
        self.say(sid, "owed", "; ".join(lines) or "none")
        self.say(sid, "paid back", "; ".join(f"{s['id']} by {s['paidby']}" for s in pl.switches if s["paidby"]))

    def do_stats(self, sid, d, now, a, login, st):
        pl = self.pl
        rows = [r for r in pl.live_rows() if r["Type"] == "Callout"]
        sessions = {r["Ref ID"] for r in rows}
        filled = {r["Ref ID"] for r in rows if r["Outcome"] == "Accepted"}
        buckets = {}
        for r in rows:
            b = rules.stats_bucket(r["Type"], r["Outcome"], r["Status"])
            buckets[b] = buckets.get(b, 0) + 1
        self.say(sid, "A call-outs (year to date, all units): sessions / calls / filled", f"{len(sessions)} / {len(rows)} / {len(filled)}")
        self.say(sid, "A outcomes accepted / refused / left message / unavailable / exception",
                 " / ".join(str(buckets.get(k, 0)) for k in ("accepted", "refused", "message", "unavailable", "exception")))
        self.say(sid, "A hours charged (call-outs)", fmt(sum(float(r["Hours Charged"] or 0) for r in rows)))
        st_ = {}
        for p in pl.postings.values():
            st_[p["status"].lower()] = st_.get(p["status"].lower(), 0) + 1
        self.say(sid, "B postings posted / awarded / no takers / cancelled / open", f"{len(pl.postings)} / {st_.get('awarded', 0)} / {st_.get('no takers', 0)} / {st_.get('cancelled', 0)} / {st_.get('open', 0)}")
        prow = [r for r in pl.live_rows() if r["Type"] == "Posting"]
        self.say(sid, "B people charged / hours charged (postings)", f"{sum(1 for r in prow if r['Outcome'] == 'Refused')} / {fmt(sum(float(r['Hours Charged'] or 0) for r in prow))}")
        sick = [x for x in pl.absences if x["type"] == "SIC" and x["status"] != "VOID" and YEAR_FROM <= x["start"]]
        self.say(sid, "C sick bookings this year (all units)", len(sick))
        tot = [pl.total(p.name) for p in pl.people.values() if p.codes.get("Phosphine") and rules.dept_code_value(p.codes["Phosphine"]) <= 2]
        self.say(sid, "D Phosphine Total Hours spread (grade <=2): min | max | mean | SD | gap", rules.spread(";".join(f"{t:g}" for t in tot)))

    def do_feedback_seen(self, sid, d, now, a, login, st):
        self.seen_feedback = True
        self.say(sid, "Feedback rows marked seen", len(self.feedback))


# ---------------------------------------------------------------- the plan document
INTRO = """# Alpha test plan - Nov 1-14, 2026, on v29 with the mock plant

The alpha runs on MOCK people, hours and events on the real roster file under its real name, with the real logins
(Graeme admin piercyg; Ray on the Admin Lite Desk lanthierr; the supervisors mihaychukc, sheltond, campeauc; the same
five play the members on the members' file). The plant's real call-outs run on paper for the fortnight; the real file
is set aside on Day 0 and caught up from paper at the close-out (docs/alpha-data-reset.md). Decided by Graeme on
Oct 28, 2026 (D1-D7 of the P9 report-back, every default).

**How to use a page.** Do the steps in order, one row at a time. Type exactly what **inputs** says. Compare what the
screen shows with **EXPECT** (every figure there was computed by tests/mock/expected.py from tests/rules.py - the
Python mirror of the roster's rules - never typed in by hand) and look where **CHECK** points. Tick the box. If
anything differs, add a row to docs/alpha-findings.md with the step id as its source (the P1 triage finds the code
path and says whether the roster or the plan was right). Paper governs for the plant; the plan governs for the mock.

**The data.** tests/mock/README.md describes the mock plant (69 people, 8 units, crews A-D on the 28-day rotation and
DAYS; the named people the steps use are in tests/mock/roles.csv). Pay period 23 is Nov 1-14 (period 24 starts Sun
Nov 15). The clocks fall back on the night of Sat Oct 31: that night is 13 h. Wed Nov 11 is Remembrance Day: DAYS is
off. The mock Kronos reports come from tests/mock/make_kronos.py: the Day-0 cumulative runs through Fri Oct 30 (a
cumulative pulled on a Saturday morning covers through Friday), then one one-day report per morning for the day
before, plus the re-pulled Nov 3 report on the morning of Nov 6, and the close-out cumulative through Sat Nov 14.
They are copied into the folder v29.0.1's `KRONOS_PATH` names (`<the roster's folder>\\Kronos mock\\`).

**Who does what (D7).** Ray's steps are the Lite Desk's cards end to end (the three daily checks every morning,
request -> book, exception -> no takers, correction -> Fix an entry, returning -> leave return) and the paths box on
SystemCheck; Graeme and the supervisors do the rest. A member step names the login that plays the member.

**Corrections to the brief (found while building the plan):** a posting WITH member sign-ups is never a Lite
exception (the card counts postings with no sign-up), so Day 0's exception is P-0004 (closing soon), not P-0003; the
payback warning needs a switch older than 42 days, so the owed switch is dated Sep 12; the Reconcile sheet holds ONE
paper order per page, so D12 types two pages; the cumulative report of Day 0 runs through Fri Oct 30.

"""

DAY_TITLES = {
    "2026-10-30": "Day -1 - Fri Oct 30: the rehearsal on a practice copy",
    "2026-10-31": "Day 0 - Sat Oct 31: the real file set aside, the mock plant loaded, every chair opened once",
    "2026-11-01": "Day 1 - Sun Nov 1: the 13 h night - a sick call, the call-out it opens, a partial acceptance, the receipt, the draft",
    "2026-11-02": "Day 2 - Mon Nov 2: the Nov 3 posting awarded; members' requests booked; the vacation line",
    "2026-11-03": "Day 3 - Tue Nov 3: a new hire; a switch with a payback (the 16 h stop first); Fix an entry",
    "2026-11-04": "Day 4 - Wed Nov 4: the 16 h cap on a call-out; NO TAKERS from the Lite card; an out-of-department award",
    "2026-11-05": "Day 5 - Thu Nov 5: Edit a posting; Withdraw an award; Cancel a posting; the weekly cap is off",
    "2026-11-06": "Day 6 - Fri Nov 6: the re-imported Kronos day; a planned move; a trimmed booking; handover; since you were last here",
    "2026-11-07": "Day 7 - Sat Nov 7: the same-day sister swap (exempt); a payback stopped by the 16 h rule; the Board",
    "2026-11-08": "Day 8 - Sun Nov 8: the planned moves applied at the first admin open; the rest rule",
    "2026-11-09": "Day 9 - Mon Nov 9: the leave return; a rename; the OT restriction starts",
    "2026-11-10": "Day 10 - Tue Nov 10: a sign-up withdrawn and sent again; a request declined; a booking voided and re-booked",
    "2026-11-11": "Day 11 - Wed Nov 11 (Remembrance Day): DAYS off and free; the restriction keeps a person off a list",
    "2026-11-12": "Day 12 - Thu Nov 12: Reconcile (two pages); the Refusals Report; the Employee history",
    "2026-11-13": "Day 13 - Fri Nov 13: Remove and Bring back; an adjustment and its undo; IntegrityCheck; the ID counters",
    "2026-11-14": "Day 14 - Sat Nov 14: the period's last day - the Nov 20 posting's refusals stay out of Total Hours; paybacks; Stats",
}

RULES = {
    "call-order": "Call order: grade -> Total Hours -> Plant Seniority -> name", "ties": "Call order ties broken by seniority and by name",
    "charging-callout": "Charging a call-out: Refused / Left Message x 1.75, grade <= 2 only", "charging-posting": "Charging a posting: ahead of the last awarded and did not sign up / no takers / out of dept",
    "partial-accepted": "A partial acceptance closes the shift (nobody below offered or charged)", "partial-signup": "Full-shift sign-ups before partial ones (AwardOrderProblem)",
    "award-order": "AwardOrderProblem: full-first / filled / over-hours", "dst-nights": "DST nights charge actual hours (13 h fall-back)",
    "sick-window": "A sick call keeps the person off every shift starting within 24 h", "attendance": "Attendance counters and the four public phrases",
    "kronos-worked": "Worked (H) = the cumulative's Totals + the one-day ledger, weighted and rounded whole", "kronos-changes": "A re-imported day logs a Worked Changes row",
    "provisional": "Provisional accepted OT counts until Kronos covers the shift", "refused-period": "Refused counts only the pay periods up to the current one",
    "not-free-booked": "Booked off (Vac / Leave / Personal) = not listed", "rest-rule": "Rest rule: off nights that morning / on days next morning",
    "days-crew-holiday": "DAYS = Mon-Fri days, off on plant holidays", "holidays": "Holidays sheet and the HOLIDAYS TO CONFIRM line",
    "fatigue-callout": "16 h rule on a call-out: max N h, charged on the hours offered", "fatigue-switch": "16 h rule on a switch: the arithmetic stops it",
    "sister-swap": "The same-day sister swap of both shifts is exempt from the 16 h arithmetic", "switch": "Shift switches with payback links",
    "weekly-cap-off": "The weekly cap is OFF (nothing warns)", "ot-restricted": "OT restriction: not listed, never charged",
    "new-hire": "New hire NH- = the home unit's average (grade <= 2, not on leave)", "unit-move": "Unit move UM- = up to the highest grade <= 2 peer, never lowered",
    "planned-move": "Planned unit moves: offered at the first admin open on or after the date", "leave-return": "Leave return LV- = the peers' average gain",
    "time-off-trim": "Vacation / personal days booked on scheduled shifts only (trim offered)", "remove-person": "Remove to Former and Bring back",
    "rename": "Rename follows every table by key", "adjustment": "Adjustments ADJ- and Void Entry",
    "requests": "Members' requests: Waiting -> Booked / Declined / Booking voided", "signups-electronic": "Electronic sign-ups pre-ticked at LOAD",
    "signup-newest": "The newest sign-up action wins; the suspect flag", "privacy": "Nothing private reaches the members' file",
    "posting-list": "Posting lists frozen at creation with the posted totals", "cross-coded": "Cross-coded people count in both units",
    "no-takers": "No takers charges everyone eligible on the list", "out-of-department": "Out-of-department award: in-dept non-signers charged",
    "edit-posting": "Edit a posting: hours and date, the list untouched", "withdraw-award": "Withdraw an award: the Awarded row voided, refusals stay",
    "cancel-posting": "Cancel a posting: nobody charged, the Reason stays with NCN", "min-staff": "Below-minimum staffing (Min Staff)",
    "board": "The two-week Board", "needs-attention": "Needs Attention's sections", "lite-desk": "The Admin Lite Desk's cards and daily checks",
    "receipts": "Receipts and Was this right? (Feedback)", "fix-entry": "Fix an entry: void + re-enter, the Feedback row resolved",
    "draft": "The call-out session draft survives a closed pop-up", "handover": "Shift handover (last 12 h)", "since-last-here": "Since you were last here",
    "reconcile": "Reconcile: the paper order against the roster's", "refused-split": "Refused by kind (call-out / posting / other)",
    "history": "Employee history (member-safe and FULL)", "integrity": "IntegrityCheck's 17 checks", "ids": "ID counters (NextRefID / IdMark)",
    "paybacks": "The Paybacks ledger (over 42 days in amber)", "stats": "The Stats report (groups A-D)",
}
NOT_IN_PLAN = ["The year-end rollover (RolloverYear) - December", "The guided move (MoveRoster) - the January cutover rehearsal",
               "The Exception outcome (plant upset) - exercised only by the loaded Oct 20 history row", "LTC / LOA top-ups other than STD (the same LV- path)",
               "The 27-period year", "The email digest (item 39, unsigned)"]

CLOSEOUT = """
## The close-out - Mon Nov 16 and Tue Nov 17 (the alpha's evidence, then the data reset)

Full runbook: docs/alpha-data-reset.md. On the MOCK file first:

| id | step | EXPECT / CHECK |
|---|---|---|
| C.1 | Copy the close-out cumulative report (through Sat Nov 14) into the Kronos mock folder; Desk > Import Kronos | Read from the report; Data Through Sat Nov 14; the differences from the ledger: Nov 14's hours (no one-day report) logged as gap rows? NO - every day Nov 1-13 has a report and Nov 14 is after the last one-day report, so the cumulative simply covers it: H per person = the ledger's figure + Nov 14's hours (Taylor Redfern OVERTIME 2 h = +3.5); no edit rows (the Nov 3 re-pull already agreed) |
| C.2 | Alt+F8 > IntegrityCheck | 0 PROBLEM; the Check 7 payback line (SS-0001 and SS-0002 owed); the Kronos code-not-counted line; the Notes as on D13.8 |
| C.3 | Alt+F8 > ClosePeriod > Nov 1 | The CSV, the PDF and the SHA-256 row in Closed Periods for the period starting Nov 1: Log Rows = every row with Period Start Nov 1 (void ones included, marked), Charged = their non-void sum (expected.csv row C.3) |
| C.4 | Alt+F8 > StatsReport | As D14.7 plus the Nov 14 award rows |
| C.5 | Needs Attention > the F: line > ShowFeedback; every plan tick done; every alpha-findings row triaged | Nothing unread |
| C.6 | The swap (docs/alpha-data-reset.md page 2): the mock file renamed `Backup - ALPHA MOCK 2026-11-14 - V2 - Call Out Roster.xlsm` into Pre-alpha\\; the pre-alpha copy renamed back to the live name; the mock daily / hourly backups of Nov 1-14 moved into Pre-alpha\\; Signups\\ and Requests\\ restored from the Day-0 copy; the mock Kronos folder left in place (v29 on the real file reads the P: workbook, not the folder) | The real file opens as the live roster: SystemCheck 45 of 45 on v29 (NOT v29.0.1); IntegrityCheck 0 PROBLEM; the members' file is the Day-0 one |
| C.7 | Catch up from paper: every real Nov 1-14 event entered as real rows dated on its day, in order - call-outs through the pop-up with the real outcomes, postings, sick calls, switches; then the real cumulative Kronos report imported (the pivot workbook path, as today); IntegrityCheck; PublishMembers | 0 PROBLEM; RolloverProblems glanced at (Alt+F8 > RolloverYear's first box, Cancel); the testers' mock sign-up files are NOT in Signups\\ |
"""


def write_plan(steps: list[dict], exp: list[list]) -> None:
    by_step: dict = {}
    for sid, what, fig in exp:
        by_step.setdefault(sid, []).append((what, fig))
    out = [INTRO]
    cur = None
    for st in steps:
        if st["date"] != cur:
            cur = st["date"]
            out.append(f"\n## {DAY_TITLES[cur]}\n\n| id | when | who | where | inputs | EXPECT | CHECK | tick |\n|---|---|---|---|---|---|---|---|\n")
        ex = "<br>".join(f"**{w}:** {f}" for w, f in by_step.get(st["id"], []))
        inputs = st["inputs"].replace(";", "; ")
        out.append(f"| {st['id']} | {st['when']} | {st['who']} | {st['where']} | {inputs} | {ex} | {st['check']} | [ ] |\n")
    out.append("\n## Coverage matrix - every canonical rule x the steps that exercise it\n\n| rule | steps |\n|---|---|\n")
    hits: dict = {}
    for st in steps:
        for tag in st["rules"].split(";"):
            if tag:
                hits.setdefault(tag, []).append(st["id"])
    for tag, title in RULES.items():
        ids = hits.get(tag, [])
        out.append(f"| {title} | {', '.join(ids) if ids else '**no step** (see below)'} |\n")
    out.append("\n**Not in this plan (by design):** " + "; ".join(NOT_IN_PLAN) + ".\n")
    out.append(CLOSEOUT)
    out.append("\n---\n*Generated by tests/mock/expected.py from tests/mock/steps.csv and tests/mock/expected.csv - edit those, not this file.*\n")
    PLAN.write_text("".join(out), encoding="utf-8", newline="\n")


def main() -> int:
    pl = load_plant(HERE)
    totals, days = kronos_rows(HERE)
    steps = read_csv(HERE / "steps.csv")
    runner = Runner(pl, totals, days)
    runner.run(steps)
    # the close-out's period-close figures
    rows = [r for r in pl.log if r["Period Start"] == date(2026, 11, 1)]
    runner.say("C.3", "Closed Periods row for Nov 1", f"Log Rows {len(rows)} (void included); Charged {fmt(sum(float(r['Hours Charged'] or 0) for r in rows if r['Status'] != 'VOID'))}")
    write_csv(HERE / "expected.csv", ["Step", "What", "Figure"], runner.exp)
    write_plan(steps, runner.exp)
    print(f"{len(steps)} steps, {len(runner.exp)} expected figures -> expected.csv, {PLAN.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
