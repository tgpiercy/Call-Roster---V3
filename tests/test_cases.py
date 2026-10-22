"""Known-answer tests shared with the workbook's SelfTest.

cases.csv is the single source of expected answers. The same rows are
pasted into modSelfTest.bas by gen_selftest_cases.py, so the VBA and the
Python mirror are checked against identical values.

Row format: suite,case,in1,in2,in3,in4,in5,expected. No commas inside a
field (lists use '/', records use ';', fields inside a record use ':').
"""
from __future__ import annotations

import csv
from datetime import date, datetime, timedelta
from pathlib import Path

import pytest

import rules

CASES = Path(__file__).with_name("cases.csv")
YEAR_FROM = date(2025, 12, 28)      # provcount suite: this roster's year
YEAR_TO = date(2026, 12, 26)


def d(s: str) -> date:
    return date(int(s[0:4]), int(s[5:7]), int(s[8:10]))


def iso(x: date) -> str:
    return x.isoformat()


def load_cases() -> list[dict]:
    with CASES.open(newline="") as f:
        rows = list(csv.DictReader(f))
    assert rows, "cases.csv is empty"
    return rows


def parse_map(spec: str) -> dict[str, str]:
    """'name=12|21|21;name2=...' -> {name: '12|21|21'} (modSelfTest.CaseMap)."""
    out: dict[str, str] = {}
    for part in spec.split(";"):
        p = part.find("=")
        if p > 0:
            out[part[:p].strip(" ")] = part[p + 1:]
    return out


def parse_people(spec: str) -> list[rules.Emp]:
    out = []
    for rec in spec.split(";"):
        name, grade, hrs, sen = rec.split(":")
        out.append(rules.Emp(name, float(grade), float(hrs), d(sen) if sen else None))
    return out


def parse_listed(spec: str) -> list[rules.Listed]:
    out = []
    for rec in spec.split(";"):
        f = rec.split(":")                      # rank:grade:signed:awarded[:X = exempt]
        out.append(rules.Listed(int(f[0]), float(f[1]), f[2] == "Y", f[3] == "Y",
                                exempt=(len(f) > 4 and f[4] == "X")))
    return out


def parse_crew_hist(spec: str) -> list[tuple[str, date]]:
    """'C:2026-12-27;A:2027-02-01' -> [(crew, effective from), ...] in table order."""
    out = []
    for rec in spec.split(";"):
        if rec:
            crew, frm = rec.split(":")
            out.append((crew, d(frm)))
    return out


def parse_signups(spec: str) -> list[rules.Listed]:
    """'rank:signup:awarded;...' with signup = N (didn't sign up), F (whole shift) or the hours."""
    out = []
    for rec in spec.split(";"):
        rank, signup, awarded = rec.split(":")
        out.append(rules.Listed(int(rank), 1.0, signup != "N", awarded == "Y",
                                signup_hrs=0.0 if signup in ("N", "F") else float(signup)))
    return out


def fmt_num(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else repr(float(x))


def case_value(spec: str):
    """Value specs shared with modSelfTest.CaseValue: empty, err, d:date, n:number, s:text."""
    if spec == "empty":
        return None
    if spec == "err":
        return ("err",)
    if spec.startswith("d:"):
        return d(spec[2:])
    if spec.startswith("n:"):
        return float(spec[2:])
    if spec.startswith("s:"):
        return spec[2:]
    raise ValueError(spec)


def charged_text(lst: list[rules.Listed]) -> str:
    return ";".join(f"{p.rank}={fmt_num(p.charged)}" for p in sorted(lst, key=lambda p: p.rank))


def evaluate(row: dict) -> str:
    """One case -> the actual answer as text (same formats as modSelfTest)."""
    suite = row["suite"]
    i1, i2, i3, i4, i5 = (row[k] for k in ("in1", "in2", "in3", "in4", "in5"))
    if suite == "rotation":                      # date -> days crew / nights crew
        return f"{rules.days_crew(d(i1))}/{rules.nights_crew(d(i1))}"
    if suite == "crewshift":                     # code, date -> Days / Nights / off
        return rules.crew_shift_on(i1, d(i2)) or "off"
    if suite == "avail":                         # date, D|N -> crews eligible
        return rules.available_crew(d(i1), i2 == "N")
    if suite == "holiday":                       # name, year -> date
        return iso(rules.holiday_date(i1, int(i2)))
    if suite == "holoff":                        # date, fixture 'name|year|observed' -> off / on
        observed = {}
        if i2:
            nm, yy, obs = i2.split("|")
            observed[(int(yy), nm)] = d(obs)
        return "off" if rules.days_holiday_off(d(i1), observed) else "on"
    if suite == "dst":                           # year, fall|spring -> date
        return iso(rules.dst_fall_back(int(i1)) if i2 == "fall" else rules.dst_spring_forward(int(i1)))
    if suite == "dsthours":                      # date -> 13 / 11 / 12
        return fmt_num(rules.dst_night_hours(d(i1)))
    if suite == "typeddate":                     # text (_ = space), today -> yyyy-mm-dd or the hint word (v25)
        return rules.typed_date(i1.replace("_", " "), d(i2))
    if suite == "trimkind":                      # start, end, first, last -> none / start / end / both (v25)
        return rules.trim_kind(d(i1), d(i2), d(i3), d(i4)) or "none"
    if suite == "partialhrs":                    # typed hours, posting hours -> whole / hours / bad (v25)
        return rules.partial_hours_value(i1, float(i2))
    if suite == "postclosed":                    # closing date, as-of -> open / closed (v25)
        return "closed" if rules.posting_closed(d(i1), d(i2)) else "open"
    if suite == "crewsoff":                      # date -> the crews off that day, e.g. C/D (v25)
        return rules.crews_off(d(i1))
    if suite == "shifthours":                    # date, day|night -> the default hours (v25)
        return fmt_num(rules.default_shift_hours(d(i1), i2 == "night"))
    if suite == "shiftdate":                     # v26.5.5: date, clock time hh:mm -> the date a shift logged then belongs to
        dd = d(i1)
        return rules.default_shift_date(datetime(dd.year, dd.month, dd.day, int(i2[0:2]), int(i2[3:5]))).isoformat()
    if suite == "accepted":                      # offer, taken -> hours accepted (v25)
        return fmt_num(rules.accepted_hours(float(i1), float(i2)))
    if suite == "callorder":                     # people -> names in order
        return "/".join(e.name for e in rules.call_order(parse_people(i1)))
    if suite == "listinsert":                    # frozen list in its order, newcomer -> index 1..n+1
        return str(rules.list_insert_index(parse_people(i1), parse_people(i2)[0]))
    if suite == "callout":                       # outcome, grade, hours -> charged
        return fmt_num(rules.callout_charge(i1, float(i2), float(i3)))
    if suite == "posting":                       # mode, hours, list -> rank=charged;...
        lst = parse_listed(i3)
        hrs = float(i2)
        if i1 == "award":
            rules.compute_award(hrs, lst)
        elif i1 == "notakers":
            rules.compute_no_takers(hrs, lst)
        elif i1 == "outside":
            rules.compute_outside(hrs, lst)
        elif i1 == "cancel":
            rules.compute_cancel(lst)
        else:
            raise ValueError(i1)
        return charged_text(lst)
    if suite == "eligibility":                   # code, date, shift, fixture -> reason
        flags = {}
        if i4.startswith("off:"):
            flags["booked_off"] = i4[4:]
        elif i4 == "ss":
            flags["ss_working"] = True
        elif i4 == "ssoff":
            flags["ss_gave_up"] = True
        elif i4 == "ssrest":
            flags["ss_rest"] = True
        elif i4 == "ssgaveup":
            flags["ss_gave_up_adjoining"] = True
        elif i4:
            raise ValueError(i4)
        return rules.not_free_reason(i1, d(i2), i3, **flags) or "free"
    if suite == "prov":                          # shift date, D|N, through -> covered / counts
        return "covered" if rules.provisional_covered(d(i1), i2 == "N", d(i3)) else "counts"
    if suite == "provcount":                     # outcome, status, shift date, D|N, through -> counts / no
        return "counts" if rules.provisional_counts(i1, i2, d(i3), i4 == "N", d(i5), YEAR_FROM, YEAR_TO) else "no"
    if suite == "sick":                          # sick date, D|N, candidate date, D|N -> blocked / free
        return "blocked" if rules.sick_blocks(d(i1), i2 == "N", d(i3), i4 == "N") else "free"
    if suite == "hoursasof":                     # hh 'through:hours;...', log 'date:outcome:status:hours:charged:D|N;...', as-of -> total|through|est
        hh = [(d(p.split(":")[0]), float(p.split(":")[1])) for p in i1.split(";") if p]
        lg = []
        for rec in (p for p in i2.split(";") if p):
            dt, outcome, status, hrs, chg, sh = rec.split(":")
            lg.append(dict(date=d(dt), outcome=outcome, status=status, hours=float(hrs), charged=float(chg), night=(sh == "N")))
        total, through, est = rules.hours_as_of(hh, lg, d(i3), YEAR_FROM, YEAR_TO)
        return f"{fmt_num(total)}|{iso(through) if through else 'none'}|{'yes' if est else 'no'}"
    if suite == "leavepeer":                     # grade, returner Y/N, overlapping leave Y/N, history before Y/N -> peer / excluded
        return "peer" if rules.peer_eligible(float(i1), i2 == "Y", i3 == "Y", i4 == "Y") else "excluded"
    if suite == "leaveavg":                      # 'grade:returner:onleave:hist:delta;...' -> average of the eligible
        deltas = []
        for rec in (p for p in i1.split(";") if p):
            g, r, o, h, delta = rec.split(":")
            if rules.peer_eligible(float(g), r == "Y", o == "Y", h == "Y"):
                deltas.append(float(delta))
        return fmt_num(rules.leave_average(deltas))
    if suite == "attend":                        # crew, sick 'date:D|N;...', personal 'date;...', as-of -> used/paid/LEFT/occ/per/ENT|WAIT date/PAID|UNPAID|-
        sick = [(d(p.split(":")[0]), p.split(":")[1] == "N") for p in i2.split(";") if p]
        per = [d(p) for p in i3.split(";") if p]
        a = rules.attendance(i1, sick, per, d(i4))
        ent = "ENT" if a["entitled"] else "WAIT " + iso(a["wait_until"])
        last = "-" if a["last_sick"] is None else ("PAID" if a["last_paid"] else "UNPAID")
        return f"{a['sick_used']}/{a['paid_used']}/{a['paid_left']}/{a['occurrences']}/{a['personal_used']}/{ent}/{last}"
    if suite == "year":                          # anchor, year -> first period start|periods|roster year
        st = rules.first_period_start(d(i1), int(i2))
        return f"{iso(st)}|{rules.periods_in_year(d(i1), int(i2))}|{rules.roster_year_of(st)}"
    if suite == "nextid":                        # highest in the log, mark -> next number
        return str(rules.next_number(int(i1), int(i2)))
    if suite == "cell":                          # F|V, LOG|OTHER|'', value spec -> ok / code
        text = {"LOG": "=SUMIFS(tblOTLog[x])", "OTHER": "=SUM(A1:A2)"}.get(i2, "")
        return rules.period_cell_code(i1 == "F", text, case_value(i3)) or "ok"
    if suite == "dups":                          # 'a;b;a' -> 'a x2' / none
        return rules.duplicates_in(i1.split(";")) or "none"
    if suite == "archive":                       # period start spec, cut date -> archive / keep
        return "archive" if rules.goes_to_archive(case_value(i1), d(i2)) else "keep"
    if suite == "privacy":                       # one published value -> ok / email / phone / number
        return rules.private_value_kind(case_value(i1)) or "ok"
    if suite == "attendpriv":                    # count|status, value -> ok / stop
        if i1 == "count":
            return "ok" if rules.attend_count_ok(case_value(i2)) else "stop"
        return "ok" if rules.attend_status_ok(i2[2:]) else "stop"
    if suite == "leavewords":                    # v24.3.1: published free text -> the offending word / ok
        return rules.leave_words_in(i1) or "ok"
    if suite == "paytext":                       # a cell value -> the Payroll # as text / blank
        return rules.payroll_text(case_value(i1)) or "blank"
    if suite == "payhdr":                        # header text -> yes / no
        return "yes" if rules.is_payroll_header(i1) else "no"
    if suite == "krhdr":                         # v24.3.1: heading cell, wanted heading -> yes / no
        return "yes" if rules.is_header_match(i1, i2) else "no"
    if suite == "payroll":                       # roster number, in Kronos, names agree -> name / number / mismatch
        return rules.payroll_decision(i1, i2 == "Y", i3 == "Y")
    if suite == "kmult":                         # v27: a Kronos pay code -> its multiplier / none
        km = rules.premium_multiplier(i1)
        return "none" if km < 0 else fmt_num(km)
    if suite == "kweight":                       # v27: "code=hours;code=hours" -> the weighted whole-hour total | codes not counted
        return rules.weighted_total(i1)
    if suite == "kperiod":                       # v27: the report's Time Period line -> yyyy-mm-dd|yyyy-mm-dd or the problem
        return rules.time_period(i1)
    if suite == "kkind":                         # v27: from, to, the roster year's first and last day -> day / cumulative / range / outside
        return rules.report_kind(d(i1), d(i2), d(i3), d(i4))
    if suite == "firstref":                      # v27 (item 8): a block of text -> its first Ref ID, or '-' for none
        return rules.first_ref(i1) or "-"
    if suite == "reconcile":                     # v27: the paper's names (;), the roster's names (;) -> positions that differ|not listed|not on paper
        return rules.reconcile_diff(i1, i2)
    if suite == "squarter":                      # v27 (Stats): a date, the roster year's Jan 1 -> Q1..Q4
        return rules.stats_quarter(d(i1), d(i2))
    if suite == "sbucket":                       # v27 (Stats): Type, Outcome, Status -> the bucket
        return rules.stats_bucket(i1, i2, i3)
    if suite == "sband":                         # v27 (Stats): an hour -> the band
        return rules.time_of_day_band(int(i1))
    if suite == "sspread":                       # v27 (Stats): "h;h;h" -> min|max|mean|sd|gap
        return rules.spread(i1)
    if suite == "bkeep":                         # v27 (backups): kind, file date, today, period start Y/N -> keep / drop
        return "keep" if rules.backup_keep(i1, d(i2), d(i3), i4 == "Y") else "drop"
    if suite == "bhourly":                       # v27 (backups): file name -> yyyy-mm-dd hh / none
        bh = rules.hourly_backup_date(i1)
        return bh.strftime("%Y-%m-%d %H") if bh else "none"
    if suite == "dryname":                       # file name -> yes / no
        return "yes" if rules.is_dry_run_name(i1) else "no"
    if suite == "backupname":                    # v24.3.1: file name -> yes / no (a backup, FINAL or pre-change copy)
        return "yes" if rules.is_backup_copy_name(i1) else "no"
    if suite == "viewmode":                      # v25.3: login, ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS (v25.5; ; for ,) -> desk / admin / supervisor / user
        return rules.view_mode_for(i1, i2.replace(";", ","), i3.replace(";", ","), i4.replace(";", ","))
    if suite == "backupdate":                    # v25: file name -> the date a daily backup carries / none
        bd = rules.backup_file_date(i1)
        return iso(bd) if bd else "none"
    if suite == "deptcode":                      # v25: grade code (_ = space) -> numeric part (999 = none)
        return fmt_num(rules.dept_code_value(i1.replace("_", " ")))
    if suite == "hourstext":                     # v25: hours -> the text every message shows
        return rules.hours_text(float(i1))
    if suite == "payback":                       # v26: days owed, warn days -> the Needs Attention status
        return rules.owed_status_text(int(i1), int(i2))
    if suite == "idmark":                        # v26 (I8): Ref IDs joined by | -> highest number per prefix, or none
        return rules.ref_prefix_max_text(i1) or "none"
    if suite == "refthrough":                    # v26.5: period start, the current period's start -> counts / waits
        return "counts" if rules.refused_counts(d(i1), d(i2)) else "waits"
    if suite == "minstaff":                      # v26: scheduled, off, covered, minimum -> '' (shown as ok) / SHORT staffed/minimum
        return rules.staff_short(int(i1), int(i2), int(i3), int(i4)) or "ok"
    if suite == "practice":                      # v26.5.5: a file name -> the date and time a practice copy carries / none
        pr = rules.practice_copy_date(i1)
        return pr.strftime("%Y-%m-%d %H:%M") if pr else "none"
    if suite == "nextstep":                      # v26.5.5: accepted, partial hours, left messages, admin Y/N, admin name -> the last line (; for ,)
        return rules.next_step_text(int(i1), float(i2 or 0), int(i3), i4 == "Y", i5).replace(",", ";")
    if suite == "fatblock":                      # v26.5.5: hours on the four shifts around the offer -> block|hours to next
        wb, htn = rules.fatigue_block_core(float(i1), float(i2), float(i3), float(i4))
        return fmt_num(wb) + "|" + fmt_num(htn)
    if suite == "postcap":                       # v26.5.5: posting hours, the 16 h max, max known Y/N -> the hours charged on
        return fmt_num(rules.capped_posting_hours(float(i1), float(i2), i3 == "Y"))
    if suite == "fatigue":                       # v26.5.5: worked before, hours offered, hours to the next work -> ok / 16 h rule | the max
        mx = rules.fatigue_max_hours(float(i1), float(i3))
        return (rules.fatigue_problem(float(i1), float(i2), float(i3)) or "ok") + "|" + (str(int(mx)) if mx == int(mx) else str(mx))
    if suite == "chargewhy":                     # v26: type, outcome, call order, 'hours|charged', code -> the sentence
        hc = i4.split("|")
        return rules.charge_why(i1, i2, int(i3 or 0), float(hc[0] or 0), float(hc[1] or 0), i5).replace(",", ";")   # ; for ,
    if suite == "prechange":                     # v26: file name, roster file name -> the date a pre-change copy carries / none
        pd = rules.pre_change_copy_date(i1, i2)
        return iso(pd) if pd else "none"
    if suite == "typetag":                       # v25: absence code -> calendar tag / what members get
        return rules.type_tag(i1) + "/" + rules.public_type(i1)
    if suite == "planid":                        # v25: an Exempt note -> the planned move's ID / none
        return rules.plan_id_in(i1) or "none"
    if suite == "namekey":                       # v25: a name (_ = space) -> the match key (, shown as /)
        return rules.name_key(i1.replace("_", " ")).replace(",", "/")
    if suite == "barlist":                       # v25: in|trim, list (_ = space), item -> yes / no, or the cleaned list
        if i1 == "in":
            return "yes" if rules.in_bar_list(i2.replace("_", " "), i3.replace("_", " ")) else "no"
        return rules.trim_bar_list(i2.replace("_", " ")) or "empty"
    if suite == "matchname":                     # v25: text (_ = space), roster 'Last/First|...' -> name / ? / none
        roster = i2.replace("/", ", ").replace("_", " ").split("|")
        return rules.match_import_name(i1.replace("_", " "), roster).replace(", ", "/") or "none"
    if suite == "attendstatus":                  # v25: entitled Y|N, paid left, occurrences, wait-until -> the phrase (, as .)
        return rules.attend_public_status(i1 == "Y", int(i2), int(i3), d(i4) if i4 else None).replace(",", ".")
    if suite == "totalsdiff":                    # v25: before, after 'name=G|I|periods;...', skip -> the lines squeezed to one
        s = rules.totals_difference(parse_map(i1), parse_map(i2), i3)
        s = s.replace("\n  ", ";").replace("\n", "").strip(" ").replace(", ", "/")
        return s or "same"
    if suite == "snapyear":                      # through, snapshot year (0 = none), roster year, year from -> through to use
        return iso(rules.snapshot_through_for_year(d(i1), int(i2), int(i3), d(i4)))
    if suite == "signups":                       # 'name:S|W:hours:yyyy-mm-dd-hh;...', closing 'yyyy-mm-dd-hh' -> name=F|hours[L];...
        def stamp(s: str) -> datetime:
            return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]))
        rows = [(f[0], "signup" if f[1] == "S" else "withdraw", float(f[2]), stamp(f[3]))
                for f in (r.split(":") for r in i1.split(";") if r)]
        red = rules.reduce_signups(rows, stamp(i2) if i2 else None)
        return ";".join(f"{n}={'F' if h == 0 else fmt_num(h)}{'L' if late else ''}" for n, h, late, _s, _m in red) or "none"
    if suite == "sgsuspect":                     # v27 (item 14): stamp, file time (yyyy-mm-dd-hh-nn; 0 = unknown) -> suspect / ok
        def stamp_min(s: str):
            s = s.strip()
            if len(s) < 16:
                return None
            return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]), int(s[14:16]))
        return "suspect" if rules.signup_suspect(stamp_min(i1), stamp_min(i2), i3.upper() == "Y") else "ok"
    if suite == "sgmix":                         # v27 (item 14): 'name:S|W:hours:yyyy-mm-dd-hh:login;...' -> name[!];... / none
        def stamp3(s: str) -> datetime:
            return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]))
        rows5 = [(f[0], "signup" if f[1] == "S" else "withdraw", float(f[2]), stamp3(f[3]), f[4] if len(f) > 4 else "")
                 for f in (r.split(":") for r in i1.split(";") if r)]
        red = rules.reduce_signups(rows5, None)
        return ";".join(f"{n}{'!' if m else ''}" for n, _h, _l, _s, m in red) or "none"
    if suite == "withdrawn":                     # v26 (M3): 'name:S|W:hours:yyyy-mm-dd-hh:login;...' -> name=login[!];... / none
        def stamp2(s: str) -> datetime:
            return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]))
        rows5 = [(f[0], "signup" if f[1] == "S" else "withdraw", float(f[2]), stamp2(f[3]), f[4] if len(f) > 4 else "")
                 for f in (r.split(":") for r in i1.split(";") if r)]
        wd = rules.withdrawn_signups(rows5)
        return ";".join(f"{n}={lg}{'!' if m else ''}" for n, lg, _st, m in wd) or "none"
    if suite == "calloffer":                     # hours, 'outcome:taken;...' (grade 1) -> offered=charged;...
        calls = [(f[0], float(f[1])) for f in (r.split(":") for r in i2.split(";") if r)]
        offers = rules.callout_offers(float(i1), calls)
        return ";".join(f"{fmt_num(o)}={fmt_num(rules.callout_charge(c[0], 1, o))}" for o, c in zip(offers, calls))
    if suite == "attendocc":                     # crew, sick, personal, as-of -> first..last:days:paid;... / none
        sick = [(d(p.split(":")[0]), p.split(":")[1] == "N") for p in i2.split(";") if p]
        per = [d(p) for p in i3.split(";") if p]
        a = rules.attendance(i1, sick, per, d(i4))
        return ";".join(f"{iso(o[0])}..{iso(o[1])}:{o[2]}:{o[3]}" for o in a["occs"]) or "none"
    if suite == "refsplit":                      # 'type:hours:V|A:Y|N;...' -> callout|posting|other
        rows = [(f[0], float(f[1]), f[2] == "V", f[3] == "Y") for f in (r.split(":") for r in i1.split(";") if r)]
        c, p_, o = rules.refused_split(rows)
        return f"{fmt_num(c)}|{fmt_num(p_)}|{fmt_num(o)}"
    if suite == "awardorder":                    # hours, positions, sign-ups -> ok / the problem
        return rules.award_order_problem(float(i1), int(i2), parse_signups(i3)) or "ok"
    if suite == "awardcover":                    # hours, positions, sign-ups -> covered|uncovered
        cov, unc = rules.award_coverage(float(i1), int(i2), parse_signups(i3))
        return f"{fmt_num(cov)}|{fmt_num(unc)}"
    if suite == "newhire":                       # peers 'grade:onleave:total;...' -> starting hours
        peers = [(float(g), lv == "Y", float(tot)) for g, lv, tot in (r.split(":") for r in i1.split(";") if r)]
        return fmt_num(rules.new_hire_average(peers))
    if suite == "unitmove":                      # own total, peers 'grade:mover:total;...' -> hours added
        peers = [(float(g), mv == "Y", float(tot)) for g, mv, tot in (r.split(":") for r in i2.split(";") if r)]
        return fmt_num(rules.unit_move_top_up(float(i1), rules.highest_peer_total(peers)))
    if suite == "plancode":                      # plans 'unit:code:clear:from:status[:other|lower];...', unit, date, roster code -> code / none
        plans = []
        for r in (x for x in i1.split(";") if x):
            p = r.split(":")
            who = "Test Person"
            if len(p) > 5 and p[5] == "other":
                who = "Someone Else"
            if len(p) > 5 and p[5] == "lower":
                who = "test person"
            plans.append((who, p[0], p[1], p[2], d(p[3]), p[4]))
        return rules.planned_code(plans, "Test Person", i2, d(i3), i4) or "none"
    if suite == "planhours":                     # own total, others 'grade:mover:total;...' -> the Total the mover is listed at
        peers = [(float(g), mv == "Y", float(tot)) for g, mv, tot in (r.split(":") for r in i2.split(";") if r)]
        return fmt_num(rules.planned_list_hours(float(i1), peers))
    if suite == "renamelist":                    # text, old name, new name -> text
        return rules.rename_in_list(i1, i2, i3)
    if suite == "nametoken":                     # text, old short name, new short name -> text
        return rules.replace_name_token(i1, i2, i3)
    if suite == "crewon":                        # history 'crew:from;...', date, fallback -> crew / none
        return rules.crew_on(parse_crew_hist(i1), d(i2), i3) or "none"
    if suite == "crewsched":                     # history, date -> Days / Nights / off on that date
        return rules.crew_shift_on(rules.crew_on(parse_crew_hist(i1), d(i2)), d(i2)) or "off"
    if suite == "eligcrew":                      # history, date, shift -> scheduled / rest / free
        h = parse_crew_hist(i1)
        one = timedelta(days=1)
        return rules.rotation_free_reason3(rules.crew_on(h, d(i2) - one), rules.crew_on(h, d(i2)),
                                           rules.crew_on(h, d(i2) + one), d(i2), i3) or "free"
    if suite == "merge":                         # 'date;date;...' -> 'start..end;...'
        runs = rules.merge_runs([d(p) for p in i1.split(";") if p])
        return ";".join(f"{iso(a)}..{iso(b)}" for a, b in runs)
    raise ValueError(f"unknown suite {suite}")


@pytest.mark.parametrize("row", load_cases(), ids=lambda r: f"{r['suite']}:{r['case']}")
def test_case(row):
    assert evaluate(row) == row["expected"], f"{row['suite']} {row['case']}: {row}"


def test_known_facts():
    """A few anchors the brief and the plant schedule state outright."""
    assert rules.days_crew(date(2025, 12, 28)) == "B"
    assert rules.nights_crew(date(2025, 12, 28)) == "A"
    assert rules.dst_fall_back(2026) == date(2026, 10, 31)
    assert rules.dst_spring_forward(2027) == date(2027, 3, 13)
    assert rules.holiday_date("Thanksgiving", 2026) == date(2026, 10, 12)
    assert rules.callout_charge("Refused", 1, 13) == 22.75
