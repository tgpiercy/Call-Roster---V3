"""The contract between the master's publish and the members' file (v27, Members M4 item 38, review C2).

The members' workbook holds no rules: it reads the tables the master writes, by column. These tests tie the two
sides together so a column moved on one side fails here instead of in a member's hands:
- every data sheet the master writes (the DS_* constants) is one the members' template makes, and back;
- the header row each Build* function writes has, at the column the members' code reads, the name it expects;
- the members' roster column constants (RC_*) sit on the master's roster headers;
- every key the master's sign-up reader understands is one the members' file writes.
(The brief asked for the members' code to read every table through ColOf(v, header); it still reads by column, so
this test is the contract instead - a moved column fails pytest before any install.)
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PUB = (ROOT / "modules" / "modPublish.bas").read_text(encoding="ascii")
SIGN = (ROOT / "modules" / "modSignups.bas").read_text(encoding="ascii")
REQ = (ROOT / "modules" / "modRequests.bas").read_text(encoding="ascii")      # v28 (M5): members' requests
MEM = (ROOT / "member" / "modMember.bas").read_text(encoding="ascii")


def proc_body(text: str, name: str) -> str:
    m = re.search(r"(?:Private|Public) (?:Function|Sub) " + re.escape(name) + r"\b.*?\nEnd (?:Function|Sub)", text, re.S)
    assert m, f"{name} not found"
    return m.group(0)


def headers(text: str, proc: str, var: str) -> list[str]:
    """The header row a Build* function writes into var(1, k)."""
    body = proc_body(text, proc)
    cells = {int(k): v for k, v in re.findall(re.escape(var) + r"\(1, (\d+)\) = \"([^\"]*)\"", body)}
    assert cells, f"{proc}: no {var}(1, k) header cells"
    return [cells[k] for k in sorted(cells)]


def test_the_data_sheets_are_one_list_on_both_sides():
    written = set(re.findall(r'Public Const DS_\w+ As String = "(\w+)"', PUB))
    m = re.search(r'For Each nm In Array\(("dMeta".*?)\)', MEM)
    assert m, "BuildMemberTemplate's data-sheet list not found"
    made = set(re.findall(r'"(\w+)"', m.group(1)))
    assert written == made, f"master writes {sorted(written - made)} the template doesn't make; template makes {sorted(made - written)} the master doesn't write"


# (table, the Build* function, its array, {column: the header the members' code expects there})
CONTRACT = [
    ("dRoster", "BuildRoster", "out", {1: "Name", 2: "Crew", 3: "Home Unit", 4: "Total", 5: "Worked", 6: "Refused", 7: "Provisional",
                                       8: "Refused call-out", 9: "Refused posting", 10: "Refused other"}),
    ("dNext", "BuildNext", "out", {1: "Unit", 2: "Date", 3: "Shift", 4: "Rank", 5: "Name", 6: "Total", 7: "Provisional", 8: "Code", 9: "EstHrs"}),
    ("dNotFree", "BuildNext", "nfOut", {1: "Unit", 2: "Date", 3: "Shift", 4: "Name", 5: "Reason"}),
    ("dPost", "BuildPostings", "p", {1: "Posting ID", 2: "Unit", 3: "Date", 4: "Shift", 5: "Hours", 8: "Closes"}),
    ("dPostList", "BuildPostings", "l", {1: "Posting ID", 2: "Rank", 3: "Name", 5: "Signed", 6: "Exempt"}),
    ("dAttend", "BuildAttend", "out", {1: "Name", 2: "Sick days used", 3: "Paid sick days left", 4: "Occurrences used",
                                       5: "Personal days used", 6: "Status"}),
    ("dWorked", "BuildWorked", "out", {1: "Employee", 2: "Period Start", 3: "Hours"}),
]


def test_each_table_has_the_columns_the_members_read():
    for table, proc, var, want in CONTRACT:
        got = headers(PUB, proc, var)
        for col, name in want.items():
            assert col <= len(got) and got[col - 1] == name, f"{table}: column {col} is {got[col - 1] if col <= len(got) else '(none)'!r}, the members' file reads it as {name!r}"


def test_the_members_roster_constants_sit_on_the_master_headers():
    got = headers(PUB, "BuildRoster", "out")
    consts = {k: int(v) for k, v in re.findall(r"Private Const (RC_\w+) As Long = (\d+)", MEM)}
    names = {"RC_NAME": "Name", "RC_CREW": "Crew", "RC_HOME": "Home Unit", "RC_TOTAL": "Total", "RC_WORKED": "Worked",
             "RC_REFUSED": "Refused", "RC_PROV": "Provisional", "RC_RCALL": "Refused call-out", "RC_RPOST": "Refused posting",
             "RC_ROTHER": "Refused other"}
    for c, name in names.items():
        assert got[consts[c] - 1] == name, f"{c} = {consts[c]} reads {got[consts[c] - 1]!r}, not {name!r}"
    assert consts["RC_FIXED"] == len(got), "RC_FIXED must be the count of the roster's fixed columns (the unit pairs follow it)"
    assert '"Refused call-out"' in MEM, "the members' DataProblem check names the call-out column"


def test_every_signup_key_the_master_reads_is_written_by_the_members_file():
    body = proc_body(SIGN, "ReadOneSignup")
    read = set(re.findall(r'Case "(\w+)"', body))
    written = set(re.findall(r'Print #ff, "(\w+)="', MEM))
    assert read, "ReadOneSignup's keys not found"
    assert read <= written, f"the master reads {sorted(read - written)} that the members' file never writes"


# ---- v28 (Members M5): the requests contract ----
def test_drequests_columns_match_the_members_constants():
    """modRequests.PublishRequestsTable writes out(1, k); the members' DQ_* constants read dRequests by those columns."""
    got = headers(REQ, "PublishRequestsTable", "out")
    want = ["Employee", "Kind", "From", "To", "Shift", "Partner", "Status", "Answer", "Decided"]
    assert got == want, got
    dq = {k: int(v) for k, v in re.findall(r"Private Const DQ_(\w+) As Long = (\d+)", MEM)}
    assert dq == {"EMP": 1, "KIND": 2, "FROM": 3, "TO": 4, "SHIFT": 5, "PARTNER": 6, "STATUS": 7, "ANSWER": 8, "DECIDED": 9}, dq


def test_every_request_key_the_master_reads_is_written_by_the_members_file():
    body = proc_body(REQ, "RequestFileRead")
    read = set(re.findall(r'Case "(\w+)"', body))
    wbody = proc_body(MEM, "WriteRequestFile")
    written = set(re.findall(r'Print #ff, "(\w+)="', wbody))
    assert read, "RequestFileRead's keys not found"
    assert read <= written, f"the master reads {sorted(read - written)} that the members' file never writes"
    # and the members' own reader (the partner's "asked of me") understands every key it writes but machine
    mbody = proc_body(MEM, "ReadRequestFile")
    mread = set(re.findall(r'Case "(\w+)"', mbody))
    assert written - {"machine"} <= mread, sorted((written - {"machine"}) - mread)


def test_request_status_words_are_one_list_on_both_sides():
    master = set(re.findall(r'Public Const RQ_ST_\w+ As String = "([^"]+)"', REQ))
    shown = set(re.findall(r'(?:st = |Or st = |st <> )"([^"]+)"', proc_body(MEM, "MemberShowRequests")))
    # the members' file adds one word of its own for a request the roster has not read yet
    assert shown - master == {"Sent - waiting for the roster"}, sorted(shown - master)
    # v29 (review LOW-7): every status word the master can publish is named in the members' legend ("Statuses: ...")
    legend = re.search(r'"Statuses: ([^"]*)"', proc_body(MEM, "MemberShowRequests"))
    assert legend, "the members' Statuses legend was not found"
    missing = [w for w in master if w not in legend.group(1)]
    assert not missing, f"status words the master publishes but the members' legend does not name: {missing}"


def test_the_leave_words_are_one_list_on_both_sides():
    """The members' file screens a request's note with the master's leave-word list (LeaveWordsIn) mirrored as text."""
    master_body = proc_body(PUB, "LeaveWordsIn")
    master_line = re.search(r'Case ("STD"[^\r\n]*)', master_body).group(1)
    master = set(master_line.replace('"', "").replace(" ", "").split(","))
    member_body = proc_body(MEM, "LeaveWordIn")
    member = set(re.search(r'Array\((.*?)\)', member_body).group(1).replace('"', "").replace(" ", "").split(","))
    assert master == member, (sorted(master - member), sorted(member - master))
    assert "LEAVE OF ABSENCE" in member_body and "LEAVE OF ABSENCE" in master_body
