"""v27.1 builder B1, commit 1: the modCalendar split, made by script (never retyped).

Run from the repo root on ab858ac (modCalendar.bas as at v27, 6,312 lines):

    py -3.14 tests/refactor/b1_move.py

It is exactly this command:

    py -3.14 tests/splitmodule.py . modCalendar modBoard=3520-3730 modMinStaff=95-98,3731-4076,6230-6290
        modCalImport=4090-4318 modNeeds=106-115,4541-4673,4686-5196,5208-5393
        modCalMenus=99-99,104-104,5394-5561,5582-6168

The cut table in docs/v27.1-report-back.md ends three ranges one line later (4674, 5197, 5562). Those three lines
are the one-line header comments of SlotUnit, CovHasSourceCached and Dt2, which STAY in modCalendar, so the ranges
here stop one line short and each comment stays with its procedure. Nothing else differs.

Then prove it:  py -3.14 tests/movecheck.py ab858ac
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import splitmodule  # noqa: E402

ARGS = [
    ".",
    "modCalendar",
    "modBoard=3520-3730",
    "modMinStaff=95-98,3731-4076,6230-6290",
    "modCalImport=4090-4318",
    "modNeeds=106-115,4541-4673,4686-5196,5208-5393",
    "modCalMenus=99-99,104-104,5394-5561,5582-6168",
]

if __name__ == "__main__":
    sys.exit(splitmodule.main(ARGS))
