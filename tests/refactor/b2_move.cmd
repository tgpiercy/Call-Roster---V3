@echo off
rem v27.1 builder B2, commit 1: the move of modPeople into modDates / modFormer / modPlannedMoves, by script only.
rem Line numbers are modPeople.bas as at ab858ac. Run from the repo root. Prove it with: py -3.14 tests/movecheck.py ab858ac
rem modFormer starts at 56 (not 57): line 56 is the comment that heads the FM_* constants.
py -3.14 tests/splitmodule.py . modPeople modDates=652-834 modFormer=56-75,1147-1326,3006-3197 modPlannedMoves=127-129,1774-2661
