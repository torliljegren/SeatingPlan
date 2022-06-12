# encoding: UTF-8

import PlanWin as pw
from pathlib import Path
from Constants import PREV_FILES_FILENAME

# open and read a file of previous opened files
checkfile = Path(PREV_FILES_FILENAME)
if checkfile.is_file():
    with open(PREV_FILES_FILENAME, 'r', encoding='utf-8') as f:
        prev = f.readline()
        prevslist: list[str] = []
        if prev:
            prevslist = prev.strip(";").split(";")
    window = pw.PlanWin(prevslist)
else:
    with open(PREV_FILES_FILENAME, 'w', encoding='utf-8') as f:
        f.write("")
    window = pw.PlanWin()