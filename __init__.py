# encoding: UTF-8

import PlanWin as pw
from pathlib import Path
from Constants import PREV_FILES_PATH

# open and read the file of previous opened files
checkfile = Path(PREV_FILES_PATH)
if checkfile.is_file():
    with open(PREV_FILES_PATH, 'r', encoding='utf-8') as f:
        prev = f.readline()
        prevslist: list[str] = []
        if prev:
            prevslist = prev.strip(";").split(";")
    window = pw.PlanWin(prevslist)
else: # create a empty previous file if none exists
    with open(PREV_FILES_PATH, 'w', encoding='utf-8') as f:
        f.write("")
    window = pw.PlanWin()