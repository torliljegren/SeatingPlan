# encoding: UTF-8

from platform import system
from pathlib import Path
from tkinter.messagebox import showwarning

OP_SYS = system().lower()
HOME_FOLDER = ''
try:
    HOME_FOLDER = str(Path.home()) + '/'
except RuntimeError as error:
    showwarning(title='Filfel', message='Hemkatalogen kunde inte öppnas. Felmeddelande:\n' + error.__str__())

PREV_FILES_PATH = HOME_FOLDER + 'tidigare.txt'


# converts Excel style column letters to 1-indexed ints
COLNUM = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M',
          14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X'}

ACTIVE_COLOR: str = "#cdaa7d" # "burlywood3"
ACTIVE_COLOR_HOOVER: str = 'burlywood2'
PASSIVE_COLOR: str = "floral white"
SELECTED_COLOR: str = "dodger blue"
SELECTED_COLOR_HOOVER: str = 'SteelBlue1'
BOARD_COLOR = "#fffaf0" # "floral white"
BUTTON_RELIEF = 'solid' if OP_SYS == 'windows' else 'flat'
BUTTON_BORDERWIDTH = 1
SEAT_FONT = (None, 18) if OP_SYS == 'darwin' else (None, 13)
SEAT_SPACING = 1 if OP_SYS == 'windows'else 0

SEAT_WIDTH: int = 110
SEAT_HEIGHT: int = 60

SEAT_CHARS_HEIGHT = SEAT_HEIGHT if OP_SYS == 'darwin' else 2
SEAT_CHARS_WIDTH = SEAT_WIDTH if OP_SYS == 'darwin' else 11

TOTAL_SEATS_X: int = 10
TOTAL_SEATS_Y: int = 12

FRAME_WIDTH = SEAT_WIDTH * TOTAL_SEATS_X
FRAME_HEIGHT = SEAT_HEIGHT * TOTAL_SEATS_Y

# Platform dependant scrolling constant
SCR_SPEED = 1 if OP_SYS == "darwin" else 120

## SHARED DATA FOR EDIT MODE ##
edit_mode = False
taken_seats = 0
seat1 = None
seat2 = None