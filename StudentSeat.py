# encoding: UTF-8

import Constants
import platform
from tkinter import StringVar, BooleanVar
import PlanWin

# import the right kind of button depending on OS
if Constants.OP_SYS == 'darwin':
    from tkmacosx import Button
else:
    from tkinter import Button


class StudentSeat(Button):

    def __init__(self, parent, xpos: int, ypos: int, win, editvar: BooleanVar, name: str = '', callback=None):
        super().__init__(parent, width=Constants.SEAT_CHARS_WIDTH, height=Constants.SEAT_CHARS_HEIGHT)

        self.varname = StringVar(self, value=name)
        self.active: bool = False
        self.config(relief=Constants.BUTTON_RELIEF, borderwidth=Constants.BUTTON_BORDERWIDTH)
        self.xpos = xpos
        self.ypos = ypos
        self.editvar = editvar
        self.callback = callback
        self.win: PlanWin.PlanWin = win
        self.has_swapped = False

        self.bind('<Enter>', lambda e: StudentSeat.on_enter(e, self))
        #self.bind('<Leave>', lambda e: StudentSeat.on_leave(e, self))

        self.config(command=self.pressed, textvariable=self.varname, bg=Constants.PASSIVE_COLOR, borderwidth=0,
                    font=Constants.SEAT_FONT, activebackground=Constants.ACTIVE_COLOR_HOOVER)
        if Constants.OP_SYS == 'darwin':
            self.config(focuscolor='', bordercolor=Constants.PASSIVE_COLOR, borderwidth=0)

    def activate(self):
        self.configure(bg=Constants.ACTIVE_COLOR, activebackground=Constants.ACTIVE_COLOR)#, borderwidth=1)
        self.active = True
        # if Constants.OP_SYS == 'darwin' and not self.varname.get():
        #     self.name_set('*AKTIV*')
        #    # self.configure(borderwidth=2)

    def deactivate(self):
        self.configure(bg=Constants.PASSIVE_COLOR, activebackground=Constants.PASSIVE_COLOR)#, borderwidth=1)
        self.active = False
        self.name_set("")
        # if platform.system().lower() == 'darwin':
            # self.configure(borderwidth=1)


    def pressed(self):
        if self.editvar.get() and self.active:
            self.configure(bg=Constants.SELECTED_COLOR)
            self.win.seat_callback(self)
        else:
            self.win.dirty = True
            if self.active:
                self.deactivate()
            else:
                self.activate()

    def name_get(self) -> str:
        return self.varname.get()

    def name_set(self, name: str):
        self.varname.set(name)

    def restore_color(self):
        if self.active:
            self.configure(bg=Constants.ACTIVE_COLOR)
        else:
            self.configure(bg=Constants.PASSIVE_COLOR)

    @staticmethod
    def on_enter(event, seat):
        color = None
        if seat.active and seat.editvar.get():
            color = Constants.SELECTED_COLOR_HOOVER
        elif seat.active:
            color = Constants.ACTIVE_COLOR_HOOVER
        else:
            color = Constants.ACTIVE_COLOR_HOOVER
        seat.configure(activebackground=color)#, borderwidth=1)

    @staticmethod
    def on_leave(event, seat):
        color = Constants.ACTIVE_COLOR if seat.active else Constants.PASSIVE_COLOR
        seat.configure(activebackground=color)#, borderwidth=0)
    # def __str__(self) -> str:
    #    return "Seat "+str(self.xpos)+", "+str(self.ypos)