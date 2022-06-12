import tkinter.ttk
from tkinter import *
from tkinter.ttk import *
from StudentSeat import  *
from ScrolledFrame import *
from Constants import *
import PlanWin

class ManualPlaceWin(object):

    def __init__(self, master, planwin):
        super().__init__()
        self.planwin: PlanWin.PlanWin = planwin
        self.win = Toplevel(master, bg='white')
        self.win.title('Lista')
        self.win.wm_protocol('WM_DELETE_WINDOW', lambda: self.win.iconify())
        self.buttons: list[StudentSeat] = []
        self.names: list[str] = []
        #self.titleframe = Frame(self.win)
        #Label(self.titleframe, text='Klasslista').pack()
        #self.titleframe.pack(side=TOP)
        self.sf = ScrolledFrame(self.win)
        self.win.rowconfigure(1,weight=1)
        self.win.columnconfigure(0,weight=1)
        self.sf.grid(row=1, column=0,sticky=NSEW)
        self.mainframe = self.sf.scrolled_frame
        self.labelframe = tkinter.ttk.Frame(self.win)
        self.labelframe.grid(row=0, column=0, sticky=EW)
        Label(self.labelframe, text='Ej placerade').pack(side=TOP)
        self.empty_label = Label(self.mainframe, text='Listan är tom')


        self.update_names(only_unplaced=True)

        if OP_SYS  == 'linux':
            self.win.bind_all('<4>', self.scrolling_linux)
            self.win.bind_all('<5>', self.scrolling_linux)
        else:
            self.win.bind_all('<MouseWheel>', self.scrolling)


        self.win.geometry(str(int(self.planwin.seats[0].winfo_width()*1.2))+'x'+str(int(self.win.winfo_screenheight()*0.6)))

        #self.win.geometry('100x500')

    def update_names(self, only_unplaced=False):
        # remove all buttons
        if self.buttons:
            for button in self.buttons:
                button.grid_forget()
                button.destroy()
            self.buttons = []

        # retrieve updated list and create new buttons
        self.names = list(self.planwin.name_tuple())

        # create list with unfilled names
        if only_unplaced:
            placed_names = [x.name_get() for x in self.planwin.active_seats()]
            unplaced_names = [x for x in self.names if x not in placed_names]
            self.names = unplaced_names

        # place a label explaining the list is empty if needed
        if not self.names:
            self.empty_label.grid()
        else:
            self.names.sort()
            self.empty_label.grid_forget()

        buttonrow = 1
        for studentname in self.names:
            newseat = StudentSeat(self.mainframe, buttonrow, 0, self.planwin, self.planwin.editvar, studentname)
            self.buttons.append(newseat)
            newseat.grid(row=buttonrow, column=0)
            newseat.activate()
            buttonrow += 1

    def scrolling(self, event):
        self.sf.canvas.yview_scroll(int(-1*event.delta/SCR_SPEED), "units")

    def scrolling_linux(self, event):
        if event.num == 4:
            self.sf.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.sf.canvas.yview_scroll(1, "units")
