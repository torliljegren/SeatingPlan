# enconding: UTF-8
import random
import sys
import time
import tkinter
import pyperclip
import tkinter.ttk as ttk
from pathlib import Path
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename, askopenfilename
from tkinter.scrolledtext import ScrolledText
from tkinter.messagebox import askyesnocancel
from threading import Thread

import Constants
from Constants import *
from ManualPlaceWin import *
from tktooltip import ToolTip

if OP_SYS == 'linux':
    from ttkthemes.themed_tk import ThemedTk

from StudentSeat import *
from xlsxwriter import *


class PlanWin(object):

    def __init__(self, prev_files: list[str]=None):
        self.seats: list[StudentSeat] = []
        self.editclickse = 0
        self.filepath = ""
        self.dirty = False
        self.run_thread = True
        self.manwin: ManualPlaceWin = None
        self.first_edit_click = True

        if OP_SYS == 'linux':
            self.root = ThemedTk()
        else:
            self.root = Tk()
        self.root.title('Ny placering')

        self.bgframe = ttk.Frame(self.root)
        self.bgframe.pack()

        self.notebook = ttk.Notebook(self.bgframe)
        self.notebook.pack(expand=True, side=tkinter.BOTTOM, pady=(10,0))

        self.buttonframe = ttk.LabelFrame(self.bgframe, text='Verktyg')
        self.buttonframe.pack(expand=True, fill='x', side=tkinter.TOP, pady=(5,0), padx=10)

        self.countvar = StringVar(self.bgframe, value='Antal placerade: 0   Antal i klasslistan: 0')
        self.countlabel = Label(self.bgframe, textvariable=self.countvar)
        self.countlabel.pack(expand= True, fill='x', side=tkinter.BOTTOM, pady=5, padx=15)

        imrand = PhotoImage(file='rand.png')
        self.randbutton = ttk.Button(self.buttonframe, image=imrand, command=self.cmd_rand)
        self.randbutton.grid(row=0, column=1, padx=10, pady=10)
        ToolTip(self.randbutton, 'Slumpa sittplatser', 1.5, follow=False)

        imsort = PhotoImage(file='sort.png')
        self.sortbutton = ttk.Button(self.buttonframe, image=imsort, command=self.cmd_sort)
        self.sortbutton.grid(row=0, column=2, padx=10, pady=10)
        ToolTip(self.sortbutton, 'Sortera sittplatser A-Ö', 1.5, follow=False)

        imerease = PhotoImage(file='erease.png')
        self.ereasebutton = ttk.Button(self.buttonframe, image=imerease, command=self.cmd_unplace)
        self.ereasebutton.grid(row=0, column=3, padx=10, pady=10)
        ToolTip(self.ereasebutton, 'Ta bort alla namn från placeringen', 1.5, follow=False)

        imclear = PhotoImage(file='clear.png')
        self.clearbutton = ttk.Button(self.buttonframe, image=imclear, command=self.cmd_clear)
        self.clearbutton.grid(row=0, column=4, padx=10, pady=10)
        ToolTip(self.clearbutton, 'Ta bort alla sittplatser', 1.5, follow=False)

        imsave = PhotoImage(file='save.png')
        self.savebutton = ttk.Button(self.buttonframe, image=imsave, command=self.cmd_save)
        self.savebutton.grid(row=0, column=5, padx=10, pady=10)
        ToolTip(self.savebutton, 'Spara sittplacering till fil', 1.5, follow=False)

        imsaveas = PhotoImage(file='saveas.png')
        self.saveasbutton = ttk.Button(self.buttonframe, image=imsaveas, command=lambda:self.cmd_save(saveas=True))
        self.saveasbutton.grid(row=0, column=6, padx=10, pady=10)
        ToolTip(self.saveasbutton, 'Spara som ny fil', 1.5, follow=False)

        imload = PhotoImage(file='open.png')
        self.loadbutton = ttk.Button(self.buttonframe, image=imload, command=self.cmd_open)
        self.loadbutton.grid(row=0, column=7, padx=10, pady=10)
        ToolTip(self.loadbutton, 'Öppna en sparad sittplacering', 1.5, follow=False)

        imexcel = PhotoImage(file='excel.png')
        self.export_button = ttk.Button(self.buttonframe, image=imexcel, command=self.cmd_export)
        self.export_button.grid(row=0, column=8, padx=(10,20), pady=10)
        ToolTip(self.export_button, 'Spara som Excel arbetsbok', 1.5, follow=False)

        self.editvar = tkinter.BooleanVar(value=False)
        self.editbutton = ttk.Checkbutton(self.buttonframe, text="Ändra platser", variable=self.editvar,
                                          command=self.cmd_editmode)
        self.editbutton.grid(row=0, column=9, padx=(20,20), pady=10)

        ttk.Label(self.buttonframe, text='Tavlan:').grid(row=0, column=10, padx=(10,0), pady=10)
        self.upvar = tkinter.StringVar(value='n')
        ttk.Radiobutton(self.buttonframe, text='Norr', value='n', variable=self.upvar,
                        command=self.cmd_orientation).grid(row=0,column=11, padx=(0,5), pady=10)
        ttk.Radiobutton(self.buttonframe, text='Söder', value='s', variable=self.upvar,
                        command=self.cmd_orientation).grid(row=0, column=12, pady=10)

        im1 = PhotoImage(file='addrow.png')
        self.plusrowbutton = ttk.Button(self.buttonframe, image=im1, command=lambda: self.change_grid('+r'), width=2)
        self.plusrowbutton.grid(row=0, column=13, padx=(40,0), sticky='e')
        ToolTip(self.plusrowbutton, 'Lägg till en rad längst ned', 1.5, follow=False)

        im2 = PhotoImage(file='delrow.png')
        self.minusrowbutton = ttk.Button(self.buttonframe, image=im2, command=lambda: self.change_grid('-r'), width=2)
        self.minusrowbutton.grid(row=0, column=14, padx=(0,0), sticky='e')
        ToolTip(self.minusrowbutton, 'Ta bort raden längst ned', 1.5, follow=False)

        im3 = PhotoImage(file='addcol.png')
        self.pluscolbutton = ttk.Button(self.buttonframe, image=im3, command=lambda: self.change_grid('+c'), width=2)
        self.pluscolbutton.grid(row=0, column=15, padx=(0, 0), sticky='e')
        ToolTip(self.pluscolbutton, 'Lägg till en kolumn längst till höger ', 1.5, follow=False)

        im4 = PhotoImage(file='delcol.png')
        self.minuscolbutton = ttk.Button(self.buttonframe, image=im4, command=lambda: self.change_grid('-c'), width=2)
        self.minuscolbutton.grid(row=0, column=16, padx=(0, 10), sticky='e')
        ToolTip(self.minuscolbutton, 'Ta bort kolumnen längst till höger', 1.5, follow=False)

        # spacer label
        # ttk.Label(self.buttonframe, text='').grid(row=0, column=15)
        self.buttonframe.columnconfigure(17, weight=1)

        self.search_var = StringVar(value='Sök')
        self.search_var.trace_add('write', self.cmd_search)
        self.search_entry = ttk.Entry(self.buttonframe, textvariable=self.search_var, state=tkinter.DISABLED,
                                      name='searchEntry')
        self.search_entry.grid(row=0, column=17, padx=(30,20))
        self.search_entry.bind('<Enter>', self.search_on_enter)
        self.search_entry.bind('<FocusOut>', self.search_on_fout)
        self.search_entry.bind('<Leave>', self.search_on_fout)

        self.buttonframe.columnconfigure(19, weight=1)

        ttk.Label(self.buttonframe, text='Sparade:').grid(row=0, column=18, padx=(10, 0), pady=10, sticky='e')
        self.combovar = tkinter.StringVar()
        self.previouscombobox = ttk.Combobox(self.buttonframe, textvariable=self.combovar, state='readonly')
        self.previouscombobox.grid(row=0, column=19, padx=(5, 0), sticky='e', pady=10)
        self.previouscombobox.bind('<<ComboboxSelected>>', self.combobox_event)

        imunpin = PhotoImage(file='unpin.png')
        self.unpin_button = ttk.Button(self.buttonframe, image=imunpin, command=self.remove_selected_from_combobox)
        self.unpin_button.grid(row=0, column=20, padx=(2,10), pady=10)
        ToolTip(self.unpin_button, 'Ta bort från listan', 1.5, follow=False)

        self.seatframe = ttk.Frame(self.notebook, width=FRAME_WIDTH, height=FRAME_HEIGHT)
        self.nameframe = ttk.Frame(self.notebook, width=FRAME_WIDTH, height=FRAME_HEIGHT)

        self.seatframe.pack(fill='both', expand=True, side=LEFT)
        self.nameframe.pack(fill='both', expand=True, side=LEFT)

        self.notebook.add(self.seatframe, text='Placering')
        self.notebook.add(self.nameframe, text='Klasslista')

        # TODO: change this widget to my own per here:
        # https://stackoverflow.com/questions/64774411/is-there-a-ttk-equivalent-of-scrolledtext-widget-tkinter
        self.textarea = ScrolledText(self.nameframe, name='textarea')
        self.textarea.pack(expand=True, side=TOP, fill='both')
        self.textarea.insert(1.0, 'Högerklicka för att klistra in.')

        if prev_files:
            self.prev_files = prev_files
            self.update_combobox()
        else:
            self.prev_files: list[str] = []

        self.root.wm_protocol('WM_DELETE_WINDOW', self.on_close)

        self.root.bind('<Control-f>', self.search_on_enter)
        self.root.bind('<Control-s>', self.cmd_save)
        self.root.bind('<Control-o>', self.cmd_open)
        self.root.bind('<Command-s>', self.cmd_save)
        self.root.bind('<Command-o>', self.cmd_open)
        self.root.bind('<Escape>', self.search_on_leave)
        self.root.bind_all('<KeyPress>', self.keypress)

        self.textarea.bind('<Button-2>', self.cmd_paste)
        self.textarea.bind('<Button-3>', self.cmd_paste)

        self.setup_grid()

        if OP_SYS == 'linux':
            s = ttk.Style()
            s.theme_use('plastik')

        self.update_thread = Thread(target=self.periodic_stucount_update, daemon=True,
                                    args=(self.root, self.update_student_count, self.run_thread))
        self.update_thread.start()

        self.root.mainloop()

    # creates the buttons representing seats and grids them to self.buttonframe
    def setup_grid(self):
        for y in range(0, TOTAL_SEATS_Y):
            for x in range(0, TOTAL_SEATS_X):
                # s = Button(self.seatframe, text=str(x)+str(y))
                s = StudentSeat(self.seatframe, x, y, self, self.editvar)#, name=str(x)+", "+str(y))
                # s.callback = lambda: self.seat_callback(s)
                self.seats.append(s)
                s.grid(row=y, column=x, padx=SEAT_SPACING, pady=SEAT_SPACING)

    # finds the maximum x- and y-coordinate where seats are active, e.g. the bottom rightmost edge of the classroom
    def seat_bounds(self) -> tuple:
        xmax = 0
        ymax = 0

        for seat in self.seats:
            if seat.xpos > xmax and seat.active:
                xmax = seat.xpos
            if seat.ypos > ymax and seat.active:
                ymax = seat.ypos
        # print("xmax ymax =", xmax, ymax)
        return (xmax, ymax)


    # Threading function for updating student count
    def periodic_stucount_update(self, window, upd, run):
        while run:
            time.sleep(0.5)
            t1 = time.perf_counter_ns()
            upd()
            t2 = time.perf_counter_ns()
            # print('%f ms'%(float((t2-t1)*10**(-6))))

    ########################
    #     BUTTON CALLS     #
    ########################

    # paste into textarea
    def cmd_paste(self, e):
        # prepare and clean up the list
        namelist_1 = [n.strip() for n in pyperclip.paste().split('\n') if '\n' not in n and n != '']
        print(namelist_1)
        namelist_2 = [n for n in namelist_1 if n != '']

        # if platform is windows, sometimes åäö and accented chars get messed up. Try to encode then decode to resolve.
        if OP_SYS == 'windows':
            try:
                print('Attempting to fix utf8 chars')
                templist = [n.encode('utf-8').decode('utf-8') for n in namelist_2]
                print('Success!')
                namelist_2 = templist
            except UnicodeEncodeError:
                print('UTF-8 encoding error in string:', [s for s in namelist_2 if s not in namelist_2])

        # write the names to the textarea
        self.textarea.delete(1.0, tkinter.END)
        if not namelist_2:
            namelist_2 = ['']
        namelist_2.sort()
        self.textarea.insert(1.0, '\n'.join(namelist_2))

    # removes all seats
    def cmd_clear(self):
        self.dirty = True
        for seat in self.seats:
            seat.deactivate()
            #seat.name_set('')
        if self.manwin:
            self.manwin.update_names(only_unplaced=True)


    # removes all the names from the seats but keeps the seats
    def cmd_unplace(self):
        self.dirty = True
        for seat in self.seats:
            seat.name_set('')
        if self.manwin:
            self.manwin.update_names()


    def cmd_save(self, e=None, saveas=False):
        filepath = ''
        if not self.filepath or saveas:
            # print('self.filepath = ', self.filepath, '   saveas = ', saveas)
            filepath = asksaveasfilename(parent=self.root, defaultextension=".dat",
                                         filetypes=[("Placeringsdata","*.dat")])
        else:
            filepath = self.filepath

        if not filepath:
            return
        else:
            self.root.title(filepath.split('/')[-1])
            if filepath not in self.prev_files:
                self.prev_files.append(filepath)
                self.update_combobox(filepath)
                with open(PREV_FILES_PATH, 'a', encoding='utf-8') as f:
                    f.write(filepath+";")
                self.filepath = filepath
            self.write_data(filepath)
            self.dirty = False


    def cmd_open(self, e=None):
        filepath = askopenfilename(parent=self.root, filetypes=[("Planeringsdata", "*.dat")])
        if not filepath:
            return
        else:
            self.filepath = filepath
            self.root.title(filepath.split('/')[-1])
            if filepath not in self.prev_files:
                self.prev_files.append(filepath)
                self.update_combobox(filepath)
                with open(PREV_FILES_PATH, 'a', encoding='utf-8') as f:
                    f.write(filepath+';')
            self.load_data(filepath)
            self.dirty = False


    def cmd_sort(self):
        for seat in self.seats:
            seat.name_set('')
        sorted_names = tuple(sorted(self.name_tuple()))
        self.textarea.delete(1.0, tkinter.END)
        self.textarea.insert(1.0, '\n'.join(sorted_names))
        self.place_names(sorted_names)
        self.cmd_search(None)
        if self.manwin:
            self.manwin.update_names(only_unplaced=True)
        self.dirty = True


    def cmd_rand(self):
        # clear names
        # print("Window height = ", self.root.winfo_height())
        for seat in self.seats:
            seat.name_set("")

        # get names
        names = list(self.name_tuple())
        # print("Original:",names)

        # rand names
        random.shuffle(names)
        # print("Shuffled:", names)

        # place names
        self.place_names(tuple(names))
        self.cmd_search(None)

        if self.manwin:
            self.manwin.update_names(only_unplaced=True)

        self.dirty = True


    def cmd_editmode(self):
        if not self.manwin:
            # init the manual placement window if its not inited already
            self.manwin = ManualPlaceWin(self.root, self)

            # fixes a bug where editbutton has to be invoked twice for editmode to work when used for the first time
            # in a session
            # this work around is placed here so that it toggles twice only the first time the user clicks it.
            self.editbutton.invoke()
            self.editbutton.invoke()

        # reset edit mode when checkbutton unchecks
        if not self.editvar.get():
            self.editclicks = 0
            if Constants.seat1:
                Constants.seat1.configure(bg=Constants.ACTIVE_COLOR)
            # hide the manual placement window
            self.manwin.win.iconify()
        else:
            # show the manual placement window and update it
            self.manwin.update_names(only_unplaced=True)
            self.manwin.win.deiconify()


    def cmd_export(self):
        filepath = asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel 2007-","*.xlsx")])
        if not filepath:
            return
        else:
            self.export_xlsx(filepath)


    def cmd_orientation(self):
        self.rotate_seats_180()
        self.dirty = True


    def on_close(self):
        # prompt to save if changes were made
        if self.dirty:
            result = askyesnocancel(title='Spara?', message='Vill du spara ändringarna innan du avslutar?',
                                    master=self.root)
            if result is None:
                return
            elif result:
                self.cmd_save(e=None, saveas=False)

        # write list of previous opened files to file
        with open(PREV_FILES_PATH, mode='w', encoding='utf-8') as f:
            if self.prev_files:
                # print('Writing to tidigare.txt:', self.prev_files)
                for filepath in self.prev_files:
                    f.write(filepath+';')
            else:
                f.write('')

        # terminate thread then exit
        self.run_thread = False
        #self.update_thread.join()
        self.root.destroy()
        sys.exit(0)


    def cmd_search(self, *args):
        searchtext = self.search_var.get()
        if not searchtext:
            self.search_on_leave(None)
        for seat in self.seats:
            if seat.name_get().lower().startswith(searchtext.lower()) and searchtext:
                seat.configure(bg='spring green')
            else:
                seat.restore_color()


    def keypress(self, e):
        if not e.char.isalpha() or self.notebook.tab(self.notebook.select(), 'text') == 'Klasslista':#self.notebook.select()):
            return

        if 'searchEntry' not in str(self.root.focus_get()):
            self.search_on_enter(None)
            self.search_var.set(e.char)
            self.search_entry.icursor(tkinter.END)
            self.cmd_search(None)
            self.search_entry.focus_set()




    ########################
    # CONTROLLER FUNCTIONS #
    ########################

    def update_student_count(self):
        # total number of students in list
        tot_stus = len(self.name_tuple())

        # total number of placed students
        placed_stus = self.num_active(only_taken=True)

        self.countvar.set('Antal placerade: %s   Antal i klasslistan: %s'%(placed_stus, tot_stus))

    def search_on_enter(self, e):
        if self.search_var.get() == 'Sök':
            self.search_var.set('')
        self.search_entry.configure(state=tkinter.NORMAL)
        # self.search_entry.focus_set()


    def search_on_leave(self, e):
        self.search_entry.configure(state=tkinter.DISABLED)
        self.search_var.set('Sök')
        try:
            self.buttonframe.focus_set()
        except:
            pass


    def search_on_fout(self, e):
        if not self.search_var.get() or self.search_var.get() == 'Sök':
            # print('Clearing search')
            self.search_on_leave(None)
        # else:
            # print('Searchtext is', self.search_var.get())


    def change_grid(self, change: str):
        max_x = Constants.TOTAL_SEATS_X
        max_y = Constants.TOTAL_SEATS_Y

        if change == '+r':
            for i in range(max_x):
                s = StudentSeat(self.seatframe, i, max_y, self.root, self.editvar)
                s.grid(row=max_y, column=i, padx=SEAT_SPACING, pady=SEAT_SPACING)
                self.seats.append(s)
            Constants.TOTAL_SEATS_Y += 1
        elif change == '-r':
            i = 0
            while i < len(self.seats):
                s = self.seats[i]
                if s.ypos == max_y-1:
                    s.grid_forget()
                    s.destroy()
                    self.seats.pop(i)
                    continue
                i += 1
            Constants.TOTAL_SEATS_Y -= 1
        elif change == '+c':
            for i in range(max_y):
                s = StudentSeat(self.seatframe, max_x, i, self.root, self.editvar)
                s.grid(row=i, column=max_x, padx=SEAT_SPACING, pady=SEAT_SPACING)
                self.seats.append(s)
            Constants.TOTAL_SEATS_X += 1
        elif change == '-c':
            i = 0
            while i < len(self.seats):
                s = self.seats[i]
                if s.xpos == max_x-1:
                    s.grid_forget()
                    s.destroy()
                    self.seats.pop(i)
                    continue
                i += 1
            Constants.TOTAL_SEATS_X -= 1


    def ispropername(self, name: str):
        if 'Högerklicka' in name:
            return False
        elif name.isalnum():
            return True
        elif ' ' in name.strip():
           return True 
        else:
            return False
    
    # read the names from the textbox in the GUI and return them in a tuple
    def name_tuple(self) -> tuple:
        namesstr = self.textarea.get(1.0, END).strip()
        tempnames = str.split(namesstr, "\n")
        clean_names: list[str] = [name for name in tempnames if self.ispropername(name)]

        # print('name_tuple():', tempnames)
        return tuple(clean_names)


    def place_names(self, names: tuple):
        act_sts = self.active_seats()
        if not act_sts:
            return

        i = 0
        for seat in act_sts:
            if i >= len(names):
                return
            else:
                seat.name_set(names[i])
                i += 1


    def num_active(self, only_taken=False) -> int:
        n = 0
        if only_taken:
            for seat in self.seats:
                if seat.active and self.ispropername(seat.name_get()):
                    n += 1
        else:
            for seat in self.seats:
                if seat.active:
                    n += 1
        # print("Active:",n)
        return n


    def active_seats(self) -> list[StudentSeat]:
        act_sts = []
        for seat in self.seats:
            if seat.active:
                act_sts.append(seat)
        return act_sts


    def seat_callback(self, btn: StudentSeat):
        if self.editclicks == 1:
            Constants.seat2 = btn
            # print("Seat1:", Constants.seat1, "Seat2:", Constants.seat2)
            # print("Swapping", name1, "and", name2)
            self.swap_seats(Constants.seat1, Constants.seat2)
            Constants.seat1.configure(bg=Constants.ACTIVE_COLOR)
            Constants.seat2.configure(bg=Constants.ACTIVE_COLOR)
            self.cmd_search(None)
            Constants.seat1 = None
            Constants.seat2 = None
            self.editclicks = 0
        elif self.editclicks == 0:
            self.editclicks += 1
            Constants.seat1 = btn
            # print("Seat1:", Constants.seat1, "Seat2:", Constants.seat2)


    def swap_seats(self, seat1: StudentSeat, seat2: StudentSeat):
        self.dirty = True
        name1 = seat1.name_get()
        name2 = seat2.name_get()
        act1 = seat1.active
        act2 = seat2.active
        # print("Swapping", name1, "and", name2)
        seat1.name_set(name2)
        seat2.name_set(name1)

        if act1:
            seat2.activate()
        else:
            seat2.deactivate()

        if act2:
            seat1.activate()
        else:
            seat1.deactivate()


    def edit_seat(self, x:int, y:int, newname:str, newactive:bool):
        for seat in self.seats:
            if seat.xpos == x and seat.ypos == y:
                seat.name_set(newname)
                if newactive:
                    seat.activate()
                else:
                    seat.deactivate()


    def write_data(self, filepath: str):
        # prepare strings from student list
        students = self.textarea.get(1.0, tkinter.END).split("\n")

        # prepare strings of taken seats.
        taken_seats: list[str] = []
        for seat in self.seats:
            if seat.active:
                #                       row                    #col                    #name
                taken_seats.append(str(seat.ypos) + "," + str(seat.xpos) + "," + seat.name_get() + ";")

        with open(filepath, 'w', encoding='utf-8') as f:
            # write list of students, like so:
            # name1;name2;name3...
            f.write("STUDENT LIST\n")
            for student in students:
                if student:
                    f.write(student+";")

            # write list of active seats with coords and name, like so:
            # row,col,name;row,col,name
            f.write("\nACTIVE SEATS\n")
            for seat in taken_seats:
                f.write(seat)


    def update_combobox(self, newpath: str=""):
        filenames = []
        for filepath in self.prev_files:
            filenames.append(filepath.split("/")[-1])

        self.previouscombobox['values'] = tuple(filenames)
        self.combovar.set(newpath.split("/")[-1])


    def combobox_event(self, e):
        sel = self.combovar.get()
        for filepath in self.prev_files:
            if filepath.split("/")[-1] == sel and sel and sel!="\n":
                self.load_data(filepath)
                self.root.title(filepath.split('/')[-1])

    def remove_selected_from_combobox(self):
        self.prev_files.remove(self.filepath)
        self.update_combobox()


    def load_data(self, filepath: str):
        stus = ""
        acts = ""
        test_path = Path(filepath)
        if not test_path.is_file():
            ans = messagebox.askyesno(title='Filfel', message='Kunde inte öppna filen\n'+ filepath +
                                                              '.\nDen kanske har blivit borttagen eller flyttad.\n\n'+
                                                              'Ta bort filen från listan?')
            if ans:
                self.prev_files.remove(filepath)
                self.update_combobox()
            return

        with open(filepath, 'r') as f:
            f.readline()
            stus = f.readline()
            f.readline()
            acts = f.readline()

        # remove trailing ;
        stus = stus[0:-1] if stus[-1]==";" else stus
        acts = acts[0:-1] if acts[-1]==";" else acts

        stulist = stus.split(";")
        actslist = acts.split(";")
        # print("Students:",stulist,"\nSeats:",actslist)

        # replace textarea with loaded students
        stustring = ""
        for stu in stulist:
            stustring += stu + "\n"

        self.textarea.delete(1.0, tkinter.END)
        self.textarea.insert(1.0, stustring)

        # clear seats and insert the loaded ones
        self.cmd_clear()
        # check if grid needs to be enlarged
        # then insert loaded ones
        for act in actslist:
            actsplit = act.split(",")
            while int(actsplit[0]) >= Constants.TOTAL_SEATS_Y:
                self.change_grid('+r')
            while int(actsplit[1]) >= Constants.TOTAL_SEATS_X:
                self.change_grid('+c')
            self.edit_seat(int(actsplit[1]), int(actsplit[0]), actsplit[2], True)
        self.filepath = filepath
        self.dirty = False

        # close the manual placement window if its opened
        if self.editvar.get():
            self.editbutton.invoke()



    def rotate_seats_180(self):
        xmax, ymax = self.seat_bounds()

        new_order = []

        for seat in self.seats:
            newx = xmax - seat.xpos
            newy = ymax - seat.ypos

            for newseat in self.seats:
                if newseat.xpos == newx and newseat.ypos == newy and not seat.has_swapped:# and seat.active:
                    seat.has_swapped = True
                    newseat.has_swapped = True
                    self.swap_seats(seat, newseat)

        for seat in self.seats:
            seat.has_swapped = False

        self.cmd_search(None)


    def export_xlsx(self, filepath: str):
        # print('Saving workbook:', filepath)
        wb = Workbook(filepath)
        ws = wb.add_worksheet()
        ws.set_landscape()

        # create a heading (whiteboard)
        headingformat = wb.add_format({'bold':True, 'align':'center', 'valign':'vcenter',
                                       'fg_color':Constants.BOARD_COLOR,
                                       'font_size':'20'})
        # convert seats x coord. to letter
        xymax = self.seat_bounds()
        xmax = Constants.COLNUM[xymax[0]]
        ymax = xymax[1]
        headingformat.set_bottom()
        headingformat.set_top()
        headingformat.set_left()
        headingformat.set_right()
        headingformat.set_border(2)

        ortn = self.upvar.get()

        if ortn == 'n':
            ws.merge_range('B1:'+xmax+'1', 'Tavla', headingformat)
        else:
            # print('ymax is', ymax)
            ymax = str(int(ymax) + 4)
            ws.merge_range('B' + ymax + ':' + xmax + ymax, 'Tavla', headingformat)

        tableformat = wb.add_format({'fg_color':'#E3D0B7',#a lighter Constants.ACTIVE_COLOR,
                                     'align':'center', 'valign':'vcenter',
                                     'text_wrap':True, 'font_size':20})
        tableformat.set_bottom()
        tableformat.set_top()
        tableformat.set_left()
        tableformat.set_right()
        buf_spc = 3 if ortn=='n' else 0
        for seat in self.seats:
            if seat.active:
                            #row        col
                ws.write(seat.ypos+buf_spc, seat.xpos, seat.name_get(), tableformat)
                ws.set_row(seat.ypos + buf_spc, 45)
                ws.set_column(seat.xpos, seat.xpos, 16)

        wb.close()
