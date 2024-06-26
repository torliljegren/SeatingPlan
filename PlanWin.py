# enconding: UTF-8
import random
import sys
import time
import tkinter
import pyperclip
# import tkinter.ttk as ttk
# from pathlib import Path
# from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename, askopenfilename
# from tkinter.scrolledtext import ScrolledText
from tkinter.messagebox import askyesnocancel
from threading import Thread

# import Constants
# import Constants
from ManualPlaceWin import *
from TtkTextArea import TtkTextArea
from tktooltip import ToolTip

if OP_SYS == 'linux':
    # print("Importing ThemedTk")
    from ttkthemes.themed_tk import ThemedTk

from StudentSeat import StudentSeat
from xlsxwriter import *


class PlanWin(object):

    def __init__(self, prev_files: list[str] = None):
        tcl_v = tkinter.Tcl().eval('info patchlevel')
        print(f'TCL version: {tcl_v}')
        self.editclicks = 0
        self.seats: list[StudentSeat] = []
        self.filepath = ""
        self.dirty = False
        self.run_thread = True
        self.manwin: ManualPlaceWin = None
        self.first_edit_click = True

        self.root = ThemedTk(theme='plastik') if OP_SYS == 'linux' else Tk()
        self.root.title('Ny placering')

        self.bgframe = ttk.Frame(self.root)
        self.bgframe.pack()

        self.notebook = ttk.Notebook(self.bgframe)
        self.notebook.grid(row=3, column=0, columnspan=2, pady=(0, 20))

        self.buttonframe = ttk.LabelFrame(self.bgframe, text='Verktyg')
        self.buttonframe.grid(row=0, column=0, columnspan=2, padx=10)

        self.whiteboard_and_seatsframe = ttk.Frame(self.notebook)

        self.whiteboardframe = ttk.Frame(self.whiteboard_and_seatsframe)
        self.whiteboardframe.grid(row=0, column=0, pady=(10,15))
        self.whiteboardlabel = tk.Label(self.whiteboardframe, relief=SOLID, text=' ' * 30 + 'Tavla' + ' ' * 30,
                                        bg='white', font=('Helvetica', 12, 'bold'))
        self.whiteboardlabel.pack()

        self.countvar = StringVar(self.bgframe, value='Antal placerade: 0   Antal i klasslistan: 0')
        self.countlabel = Label(self.bgframe, textvariable=self.countvar)
        self.countlabel.grid(row=1, column=0, pady=(5,0), padx=15)

        imrand = PhotoImage(file='rand.png')
        self.randbutton = ttk.Button(self.buttonframe, image=imrand, command=self.cmd_rand)
        self.randbutton.grid(row=0, column=1, padx=10, pady=10)
        ToolTip(self.randbutton, 'Slumpa sittplatser', 1.5, follow=False)

        imsort = PhotoImage(file='sort.png')
        self.sortbutton = ttk.Button(self.buttonframe, image=imsort, command=self.cmd_sort_columnwise_cluster)
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
        self.saveasbutton = ttk.Button(self.buttonframe, image=imsaveas, command=lambda: self.cmd_save(saveas=True))
        self.saveasbutton.grid(row=0, column=6, padx=10, pady=10)
        ToolTip(self.saveasbutton, 'Spara som ny fil', 1.5, follow=False)

        imload = PhotoImage(file='open.png')
        self.loadbutton = ttk.Button(self.buttonframe, image=imload, command=self.cmd_open)
        self.loadbutton.grid(row=0, column=7, padx=10, pady=10)
        ToolTip(self.loadbutton, 'Öppna en sparad sittplacering', 1.5, follow=False)

        imexcel = PhotoImage(file='excel.png')
        self.export_button = ttk.Button(self.buttonframe, image=imexcel, command=self.cmd_export)
        self.export_button.grid(row=0, column=8, padx=(10, 20), pady=10)
        ToolTip(self.export_button, 'Spara som Excel arbetsbok', 1.5, follow=False)

        self.editvar = tkinter.BooleanVar(value=False)
        self.editbutton = ttk.Checkbutton(self.buttonframe, text="Ändra platser", variable=self.editvar,
                                          command=self.cmd_editmode)
        self.editbutton.grid(row=0, column=9, padx=(20, 20), pady=10)

        ttk.Label(self.buttonframe, text='Tavlan:').grid(row=0, column=10, padx=(10, 0), pady=10)
        self.orientationvar = tkinter.StringVar(value='n')
        self.northbutton = ttk.Radiobutton(self.buttonframe, text='Norr', value='n', variable=self.orientationvar,
                                           command=self.cmd_orientation_n, state='disabled')
        self.northbutton.grid(row=0, column=11, padx=(0, 5), pady=10)
        self.southbutton = ttk.Radiobutton(self.buttonframe, text='Söder', value='s', variable=self.orientationvar,
                                           command=self.cmd_orientation_s)
        self.southbutton.grid(row=0, column=12, pady=10)

        im1 = PhotoImage(file='addrow.png')
        self.plusrowbutton = ttk.Button(self.buttonframe, image=im1, command=lambda: self.change_grid('+r'), width=2)
        self.plusrowbutton.grid(row=0, column=13, padx=(40, 0), sticky='e')
        ToolTip(self.plusrowbutton, 'Lägg till en rad längst ned', 1.5, follow=False)

        im2 = PhotoImage(file='delrow.png')
        self.minusrowbutton = ttk.Button(self.buttonframe, image=im2, command=lambda: self.change_grid('-r'), width=2)
        self.minusrowbutton.grid(row=0, column=14, padx=(0, 0), sticky='e')
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
        self.search_entry.grid(row=0, column=17, padx=(30, 20))
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
        self.unpin_button.grid(row=0, column=20, padx=(2, 10), pady=10)
        ToolTip(self.unpin_button, 'Ta bort från listan', 1.5, follow=False)

        self.seatframe = ttk.Frame(self.whiteboard_and_seatsframe, width=FRAME_WIDTH, height=FRAME_HEIGHT)
        self.nameframe = ttk.Frame(self.notebook, width=FRAME_WIDTH, height=FRAME_HEIGHT)

        # self.seatframe.pack(fill='both', expand=True, side=LEFT)
        # self.nameframe.pack(fill='both', expand=True, side=LEFT)

        self.seatframe.grid(row=1, column=0)
        self.nameframe.grid(row=0, column=0)

        self.notebook.add(self.whiteboard_and_seatsframe, text='Placering')
        self.notebook.add(self.nameframe, text='Klasslista')

        # self.textarea = ScrolledText(self.nameframe, name='textarea')
        self.textarea = TtkTextArea(self.nameframe, name='textarea')
        self.textarea.pack(expand=True, side=TOP, fill='both')
        self.textarea.insert(1.0, 'Högerklicka här för att klistra in.')

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

        self.textarea.text.bind('<Button-1>', self.textbox_click_event)
        self.textarea.text.bind('<Button-2>', self.cmd_paste)
        self.textarea.text.bind('<Button-3>', self.cmd_paste)
        self.textarea.text.bind('<Control-v>', self.cmd_paste)

        self.setup_grid()

        # if OP_SYS == 'linux':
        #     s = ttk.Style()
        #     s.theme_use('clam')

        self.update_thread = Thread(target=self.periodic_stucount_update, daemon=True,
                                    args=(self.root, self.update_student_count, self.run_thread))
        self.update_thread.start()

        self.root.mainloop()

    # creates the buttons representing seats and grids them to self.buttonframe
    def setup_grid(self):
        for y in range(0, TOTAL_SEATS_Y):
            for x in range(0, TOTAL_SEATS_X):
                # s = Button(self.seatframe, text=str(x)+str(y))
                s = StudentSeat(self.seatframe, x, y, self, self.editvar)  # , name=str(x)+", "+str(y))
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
        # # print("xmax ymax =", xmax, ymax)
        return xmax, ymax

    # Threading function for updating student count
    def periodic_stucount_update(self, window, upd, run):
        while run:
            time.sleep(0.5)
            # t1 = time.perf_counter_ns()
            upd()
            # t2 = time.perf_counter_ns()
            # # print('%f ms'%(float((t2-t1)*10**(-6))))

    ########################
    #     BUTTON CALLS     #
    ########################

    # paste into textarea
    def cmd_paste(self, e):
        # prepare and clean up the list
        namelist_1 = [n.strip() for n in pyperclip.paste().split('\n') if '\n' not in n and n != '']
        # print('Namelist 1 len=', len(namelist_1), namelist_1)
        # print(namelist_1)
        namelist_2 = [n for n in namelist_1 if n != '']
        # print('Namelist 2 len=', len(namelist_2), namelist_2)

        # if platform is windows, sometimes åäö and accented chars get messed up. Try to encode then decode to resolve.
        if OP_SYS == 'windows':
            try:
                # print('Attempting to fix utf8 chars')
                templist = [n.encode('utf-8').decode('utf-8') for n in namelist_2]
                # print('Success!')
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
            # seat.name_set('')
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
            # # print('self.filepath = ', self.filepath, '   saveas = ', saveas)
            filepath = asksaveasfilename(parent=self.root, defaultextension=".dat",
                                         filetypes=[("Placeringsdata", "*.dat")])
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
                    f.write(filepath + ";")
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
                    f.write(filepath + ';')
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

    def room_bounds(self) -> tuple:
        max_x = max([s.xpos for s in self.seats])
        max_y = max([s.ypos for s in self.seats])
        return max_x, max_y


    # Finds clustered active seats near seat at x_pos, y_pos
    def seat_cluster(self, x_pos, y_pos, col_list) -> list[StudentSeat]:
        # the col_list is to avoid unnecesary overhead by calling seats_columnwise() repeatedly
        focus = self.seat_at(x_pos, y_pos)
        # print(f'Searching around {focus.name_get()} col:{x_pos}, row:{y_pos}')
        cluster = [focus]

        # search vertically first and for each vertical neighbour find its horizontal neighbours
        # it is sufficient to find only right neighbours
        v_neighbours = self.vertical_neighbours(x_pos, y_pos, col_list)
        if len(v_neighbours) > 0:
            cluster.extend(v_neighbours)
        # now cluster contains all vertical neighbours from x_pos, y_pos

        # search for all horizontal neighbours of the vertical neighbours and append them to the cluster
        h_cluster = list()
        for seat in cluster:
            # print(f'Calling v_n() with x:{seat.xpos}, y:{seat.ypos}')
            h_neighbours = self.horizontal_neighbours(seat.xpos, seat.ypos, col_list)
            if len(h_neighbours) > 0:
                h_cluster.extend(h_neighbours)
        cluster.extend(h_cluster)

        # search for all vertical neighbours of the horizontal neighbours and append them to the cluster
        for seat in h_cluster:
            v_neighbours = self.vertical_neighbours(seat.xpos, seat.ypos, col_list)
            for v_seat in v_neighbours:
                if v_seat in cluster:
                    v_neighbours.remove(v_seat)
            if len(v_neighbours) > 0:
                cluster.extend(v_neighbours)

        # search for all vertical neighbours again of and append them to the cluster
        # for seat in cluster:
        #     v_neighbours = self.vertical_neighbours(seat.xpos, seat.ypos, col_list)
        #     for v_seat in v_neighbours:
        #         if v_seat in cluster:
        #             v_neighbours.remove(v_seat)
        #     if len(v_neighbours) > 0:
        #         cluster.extend(v_neighbours)

        return cluster


    def vertical_neighbours(self, x_pos, y_pos, col_list):
        cluster = list()
        y_0 = y_pos
        # search downwards
        while y_pos < TOTAL_SEATS_Y - 1:
            neighb = col_list[x_pos][y_pos + 1]
            if neighb.active:
                # print(f'v_n(): found {neighb.varname.get()}')
                cluster.append(neighb)
                y_pos += 1
            else:
                break

        # search upwards
        y_pos = y_0
        while y_pos > 0:
            neighb = col_list[x_pos][y_pos - 1]
            if neighb.active and neighb not in cluster:
                cluster.append(neighb)
                y_pos -= 1
            else:
                break

        return cluster

    def horizontal_neighbours(self, x_pos, y_pos, col_list):
        cluster = list()
        while x_pos < TOTAL_SEATS_X - 1:
            # print(f'h_n(): x_pos:{x_pos} y_pos:{y_pos}')
            neighb = col_list[x_pos + 1][y_pos]
            if neighb.active:
                cluster.append(neighb)
                x_pos += 1
            else:
                break
        return cluster

    def get_seat_clusters(self):
        active_seats_cols = self.active_seats_columnwise()
        active_seats = self.active_seats()
        all_seats_cols = self.seats_columnwise()
        clusters = list()

        for col in active_seats_cols:
            for seat in col:
                if seat in active_seats:
                    cluster = self.seat_cluster(seat.xpos, seat.ypos, all_seats_cols)
                    clusters.append(cluster)
                    c_l = [s.seat_coords() for s in cluster]
                    # print(f'found cluster: {c_l}')
                    for c_seat in cluster:
                        try:
                            active_seats.remove(c_seat)
                        except:
                            print(f'ERROR: {c_seat.name_get()} could not be found, proceeding...')

        return clusters

    def seat_at(self, x_pos, y_pos):
        for seat in self.seats:
            if seat.xpos == x_pos and seat.ypos == y_pos:
                return seat
        return None



    def cmd_sort_columnwise(self):
        # sort the lists and use it to search and fill the active seats
        # for each column in the list
        #   for each element in the column
        #       insert a name from sorted namelist
        columns = self.seats_columnwise()
        names = list(self.name_tuple())
        names.sort()
        seatnr = 0
        self.cmd_unplace()
        for column in columns:
            for seat in column:
                seat.name_set(names[seatnr])
                seatnr += 1

    def cmd_sort_columnwise_cluster(self):
        names = list(self.name_tuple(allow_improper=True))
        # dont place improper names first - pop them an add last
        improper = list()
        i = 0
        for name in names:
            if '.' in name:
                print(f'name {name} contains .')
                improper.append(names.pop(i)[1:])
            i += 1
        i = 0
        # due to a bug somewhere in python the char '.' is not always recognized in the first loop, so run it again
        for name in names:
            if '.' in name:
                print(f'name {name} contains .')
                improper.append(names.pop(i)[1:])
            i += 1
        names.sort()
        improper.sort()
        names.extend(improper)
        print(improper)
        print(names)

        self.fill_textarea(names)

        n_active = self.num_active()
        if n_active == 0 or not names:
            return
        clusters = self.get_seat_clusters()
        self.cmd_unplace()
        name_nr = 0
        for cluster in clusters:
            for seat in cluster:
                seat.name_set(names[name_nr])
                name_nr += 1
                # stop placing names when the namelist is depleted or the active seats are depleted
                if name_nr >= len(names) or name_nr >= n_active:
                    return

    def active_seats_columnwise(self):
        active_seats = self.active_seats()
        bounds = self.seat_bounds()

        # make a list of lists with seats row in each column. It will be unordered
        # [
        #    col1       col2    ....
        #    [3,        [5,
        #     0,         2,
        #     4,         0,
        #     ...]       ...]
        # ]

        columns = list()

        for i in range(bounds[0]+1):
            column = list()
            for seat in active_seats:
                if seat.xpos == i:
                    column.append(seat)
            column.sort(key=lambda s: s.ypos)
            columns.append(column)
        return columns

    def seats_columnwise(self):
        bounds = (TOTAL_SEATS_X - 1, TOTAL_SEATS_Y - 1)

        # make a list of lists with seats row in each column. It will be unordered
        # [
        #    col1       col2    ....
        #    [3,        [5,
        #     0,         2,
        #     4,         0,
        #     ...]       ...]
        # ]

        columns = list()

        for i in range(bounds[0]+1):
            column = list()
            # print(f'Column {i}')
            for seat in self.seats:
                if seat.xpos == i:
                    # print(f'[{seat.name_get()}]', end=", ")
                    column.append(seat)
            # print("")
            column.sort(key=lambda s: s.ypos)
            columns.append(column)
        return columns


    def cmd_rand(self):
        # clear names
        for seat in self.seats:
            seat.name_set("")

        # get names
        names = list(self.name_tuple())

        # rand names
        random.shuffle(names)

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
            if len(self.manwin.names) > 0:
                self.manwin.win.deiconify()

    def cmd_export(self):
        filepath = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2007-", "*.xlsx")])
        if not filepath:
            return
        else:
            self.export_xlsx(filepath)

    def cmd_orientation_n(self):
        self.northbutton.config(state='disabled')
        self.southbutton.config(state='enabled')
        self.rotate_seats_180()
        self.flip_whiteboard()
        self.dirty = True

    def cmd_orientation_s(self):
        self.southbutton.config(state='disabled')
        self.northbutton.config(state='enabled')
        self.rotate_seats_180()
        self.flip_whiteboard()
        self.dirty = True

    def flip_whiteboard(self):
        orientation = self.orientationvar.get()
        self.whiteboardframe.grid_forget()
        if orientation == 'n':
            self.whiteboardframe.grid(row=0, column=0, pady=(10,15))
        else:
            self.whiteboardframe.grid(row=2, column=0, columnspan=2, pady=(15, 0))

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
                for filepath in self.prev_files:
                    f.write(filepath + ';')
            else:
                f.write('')

        # terminate thread then exit
        self.run_thread = False
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
        # only push alpa chars to searchbox
        if not e.char.isalpha() or self.notebook.tab(self.notebook.select(),
                                                     'text') == 'Klasslista':  # self.notebook.select()):
            return

        # have to catch exception 'cause the save file dialog causes exceptions to be thrown when the key listener
        # asks for what's in focus
        try:
            # set focus on the searchbox
            if 'searchEntry' not in str(self.root.focus_get()):
                self.search_on_enter(None)
                self.search_var.set(e.char)
                self.search_entry.icursor(tkinter.END)
                self.cmd_search(None)
                self.search_entry.focus_set()
        except Exception:
            pass

    ########################
    # CONTROLLER FUNCTIONS #
    ########################

    def update_student_count(self):
        # total number of students in list
        tot_stus = len(self.name_tuple())

        # total number of placed students
        placed_stus = self.num_active(only_taken=True)

        self.countvar.set('Antal placerade: %s   Antal i klasslistan: %s' % (placed_stus, tot_stus))

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
            self.search_on_leave(None)

    def change_grid(self, change: str):
        """Add or remove rows at the bottom and columns to the right"""
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
                if s.ypos == max_y - 1:
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
                if s.xpos == max_x - 1:
                    s.grid_forget()
                    s.destroy()
                    self.seats.pop(i)
                    continue
                i += 1
            Constants.TOTAL_SEATS_X -= 1

    def is_proper_name(self, name: str):
        if 'Högerklicka' in name:
            return False
        elif name.isalnum():
            return True
        elif ' ' in name.strip():
            return True
        elif '-' in name.strip():
            return True
        else:
            return False

    # read the names from the textbox in the GUI and return them in a tuple
    def name_tuple(self, allow_improper=False) -> tuple:
        namesstr = self.textarea.get(1.0, END).strip()
        tempnames = str.split(namesstr, "\n")

        clean_names: list[str] = [name for name in tempnames if self.is_proper_name(name)]

        return (name for name in tempnames if name) if allow_improper else tuple(clean_names)

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
                if seat.active and self.is_proper_name(seat.name_get()):
                    n += 1
        else:
            for seat in self.seats:
                if seat.active:
                    n += 1
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

    def swap_seats(self, seat1: StudentSeat, seat2: StudentSeat):
        self.dirty = True
        name1 = seat1.name_get()
        name2 = seat2.name_get()
        act1 = seat1.active
        act2 = seat2.active
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

    def edit_seat(self, x: int, y: int, newname: str, newactive: bool):
        for seat in self.seats:
            if seat.xpos == x and seat.ypos == y:
                seat.name_set(newname)
                if newactive:
                    seat.activate()
                else:
                    seat.deactivate()

    def fill_textarea(self, stulist):
        # replace textarea with loaded students
        stustring = ""
        for stu in stulist:
            stustring += stu + "\n"

        self.textarea.delete(1.0, tkinter.END)
        self.textarea.insert(1.0, stustring)

    def write_data(self, filepath: str):
        # prepare strings from student list
        students = self.textarea.get(1.0, tkinter.END).split("\n")

        # rotate the classroom if orientation is south, to always save in north position
        south_facing = self.orientationvar.get() == 's'
        if south_facing:
            self.northbutton.invoke()

        # prepare strings of taken seats.
        taken_seats: list[str] = []
        for seat in self.seats:
            if seat.active:
                #                       row                    #col                    #name
                taken_seats.append(str(seat.ypos) + "," + str(seat.xpos) + "," + seat.name_get() + ";")

        with open(filepath, 'w') as f:
            # write list of students, like so:
            # name1;name2;name3...
            f.write("STUDENT LIST\n")
            for student in students:
                if student:
                    f.write(student + ";")

            # write list of active seats with coords and name, like so:
            # row,col,name;row,col,name
            f.write("\nACTIVE SEATS\n")
            for seat in taken_seats:
                f.write(seat)

            # return to south
            if south_facing:
                self.southbutton.invoke()

            # write the orientation
            f.write("\nORIENTATION\n")
            f.write(self.orientationvar.get())

    def update_combobox(self, newpath: str = ""):
        filenames = []
        for filepath in self.prev_files:
            filenames.append(filepath.split("/")[-1])
        filenames.sort()

        self.previouscombobox['values'] = tuple(filenames)
        self.combovar.set(newpath.split("/")[-1])

    def combobox_event(self, e):
        sel = self.combovar.get()
        for filepath in self.prev_files:
            if filepath.split("/")[-1] == sel and sel and sel != "\n":
                self.load_data(filepath)
                self.root.title(filepath.split('/')[-1])

    def textbox_click_event(self, e):
        namelist = self.name_tuple()
        # use the fact that the initial text "Högerklicka för att..." isnt a proper name and gives a 0 length name tuple
        if len(namelist) == 0:
            self.textarea.delete('1.0', tkinter.END)
        else:
            self.textarea.focus_set()

    def remove_selected_from_combobox(self):
        self.prev_files.remove(self.filepath)
        self.update_combobox()

    def load_data(self, filepath: str):
        stus = ""
        acts = ""
        orient = ""
        test_path = Path(filepath)
        if not test_path.is_file():
            ans = messagebox.askyesno(title='Filfel', message='Kunde inte öppna filen\n' + filepath +
                                                              '.\nDen kanske har blivit borttagen eller flyttad.\n\n' +
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

            # must see if we're opening legacy file without orientation data
            f.readline() # discard heading
            orient = f.readline()
            if orient:
                print(f'orient={orient}')
            else:
                print("Opening a legacy file without orientation data")
                print("defaulting to north")

            if orient not in ("n", "s"):
                # print(f"Orientation data is corrput: orient={orient}")
                # print("defaulting to north")
                if self.orientationvar.get() == "s":
                    self.northbutton.invoke()

        # remove trailing ;
        if len(stus) >= 2:
            stus = stus[0:-1] if stus[-1] == ";" else stus
        if len(acts) >= 2:
            acts = acts[0:-1] if acts[-1] == ";" else acts

        stulist = stus.split(";")
        actslist = acts.split(";")

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
            try:
                y = int(actsplit[0])
                x = int(actsplit[1])
            except:
                continue
            while y >= Constants.TOTAL_SEATS_Y:
                self.change_grid('+r')
            while x >= Constants.TOTAL_SEATS_X:
                self.change_grid('+c')
            self.edit_seat(int(actsplit[1]), int(actsplit[0]), actsplit[2], True)
        self.filepath = filepath
        self.dirty = False

        # close the manual placement window if its opened
        if self.editvar.get():
            self.editbutton.invoke()

        # rotate the classroom if it was saved in orientation south
        if orient == "s" and self.orientationvar.get() == "n":
            self.southbutton.invoke()
        elif orient == "n" and self.orientationvar.get() == "s":
            self.northbutton.invoke()
            self.rotate_seats_180() # for some reason the return to n misses this rotation. add it back.

    def rotate_seats_180(self):
        xmax, ymax = self.seat_bounds()

        new_order = []

        for seat in self.seats:
            newx = xmax - seat.xpos
            newy = ymax - seat.ypos

            for newseat in self.seats:
                if newseat.xpos == newx and newseat.ypos == newy and not seat.has_swapped:  # and seat.active:
                    seat.has_swapped = True
                    newseat.has_swapped = True
                    self.swap_seats(seat, newseat)

        for seat in self.seats:
            seat.has_swapped = False

        self.cmd_search(None)

    def export_xlsx(self, filepath: str):
        wb = Workbook(filepath)
        ws = wb.add_worksheet()
        ws.set_landscape()

        # create a heading (whiteboard)
        headingformat = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                       'fg_color': Constants.BOARD_COLOR,
                                       'font_size': '20'})
        # convert seats x coord. to letter
        xymax = self.seat_bounds()
        xmax = Constants.COLNUM[xymax[0]]
        ymax = xymax[1]
        headingformat.set_bottom()
        headingformat.set_top()
        headingformat.set_left()
        headingformat.set_right()
        headingformat.set_border(2)

        ortn = self.orientationvar.get()

        if ortn == 'n':
            ws.merge_range('B1:' + xmax + '1', 'Tavla', headingformat)
        else:
            # # print('ymax is', ymax)
            ymax = str(int(ymax) + 4)
            ws.merge_range('B' + ymax + ':' + xmax + ymax, 'Tavla', headingformat)

        tableformat = wb.add_format({'fg_color': '#E3D0B7',  # a lighter Constants.ACTIVE_COLOR,
                                     'align': 'center', 'valign': 'vcenter',
                                     'text_wrap': True, 'font_size': 20})
        tableformat.set_bottom()
        tableformat.set_top()
        tableformat.set_left()
        tableformat.set_right()
        buf_spc = 3 if ortn == 'n' else 0
        for seat in self.seats:
            if seat.active:
                # row        col
                ws.write(seat.ypos + buf_spc, seat.xpos, seat.name_get(), tableformat)
                ws.set_row(seat.ypos + buf_spc, 45)
                ws.set_column(seat.xpos, seat.xpos, 16)

        wb.close()
