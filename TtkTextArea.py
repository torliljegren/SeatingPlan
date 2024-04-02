import tkinter as tk
import tkinter.ttk as ttk

class TtkTextArea(ttk.LabelFrame):
    def __init__(self, master, name=None):
        self.lbl = ttk.Label(master=master, text='Ett namn per rad')
        super().__init__(master, name=name, labelwidget=self.lbl)
        self.text = tk.Text(self, wrap="none")
        self.vsb = ttk.Scrollbar(self, command=self.text.yview, orient="vertical")
        # self.hsb = ttk.Scrollbar(self, command=self.text.xview, orient="horizontal")
        self.text.configure(yscrollcommand=self.vsb.set)#, xscrollcommand=self.hsb.set)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.vsb.grid(row=0, column=1, sticky="ns")
        # self.hsb.grid(row=1, column=0, sticky="ew")
        self.text.grid(row=0, column=0, sticky="nsew")

    def insert(self, index: float, chars: str):
        self.text.insert(index, chars)

    def delete(self, index1, index2):
        self.text.delete(index1, index2)

    def get(self, index1, index2):
        return self.text.get(index1, index2)

