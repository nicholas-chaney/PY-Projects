import runpy
from tkinter import *
from openpyxl import *
from tkinter.ttk import *
import os
import subprocess
import runpy

root = Tk()                         #used for the user interface


topframe = Frame(root)              #used for the user interface
topframe.pack(side=TOP)             #used for the user interface

NewChart_Button = Button(topframe, text="Open a Batter's Spraychart", width=25, command=lambda: runpy.run_path('graph.py'))
NewChart_Button.pack(side=LEFT)

NewChart_Button = Button(topframe, text="Open the pitcher's charts", width=25, command=lambda: runpy.run_path('PitcherGraphs.py'))
NewChart_Button.pack(side=LEFT)
root.mainloop()
