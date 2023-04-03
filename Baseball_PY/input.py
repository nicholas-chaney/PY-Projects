from tkinter import *
from openpyxl import *
from tkinter.ttk import *
import sys
import subprocess
import runpy

wb = load_workbook('file.xlsx')     #opens an excel workbook with the name in quotes
ws = wb.active                      #sets the first worksheet of the workbook as the active sheet/the sheet getting edited

root = Tk()                         #used for the user interface

topframe = Frame(root)              #used for the user interface
topframe.pack(side=TOP)             #used for the user interface



def inputPitchInfo():
    global result
    #Gets the input from textboxes and stores it to its own variable
    pitchname = textbox.get("1.0", "end-1c")        #gets pitcher name
    batname = textbox1.get("1.0", "end-1c")         #gets batter name
    pitch = textbox2.get("1.0", "end-1c")           #gets the pitch type
    pitchspd = textbox3.get("1.0", "end-1c")        #gets the pitch speed(mph)
    pitchlocation = textbox4.get("1.0", "end-1c")   #gets the location of the pitch(HI, HM, HA, MI, MM, MA, LI, LM, LA)
    result = textbox5.get("1.0", "end-1c")          #gets the outcome of the pitch(EX. foul, strike looking, strikeout)
    count = textbox6.get("1.0", "end-1c")           #gets the count that the pitch was thrown in

    ws['B1'] = pitchname.lower()
    ws.append([batname.lower(), pitch.upper(), pitchlocation.upper(), pitchspd, count, result.lower()])
    wb.save("file.xlsx")
    checkforhit()

    #textbox1.delete("1.0",END)
    #textbox2.delete("1.0", END)
    textbox3.delete("1.0", END)
    textbox4.delete("1.0", END)
    textbox5.delete("1.0", END)
    textbox6.delete("1.0", END)

def checkforhit():
    if result == "hit":
        wb.save('file.xlsx')
        wb.close()
        root.destroy()
        runpy.run_path("field.py")









pitcherName = Label(topframe, width=20, text="Pitcher Name:")
pitcherName.pack(side=TOP)

textbox = Text(topframe, height=1, width=20)
textbox.pack(side=TOP)

battername = Label(topframe, width=20, text="Batter Name:")
battername.pack(side=TOP)

textbox1 = Text(topframe, height=1, width=20)
textbox1.pack(side=TOP)

pitchtype = Label(topframe, width=20, text="Pitch Type:")
pitchtype.pack(side=TOP)

textbox2 = Text(topframe, height=1, width=20)
textbox2.pack(side=TOP)

pitchspeed = Label(topframe, width=20, text="Pitch Speed:")
pitchspeed.pack(side=TOP)

textbox3 = Text(topframe, height=1, width=20)
textbox3.pack(side=TOP)

pitchloc = Label(topframe, width=20, text="Pitch Location:")
pitchloc.pack(side=TOP)

textbox4 = Text(topframe, height=1, width=20)
textbox4.pack(side=TOP)

res = Label(topframe, width=20, text="Result:")
res.pack(side=TOP)

textbox5 = Text(topframe, height=1, width=20)
textbox5.pack(side=TOP)

cnt = Label(topframe, width=20, text="Count:")
cnt.pack(side=TOP)

textbox6 = Text(topframe, height=1, width=20)
textbox6.pack(side=TOP)

enterpitch = Button(topframe, text="enter pitch", width=25, command=lambda: inputPitchInfo())
enterpitch.pack(side=LEFT)

bottomframe = Frame(root)
bottomframe.pack(side=TOP)

root.mainloop()


wb.save("file.xlsx")
#os.system("file.xlsx")
