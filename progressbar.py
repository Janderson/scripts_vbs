from Tkinter import *
import Tkinter.

t = Tk()
progval = IntVar(t)
progmsg = StringVar(t); progmsg.set("Compute in progress...")
b = Button(t, relief=LINK, text="Quit (using bwidget)", command=t.destroy)
b.pack()
c = ProgressDialog(t, title="Please wait...",
                   type="infinite",
                   width=20,
                   stop="Stop",
                   textvariable=progmsg,
                   variable=progval,
                   command=lambda: c.destroy()
                   )
def update_progress():
       progval.set(2)
       c.after(20, update_progress)

update_progress()
t.mainloop()