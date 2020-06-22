from tkinter import*
from tkinter import messagebox
from PIL import ImageTk, Image, ImageOps
import math
import datetime
import threading
import os
from xlwt import Workbook


class Tasks:

    run = False

    def __init__(self, activity, description, h, m, s, delete):
        self.activity = activity
        self.description = description
        self.h = h
        self.m = m
        self.s = s
        self.delete = delete
        self.img = Image.open("Buttons/Start.png")
        self.img = self.img.resize((45,45))
        self.start_img = ImageTk.PhotoImage(self.img)
        self.img = Image.open("Buttons/Pause.png")
        self.img = self.img.resize((45,45))
        self.pause_img = ImageTk.PhotoImage(self.img)
        self.img = Image.open("Buttons/Reset.png")
        self.img = self.img.resize((45,45))
        self.reset_img = ImageTk.PhotoImage(self.img)
        self.img = Image.open("Buttons/Arrow.png")
        self.img = self.img.resize((35,12))
        self.uparrow_img = ImageTk.PhotoImage(self.img)
        self.img = Image.open("Buttons/Arrow.png")
        self.img = self.img.resize((35,12))
        self.img = ImageOps.flip(self.img)
        self.downarrow_img = ImageTk.PhotoImage(self.img)

    def newTask(self, frame):

        self.task_1 = LabelFrame(frame, padx=5, pady=7)
        self.task_1.pack(padx=5, pady=5)

        task_1_0 = Frame(self.task_1)
        task_1_0.grid(row=0, column=0, padx=(15,0), pady=7)

        var = IntVar()
        checkbtn = Checkbutton(task_1_0, variable = var)
        checkbtn.grid(row=0, column = 0)

        task_1_1 = Frame(self.task_1)
        task_1_1.grid(row=0, column=1, padx=7, pady=7)

        task_1_2 = Frame(self.task_1)
        task_1_2.grid(row=0, column=2, padx=7, pady=7)

        task_1_3 = Frame(self.task_1)
        task_1_3.grid(row=0, column=3, padx=4, pady=4)

        task_1_4 = Frame(self.task_1)
        task_1_4.grid(row=0, column=4, padx=(0,7), pady=7)

        # Labels
        activity_1L = Label(task_1_1, text="Activity:       ", pady=5).grid(row=1, column=1)
        self.activity_1E = Entry(task_1_1, width=40)
        self.activity_1E.grid(row=1, column=2)
        self.activity_1E.insert(0, self.activity)
        description_1L = Label(task_1_1, text="   Description:    ", pady=5).grid(row=2, column=1)
        self.description_1E = Entry(task_1_1, width=40)
        self.description_1E.grid(row=2, column=2)
        self.description_1E.insert(0, self.description)


        self.timelapsed = (self.h + ":" + self.m + ":" + self.s)
        self.timeunits = math.ceil((3600 * int(self.h) + 60 * int(self.m) + int(self.s)) / 360) / 10

        self.timelapsed_1 = Label(task_1_2, text=self.timelapsed, padx=5, pady=5, font=("Calibri 15"))
        self.timelapsed_1.grid(row=0, column=3)

        self.start_1 = Button(task_1_2, image =self.start_img, border =0,  command=lambda: self.Start())
        self.start_1.grid(row=0, column=0, padx=5, pady=5,)
        self.reset_1 = Button(task_1_2, image = self.reset_img, border =0, state='disabled', command=lambda: self.Reset())
        self.reset_1.grid(row=0, column=1)
        self.lblspace = Label(task_1_2, text = " ")
        self.lblspace.grid(row=0,column=2)

        self.hourup_1 = Button(task_1_3, image = self.uparrow_img, border =0, padx=5, pady=5, command=lambda: self.Manual(0))
        self.hourup_1.grid(row=0, column=0, padx=7, pady=2)
        self.hourdown_1 = Button(task_1_3, image = self.downarrow_img, border =0, padx=5, pady=5,
                            command=lambda: self.Manual(1))
        self.hourdown_1.grid(row=2, column=0, padx=7, pady=(2,7))

        self.labelhour = Label(task_1_3, text = "1.0", font=("Calibri 18 bold"))
        self.labelhour.grid(row=1, column=0)
        self.labelminute = Label(task_1_3, text = "0.1", font=("Calibri 18 bold"))
        self.labelminute.grid(row=1, column=1)

        self.minuteup_1 = Button(task_1_3, image = self.uparrow_img, border =0, padx=5, pady=5,
                            command=lambda: self.Manual(2))
        self.minuteup_1.grid(row=0, column=1, padx=7, pady=2)
        self.minutedown_1 = Button(task_1_3, image = self.downarrow_img, border =0, padx=5, pady=5,
                              command=lambda: self.Manual(3))
        self.minutedown_1.grid(row=2, column=1, padx=7, pady=(2,7))

        self.timeunits_1 = Label(task_1_4, text=self.timeunits, width = 4, padx=5, pady=5, font=("Calibri 25 bold"))
        self.timeunits_1.grid(row=0, column=0)

        autoupdate = True

        def refresh():
            if autoupdate == True:
                self.activity = self.activity_1E.get()
                self.description = self.description_1E.get()
                if var.get() == 1:
                    self.delete = 1
                elif var.get() == 0:
                    self.delete = 0
                self.activity_1E.after(50,refresh)
        refresh()

    def clearrows(self):
        self.description_1E.delete(0, 'end')
        self.description_1E.insert(0, "")
        self.activity_1E.delete(0, 'end')
        self.activity_1E.insert(0, "")


    def Manual(self, action):
        elapsed = self.timelapsed_1['text']
        self.h, self.m, self.s = map(int, elapsed.split(":"))
        if (self.h < 99) and action == 0:
            self.h += 1
        if action == 1:
            if self.h > 0:
                self.h -=1
            elif self.h == 0:
                self.m = 0
                self.s = 0
        if action == 2:
            if self.m < 54:
                self.m = self.m + 6
            elif self.m >= 54:
                self.m = self.m - 54
                self.h +=1
        if action == 3:
            if self.h == 0:
                if self.m < 6:
                    self.m = 0
                    self.s = 0
                if self.m >= 6:
                    self.m = self.m - 6
            elif self.h > 0:
                if self.m < 6:
                    self.m = self.m + 54
                    self.h -= 1
                elif self.m >= 6:
                    self.m = self.m - 6
        if (self.h < 10):
            self.h = str(0) + str(self.h)
        else:
            self.h = str(self.h)
        if (self.m < 10):
            self.m = str(0) + str(self.m)
        else:
            self.m = str(self.m)
        if (self.s < 10):
            self.s = str(0) + str(self.s)
        else:
            self.s = str(self.s)
        self.timelapsed_1['text'] = self.h + ":" + self.m + ":" + self.s
        self.timeunits_1['text'] = math.ceil((3600 * int(self.h) + 60 * int(self.m) + int(self.s)) / 360) / 10

    def starttime(self):
        def value():
            if self.run == True:
                elapsed = self.timelapsed_1['text']
                self.h, self.m, self.s = map(int, elapsed.split(":"))
                self.h = int(self.h)
                self.m = int(self.m)
                self.s = int(self.s)
                if (self.s < 59):
                    self.s += 1
                elif (self.s == 59):
                    self.s = 0
                    if (self.m < 59):
                        self.m += 1
                    elif (self.m == 59):
                        self.h += 1
                        self.m = 0
                if (self.h < 10):
                    self.h = str(0) + str(self.h)
                else:
                    self.h = str(self.h)
                if (self.m < 10):
                    self.m = str(0) + str(self.m)
                else:
                    self.m = str(self.m)
                if (self.s < 10):
                    self.s = str(0) + str(self.s)
                else:
                    self.s = str(self.s)
                elapsed = self.h + ":" + self.m + ":" + self.s
                elaspedunits = math.ceil((3600 * int(self.h) + 60 * int(self.m) + int(self.s)) / 360) / 10
                self.timelapsed_1['text'] = elapsed
                self.timeunits_1['text'] = elaspedunits
                self.timelapsed_1.after(1000, value)
        value()

    def Start(self):
        self.run = False
        self.run = True
        self.starttime()
        self.start_1.configure(image =self.pause_img, command=lambda: self.Stop())
        self.reset_1['state'] = 'normal'

    def Stop(self):
        self.start_1.configure(image =self.start_img, command=lambda: self.Start())
        self.reset_1['state'] = 'normal'
        self.run = False


    # For Reset
    def Reset(self):
        if self.run == False:
            self.timelapsed_1['text'] = '00:00:00'
            self.timeunits_1['text'] = '0.0'
            self.Stop()
            self.m = 0
            self.h = 0
            self.s = 0
            self.reset_1['state'] = 'disabled'
        else:
            self.run = False
            self.Stop()
            self.reset_1['state'] = 'disabled'
            self.timelapsed_1['text'] = '00:00:00'
            self.timeunits_1['text'] = '0.0'

#Jobcode Class

class Jobcode:

    def __init__(self, name, description):
        self.name = name
        self.description = description

#Main Window

root = Tk()
root.title("TimeTracker")
root.geometry("850x625")
root.resizable(0,0)
root.iconbitmap("Buttons/Icon.ico")


hotbartop = Frame(root,  padx=5, pady=5)
hotbartop.pack(side="top")

#MainCanvas

def onFrameConfigure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

def onCanvasConfigure(event):
    canvas_width = event.width
    canvas.itemconfig(canvas_window, width=canvas_width)

canvas = Canvas(borderwidth=0 , highlightthickness = 0)
viewPort = Frame(canvas)
vsb = Scrollbar(orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=vsb.set)
vsb.pack(side="right", fill ="y")
canvas.pack(side="top", fill="both", expand=True)
canvas_window = canvas.create_window((4,4), window=viewPort,anchor = "nw", tags="viewPort")

viewPort.bind("<Configure>", onFrameConfigure)
canvas.bind("<Configure>", onCanvasConfigure)
onFrameConfigure(None)

def _on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
canvas.bind_all("<MouseWheel>", _on_mousewheel)

def client_exit():
    exit()

#MainFrames


hotbarbottom = Frame(root,  padx=5, pady=5)
hotbarbottom.pack(side="bottom")

activities = Frame(viewPort)
activities.pack(side="top")

insertbtn = Button(hotbartop, text="New Task",  padx=3, pady=3, width = 10, command=lambda: addTask(activities))
insertbtn.grid(row = 0, column = 0, padx= 3, pady= (6,3))

#menubar
menu = Menu(root)
root.config(menu=menu)
file = Menu(menu)
file.add_command(label="New", command=lambda: addTask(activities))
file.add_command(label="Save", command=lambda: save())
file.add_command(label="View History", command=lambda: historical())
file.add_command(label="Exit", command=lambda: client_exit())

#edit = Menu(menu)
#menu.add_cascade(label="File", menu = file)
#menu.add_cascade(label="Edit", menu = edit)

#newtasks

task = []
now = str(datetime.datetime.now())
now = now[:-7]
date = now[:-9]
task_no = 1

#jcvariables
jcnum = 0
jclist = []

try:
    os.mkdir(os.getenv('APPDATA')+"\TimeTracker")
except OSError:
    pass

#Loading
try:
    i = 0
    f = open(os.getenv('APPDATA')+"\TimeTracker\\"+date + "-timetracker.txt", "r")
    info = f.readlines()
    info = list(filter(('00:00:00::\n').__ne__, info))
    while i < len(info):
        h, m, s, a, d = map(str, str(info[i]).strip("\n").split(":"))
        task.append(Tasks(a, d, h, m, s, 0))
        i += 1
        task[task_no - 1].newTask(activities)
        task_no += 1
    f.close()
except FileNotFoundError:
    pass
except ValueError:
    pass
try:
    i = 0
    f = open(os.getenv('APPDATA') + "\TimeTracker\Jobcodes.txt")
    info = f.readlines()
    while i < len(info):
        j, d = map(str, str(info[i]).strip("\n").split(":"))
        jclist.append(Jobcode(j, d))
        i += 1
        jcnum += 1
    f.close()
except FileNotFoundError:
    pass
except ValueError:
    pass

def addTask (frame):
    global task_no
    if task_no < 21:
        task.append(Tasks("", "", "00", "00", "00", 0))
        task[task_no-1].newTask(activities)
        task_no += 1

lblblank2 = Label(hotbartop, text = "AutoSave: Enabled", width = 45, anchor="e")

def savetext():
    try:
        lblblank2.configure(text="AutoSave: Enabled", width = 45, anchor="e")
    except RuntimeError:
        pass


def save():
    now = str(datetime.datetime.now())
    now = now[:-7]
    date = now[:-9]
    f = open(os.getenv('APPDATA')+"\TimeTracker\\"+date + "-timetracker.txt", "w+")
    i = 0
    output_list = []
    while i < task_no - 1:
        output_list.append(str(task[i].h) + ":" + str(task[i].m) + ":" + str(task[i].s) + ":" + str(task[i].activity) + ":" + str(task[
            i].description)+"\n")
        i += 1
    f.writelines(output_list)
    f.close
    jcf = open(os.getenv('APPDATA')+"\TimeTracker\Jobcodes.txt", "w+")
    i = 0
    outputjc_list = []
    while i < jcnum:
        outputjc_list.append(str(jclist[i].name)+":"+str(jclist[i].description)+"\n")
        i += 1
    jcf.writelines(outputjc_list)
    lblblank2.configure(text = "Tasks Saved", width = 45, anchor="e")
    threading.Timer(3,savetext).start()

totaltimelbl = Label(hotbarbottom, text = "Total Time: 0.0", font=("Calibri 15 bold"))
totaltimelbl.pack()

def totaltime():
    i = 0
    totalunits = 0.0
    while i < task_no-1:
        totalunits = totalunits + math.ceil((3600 * float(task[i].h) + 60 * int(task[i].m) + int(task[i].s)) / 360) / 10
        i += 1
    totalunits = round(totalunits,1)
    totaltimelbl['text'] = "Total Time: " + str(totalunits)
    totaltimelbl.after(100,totaltime)
totaltime()

saveauto = True

def autosave():
    global root
    global saveauto
    save()
    if saveauto == True:
        root.after(60000, autosave)
autosave()

def jcwindow():

    global jcwindow
    global jcnum
    global jclist

    btnjobcode['state'] = 'disabled'

    jc = Toplevel()
    jc.wm_title("Jobcodes")
    jc.geometry("410x500")
    jc.resizable(0, 0)
    jc.iconbitmap("Buttons/Icon.ico")

    input = Frame(jc)
    input.grid(row=0, column =0, padx = 5, pady = 5)

    input1 = Frame(input)
    input1.grid(row=0, column =0)

    input2 = Frame(input)
    input2.grid(row=0, column =1)

    input3 = Frame(jc)
    input3.grid(row=1, column = 0)

    output = Frame (jc)
    output.grid(row=2, column =0, padx = 20, pady = 5)

    output2 = Frame (jc)
    output2.grid(row=3, column =0, padx = 20, pady = 5)

    def onselectjc(event):
        outputdes.selection_clear(0,END)
        outputdes.selection_set(outputjc.curselection())

    def onselectd(event):
        outputjc.selection_clear(0,END)
        outputjc.selection_set(outputdes.curselection())

    def OnMouseWheel(event):
        outputjc.yview("scroll", event.delta, "units")
        outputdes.yview("scroll", event.delta, "units")
        return "break"

    outputjc = Listbox(output, width = 25, height = 20, exportselection=0)
    outputjc.MultiSelect = 0
    outputjc.bind('<<ListboxSelect>>', onselectjc)
    outputjc.bind("<MouseWheel>", OnMouseWheel)
    outputjc.grid(row = 1, column =0)
    outputjclbl = Label (output, text = "Jobcode", anchor = "w", background = 'blue', foreground = 'white', width = 21)
    outputjclbl.grid(row=0,column =0)

    outputdes = Listbox(output, width = 35, height = 20, exportselection=0)
    outputdes.grid(row=1, column=1)
    outputdes.MultiSelect = 0
    outputdes.bind('<<ListboxSelect>>', onselectd)
    outputdes.bind("<MouseWheel>", OnMouseWheel)
    outputjcdes = Label (output, text = "Description", anchor = "w", background = 'blue', foreground = 'white', width = 30)
    outputjcdes.grid(row=0,column =1)

    jclbl = Label (input1, width = 9, text = "Jobcode:", anchor = "w")
    jclbl.grid(row = 0, column = 0, padx = 5, pady = 5)
    jcdlbl = Label (input1, width = 9, text = "Description:", anchor = "w")
    jcdlbl.grid(row = 1, column = 0, padx = 5, pady = 5)

    jcentry = Entry(input1, width = 25)
    jcdentry = Entry(input1, width = 25)

    lblmsg = Label(input3, width = 25, text = "")
    lblmsg.pack(pady = (0,5))

    jcentry.grid(row = 0, column = 1, padx = 5, pady = 5)
    jcdentry.grid(row = 1, column = 1, padx = 5, pady = 5)

    def loadjc():
        global jcnum
        global jclist
        i = 0
        while i < jcnum:
            outputjc.insert("end", jclist[i].name)
            outputdes.insert("end", jclist[i].description)
            i += 1

    loadjc()

    def insertjc ():
        global jcnum
        global jclist
        if jcentry.get() != "" and jcdentry.get() != "":
            i = 0
            duplicatefound = False
            value = jcentry.get()
            while i < jcnum:
                if value == jclist[i].name:
                    duplicatefound = True
                    lblmsg['text'] = "Duplicated Entry."
                i += 1
            if duplicatefound == False:
                jclist.append (Jobcode(jcentry.get(),jcdentry.get()))
                outputjc.insert("end", jcentry.get())
                outputdes.insert("end", jcdentry.get())
                jcnum += 1
                lblmsg['text'] = "Entry Added."
        else:
            lblmsg['text'] = "Invalid Entry."

    def deletejc():
        global jcnum
        global jclist
        try:
            index = outputjc.curselection()
            value = outputjc.get(index)
            i = 0
            while i < jcnum:
                if jclist[i].name == value:
                    popindex = i
                i += 1
            jclist.pop(popindex)
            outputjc.delete(index)
            outputdes.delete(index)
            outputjc.selection_clear(0,END)
            outputdes.selection_clear(0,END)
            jcnum -= 1
            lblmsg['text'] = "Entry Deleted."
        except TclError:
            pass
    def copyjc():
        try:
            index = outputjc.curselection()
            value = outputjc.get(index)
            r = Tk()
            r.withdraw()
            r.clipboard_clear()
            r.clipboard_append(value)
            r.update()
            r.destroy()
            lblmsg['text'] = "Entry Copied."
        except TclError:
            pass

    btnaddjc = Button(input2, text = "Insert", width = 7, command = lambda: insertjc())
    btnaddjc.pack(padx = 7, pady = 5)
    btncopyjc = Button(output2, text = "Copy", width = 10, command = lambda: copyjc())
    btncopyjc.grid(row=0, column = 0, padx = 5)
    btndeletejc = Button(output2, text = "Delete", width = 10, command = lambda: deletejc())
    btndeletejc.grid(row=0, column = 1, padx = 5)

    def on_closing():
        btnjobcode['state'] = 'normal'
        jc.destroy()

    jc.protocol("WM_DELETE_WINDOW", on_closing)

def historical():

    his = Toplevel()
    his.wm_title("Historical Data")
    his.geometry ("600x600")
    his.resizable(0,0)
    his.iconbitmap("Buttons/Icon.ico")
    btnhistorical['state'] = 'disabled'

    input = Frame(his)
    input.pack(side = 'top')

    output = Frame(his)
    output.pack(side='top')


    today = datetime.date.today()
    sunday = today + datetime.timedelta( (6-today.weekday()) % 7 -7)
    saturday = today + datetime.timedelta( (5-today.weekday()) % 7 )

    startdatelbl = Label(input, text ="From ").grid(row=0, column =0)
    startdateentry = Entry(input, width=12)
    startdateentry.grid(row=0, column = 1)
    startdateentry.insert(0, sunday)
    startdatelbl = Label(input, text=" to ").grid(row=0, column=2)
    enddateentry = Entry(input, width=12)
    enddateentry.grid(row=0, column = 3)
    enddateentry.insert(0, saturday)

    retrievebtn = Button(input, text = "Retrieve", command = lambda: retrieve(startdateentry.get(),enddateentry.get()))
    retrievebtn.grid(row=0, column = 4, padx=5, pady=5)

    datelbl = Label(output, text="Date", width = 10, background = 'blue', foreground = "white")
    datelbl.grid(row=1, column=0)

    activitylbl = Label(output, text="Activity", width = 20,  background = 'blue', foreground = "white")
    activitylbl.grid(row=1, column=2)

    desclbl = Label(output, text="Description", width = 40,  background = 'blue', foreground = "white")
    desclbl.grid(row=1, column=4)

    timelbl = Label(output, text="Time",width = 7 , background = 'blue', foreground = "white")
    timelbl.grid(row=1, column=6)

    topdivider1 = Label(output, text= "|",background = 'blue', foreground = "white")
    topdivider1.grid(row=1, column=1)
    topdivider2 = Label(output, text="|", background='blue', foreground="white")
    topdivider2.grid(row=1, column=3)
    topdivider3 = Label(output, text="|", background='blue', foreground="white")
    topdivider3.grid(row=1, column=5)


    botdivider1 = Label(output, text="")
    botdivider1.grid(row=2, column=1)
    botdivider2 = Label(output, text="")
    botdivider2.grid(row=2, column=3)
    botdivider3 = Label(output, text="")
    botdivider3.grid(row=2, column=5)

    datedata = Label(output, text= "", width = 10)
    datedata.grid(row=2, column=0)

    activitydata = Label(output, text= "", width = 20, anchor="w")
    activitydata.grid(row=2, column=2)

    descrdata = Label(output, text= "", width = 40, anchor="w")
    descrdata.grid(row=2, column=4)

    timedata = Label(output, text= "", width = 7)
    timedata.grid(row=2, column=6)

    btnexport = Button(input, text = "Export Excel", state = 'disabled', command = lambda: exportexcel())
    btnexport.grid(row=0, column = 5)

    datelist = []
    actlist = []
    descrlist = []
    timelist = []

    def on_closing():
        btnhistorical['state'] = 'normal'
        his.destroy()

    his.protocol("WM_DELETE_WINDOW", on_closing)

    def retrieve(startdate, enddate):
        dateoutput = ""
        activityoutput = ""
        descriptionoutput = ""
        timeoutput = ""
        divideroutput = ""
        try:
            sy, sm, sd = map(int, startdate.split("-"))
            ey, em, ed = map(int, enddate.split("-"))
            if (0 < sd < 32 and
                0 < ed < 32 and
                0 < sm < 13 and
                0 < em < 13 and
                2015 < sy < 2030 and
                2015 < ey < 2030
            ):
                mydates = []
                sdate = datetime.date(int(sy),int(sm),int(sd))
                edate = datetime.date(int(ey), int(em), int(ed))
                change = edate - sdate
                for i in range(change.days+1):
                    day = sdate+datetime.timedelta(days=i)
                    mydates.append(str(day))
                datenum = len(mydates)
                i = 0
                checkfile = ""
                retrieveddata = []
                while i < datenum:
                    checkfile = os.getenv('APPDATA') + "\TimeTracker\\" + mydates[i] + "-timetracker.txt"
                    try:
                        f = open(checkfile, "r")
                        saveddata = f.readlines()
                        f.close()
                        saveddata = list(filter(('00:00:00::\n').__ne__, saveddata))
                        j = 0
                        while j < len(saveddata):
                            retrieveddata.append(saveddata[j]+":" + mydates[i])
                            j += 1
                    except FileNotFoundError:
                        pass
                    i += 1
                datafound = len(retrieveddata)
                i = 0
                while i < datafound:
                    h, m, s, a, d, date = map(str, str(retrieveddata[i]).split(":"))
                    datelist.append(str(date))
                    descrlist.append(d)
                    actlist.append(a)
                    timelist.append(str(math.ceil((3600 * int(h) + 60 * int(m) + int(s)) / 360) / 10))
                    activityoutput = activityoutput + a + "\n"
                    descriptionoutput = descriptionoutput + d.strip("\n") + "\n"
                    timeoutput = timeoutput + str(math.ceil((3600 * int(h) + 60 * int(m) + int(s)) / 360) / 10) + "\n"
                    dateoutput = dateoutput + str(date) + "\n"
                    divideroutput = divideroutput + "|\n"
                    i += 1
                btnexport['state'] = 'normal'
            else:
                messagebox.showerror("error","Invalid Dates")
        except ValueError:
             messagebox.showerror("error","Invalid Dates")
        datedata['text'] = dateoutput
        timedata['text'] = timeoutput
        activitydata ['text'] = activityoutput
        descrdata ['text'] = descriptionoutput
        botdivider1 ['text'] = divideroutput
        botdivider2['text'] = divideroutput
        botdivider3['text'] = divideroutput

    def exportexcel():
        book = Workbook()
        sh = book.add_sheet("TimeTracker")
        sh.write(0,0, "Date")
        sh.write(0,1, "Activity")
        sh.write(0,2, "Description")
        sh.write(0,3, "Time")
        i = 0
        while i < len(datelist):
            sh.write(i+1, 0, str(datelist[i]))
            sh.write(i+1, 1, str(actlist[i]))
            sh.write(i+1, 2, str(descrlist[i]))
            sh.write(i+1, 3, str(timelist[i]))
            i += 1
        timenow = str(datetime.datetime.now())[:-7]
        timenow = timenow.replace(" ", "-").replace(":", "-")
        print(timenow)
        book.save(os.getenv('APPDATA')+"\TimeTracker\\"+ timenow + " output.xls")
        os.startfile(os.getenv('APPDATA')+"\TimeTracker\\"+ timenow + " output.xls")


def delete():
    i = 0
    global task_no
    temp = task_no
    poplist = []
    while i + 1 < task_no:
        if task[i].delete == 1:
            task[i].task_1.destroy()
            poplist.append(i)
            temp -= 1
        i += 1
    task_no = temp
    i = 0
    poplist.sort(reverse=True)
    while i < len(poplist):
        task.pop(poplist[i])
        i += 1

current_date = str(datetime.date.today())

def updatetime():
    global root
    global current_date
    test_date = str(datetime.date.today())
    if test_date != current_date:
        i = 0
        global task_no
        temp = task_no
        poplist = []
        while i + 1 < task_no:
            task[i].task_1.destroy()
            poplist.append(i)
            temp -= 1
            i += 1
        task_no = temp
        i = 0
        poplist.sort(reverse=True)
        while i < len(poplist):
            task.pop(poplist[i])
            i += 1
        current_date = test_date
    now = str(datetime.datetime.now())
    now = now[:-7]
    root.title("TimeTracker "+ now)
    root.after(1000,updatetime)
updatetime()

btnhistorical = Button(hotbartop, text="Historical", padx=3, pady=3, width = 10, command=lambda: historical())
btnhistorical.grid(row = 0, column = 1, padx = 3,pady = (6,3))
btnjobcode = Button(hotbartop, text="Jobcodes", padx=3, pady=3, width = 10, command=lambda: jcwindow())
btnjobcode.grid(row = 0, column = 2, padx = 3,pady = (6,3))
savebtn = Button(hotbartop, text="Manual Save", padx=3, pady=3, width = 10, command=lambda: save())
savebtn.grid(row = 0, column = 4, padx = 3,pady = (6,3))
deletebtn = Button(hotbartop, text="Delete", padx=3, pady=3, width = 10, command=lambda: delete())
deletebtn.grid(row = 0, column = 3, padx = 3,pady = (6,3))
lblblank2.grid(row = 0, column = 5, pady = (6,3))

root.mainloop()

