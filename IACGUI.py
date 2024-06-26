#! /usr/bin/env python3
import sys, os.path, json
import tkinter as tk
from tkinter import scrolledtext, filedialog
from tkcalendar import DateEntry
from tktooltip import ToolTip
from easydict import EasyDict
from Shared.Utility import Utility
from Shared.Compiler import Compiler

class Application(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        parent.geometry("1080x540")
        parent.resizable(False,  False)
        parent.title("IAC Automated Report GUI")
        parent.eval('tk::PlaceWindow . center')
        
        # Initialize variables
        self.EnergyChartsPath = tk.StringVar(self.parent, "")
        self.RecommendationPath = tk.StringVar(self.parent, "")
        self.ReportPath = tk.StringVar(self.parent, "")
        # Read visit information from Compiler.json
        self.ReadInfo()

        self.Labelframe1 = tk.LabelFrame(self.parent)
        self.Labelframe1.place(x=20, y=10, height=510, width=340)
        self.Labelframe1.configure(text="Visit Information")

        LblInfoX=5
        LblInfoY=10
        LblInfoH=20
        LblInfoW=120
        LblInfoGapY=26

        self.Labelinfo01 = tk.Label(self.Labelframe1)
        self.Labelinfo01.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo01.configure(anchor='w',text="Report Number")
        ToolTip(self.Labelinfo01, msg=self.info.LE.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo02 = tk.Label(self.Labelframe1)
        self.Labelinfo02.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo02.configure(anchor='w',text="Visit Date")
        ToolTip(self.Labelinfo02, msg=self.info.VDATE.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo03 = tk.Label(self.Labelframe1)
        self.Labelinfo03.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo03.configure(anchor='w',text="Plant Location")
        ToolTip(self.Labelinfo03, msg=self.info.LOC.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo04 = tk.Label(self.Labelframe1)
        self.Labelinfo04.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo04.configure(anchor='w',text="SIC Code")
        ToolTip(self.Labelinfo04, msg=self.info.SIC.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo05 = tk.Label(self.Labelframe1)
        self.Labelinfo05.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo05.configure(anchor='w',text="NAICS Code")
        ToolTip(self.Labelinfo05, msg=self.info.NAICS.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo06 = tk.Label(self.Labelframe1)
        self.Labelinfo06.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo06.configure(anchor='w',text="Annual Sales ($)")
        ToolTip(self.Labelinfo06, msg=self.info.SALE.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo07 = tk.Label(self.Labelframe1)
        self.Labelinfo07.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo07.configure(anchor='w',text="No. of Employees")
        ToolTip(self.Labelinfo07, msg=self.info.EMPL.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo08 = tk.Label(self.Labelframe1)
        self.Labelinfo08.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo08.configure(anchor='w',text="Plant Area (sqft)")
        ToolTip(self.Labelinfo08, msg=self.info.AREA.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo09 = tk.Label(self.Labelframe1)
        self.Labelinfo09.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo09.configure(anchor='w',text="Principal Product")
        ToolTip(self.Labelinfo09, msg=self.info.PROD.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo10 = tk.Label(self.Labelframe1)
        self.Labelinfo10.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo10.configure(anchor='w',text="Annual Production")
        ToolTip(self.Labelinfo10, msg=self.info.ANPR.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo11 = tk.Label(self.Labelframe1)
        self.Labelinfo11.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo11.configure(anchor='w',text="Production Units")
        ToolTip(self.Labelinfo11, msg=self.info.PRUN.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo12 = tk.Label(self.Labelframe1)
        self.Labelinfo12.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo12.configure(anchor='w',text="Production Hours")
        ToolTip(self.Labelinfo12, msg=self.info.PROH.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo13 = tk.Label(self.Labelframe1)
        self.Labelinfo13.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo13.configure(anchor='w',text="Office Hours")
        ToolTip(self.Labelinfo13, msg=self.info.OFOH.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo14 = tk.Label(self.Labelframe1)
        self.Labelinfo14.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo14.configure(anchor='w',text="Lead Professor")
        ToolTip(self.Labelinfo14, msg=self.info.PROF.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo15 = tk.Label(self.Labelframe1)
        self.Labelinfo15.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo15.configure(anchor='w',text="Lead Student")
        ToolTip(self.Labelinfo15, msg=self.info.LEAD.comment)
    
        LblInfoY+=LblInfoGapY
        self.Labelinfo16 = tk.Label(self.Labelframe1)
        self.Labelinfo16.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo16.configure(anchor='w',text="Safety Student")
        ToolTip(self.Labelinfo16, msg=self.info.SAFE.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo17 = tk.Label(self.Labelframe1)
        self.Labelinfo17.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo17.configure(anchor='w',text="Participants")
        ToolTip(self.Labelinfo17, msg=self.info.PART.comment)

        LblInfoY+=LblInfoGapY
        self.Labelinfo18 = tk.Label(self.Labelframe1)
        self.Labelinfo18.place(x=LblInfoX, y=LblInfoY, height=LblInfoH, width=LblInfoW)
        self.Labelinfo18.configure(anchor='w',text="Contributors")
        ToolTip(self.Labelinfo18, msg=self.info.CONT.comment)

        EtrInfoX=125
        EtrInfoY=10
        EtrInfoH=22
        EtrInfoW=200

        self.Entryinfo01 = tk.Entry(self.Labelframe1, textvariable=self.LE)
        self.Entryinfo01.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)
        
        EtrInfoY+=LblInfoGapY
        self.Entryinfo02 = DateEntry(self.Labelframe1, textvariable=self.VDATE, date_pattern='mm/dd/yyyy',)
        self.Entryinfo02.place(x=EtrInfoX+1, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW-1)
        self.Entryinfo02.set_date(self.info.VDATE.value)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo03 = tk.Entry(self.Labelframe1, textvariable=self.LOC)
        self.Entryinfo03.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)
       
        EtrInfoY+=LblInfoGapY
        self.Entryinfo04 = tk.Entry(self.Labelframe1, textvariable=self.SIC)
        self.Entryinfo04.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo05 = tk.Entry(self.Labelframe1, textvariable=self.NAICS)
        self.Entryinfo05.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo06 = tk.Entry(self.Labelframe1, textvariable=self.SALE)
        self.Entryinfo06.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo07 = tk.Entry(self.Labelframe1, textvariable=self.EMPL)
        self.Entryinfo07.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo08 = tk.Entry(self.Labelframe1, textvariable=self.AREA)
        self.Entryinfo08.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo09 = tk.Entry(self.Labelframe1, textvariable=self.PROD)
        self.Entryinfo09.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo10 = tk.Entry(self.Labelframe1, textvariable=self.ANPR)
        self.Entryinfo10.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo11 = tk.Entry(self.Labelframe1, textvariable=self.PRUN)
        self.Entryinfo11.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo12a = tk.Entry(self.Labelframe1, textvariable=self.PROHH)
        self.Entryinfo12a.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=50)
        self.Labelinfo12a = tk.Label(self.Labelframe1, text="H/D")
        self.Labelinfo12a.place(x=EtrInfoX+50, y=EtrInfoY, height=EtrInfoH, width=30)

        self.Entryinfo12b = tk.Entry(self.Labelframe1, textvariable=self.PROHD)
        self.Entryinfo12b.place(x=EtrInfoX+80, y=EtrInfoY, height=EtrInfoH, width=30)
        self.Labelinfo12b = tk.Label(self.Labelframe1, text="D/W")
        self.Labelinfo12b.place(x=EtrInfoX+110, y=EtrInfoY, height=EtrInfoH, width=30)

        self.Entryinfo12c = tk.Entry(self.Labelframe1, textvariable=self.PRODW)
        self.Entryinfo12c.place(x=EtrInfoX+140, y=EtrInfoY, height=EtrInfoH, width=30)
        self.Labelinfo12c = tk.Label(self.Labelframe1, text="W/Y")
        self.Labelinfo12c.place(x=EtrInfoX+170, y=EtrInfoY, height=EtrInfoH, width=30)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo13a = tk.Entry(self.Labelframe1, textvariable=self.OFOHH)
        self.Entryinfo13a.place(x=EtrInfoX, y=EtrInfoY, height=EtrInfoH, width=50)
        self.Labelinfo13a = tk.Label(self.Labelframe1, text="H/D")
        self.Labelinfo13a.place(x=EtrInfoX+50, y=EtrInfoY, height=EtrInfoH, width=30)

        self.Entryinfo13b = tk.Entry(self.Labelframe1, textvariable=self.OFOHD)
        self.Entryinfo13b.place(x=EtrInfoX+80, y=EtrInfoY, height=EtrInfoH, width=30)
        self.Labelinfo13b = tk.Label(self.Labelframe1, text="D/W")
        self.Labelinfo13b.place(x=EtrInfoX+110, y=EtrInfoY, height=EtrInfoH, width=30)

        self.Entryinfo13c = tk.Entry(self.Labelframe1, textvariable=self.OFODW)
        self.Entryinfo13c.place(x=EtrInfoX+140, y=EtrInfoY, height=EtrInfoH, width=30)
        self.Labelinfo13c = tk.Label(self.Labelframe1, text="W/Y")
        self.Labelinfo13c.place(x=EtrInfoX+170, y=EtrInfoY, height=EtrInfoH, width=30)

        EtrInfoY+=LblInfoGapY
        Professors = ["Dr. Alparslan Oztekin", "Dr. Sudhakar Neti", "Dr. Ebru Demir"]
        self.Entryinfo14 = tk.OptionMenu(self.Labelframe1, self.PROF, *Professors)
        self.Entryinfo14.place(x=EtrInfoX+1, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW-1)

        EtrInfoY+=LblInfoGapY
        Students= ["Justin Caspar", "Tong Su", "Orhan Kaya", "Guanyang Xue", "Muhannad Altimemy", "Direnc Akyildiz"]
        self.Entryinfo15 = tk.OptionMenu(self.Labelframe1, self.LEAD, *Students)
        self.Entryinfo15.place(x=EtrInfoX+1, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW-1)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo16 = tk.OptionMenu(self.Labelframe1, self.SAFE, *Students)
        self.Entryinfo16.place(x=EtrInfoX+1, y=EtrInfoY, height=EtrInfoH, width=EtrInfoW-1)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo17a = tk.Label(self.Labelframe1, textvariable=self.NumPART)
        self.Entryinfo17a.place(x=EtrInfoX+10, y=EtrInfoY, height=EtrInfoH, width=80)
        self.Entryinfo17b = tk.Button(self.Labelframe1, text="Edit", command=lambda: self.EditPeople("PART"))
        self.Entryinfo17b.place(x=EtrInfoX+110, y=EtrInfoY-1, height=EtrInfoH, width=80)

        EtrInfoY+=LblInfoGapY
        self.Entryinfo18a = tk.Label(self.Labelframe1, textvariable=self.NumCONT)
        self.Entryinfo18a.place(x=EtrInfoX+10, y=EtrInfoY, height=EtrInfoH, width=80)
        self.Entryinfo18b = tk.Button(self.Labelframe1, text="Edit", command=lambda: self.EditPeople("CONT"))
        self.Entryinfo18b.place(x=EtrInfoX+110, y=EtrInfoY-1, height=EtrInfoH, width=80)


        self.Labelframe2 = tk.LabelFrame(self.parent, text="Workflow")
        self.Labelframe2.place(x=380, y=10, height=510, width=210)

        LblCompX=10
        CmdCompX=40
        LblCompW=180
        CmdCompW=120
        CmdCompH=28
        CompY=10
        CompGapY=10

        LblCompH=20
        self.Label1 = tk.Label(self.Labelframe2)
        self.Label1.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label1.configure(anchor='w')
        self.Label1.configure(text="1. Analyze EnergyChart.xlsx")
        CompY+=LblCompH

        self.Button1 = tk.Button(self.Labelframe2, command = self.LoadEnergyChart)
        self.Button1.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button1.configure(text="Locate File")
        CompY+=CmdCompH+CompGapY

        LblCompH=40
        self.Label2 = tk.Label(self.Labelframe2)
        self.Label2.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label2.configure(anchor='w',justify='left',wraplength="180")
        self.Label2.configure(text="2. Open EnergyChart.xlsx and save as Web Page (.htm)")

        CompY+=LblCompH

        self.Button2 = tk.Button(self.Labelframe2, command = self.OpenEnergyChart)
        self.Button2.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button2.configure(text="Open in Excel")
        CompY+=CmdCompH+CompGapY

        LblCompH=20
        self.Label3 = tk.Label(self.Labelframe2)
        self.Label3.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label3.configure(anchor='w')
        self.Label3.configure(text="3. ← Fill in visit information")
        CompY+=LblCompH

        self.Button3 = tk.Button(self.Labelframe2, command = self.SaveInfo)
        self.Button3.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button3.configure(text="Save info")
        CompY+=CmdCompH+CompGapY
        

        LblCompH=40
        self.Label4 = tk.Label(self.Labelframe2)
        self.Label4.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label4.configure(anchor='w',justify='left',wraplength="180")
        self.Label4.configure(text="4. Locate the folder including all recommendation .docx")
        CompY+=LblCompH
        
        self.Button4 = tk.Button(self.Labelframe2, command = self.LocateRecommendation)
        self.Button4.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button4.configure(text="Locate folder")
        CompY+=CmdCompH+CompGapY

        LblCompH=20
        self.Label5 = tk.Label(self.Labelframe2)
        self.Label5.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label5.configure(anchor='w')
        self.Label5.configure(text="5. Edit Description.docx")
        CompY+=LblCompH

        self.Button5 = tk.Button(self.Labelframe2, command = self.OpenDescription)
        self.Button5.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button5.configure(text="Open in Word")
        CompY+=CmdCompH+CompGapY

        LblCompH=20
        self.Label6 = tk.Label(self.Labelframe2)
        self.Label6.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label6.configure(anchor='w')
        self.Label6.configure(text="6. Compile the report")
        CompY+=LblCompH

        self.Button6 = tk.Button(self.Labelframe2, command=self.CompileReport)
        self.Button6.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button6.configure(text="Compile")
        CompY+=CmdCompH+CompGapY

        LblCompH=50
        self.Label7 = tk.Label(self.Labelframe2)
        self.Label7.place(x=LblCompX, y=CompY, height=LblCompH, width=LblCompW)
        self.Label7.configure(anchor='w',justify='left',wraplength="180")
        self.Label7.configure(text="7. Open the report, select all (ctrl+A); refresh (F9), OK; refresh (F9) again, OK.")
        CompY+=LblCompH

        self.Button7 = tk.Button(self.Labelframe2, command=self.OpenReport)
        self.Button7.place(x=CmdCompX, y=CompY, height=CmdCompH, width=CmdCompW)
        self.Button7.configure(text="Open in Word")
        CompY+=CmdCompH+CompGapY


        self.Labelframe3 = tk.LabelFrame(self.parent)
        self.Labelframe3.place(x=610, y=10, height=510, width=450)
        self.Labelframe3.configure(text="Command Line Output")
        self.Text1 = scrolledtext.ScrolledText(self.Labelframe3, state='disabled',wrap="word")
        self.Text1.place(x=10, y=10, height=460, width=430)
        self.Text1.tag_configure("stderr", foreground="#b22222")
        self.Laebl8 = tk.Label(self.Labelframe3)
        self.Laebl8.place(x=10, y=465, height=20, width=430)
        self.Laebl8.configure(text="Copyright © 2024 Lehigh University Industrial Assessment Center")
        sys.stdout = TextRedirector(self.Text1, "stdout")
        sys.stderr = TextRedirector(self.Text1, "stderr")

    def ReadInfo(self):
        self.info = EasyDict(json.load(open(os.path.join("Shared", "Compiler.json"))))
        self.LE = tk.StringVar(self.parent, self.info.LE.value)
        self.VDATE = tk.StringVar(self.parent, self.info.VDATE.value)
        self.LOC = tk.StringVar(self.parent, self.info.LOC.value)
        self.SIC = tk.StringVar(self.parent, self.info.SIC.value)
        self.NAICS = tk.StringVar(self.parent, self.info.NAICS.value)
        self.SALE = tk.StringVar(self.parent, self.info.SALE.value)
        self.EMPL = tk.StringVar(self.parent, self.info.EMPL.value)
        self.AREA = tk.StringVar(self.parent, self.info.AREA.value)
        self.PROD = tk.StringVar(self.parent, self.info.PROD.value)
        self.ANPR = tk.StringVar(self.parent, self.info.ANPR.value)
        self.PRUN = tk.StringVar(self.parent, self.info.PRUN.value)
        self.PROHH = tk.StringVar(self.parent, self.info.PROH.value[0])
        self.PROHD = tk.StringVar(self.parent, self.info.PROH.value[1])
        self.PRODW = tk.StringVar(self.parent, self.info.PROH.value[2])
        self.OFOHH = tk.StringVar(self.parent, self.info.OFOH.value[0])
        self.OFOHD = tk.StringVar(self.parent, self.info.OFOH.value[1])
        self.OFODW = tk.StringVar(self.parent, self.info.OFOH.value[2])
        self.PROF = tk.StringVar(self.parent, self.info.PROF.value)
        self.LEAD = tk.StringVar(self.parent, self.info.LEAD.value)
        self.SAFE = tk.StringVar(self.parent, self.info.SAFE.value)
        self.NumPART = tk.StringVar(self.parent, str(len(self.info.PART.value))+" people")
        self.NumCONT = tk.StringVar(self.parent, str(len(self.info.CONT.value))+" people")

    def SaveInfo(self):
        """
        Save visit information
        """
        self.info.LE.value = self.LE.get()
        self.info.VDATE.value = self.VDATE.get()
        self.info.LOC.value = self.LOC.get()
        self.info.SIC.value = self.SIC.get()
        self.info.NAICS.value = self.NAICS.get()
        self.info.SALE.value = int(self.SALE.get())
        self.info.EMPL.value = int(self.EMPL.get())
        self.info.AREA.value = int(self.AREA.get())
        self.info.PROD.value = self.PROD.get()
        self.info.ANPR.value = int(self.ANPR.get())
        self.info.PRUN.value = self.PRUN.get()
        self.info.PROH.value = [float(self.PROHH.get()), int(self.PROHD.get()), int(self.PRODW.get())]
        self.info.OFOH.value = [float(self.OFOHH.get()), int(self.OFOHD.get()), int(self.OFODW.get())]
        self.info.PROF.value = self.PROF.get()
        self.info.LEAD.value = self.LEAD.get()
        self.info.SAFE.value = self.SAFE.get()

        # save json
        with open(os.path.join("Shared", "Compiler.json"), 'w') as f:
            json.dump(self.info, f, indent=4)
        print("Saved visit information to Compiler.json")

    def EditPeople(self, mode):
        namelist=self.info[mode].value
        # popup a window with 8 entries
        # Create a new Tkinter window
        edit_window = tk.Toplevel()
        edit_window.title("Edit List of People")
        # Create entries
        for i in range(8):
            entry = tk.Entry(edit_window)
            entry.grid(row=i, column=0, padx=10, pady=5)
            if i < len(namelist):
                entry.insert(0, namelist[i])
        # Add a button to save the changes
        save_button = tk.Button(edit_window, text="Save", command=lambda: self.SavePeople(mode, edit_window))
        save_button.grid(row=9, column=0, columnspan=2, padx=10, pady=5)

    def SavePeople(self, mode, edit_window):
        # Update the list of names
        namelist=self.info[mode].value
        namelist.clear()
        for i in range(8):
            name = edit_window.grid_slaves(row=i, column=0)[0].get()
            if name != "":
                namelist.append(name)
        # Close the edit window
        edit_window.destroy()
        # Update the main window
        if mode == "PART":
            self.NumPART.set(str(len(namelist))+" people")
            print("Updated participants")
        else:
            self.NumCONT.set(str(len(namelist))+" people")
            print("Updated contributors")
        print(namelist)

    def LoadEnergyChart(self):
        """
        Select Energy Charts Dialog
        """
        path = filedialog.askopenfilename(initialdir=os.path.join(os.getcwd(), "EnergyCharts"), initialfile = "EnergyCharts.xlsx", title = "Select Energy Charts", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
        self.EnergyChartsPath.set(path)
        if path == "":
            return
        else:
            print("Selected energy charts:")
            print(path)
            Utility(path)
            print("Enegy info saved to Utility.json")
            print("")

    def OpenEnergyChart(self):
        """
        Open Energy Charts in Excel
        """
        Path = self.EnergyChartsPath.get()
        if Path == "":
            raise Exception("No energy charts selected")
        else:
            import subprocess, os, platform
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', Path))
            elif platform.system() == 'Windows':    # Windows
                os.startfile(Path)
            else:                                   # Linux
                subprocess.call(('xdg-open', Path))

    def LocateRecommendation(self):
        """
        Open recommendation folder
        """
        path = filedialog.askdirectory(initialdir = os.path.join(os.getcwd(), "Recommendations"), title = "Select Recommendation Folder")
        self.RecommendationPath.set(path)
        if path == "":
            raise Exception("No folder selected")
        else:
            print("Selected recommendation folder:")
            print(path)
            print("Detected .docx files:")
            for file in os.listdir(path):
                if file.endswith(".docx"):
                    print(file)
            print("")

    def OpenDescription(self):
        """
        Open Description in Word
        """
        import subprocess, os, platform
        DescriptionPath = os.path.join("Report", "Description.docx")
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', DescriptionPath))
        elif platform.system() == 'Windows':    # Windows
            os.startfile(DescriptionPath)
        else:                                   # Linux
            subprocess.call(('xdg-open', DescriptionPath))
    
    def CompileReport(self):
        """
        Compile the report
        """
        path = filedialog.asksaveasfilename(initialdir = os.getcwd(), initialfile = self.info.LE.value+".docx", title = "Save Final Report As", filetypes = (("Word files", "*.docx"), ("all files", "*.*")))
        self.ReportPath.set(path)
        if path == "":
            raise Exception("No file selected")
        else:
            print("Final report destination:")
            print(path)
            Compiler(os.path.dirname(self.EnergyChartsPath.get()), self.RecommendationPath.get(), self.ReportPath.get())

    def OpenReport(self):
        """
        Open Report in Word
        """
        path = self.ReportPath.get()
        if path == "":
            raise Exception("Report has not been compiled yet")
        else:
            import subprocess, os, platform
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', path))
            elif platform.system() == 'Windows':    # Windows
                os.startfile(path)
            else:                                   # Linux
                subprocess.call(('xdg-open', path))

class TextRedirector(object):
     def __init__(self, widget, tag="stdout"):
         self.widget = widget
         self.tag = tag

     def write(self, string):
         self.widget.configure(state="normal")
         self.widget.insert("end", string, (self.tag,))
         self.widget.configure(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    Application(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
