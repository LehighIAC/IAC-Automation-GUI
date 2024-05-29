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

        parent.geometry("1200x600")
        parent.resizable(False,  False)
        parent.title("IAC Automated Report GUI")
        parent.eval('tk::PlaceWindow . center')
        
        # Initialize variables
        self.EnergyChartsPath = tk.StringVar(self.parent, "")
        self.RecommendationPath = tk.StringVar(self.parent, "")
        self.ReportPath = tk.StringVar(self.parent, "")
        # Read visit information from Compiler.json
        self.ReadInfo()

        # Create widgets
        self.Labelframe1 = tk.LabelFrame(self.parent)
        self.Labelframe1.place(x=20, y=10, height=570, width=340)
        self.Labelframe1.configure(text="Visit Information")

        self.Labelinfo01 = tk.Label(self.Labelframe1, text="Report Number")
        self.Labelinfo01.grid(row=1, column=0, sticky='w')
        ToolTip(self.Labelinfo01, msg=self.info.LE.comment)

        self.Labelinfo02 = tk.Label(self.Labelframe1, text="Visit Date")
        self.Labelinfo02.grid(row=2, column=0, sticky='w')
        ToolTip(self.Labelinfo02, msg=self.info.VDATE.comment)

        self.Labelinfo03 = tk.Label(self.Labelframe1, text="Plant Location")
        self.Labelinfo03.grid(row=3, column=0, sticky='w')
        ToolTip(self.Labelinfo03, msg=self.info.LOC.comment)

        self.Labelinfo04 = tk.Label(self.Labelframe1, text="SIC Code")
        self.Labelinfo04.grid(row=4, column=0, sticky='w')
        ToolTip(self.Labelinfo04, msg=self.info.SIC.comment)

        self.Labelinfo05 = tk.Label(self.Labelframe1, text="NAICS Code")
        self.Labelinfo05.grid(row=5, column=0, sticky='w')
        ToolTip(self.Labelinfo05, msg=self.info.NAICS.comment)

        self.Labelinfo06 = tk.Label(self.Labelframe1, text="Annual Sales ($)")
        self.Labelinfo06.grid(row=6, column=0, sticky='w')
        ToolTip(self.Labelinfo06, msg=self.info.SALE.comment)

        self.Labelinfo07 = tk.Label(self.Labelframe1, text="No. of Employees")
        self.Labelinfo07.grid(row=7, column=0, sticky='w')
        ToolTip(self.Labelinfo07, msg=self.info.EMPL.comment)

        self.Labelinfo08 = tk.Label(self.Labelframe1, text="Plant Area (sqft)")
        self.Labelinfo08.grid(row=8, column=0, sticky='w')
        ToolTip(self.Labelinfo08, msg=self.info.AREA.comment)

        self.Labelinfo09 = tk.Label(self.Labelframe1, text="Principal Product")
        self.Labelinfo09.grid(row=9, column=0, sticky='w')
        ToolTip(self.Labelinfo09, msg=self.info.PROD.comment)

        self.Labelinfo10 = tk.Label(self.Labelframe1, text="Annual Production")
        self.Labelinfo10.grid(row=10, column=0, sticky='w')
        ToolTip(self.Labelinfo10, msg=self.info.ANPR.comment)

        self.Labelinfo11 = tk.Label(self.Labelframe1, text="Production Units")
        self.Labelinfo11.grid(row=11, column=0, sticky='w')
        ToolTip(self.Labelinfo11, msg=self.info.PRUN.comment)

        self.Labelinfo12 = tk.Label(self.Labelframe1, text="Production Hours")
        self.Labelinfo12.grid(row=12, column=0, sticky='w')
        ToolTip(self.Labelinfo12, msg=self.info.PROH.comment)

        self.Labelinfo13 = tk.Label(self.Labelframe1, text="Office Hours")
        self.Labelinfo13.grid(row=13, column=0, sticky='w')
        ToolTip(self.Labelinfo13, msg=self.info.OFOH.comment)

        self.Labelinfo14 = tk.Label(self.Labelframe1, text="Lead Professor")
        self.Labelinfo14.grid(row=14, column=0, sticky='w')
        ToolTip(self.Labelinfo14, msg=self.info.PROF.comment)

        self.Labelinfo15 = tk.Label(self.Labelframe1, text="Lead Student")
        self.Labelinfo15.grid(row=15, column=0, sticky='w')
        ToolTip(self.Labelinfo15, msg=self.info.LEAD.comment)
    
        self.Labelinfo16 = tk.Label(self.Labelframe1, text="Safety Student")
        self.Labelinfo16.grid(row=16, column=0, sticky='w')
        ToolTip(self.Labelinfo16, msg=self.info.SAFE.comment)

        self.Labelinfo17 = tk.Label(self.Labelframe1, text="Participants")
        self.Labelinfo17.grid(row=17, column=0, sticky='w')
        ToolTip(self.Labelinfo17, msg=self.info.PART.comment)

        self.Labelinfo18 = tk.Label(self.Labelframe1, text="Contributors")
        self.Labelinfo18.grid(row=18, column=0, sticky='w')
        ToolTip(self.Labelinfo18, msg=self.info.CONT.comment)

        self.Entryinfo01 = tk.Entry(self.Labelframe1, textvariable=self.LE)
        self.Entryinfo01.grid(row=1, column=1, sticky='w')
        
        self.Dateinfo02 = DateEntry(self.Labelframe1, textvariable=self.VDATE, date_pattern='mm/dd/yyyy')
        self.Dateinfo02.set_date(self.info.VDATE.value)
        self.Dateinfo02.grid(row=2, column=1)

        self.Entryinfo03 = tk.Entry(self.Labelframe1, textvariable=self.LOC)
        self.Entryinfo03.grid(row=3, column=1, sticky='w')
       
        self.Entryinfo04 = tk.Entry(self.Labelframe1, textvariable=self.SIC)
        self.Entryinfo04.grid(row=4, column=1, sticky='w')

        self.Entryinfo05 = tk.Entry(self.Labelframe1, textvariable=self.NAICS)
        self.Entryinfo05.grid(row=5, column=1, sticky='w')

        self.Entryinfo06 = tk.Entry(self.Labelframe1, textvariable=self.SALE)
        self.Entryinfo06.grid(row=6, column=1, sticky='w')

        self.Entryinfo07 = tk.Entry(self.Labelframe1, textvariable=self.EMPL)
        self.Entryinfo07.grid(row=7, column=1, sticky='w')

        self.Entryinfo08 = tk.Entry(self.Labelframe1, textvariable=self.AREA)
        self.Entryinfo08.grid(row=8, column=1, sticky='w')

        self.Entryinfo09 = tk.Entry(self.Labelframe1, textvariable=self.PROD)
        self.Entryinfo09.grid(row=9, column=1, sticky='w')

        self.Entryinfo10 = tk.Entry(self.Labelframe1, textvariable=self.ANPR)
        self.Entryinfo10.grid(row=10, column=1, sticky='w')

        self.Entryinfo11 = tk.Entry(self.Labelframe1, textvariable=self.PRUN)
        self.Entryinfo11.grid(row=11, column=1, sticky='w')

        self.Frameinfo12 = tk.Frame(self.Labelframe1)
        self.Frameinfo12.grid(row=12, column=1, sticky='w')
        self.Entryinfo12a = tk.Entry(self.Frameinfo12, textvariable=self.PROHH)
        self.Entryinfo12a.grid(row=0, column=0)
        self.Labelinfo12a = tk.Label(self.Frameinfo12, text="H/D")
        self.Labelinfo12a.grid(row=0, column=1)
        self.Entryinfo12b = tk.Entry(self.Frameinfo12, textvariable=self.PROHD)
        self.Entryinfo12b.grid(row=0, column=2)
        self.Labelinfo12b = tk.Label(self.Frameinfo12, text="D/W")
        self.Labelinfo12b.grid(row=0, column=3)
        self.Entryinfo12c = tk.Entry(self.Frameinfo12, textvariable=self.PRODW)
        self.Entryinfo12c.grid(row=0, column=4)
        self.Labelinfo12c = tk.Label(self.Frameinfo12, text="W/Y")
        self.Labelinfo12c.grid(row=0, column=5)

        self.Frameinfo13 = tk.Frame(self.Labelframe1)
        self.Frameinfo13.grid(row=13, column=1, sticky='w')
        self.Entryinfo13a = tk.Entry(self.Frameinfo13, textvariable=self.OFOHH)
        self.Entryinfo13a.grid(row=0, column=0)
        self.Labelinfo13a = tk.Label(self.Frameinfo13, text="H/D")
        self.Labelinfo13a.grid(row=0, column=1)
        self.Entryinfo13b = tk.Entry(self.Frameinfo13, textvariable=self.OFOHD)
        self.Entryinfo13b.grid(row=0, column=2)
        self.Labelinfo13b = tk.Label(self.Frameinfo13, text="D/W")
        self.Labelinfo13b.grid(row=0, column=3)
        self.Entryinfo13c = tk.Entry(self.Frameinfo13, textvariable=self.OFODW)
        self.Entryinfo13c.grid(row=0, column=4)
        self.Labelinfo13c = tk.Label(self.Frameinfo13, text="W/Y")
        self.Labelinfo13c.grid(row=0, column=5)
        
        Professors = ["Dr. Alparslan Oztekin", "Dr. Sudhakar Neti", "Dr. Ebru Demir"]
        self.Optioninfo14 = tk.OptionMenu(self.Labelframe1, self.PROF, *Professors)
        self.Optioninfo14.grid(row=14, column=1)

        Students= ["Tong Su", "Orhan Kaya", "Guanyang Xue", "Muhannad Altimemy", "Direnc Akyildiz", "Justin Caspar"]
        self.Optioninfo15 = tk.OptionMenu(self.Labelframe1, self.LEAD, *Students)
        self.Optioninfo15.grid(row=15, column=1)

        self.Optioninfo16 = tk.OptionMenu(self.Labelframe1, self.SAFE, *Students)
        self.Optioninfo16.grid(row=16, column=1)

        self.Frameinfo17 = tk.Frame(self.Labelframe1)
        self.Frameinfo17.grid(row=17, column=1, sticky='w')
        self.Entryinfo17 = tk.Label(self.Frameinfo17, textvariable=self.NumPART)
        self.Entryinfo17.grid(row=0, column=0)
        self.Buttoninfo17 = tk.Button(self.Frameinfo17, text="Edit", command=lambda: self.EditPeople("PART"))
        self.Buttoninfo17.grid(row=0, column=1)

        self.Frameinfo18 = tk.Frame(self.Labelframe1)
        self.Frameinfo18.grid(row=18, column=1, sticky='w')
        self.Entryinfo18 = tk.Label(self.Frameinfo18, textvariable=self.NumCONT)
        self.Entryinfo18.grid(row=0, column=0)
        self.Buttoninfo18 = tk.Button(self.Frameinfo18, text="Edit", command=lambda: self.EditPeople("CONT"))
        self.Buttoninfo18.grid(row=0, column=1)

        # set grid row and column weight
        for rows in range(self.Labelframe1.grid_size()[1]):
            self.Labelframe1.rowconfigure(rows, minsize=30)
            # set label padx
            if rows > 0:                
                self.Labelframe1.grid_slaves(row=rows, column=0)[0].config(padx=10)
        self.Labelframe1.rowconfigure(0, minsize=5)
        self.Labelframe1.columnconfigure(0, minsize=120)
        self.Labelframe1.columnconfigure(1, minsize=180)
        self.Dateinfo02.config(width=18)
        self.Entryinfo12a.config(width=4)
        self.Labelinfo12a.config(width=2)
        self.Entryinfo12b.config(width=2)
        self.Labelinfo12b.config(width=2)
        self.Entryinfo12c.config(width=2)
        self.Labelinfo12c.config(width=2)
        self.Entryinfo13a.config(width=4)
        self.Labelinfo13a.config(width=2)
        self.Entryinfo13b.config(width=2)
        self.Labelinfo13b.config(width=2)
        self.Entryinfo13c.config(width=2)
        self.Labelinfo13c.config(width=2)
        self.Optioninfo14.config(width=16)
        self.Optioninfo15.config(width=16)
        self.Optioninfo16.config(width=16)
        self.Frameinfo17.columnconfigure(0, minsize=90)
        self.Frameinfo17.columnconfigure(1, minsize=90)
        self.Frameinfo18.columnconfigure(0, minsize=90)
        self.Frameinfo18.columnconfigure(1, minsize=90)
        self.Buttoninfo17.config(width=6)
        self.Buttoninfo18.config(width=6)
            
        self.Labelframe2 = tk.LabelFrame(self.parent, text="Workflow")
        self.Labelframe2.place(x=380, y=10, height=570, width=210)

        self.Label1 = tk.Label(self.Labelframe2, text="1. Analyze EnergyChart.xlsx",justify='left')
        self.Label1.grid(row=0, column=0, sticky='w')

        self.Button1 = tk.Button(self.Labelframe2, text="Locate File", command = self.LoadEnergyChart)
        self.Button1.grid(row=1, column=0)

        self.Label2 = tk.Label(self.Labelframe2, text="2. Open EnergyChart.xlsx and save as Web Page (.htm)",justify='left')
        self.Label2.grid(row=3, column=0, sticky='w')

        self.Button2 = tk.Button(self.Labelframe2, text="Open in Excel", command = self.OpenEnergyChart)
        self.Button2.grid(row=4, column=0)

        self.Label3 = tk.Label(self.Labelframe2, text="3. Fill in visit information",justify='left')
        self.Label3.grid(row=6, column=0, sticky='w')

        self.Button3 = tk.Button(self.Labelframe2, text='Save information', command = self.SaveInfo)
        self.Button3.grid(row=7, column=0)
  
        self.Label4 = tk.Label(self.Labelframe2, text="4. Locate the folder including all recommendation .docx",justify='left')
        self.Label4.grid(row=9, column=0, sticky='w')
        
        self.Button4 = tk.Button(self.Labelframe2, text="Locate folder", command = self.LocateRecommendation)
        self.Button4.grid(row=10, column=0)

        self.Label5 = tk.Label(self.Labelframe2, text="5. Edit Description.docx",justify='left')
        self.Label5.grid(row=12, column=0, sticky='w')

        self.Button5 = tk.Button(self.Labelframe2, text="Open in Word", command = self.OpenDescription)
        self.Button5.grid(row=13, column=0)

        self.Label6 = tk.Label(self.Labelframe2, text="6. Compile the report",justify='left')
        self.Label6.grid(row=15, column=0, sticky='w')

        self.Button6 = tk.Button(self.Labelframe2, text="Compile", command=self.CompileReport)
        self.Button6.grid(row=16, column=0)

        self.Label7 = tk.Label(self.Labelframe2, text="7. Open the report, select all (Ctrl+A); refresh (F9), OK; refresh (F9) again, OK.",justify='left')
        self.Label7.grid(row=18, column=0, sticky='w')

        self.Button7 = tk.Button(self.Labelframe2, text="Open in Word", command=self.OpenReport)
        self.Button7.grid(row=19, column=0)

        self.Label8 = tk.Label(self.Labelframe2, text="Copyright © 2024\nLehigh University Industrial Assessment Center", justify='left')
        self.Label8.grid(row=21, column=0, sticky='w')

        for rows in range(self.Labelframe2.grid_size()[1]):
            # add padx
            if rows % 3 == 0:
                #label
                self.Labelframe2.rowconfigure(rows, minsize=24)
                #change label wraplength
                self.Labelframe2.grid_slaves(row=rows, column=0)[0].config(padx=10,wraplength=190)
            elif rows % 3 == 1:
                #button
                #change button width
                self.Labelframe2.grid_slaves(row=rows, column=0)[0].config(padx=10,width=15)
            else:
                #space
                self.Labelframe2.rowconfigure(rows, minsize=10)


        self.Labelframe3 = tk.LabelFrame(self.parent)
        self.Labelframe3.place(x=610, y=10, height=570, width=570)
        self.Labelframe3.configure(text="Command Line Output")
        self.Text1 = scrolledtext.ScrolledText(self.Labelframe3, state='disabled',wrap="word")
        self.Text1.place(x=10, y=5, height=540, width=550)
        self.Text1.tag_configure("stderr", foreground="#b22222")

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
        PROH = self.PROHH.get()
        if PROH.isdigit():
            self.info.PROH.value = [int(PROH), int(self.PROHD.get()), int(self.PRODW.get())]
        else:
            self.info.PROH.value = [float(PROH), int(self.PROHD.get()), int(self.PRODW.get())]
        OFOH = self.OFOHH.get()
        if OFOH.isdigit():
            self.info.OFOH.value = [int(OFOH), int(self.OFOHD.get()), int(self.OFODW.get())]
        else:
            self.info.OFOH.value = [float(OFOH), int(self.OFOHD.get()), int(self.OFODW.get())]
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
