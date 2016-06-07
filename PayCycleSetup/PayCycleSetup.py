# Matt Potter
# Created May 31 2016 
# Last Edited June 6 2016
# Pay Cycle Setup

import tkinter as tk
import pypyodbc
from tkinter import ttk
import re

class MainWindow(tk.Frame):

    def __init__(self, *args, **kwargs):
        # region InitSetup
        tk.Frame.__init__(self, *args, **kwargs)
        root.wm_title("Pay Cycle Setup")
        root.geometry("%dx%d%+d%+d" % (410, 225, 250, 125))
        
        fields = ('Location', 'Pay Group', 'Tip Share', 'Pay Cycle', 'ADP Store Code')
        global cursor
        cursor = connection.cursor()
        SQLCommand = ("")
        # endregion InitSetup

        # region EntryCreation
        # cycles through fields and creates corresponding response methods
        for field in fields: 
            row = tk.Frame(root)
            lab = tk.Label(row, width=20, text=field+": ", anchor='w')
                 
            # region LocationEntry
            # creates drop down and stores string response in LocationVariable
            if field == 'Location': 
                global LocationVariable
                LocationVariable = tk.StringVar()
                #LocationVariable.set("Store Name") # default value

                # get store name values from connection
                SQLCommand = ("select [SiteName] from [POSLabor].[dbo].[NBO_Sites]")
                cursor.execute(SQLCommand)
                global siteNames
                siteNames = cursor.fetchall()

                w = ttk.OptionMenu(row, LocationVariable, "Site Name", *siteNames)
            # endregion LocationEntry

            # region PayGroupEntry
            # creates drop down and stores string response in PayGroupVariable
            elif field == 'Pay Group': 
                global PayGroupVariable
                PayGroupVariable = tk.StringVar()
                #PayGroupVariable.set("Pay Group Name") # default value

                # get pay group name values from connection
                SQLCommand = ("select distinct [PayrollGroupName] from [POSLabor].[dbo].[NBO_PayGroup]")
                cursor.execute(SQLCommand)
                payGroups = cursor.fetchall()
                
                w = ttk.OptionMenu(row, PayGroupVariable, "Payroll Group Name", *payGroups)
            # endregion PayGroupEntry

            # region TipShareEntry
            # creates check button and stores boolean response in TipShareVariable
            elif field == 'Tip Share':
                global TipShareVariable
                TipShareVariable = tk.BooleanVar()
                #TipShareVariable.set(False)
                w = tk.Checkbutton(row, variable=TipShareVariable)
            # endregion TipShareEntry

            # region PayCycleEntry
            # creates drop down and stores string response in PayCycleVariable
            elif field == 'Pay Cycle': 
                global PayCycleVariable
                PayCycleVariable = tk.StringVar(root)
                #PayCycleVariable.set("Pay Cycle") # default value

                w = ttk.OptionMenu(row, PayCycleVariable, "Pay Cycle", *range(1,5))
            # endregion PayCycleEntry

            # region ADPStoreCodeEntry
            # creates entry and stores string response in ADPStoreCodeVariable
            elif field == 'ADP Store Code': 
                global ADPStoreCodeVariable
                ADPStoreCodeVariable = tk.StringVar(root)
                w = tk.Entry(row, textvariable=ADPStoreCodeVariable)
            # endregion ADPStoreCodeEntry

            row.pack(side="top", fill="both", padx=5, pady=5)
            lab.pack(side="left")
            w.pack(side="left", fill="both")
        # endregion EntryCreation
                           
        # region ButtonCreation
        # creates submit, edit, and cancel buttons
        global submitButton
        submitButton = tk.Button(root, text='Submit', command= lambda: self.submit(cursor))
        editButton = tk.Button(root, text='Edit/Add Pay Group', command=self.editAddWindow)
        cancelButton = tk.Button(root, text='Cancel', command=root.destroy)
        
        submitButton.pack(side="right", padx=10, pady =5)
        editButton.pack(side="right", padx=0, pady =5)
        cancelButton.pack(side="right", padx=10, pady=5)
        # endregion ButtonCreation

    def editAddWindow(self):
        # region TopLevelSetup
        # creates another window
        t = tk.Toplevel(self)
        t.wm_title("Edit/Add Pay Group")
        t.geometry("%dx%d%+d%+d" % (250, 100, 250, 125))
        # endregion TopLevelSetup
        
        # region WidgetCreation
        # creates entry box and buttons
        NewPayGroupVariable = tk.StringVar(t)
        lab = tk.Label(t, width=20, text="Pay Group Name: ", anchor='w')
        w = tk.Entry(t, width=35, textvariable=NewPayGroupVariable)
        submitButton = tk.Button(t, text="Submit", command= lambda: self.submitEdit(NewPayGroupVariable))
        cancelButton = tk.Button(t, text='Cancel', command=t.destroy)
        
        # pack interface
        lab.pack(side="top", fill="both", padx=5, pady=5)
        w.pack(pady=5)        
        submitButton.pack(side="right", padx=5, pady =5)        
        cancelButton.pack(side="right", padx=5, pady=5)
        #endregion WidgetCreation

    
    def submit(self, cursor):
        loc = LocationVariable.get()
        loc = loc.strip("(\",)")
        #loc = convert SiteName to corresponding siteNumber
        payg = PayGroupVariable.get()
        payg = payg.strip("(',)")
        tip = int(TipShareVariable.get())
        payc = PayCycleVariable.get()
        adp = ADPStoreCodeVariable.get()
        insertSQL(loc, payg, tip, payc, adp)        

   
    def submitEdit(self, NewPayGroupName):
        # !!! change PayGroupVariable to read-only NewPayGroupName
        PayGroupVariable.set(str(NewPayGroupName))
        # read-only?

        # !!! check if NewPayGroupName already exists in NBO_Sites
        # !!! if no, begin addition process
            # !!! create SQL statement to insert data
        # !!! if yes, begin update process
            # !!! fill entries with pre-existing data from NBO_PayCycleSetup
        flag = False
        for name in siteNames:
            if name == str(NewPayGroupName):
                flag == True
        if flag == True: # Update fields, still unsure
            pass

        # !!! create Submit Changes button, calls submitChanges
        submitButton.config(text = "Submit Changes", command = lambda: submitChanges(cursor))        
        self.destroy()

    
    def submitChanges(self, cursor):
        loc = LocationVariable.get()
        loc = loc.strip("(\",)")
        #loc = convert SiteName to corresponding siteNumber
        payg = PayGroupVariable.get()
        payg = payg.strip("(',)")
        tip = int(TipShareVariable.get())
        payc = PayCycleVariable.get()
        adp = ADPStoreCodeVariable.get()
        insertSQL(loc, payg, tip, payc, adp) 


    def insertSQL(self, loc, payg, tip, payc, adp):
        SQLCommand = "DECLARE @RT INT \
                      EXECUTE @RT = dbo.pr_NBO_PayCycleSetup_ADD "+loc+", "+payg+", "+ tip +", "+adp+", 1, "+ payc +" \
                      PRINT @RT" # command to add data 
        cursor.execute(SQLCommand)
        #connection.commit()    # save to table
        connection.rollback()   # undo command
        self.destroy()


#region Main
if __name__ == "__main__":
    # connects to sql server
    # to do: try/except to test connection
    connection = pypyodbc.connect('Driver={SQL Server};'     
                                  'Server=VSQL15\\NBOTest;'    
                                  'Database=POSLabor;'       
                                  'uid=LaborLoad;'            
                                  'pwd=B8{uwBJ!ZxZ{')
    # opens window
    root = tk.Tk()
    main = MainWindow(root)
    main.pack(side="top", fill="both", expand=True)
    root.mainloop()
    connection.close()
#endregion Main