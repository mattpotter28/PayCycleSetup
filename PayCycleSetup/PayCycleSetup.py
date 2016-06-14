# Matt Potter
# Created May 31 2016 
# Last Edited June 13 2016
# Pay Cycle Setup v3.0

import tkinter as tk
import pypyodbc
from tkinter import ttk
from tkinter import messagebox
import re

class MainWindow(tk.Frame):

    def __init__(self, *args, **kwargs):
        # region InitSetup
        tk.Frame.__init__(self, *args, **kwargs)
        root.wm_title("Pay Cycle Setup")
        root.geometry("%dx%d%+d%+d" % (325, 225, 250, 125))
        
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
                global payGroups
                payGroups = cursor.fetchall()
                global payGroupCount 
                payGroupCount = len(payGroups)
                
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
            w.pack(side="left",  expand=1)
            connection.commit()
        # endregion EntryCreation
                           
        # region ButtonCreation
        # creates submit, edit, and cancel buttons
        global submitButton
        submitButton = tk.Button(root, text='Submit', command= lambda: self.submit(cursor))
        editButton = tk.Button(root, text='Edit Pay Group', command=self.editWindow)
        addButton = tk.Button(root, text='Add Pay Group', command=self.addWindow)
        cancelButton = tk.Button(root, text='Cancel', command=root.destroy)
        
        submitButton.pack(side="right", padx=10, pady =5)
        editButton.pack(side="right", padx=0, pady =5)
        addButton.pack(side="right", padx=10, pady=5)
        cancelButton.pack(side="right", padx=0, pady=5)
        # endregion ButtonCreation
    
    def editWindow(self):
        # region TopLevelSetup
        # creates another window
        t = tk.Toplevel(self)
        t.wm_title("Edit Pay Group")
        t.geometry("%dx%d%+d%+d" % (250, 400, 250, 125))
        # endregion TopLevelSetup

        # region WidgetCreation
        # creates entry box and buttons
        OldPayGroup = tk.StringVar()
        NewPayGroupName = tk.StringVar()
        NewPayGroupID = tk.StringVar()
        lab1 = tk.Label(t, text="Pay Group to Edit:", anchor='nw')
        nameList = tk.Listbox(t, width=30, height = 15)
        count = 0
        for name in payGroups:
            name = str(name).strip("(,)")
            nameList.insert(count, name)
            count  = count + 1
        lab2 = tk.Label(t, width=20, text="New Pay Group Name: ", anchor='w')
        entry2 = tk.Entry(t, width=35, textvariable=NewPayGroupName) 
        submitButton = tk.Button(t, text="Submit Edit", command= lambda: self.submitEdit(nameList.get(tk.ACTIVE), NewPayGroupName))
        cancelButton = tk.Button(t, text='Cancel', command=t.destroy)
        
        # pack interface
        lab1.pack(side="top", fill="both", padx=5, pady=5, expand=1)
        nameList.pack(side="top")
        lab2.pack(side="top", fill="both", padx=5, pady=5, expand=1)
        entry2.pack(side="top", pady=5, expand=1) 
              
        submitButton.pack(side="right", padx=5, pady=10, expand=1)        
        cancelButton.pack(side="right", padx=5, pady=10, expand=1)
        #endregion WidgetCreation

    def addWindow(self):
        # region TopLevelSetup
        # creates another window
        t = tk.Toplevel(self)
        t.wm_title("Add Pay Group")
        t.geometry("%dx%d%+d%+d" % (250, 100, 250, 125))
        # endregion TopLevelSetup

        # region WidgetCreation
        # creates entry box and buttons
        NewPayGroupName = tk.StringVar()
        NewPayGroupID = tk.StringVar()
        lab1 = tk.Label(t, width=20, text="New Pay Group Name: ", anchor='w')
        entry1 = tk.Entry(t, width=35, textvariable=NewPayGroupName)
        submitButton = tk.Button(t, text="Add Pay Group", command= lambda: self.submitAdd(NewPayGroupName, NewPayGroupID))
        cancelButton = tk.Button(t, text='Cancel', command=t.destroy)
        
        # pack interface
        lab1.pack(side="top", fill="both", padx=5, pady=5)
        entry1.pack(pady=5)
        submitButton.pack(side="right", padx=5, pady=5)        
        cancelButton.pack(side="right", padx=5, pady=5)
        #endregion WidgetCreation
    
    def submit(self, cursor):
        # region SiteConversion
        loc = LocationVariable.get()
        loc = loc.strip("(\",)")
        loc = str(loc).replace("'","%")
        SQLCommand = ("SELECT [SiteNumber] FROM [POSLabor].[dbo].[NBO_Sites] where [SiteName] like '"+loc+"';")
        cursor.execute(SQLCommand)
        loc = cursor.fetchone()
        loc = str(loc).strip("(,)")
        # endregion SiteConversion
               
        # region PayGroupConversion
        payg = PayGroupVariable.get()
        payg = payg.strip("(',)")
        SQLCommand = ("SELECT [PayGroupID] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayrollGroupName] like '"+payg+"';")
        cursor.execute(SQLCommand)
        payg = cursor.fetchone()
        payg = str(payg).strip("(,)")
        
        if payg == "None":
            payg = str(payGroupCount + 1)
        # endregion PayGroupConversion
        
        # region VariablePreparation
        tip = int(TipShareVariable.get())
        payc = PayCycleVariable.get()
        adp = ADPStoreCodeVariable.get()
        # endregion VariablePreparation
        
        self.insertSQL(loc, payg, tip, payc, adp)      
   
    def submitEdit(self, OldPayGroupName, NewPayGroupName):       
        # check if new values are not already used, throw errors
        SQLCommand = ("SELECT [PayGroupID] FROM [POSLabor].[dbo].[NBO_PayGroup];")
        cursor.execute(SQLCommand)
        ids = cursor.fetchall()
        flag = False
        for name in payGroups:
            if str(NewPayGroupName.get()) in name:
                mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                flag = True        
        if flag == False:
            # use sql statement to UPDATE values to new values
            OldPayGroupName = str(OldPayGroupName).replace("'", "%")

            SQLCommand = ("UPDATE [POSLabor].[dbo].[NBO_PayGroup] " \
                          "SET [PayrollGroupName]='"+str(NewPayGroupName.get())+"' "+ \
                          "WHERE [PayrollGroupName] like '"+OldPayGroupName+"';")              
                
            print (SQLCommand)            
            cursor.execute(SQLCommand)
            connection.commit()
            # change PayGroupVariable to read-only NewPayGroupName
            PayGroupVariable.set(NewPayGroupName.get())
            self.destroy()

    
    def submitAdd(self, NewPayGroupName, NewPayGroupID):
        flag = False
        for name in payGroups:
            if str(NewPayGroupName.get()) in name:
                mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                flag = True
        if flag == False:
            payGroups.insert((len(payGroups)+1), NewPayGroupName.get())
            SQLCommand = ("INSERT INTO [POSLabor].[dbo].[NBO_PayGroup] (PayrollGroupName, PayGroupID) " \
                          "VALUES ('"+str(NewPayGroupName.get())+"', "+str(len(payGroups))+" );")              
                
            print (SQLCommand)         
            cursor.execute(SQLCommand)
            connection.commit()
        
        # check if values are not already used, throw errors
        
        PayGroupVariable.set(NewPayGroupName.get())
        self.destroy()


    def insertSQL(self, loc, payg, tip, payc, adp):
        try:
            SQLCommand = ( "DECLARE @RT INT "\
                           "EXECUTE @RT = dbo.pr_NBO_PayCycleSetup_ADD "+loc+", "+payg+", "+ str(tip) +", '"+adp+"', 0, "+ payc +
                           " PRINT @RT") # command to add data 
            cursor.execute(SQLCommand)
            connection.commit()    # save to table
            mBox = tk.messagebox.showinfo("Success!","Import Complete")
        except:
            mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed Command: "+SQLCommand)
            connection.rollback() # undo command
            

        root.destroy()


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