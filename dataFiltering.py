import tkinter as tk
import sys
import os
from shutil import copyfile
import xlrd
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import tkinter as tk
from tkinter import *
#from tkinter.ttk import *
from tkinter import ttk  
from tkinter import messagebox


###############
class Input(tk.Frame):

    def __init__(self, parent):

        tk.Frame.__init__(self, parent)
        self.parent = parent
        
        clients = ["","All Clients"]
        
        docType = [""," CV ", " LOR "," SOP ", "Marksheets","Others", "Reports"]
        
        viewStatus = ["", "All Unreviewed","All Reviewed"]

        docVersion = ["","v1", "v2","v3","v4","v5","v6","v7"]

        self.clients_selection = tk.StringVar()
        self.clients_selection.set(clients[0])

        self.name_label = tk.Label(root, text="Enter Client Name: ",fg="blue",bg="white")#.grid(row = 0, column = 0)
        self.name_entry = tk.Entry(root)

        self.clients_label = tk.Label(root, text="")
        self.clients_entry = tk.OptionMenu(root, self.clients_selection, *clients)

        self.submit_button = tk.Button(text="Search & Download",fg="blue",bg="white", command=self.close_window)#button

        self.name_label.grid(row=0, column=0)
        self.name_entry.grid(row=0, column=1)

        self.clients_label.grid(row=0, column=2)
        self.submit_button.grid(columnspan=2, row=3, column=0)##button

        #self.clients_entry.grid(row=0, column=3)
        ###docType
        self.docType_selection = tk.StringVar()
        self.docType_selection.set(docType[0])

        self.name_label = tk.Label(root, text="Select docType: ",fg="blue",bg="white")#.grid(row = 0, column = 0)
        #self.name_entry = tk.Entry(root)

        self.docType_label = tk.Label(root, text="")
        self.docType_entry = tk.OptionMenu(root, self.docType_selection, *docType)

        self.name_label.grid(row=0, column=4)
        #self.name_entry.grid(row=0, column=5)

        self.docType_label.grid(row=0, column=6)

        self.docType_entry.grid(row=0, column=7)

        ###
        ###viewStatus
        self.viewStatus_selection = tk.StringVar()
        self.viewStatus_selection.set(viewStatus[0])

        self.name_label = tk.Label(root, text="Select viewStatus: ",fg="blue",bg="white")#.grid(row = 0, column = 0)
        #self.name_entry = tk.Entry(root)

        self.viewStatus_label = tk.Label(root, text="")
        self.viewStatus_entry = tk.OptionMenu(root, self.viewStatus_selection, *viewStatus)

        self.name_label.grid(row=0, column=8)
        #self.name_entry.grid(row=0, column=5)

        self.viewStatus_label.grid(row=0, column=9)

        self.viewStatus_entry.grid(row=0, column=10)

        ###
        ###docVersion
        self.docVersion_selection = tk.StringVar()
        self.docVersion_selection.set(docVersion[0])

        self.name_label = tk.Label(root, text="Select docVersion: ",fg="blue",bg="white")#.grid(row = 0, column = 0)
        #self.name_entry = tk.Entry(root)

        self.docVersion_label = tk.Label(root, text="")
        self.docVersion_entry = tk.OptionMenu(root, self.docVersion_selection, *docVersion)

        self.name_label.grid(row=0, column=11)
        #self.name_entry.grid(row=0, column=5)

        self.docVersion_label.grid(row=0, column=12)

        self.docVersion_entry.grid(row=0, column=13)

        ###
    def close_window(self):
        #global name
        #global ideal_type
        self.name = self.name_entry.get()
        self.ideal_type = self.clients_selection.get()
        self.ideal1_type = self.docType_selection.get()
        self.ideal2_type = self.viewStatus_selection.get()
        self.ideal3_type = self.docVersion_selection.get()
        #self.destroy()
        self.quit()

if __name__ == '__main__':

    root = tk.Tk()
    root.geometry("1000x600+300+300")
    app = Input(root)
    #loc = ("D:\pathdata.xlsx") #path of excel file
    loc = askopenfilename(title='Please choose a .xlsx file')
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0)
    #data_name = returned_name
    #print(data_name)
    #print(sheet.nrows)
    list1=[]
    for i in range(sheet.nrows): 
        path=sheet.cell_value(i, 0)
        list1.append(path)
    print(list1)
    downloaded_files = askdirectory(title='Please Select A Folder Where you want to save The Downloaded Files ') # shows dialog box and return the path              
    ######################################

    

   
   

    ###
    root.mainloop()
    # Note the returned variables here
    # They must be assigned to external variables
    # for continued use
    returned_name = app.name
    returned_clients = app.ideal_type
    returned_docType = app.ideal1_type
    returned_viewStatus = app.ideal2_type
    returned_docVersion = app.ideal3_type
    
    ###//end of code login###
    print("Client name is: " + returned_name)
    #print("Client type is: " + returned_ideal)
    print("Doc Type is: " + returned_docType)
    print("viewStatus is: " + returned_viewStatus)
    print("docVersion is: " + returned_docVersion)
    #loc = ("D:\pathdata.xlsx") #path of excel file
    #loc = askopenfilename(title='Please choose a .xlsx file')
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0)
    data_name = returned_name
    #print(data_name)
    #print(sheet.nrows)
    pathname=sheet.cell_value(1, 0)
    path="\\".join(pathname.split('\\')[:-3])
    #print(path)
    #downloaded_files = askdirectory(title='Please Select A Folder Where you want to save The Downloaded Files ') # shows dialog box and return the path              
    ######################################
    folders = []
    
    newFileName=""
    #new_list.append(newFileName)
    #print(new_list)
    #print(path)
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for folder in f:
            folders.append(os.path.join(r, folder))
            for f in folders:
                #print(f)
                base_file_name=os.path.basename(f)
                fileName=os.path.splitext(base_file_name)[0]
                #for LORs file ###
                fileName0=fileName.replace("LOR","LOR_")
                fileName1=fileName0.replace("LoR","LoR_")
                
                ###for Marksheets
                fileName2 = fileName1.replace("12th","Marksheets_12th")
                fileName3 = fileName2.replace("10th","Marksheets_10th")
                fileName4 = fileName3.replace("_sem_","Marksheets_sem_")
                ##for report#############
                fileName5 = fileName4.replace("Minutes_","Reports_Minutes")
                fileName6 = fileName5.replace("PGA","Reports_PGA")
                fileName7 = fileName6.replace("Profile","Reports_Profile")
                fileName8 = fileName7.replace("Quarterly","Reports_Quarterly")
                fileName9 = fileName8.replace("SS_","Reports_SS_")
                ###for SOP##############
                fileName10 = fileName9.replace("_PS_","SOP_PS_")
                fileName11 = fileName10.replace("_SoP_","_SOP_")
                ##for LORs######
                fileName12 = fileName11.replace("_LoR_","_LOR_")
                ###for others#####
                fileName13 = fileName12.replace("_PAF","Others_PAF")
                fileName14 = fileName13.replace("Grad_","Others_Grad")



                newFileName = fileName14.replace("_"," ")
                #print(newFileName)
                if returned_viewStatus=="All Reviewed":
                    Newreturned_viewStatus=" r" 
                    if data_name in newFileName and returned_docType in newFileName and Newreturned_viewStatus in newFileName and returned_docVersion in newFileName:
                        copyfile(f,downloaded_files+'\\'+base_file_name)
                        print("file found at "+f+" and\n Downloaded at new loc: "+downloaded_files+'\\'+base_file_name)
                elif returned_viewStatus=="All Unreviewed":
                    Newreturned2_viewStatus=" r"
                    if data_name in newFileName and returned_docType in newFileName and Newreturned2_viewStatus not in newFileName and returned_docVersion in newFileName:
                        copyfile(f,downloaded_files+'\\'+base_file_name)
                        print("file found at "+f+" and\n Downloaded at new loc: "+downloaded_files+'\\'+base_file_name)
                elif returned_viewStatus=="":
                    Newreturned3_viewStatus=""
                    if data_name in newFileName and returned_docType in newFileName and Newreturned3_viewStatus in newFileName and returned_docVersion in newFileName:
                        copyfile(f,downloaded_files+'\\'+base_file_name)
                        print("file found at "+f+" and\n Downloaded at new loc: "+downloaded_files+'\\'+base_file_name)
    messagebox.showinfo("Completed", "All file successfully Downloaded")           
    # Should only need root.destroy() to close down tkinter
    # But need to handle user cancelling the form instead
    try:
        root.destroy()
    except:
        sys.exit(1) 
