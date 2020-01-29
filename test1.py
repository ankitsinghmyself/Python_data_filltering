import os,re
from shutil import copyfile
import xlrd
loc = ("D:\pathdata.xlsx") #path of excel file
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
data_name=input("please enter a file_name or String: ")
for i in range(sheet.nrows): 
    path=sheet.cell_value(i, 0)
    downloaded_files = "D:\\DataDown"
    folders = []

    # r=root, d=directories, f = files
    
    for r, d, f in os.walk("D:\data"):
        for folder in f:
            folders.append(os.path.join(r, folder))
        for f in folders:
            #print(f)
            base_file_name=os.path.basename(f)
            fileName=os.path.splitext(base_file_name)[0]
            fileName1=fileName.replace("_"," ")
            if data_name in fileName1 and "CV" in fileName1:
                copyfile(f,downloaded_files+'\\'+base_file_name)
                print("file found at "+f+" and\n Downloaded at new loc"+downloaded_files+'\\'+base_file_name)
            
            
'''I wrote code that taking all path address
form excel sheet and going with each address,
and searching for file or string pattern that
given by Client and storing found files into the new location.'''
