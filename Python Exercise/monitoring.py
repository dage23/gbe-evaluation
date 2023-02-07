import os
import time
import pandas as pd
import shutil
from tkinter import filedialog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

##Class Defined to be notified of an event
class MonitorFolder (FileSystemEventHandler):
    #Define only the create action
    def on_created(self, event):
        #Print the file and the event
        print(event.src_path, event.event_type)
        #Send the path to check the extension
        main.check_extension(event.src_path)

class main:
    #Main function
    def __init__ (self):
        input("Welcome, Press Enter To Select a Folder to Monitor")
        ##Open new Dialog to ask for directory
        self.directoryPath = filedialog.askdirectory(initialdir=".", title="Choose Folder")
        if (self.directoryPath == ""):
            print("No folder selected. Closing App")
            exit()
        print ("Folder Selected: " + self.directoryPath)
        #Go to Create Folders Function
        self.createFolders()
        input("Press Enter To Monitor")
        #Go to Monitor Function
        self.monitor()

    #Function to start Monitoring
    #Parameter self
    def monitor(self):
        input("Monitoring...")
        #Declare MonitorFolder class to be used
        event_handler = MonitorFolder()
        observer = Observer()
        #Define the schedule with the MonitorFolder class, path to monitor and declare to not review inner folders.
        observer.schedule(event_handler, path=self.directoryPath, recursive=False)
        #Start observing for changes in folder
        observer.start()
        try:
            #Check for changes every two seconds
            while True:
                time.sleep(2)
        except:
            #On error stop checking, and wait for threads to finish.
            print("Stopping")
            observer.stop()
            observer.join()

    #Function to Check Extensions
    #Parameter src_path: Path of file to be checked
    def check_extension(src_path):
        directory = os.path.dirname(src_path)
        fileName = os.path.basename(src_path)
        fileExtension = os.path.splitext(fileName)[1]
        timestr = time.strftime("%Y%m%d-%H%M%S")
        #Check if the file extension is .xls and if it is a file
        if (fileExtension == ".xls" and os.path.isfile(src_path)) :
            #Send the file to be processed
            main.processExcelFile(directory,fileName)
            print ("Excel document detected, consolidated.")
            #The file has been consolidated, so now it is moved to the Processed folder
            shutil.move(src_path,directory + '/Processed/'+timestr + "_"+fileName)
        #Any other file type
        elif (os.path.isfile(src_path)):      
            #It is moved to Not Applicable    
            shutil.move(src_path,directory + '/Not Applicable/'+timestr + "_"+fileName)
            print ("Not an excel document detected, moved.")

    #Function to create folders
    def createFolders(self):
        #Check if Processed folder exists
        if not os.path.exists(self.directoryPath + '/Processed'):
            #If not, create it
            os.makedirs(self.directoryPath + '/Processed')
        #Check if Not Applicable folder exists
        if not os.path.exists(self.directoryPath + '/Not Applicable'):
            #If not, create it
            os.makedirs(self.directoryPath + '/Not Applicable')
    
    #Function to process the Excel File
    def processExcelFile(directory,filename):
        #Create full names of the main Excel File and file to be processed
        outFile = directory + '/Processed/masterExcel.xlsx'
        inFile = directory + "/"+filename
        #Send files to be created and consolidated
        main.createMasterExcel(outFile)
        main.consolidateFile(inFile,outFile)

    #Function to consolidate new file to master file
    def consolidateFile(inFile, outFile):
        #Define the excel file
        excel_file = pd.ExcelFile(inFile)
        #Get the list of the sheets in the file
        sheets = excel_file.sheet_names
        #With ExcelWriter, since it is going to be writing multiple sheets
        #Define
        #   OutFile to be the master Excel File
        #   Engine to specify which engine to use when writing to the master file
        #   if_sheets_exists to new, so even if a sheet with an existing name is detected, automatically it assigns a new name
        #   mode to a, to append to the existing file
        with pd.ExcelWriter(outFile, engine='openpyxl', if_sheet_exists='new', mode='a') as writer:
            #Cycle through all the sheets in file
            for sheet in sheets:
                #The the data from the new file and sheet name
                data = pd.read_excel(inFile, sheet_name=sheet)
                #Write the data to a new sheet of the master file
                # Index false, so it doesnt write row names
                data.to_excel(writer, sheet_name=sheet, index=False)
    
    #File to create master Excel File
    def createMasterExcel (outFile):
        #If it doesnt exists
        if not os.path.exists(outFile):
            #Create new master Excel File
            writer = pd.ExcelWriter(outFile, engine='xlsxwriter')
            #Close writer
            writer.close()

if __name__ == "__main__":
    main()
