# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 14:59:33 2019

@author: akay-w
"""
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory
import os
import datetime
from feedback_tool_2_functions import get_xliff_list, parse_xliff, \
analyze, create_excel

class App(tk.Frame):    
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        self.initializeUI()
        
    def initializeUI(self):
        self.master.geometry('625x300')
        self.master.title('Feedback Tool')
        #Blank for orig file name
        self.origFileEntry = tk.Entry(
                self.master, width=50)
        self.origFileEntry.grid(row=0, column=1, padx=10, pady=10)
        #Select orig file button
        self.getOrigFile = tk.Button(
                self.master, text='Select Original File Folder', \
                command=self.selectOrigFile
                )
        self.getOrigFile.grid(row=0, column=2)
        #Blank for final file name
        self.finalFileEntry = tk.Entry(
                self.master, width=50)
        self.finalFileEntry.grid(row=1, column=1, padx=10, pady=10)
        #Select orig file button
        self.getFinalFile = tk.Button(
                self.master, text='Select Final File Folder', \
                command=self.selectFinalFile
                )
        self.getFinalFile.grid(row=1, column=2)
        #Blank for save location
        self.saveLocationEntry = tk.Entry(
                self.master, width=50)
        self.saveLocationEntry.grid(row=2, column=1, padx=10, pady=10)
        #Select save location button
        self.getSaveLocation = tk.Button(
                self.master, text='Select Save Location', \
                command=self.selectSaveLocation
                )
        self.getSaveLocation.grid(row=2, column=2, padx=10, pady=10)
        #Blank for project number/name
        self.projectNameEntry = tk.Entry(
                self.master, width=50)
        self.projectNameEntry.grid(row=3, column=1, padx=10, pady=10)
        self.projectNameLabel = tk.Label(
                self.master, text='Enter project no./name (Optional)')
        self.projectNameLabel.grid(row=3, column=2, padx=10, pady=10)
        #Label for language type
        self.langTypeLabel = tk.Label(
                self.master, text='Select target language type:')
        self.langTypeLabel.grid(row=4, column=1, padx=10, pady=10)
        #Radio buttons for language type
        self.lang = tk.IntVar()
        self.lang.set(1)
        self.Eurolang = tk.Radiobutton(self.master, text='European', 
                                       variable=self.lang, value=1
                                       )
        self.Eurolang.grid(row=4, column=2, padx=10, pady=10)
        self.Asialang = tk.Radiobutton(self.master, text='Asian', 
                                       variable=self.lang, value=2
                                       )
        self.Asialang.grid(row=4, column=3, padx=10, pady=10)
        #Exit app button
        self.exitBtn = tk.Button(
                self.master,text='Exit', command=self.quitApp,
                )
        self.exitBtn.grid(row=5, column=2, pady=10)
        #Run check button
        self.startBtn = tk.Button(
                self.master, text='Run', command=self.startApp,
                )        
        self.startBtn.grid(row=5, column=3, pady=10)
  
    def quitApp(self):
        """
        Function called by exit button to quit app  
        """
        global root
        root.destroy()
       
    def selectOrigFile(self):
        """
        Function called by select orig file button to choose original file 
        """
        origdir = askdirectory()
        self.origFileEntry.delete(0,tk.END)
        self.origFileEntry.insert(0, origdir)
        return
    
    def selectFinalFile(self):
        """
        Function called by select final file button to choose final file 
        """
        finaldir = askdirectory()
        self.finalFileEntry.delete(0,tk.END)
        self.finalFileEntry.insert(0, finaldir)
        return
    
    def selectSaveLocation(self):
        """
        Function called by select save location button to choose save location
        """
        directory = askdirectory()
        self.saveLocationEntry.delete(0,tk.END)
        self.saveLocationEntry.insert(0,directory)
        return
    
    def startApp(self):
        """
        Function called by start button to execute application.
        Gets lists of original and edited files, parses all files for text, 
        finds edits, and creates an Excel report showing the differences.
        """
        origdir = self.origFileEntry.get()
        origfiles = get_xliff_list(origdir)
        origfiles.sort()            
        finaldir = self.finalFileEntry.get()
        finalfiles = get_xliff_list(finaldir)
        finalfiles.sort()
                        
        if not origfiles:
             messagebox.showwarning("Feedback Tool", \
            "No sdlxliff files found in original folder.")
             return
         
        if not finalfiles:
            messagebox.showwarning("Feedback Tool", \
            "No sdlxliff files found in final folder.")
            return
         
        if len(origfiles) != len(finalfiles):
            messagebox.showwarning("Feedback Tool", \
            "Different number of original and final files.")
            return
        
        directory = self.saveLocationEntry.get()
        
        if not directory:
           messagebox.showwarning("Feedback Tool", \
           "Please enter a save location.")
           return
       
        if not os.path.exists(directory):
            messagebox.showwarning("Feedback Tool", \
            "The save location does not exist, or you do not have access permissions.")
            return
       
        lang = self.lang.get()
        data_list = []
        for file in range(len(origfiles)):
            filename = os.path.basename(origfiles[file])
            origtransinfo = parse_xliff(origfiles[file])
            edittransinfo = parse_xliff(finalfiles[file])
            data_list += analyze(filename, origtransinfo, edittransinfo, lang)
        
        
        
        savedate = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        projectname = self.projectNameEntry.get()
        if projectname:
            forbiddenChar = r'<>:"/\|?*'
            for char in forbiddenChar:
                if char in projectname:
                    messagebox.showwarning("Feedback Tool", \
                                           "Project name includes forbidden characters.")
                    return
            savename = projectname + "_" + "Translation_Edits_" + savedate + ".xlsx"
        else:
            savename = "Translation_Edits_" + savedate + ".xlsx"
        savename = os.path.join(os.path.normpath(directory), savename)
        create_excel(savename, data_list, lang)
        messagebox.showinfo('Feedback Tool', 'Finished!')        

if __name__ == '__main__':   
    root = tk.Tk()
    app = App(root)
    root.mainloop()