from tkinter import filedialog
import shutil
import os
import glob
import pandas as pd
import numpy as np
from datetime import datetime
from tkinter import messagebox
from tkinter import *


class csv_excel_conv:
    def file_ip(self):
        files = filedialog.askopenfilenames(initialdir="/", title = "Select files", multiple=True)
        for i in range(len(files)):
            shutil.copy(files[i],"file_source")
        temp = os.listdir(r'file_source')
        maxlimit = len([entry for entry in temp])
        if maxlimit == 0:
            exit()
        for file_name in temp:
            if (len(file_name) > 25):
                delfile = r'file_source' + '\\' + file_name
                os.remove(delfile)
                print('Please rename your file with a name having < 25 characters. Program will continue without:' + delfile)
                messagebox.showinfo('Warning','Please rename your file with a name having < 25 characters. Program will continue without:' + delfile)

    def remove_header(self):
        maxlimit = len([entry for entry in os.listdir(r"file_source")])
        temp = os.listdir(r"file_source")
        src_folder = r"file_source"
        for i in range(maxlimit):
            filenm = src_folder + '\\' + temp[i]
            with open(filenm, 'r', encoding='utf-8') as file:
                data = file.readlines()
            data[0] = "\n"
            with open(filenm, 'w', encoding='utf-8') as file:
                file.writelines(data)
    def remove_footer(self):
        maxlimit = len([entry for entry in os.listdir(r"file_source")])
        temp = os.listdir(r"file_source")
        src_folder = r"file_source"
        for i in range(maxlimit):
            filenm = src_folder + '\\' + temp[i]
            with open(filenm, 'r+') as fp:
                # read and store all lines into list
                lines = fp.readlines()
                # move file pointer to the beginning of a file
                fp.seek(0)
                # truncate the file
                fp.truncate()
                # start writing lines except the last line
                # lines[:-1] from line 0 to the second last line
                fp.writelines(lines[:-1])

    def merge_header(self):
        maxlimit = len([entry for entry in os.listdir(r"file_source")])
        temp = os.listdir(r"file_source")
        for i in range(maxlimit):
            tmp = r'file_source' + '\\' + temp[i]
            filenames = [r'FILEfields.txt', tmp]
#            opfile = r'file_merge' + '\\' + 'm_' + temp[i]
            opfile = r'file_merge' + '\\' + temp[i]
            with open(opfile, 'w') as outfile:
                for names in filenames:
                    with open(names) as infile:
                        outfile.write(infile.read())
                    outfile.write("\n")
            i = i + 1
    def xlconvert(self):
        src_folder = r"file_merge"
        dst_folder = r"file_dest"
        for file_name in os.listdir(src_folder):
            src = src_folder + '/' + file_name
            dst = dst_folder + '/' + file_name
            if os.path.isfile(src):
                shutil.move(src, dst)
        maxlimit = len([entry for entry in os.listdir(dst_folder)])
        temp = os.listdir(dst_folder)
        writer = pd.ExcelWriter('FILE_converted.xlsx', engine='xlsxwriter')
        fname = r'file_dest'
        for i in range(maxlimit):
            filename = fname + '/' + temp[i]
            testdata = pd.read_csv(filename, dtype="string")
            testdata.index = np.arange(1,len(testdata) + 1)
            testdata.to_excel(writer, sheet_name=temp[i])
        writer.close()
    def file_bkp(self):
        if os.path.exists("FILE_converted.xlsx"):
            src = r"FILE_converted.xlsx"
            dst = r"file_dest"
            shutil.move(src,dst)
            os.remove(r'FILEfields.txt')
            maxlimit = len([entry for entry in os.listdir(r'file_source')])
            temp = os.listdir(r'file_source')
            for i in range(maxlimit):
                delfiles = r"file_source\\" + temp[i]
                os.remove(delfiles)
            now = datetime.now()
            currentdate = 'FILE_'+ now.strftime('%Y') + now.strftime('%m') + now.strftime('%d') + '_' + now.strftime('%H%M%S')
            directory = filedialog.askdirectory(initialdir="/", title = "select output location")
            if len(directory) == 0:
               askuser = messagebox.askyesno(title='confirmation',message='Do you want to exit?')
               if askuser:
                   xlpath = r'file_dest'
                   for f in glob.iglob(xlpath + '\\' + "FILE_converted.xlsx", recursive=True):
                       os.remove(f)
                   temp2 = os.listdir(r'file_dest')
                   for j in range(len([entry for entry in os.listdir(r'file_dest')])):
                       delfiles2 = r"file_dest" + '\\' + temp2[j]
                       os.remove(delfiles2)
                       exit()
            path = os.path.join(directory,currentdate)
            os.mkdir(path)
            for file_name in os.listdir(r"file_dest"):
                src = "file_dest" + '\\' + file_name
                shutil.move(src,path)
        os.startfile(path)
# REPLACE FIELD 1, FIELD 2 etc with column names of your choice
# For fields with similar names, you can use the for loops. Replace the number in range()
# to change the total no of subfields
class userfilechoice():
    def ctafheader(self):
        reportlist=list()
        reportlist.append('FIELD 1,')
        reportlist.append('FIELD 2,')
        reportlist.append('FIELD 3,')
        for i in range(24):
            reportlist.append('SUBFIELD 1,')
        reportlist.append('FIELD 4,')
        for j in range(24):
            reportlist.append('SUBFIELD 2,')
        reportlist.append('FIELD 5,')
        reportlist.append('FIELD 6,')
        for k in range(29):
            reportlist.append('SUBFIELD 3,')
        reportlist.append('FIELD 7,')        
        for l in range(29):
            reportlist.append('SUBFIELD 4,')
        reportlist.append('FIELD 8,')
#        print('reportlist=', reportlist, 'reportlistlen')
#        print('listlength=', len(reportlist))
        ipfile = r'FILEfields.txt'
        with open(ipfile, 'w') as fp:
            for i in range(len(reportlist)):
                pass
                fp.write(reportlist[i])
    def customreport(self):
            reportlist=list()
            reportlist.append('FIELD 1,')
            reportlist.append('FIELD 2,')
            reportlist.append('FIELD 3,')
            for i in range(24):
                reportlist.append('SUBFIELD 1,')
            reportlist.append('FIELD 4,')
            for j in range(24):
                reportlist.append('SUBFIELD 2,')
            reportlist.append('FIELD 5,')
            reportlist.append('FIELD 6,')
            for k in range(29):
                reportlist.append('SUBFIELD 3,')
            reportlist.append('FIELD 7,')
            for l in range(29):
                reportlist.append('SUBFIELD 4,')
            reportlist.append('FIELD 8,')
            reportlist.append('FIELD 9,')
            top = Tk()
        # create listbox object
            listbox = Listbox(top, height=100, width=30, selectmode=MULTIPLE, bg="white",
                  activestyle='dotbox',
                  font="Arial",
                  fg="black")
        # Define the size of the window.
            top.geometry("300x250")
        # Define a label for the list.
            label = Label(top, text = "CUSTOM REPORT")
        # insert elements by their index and names.
            for m in range(len(reportlist)):
                listbox.insert(m,reportlist[m])
            def printopt():
                ipfile = r'FILEfields.txt'
                with open(ipfile, 'w') as fp:
                    for z in listbox.curselection():
                        pass
                        fp.write(listbox.get(z))
            btn = Button(top, text='Print Selected', command=lambda:printopt())
        # pack the widgets
            label.pack()
            listbox.pack()
            btn.place(x=10,y=20)
        # Display until User exits themselves.
            top.mainloop()
#userfilechoice().customreport()
