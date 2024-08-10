# import xlwings as xlw

import CSVconverter2 as fileconv
from tkinter import *
def fileformat1():
    fileconv.csv_excel_conv().file_ip()
    fileconv.csv_excel_conv().remove_header()
    fileconv.csv_excel_conv().remove_footer()

def fileformat2():
    fileconv.csv_excel_conv().merge_header()
    fileconv.csv_excel_conv().xlconvert()
    fileconv.csv_excel_conv().file_bkp()

def clickcheck(input):
    buttonclick = False
    buttonclick = not buttonclick
    if buttonclick == True and input == 1:
        fileformat1()
        fileconv.userfilechoice().ctafheader()
        fileformat2()
    if buttonclick == True and input == 2:
 #       print('custom report logic here')
        fileformat1()
        fileconv.userfilechoice().customreport()
        fileformat2()



window=Tk()
#bg=PhotoImage(file='bgp3.png')
#bglabl=Label(window,image=bg)
#bglabl.place(x=0,y=0)
btnctaf=Button(window, text="CTAF", fg='black',command=lambda:clickcheck(1))
btnctaf.place(x=100, y=120)
btncustrep=Button(window, text='Custom Report', fg='black',command=lambda:clickcheck(2))
btncustrep.place(x=150, y=120)
lbl=Label(window, text="Choose your file- convert it into an excel workbook", fg='black', font=("Arial", 16))
lbl.place(x=60, y=50)
window.title('CSV file to Excel formatter')
window.geometry("600x400+10+10")
window.mainloop()


