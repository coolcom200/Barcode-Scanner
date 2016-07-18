from tkinter import *
import threading
import time
from sys import exit as Stp

e = ['Wilson', 'Winters', 'Wise', 'Witt', 'Wolf', 'Wolfe', 'Wong', 'Wood', 'Woodard', 'Woods', 'Woodward', 'Wooten',
     'Workman', 'Wright', 'Wyatt', 'Wynn', 'Yang', 'Yates', 'York', 'Young', 'Zamora', 'Zimmerman']

root = Tk()

m1 = PanedWindow(height=650, width=1000, orient=VERTICAL)
m1.pack(fill=BOTH, expand=1)

# Barcode Entry
top = Frame(m1)
m1.add(top)

barcodeEntry = Entry(top, font='Times 24 bold', width=50)
barcodeEntry.pack(side=LEFT)


def check_entry_length():
    while True:
        lengthE = len(barcodeEntry.get())
        # print(lengthE)
        if lengthE == 9:
            retrieve_entry()
        time.sleep(0.09)


thr = threading.Thread(target=check_entry_length)
thr.daemon = True
thr.start()


def retrieve_entry():
    x = barcodeEntry.get()
    barcodeEntry.delete(0, END)
    print(x)


sub = Button(top, text='Clear', command=lambda: barcodeEntry.delete(0, END)).pack(side=LEFT)

right = Frame(m1)
m1.add(right)

# Members List
scBr = Scrollbar(right)
scBr.pack(side=RIGHT, fill=Y)

dataList = Listbox(right, selectmode=SINGLE, font='Times 15 bold', yscrollcommand=scBr.set)
scBr.configure(command=dataList.yview)
dataList.insert(END, *e)
dataList.pack(fill=BOTH)

# Member information
bottom = Frame(m1)
m1.add(bottom)


def display_member_info(memberIndex):


    def member_info_template(name='Not Applicable', number='Not Applicable', paid='NO'):
        global nameLab, numLab, payLab

        color1, color2, color3 = 'white', 'Red', 'Red'
        if name == 'Not Applicable':
            color1 = 'Red'
        if paid != 'NO':
            color3 = 'lightGreen'
        if number != 'Not Applicable':
            color2 = 'lightGreen'

        nameLab = Label(bottom, text='Name: ' + name, font='Times 15 bold', bg=color1)
        nameLab.grid(row=1, column=1, sticky=W)
        numLab = Label(bottom, text='Number: ' + str(number), font='Times 15 bold', bg=color2)
        print(numLab)
        numLab.grid(row=2, column=1, sticky=W)
        payLab = Label(bottom, text='Paid: ' + paid, font='Times 15 bold', bg=color3)
        payLab.grid(row=3, column=1, sticky=W)

    memberIndex = (str(memberIndex))
    memberIndex = memberIndex.replace('(', '')
    memberIndex = memberIndex.replace(')', '')
    memberIndex = memberIndex.replace(',', '')
    if memberIndex is '':
        pass
    else:
        memberIndex = int(memberIndex)
        member_info_template(e[memberIndex])


def check_selection():
    def clean_mem_info_panel():
        numLab.destroy()
        nameLab.destroy()
        payLab.destroy()
    lastPos = 0
    while True:
        curSelect = dataList.curselection()
        if curSelect != lastPos:

            display_member_info(curSelect)
            clean_mem_info_panel()
        else:
            pass
        time.sleep(0.1)
        lastPos = curSelect


listboxThr = threading.Thread(target=check_selection)
listboxThr.daemon = True
listboxThr.start()


def quit_handler():
    root.destroy()
    Stp()


barcodeEntry.focus_force()
root.protocol("WM_DELETE_WINDOW", quit_handler)
root.mainloop()
