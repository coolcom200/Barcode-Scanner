from tkinter import *
import threading
import time
from sys import exit as Stp

e = ['Wilson', 'Winters', 'Wise', 'Witt', 'Wolf', 'Wolfe', 'Wong', 'Wood', 'Woodard', 'Woods', 'Woodward', 'Wooten',
     'Workman', 'Wright', 'Wyatt', 'Wynn', 'Yang', 'Yates', 'York', 'Young', 'Zamora', 'Zimmerman']

root = Tk()

menubar = Menu(root)
menuB = Menu(menubar, tearoff=0)
menuB2 = Menu(menubar, tearoff=0)
menuB3 = Menu(menubar, tearoff=1)
menuB.add_command(label="Exit", command=lambda: quit_handler())
menuB2.add_command(label="Attendance Graph")
menuB2.add_command(label='Total Member present')
menuB2.add_command(label='Credit Eligible Members')
menubar.add_cascade(label="File", menu=menuB)
menubar.add_cascade(label='Analyze', menu=menuB2)
menubar.add_cascade(label='Help', menu=menuB3)
root.configure(menu=menubar)

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
        global a
        if name == paid == number and name == 'None':
            color1, color2, color3 = 'white', 'white', 'white'
        else:
            color1, color2, color3 = 'white', 'lightgreen', 'lightgreen'

        if name == 'Not Applicable':
            color1 = 'Red'

        if paid == 'NO':
            color3 = 'Red'

        if number == 'Not Applicable':
            color2 = 'red'

        nameLab = Label(bottom, text='Name: ' + name, font='Times 15 bold', bg=color1)
        nameLab.grid(row=1, column=1, sticky=W)
        numLab = Label(bottom, text='Number: ' + str(number), font='Times 15 bold', bg=color2)
        print(numLab)
        numLab.grid(row=2, column=1, sticky=W)
        payLab = Label(bottom, text='Paid: ' + paid, font='Times 15 bold', bg=color3)
        payLab.grid(row=3, column=1, sticky=W)
        return [nameLab, numLab, payLab]

    memberIndex = (str(memberIndex))
    memberIndex = memberIndex.replace('(', '')
    memberIndex = memberIndex.replace(')', '')
    memberIndex = memberIndex.replace(',', '')
    if memberIndex is '':
        labelIDs = member_info_template('None', 'None', 'None')

    else:
        memberIndex = int(memberIndex)
        labelIDs = member_info_template(e[memberIndex])
    return labelIDs


def clean_mem_info_panel(labelIDList):
    if labelIDList is None:
        pass
    else:
        l1, l2, l3 = labelIDList[0], labelIDList[1], labelIDList[2]
        l1.destroy()
        l2.destroy()
        l3.destroy()


def check_selection():
    b = None
    lastPos = 0
    while True:
        curSelect = dataList.curselection()
        if curSelect != lastPos:
            clean_mem_info_panel(b)
            b = display_member_info(curSelect)
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
