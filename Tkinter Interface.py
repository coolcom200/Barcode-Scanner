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
            retrieve_del_entry()
        time.sleep(0.09)


thr = threading.Thread(target=check_entry_length)
thr.daemon = True
thr.start()


def retrieve_del_entry():
    x = barcodeEntry.get()
    barcodeEntry.delete(0, END)
    print(x)


sub = Button(top, text='Clear', command=lambda: barcodeEntry.delete(0, END)).pack(side=LEFT)

right = Frame(m1)
m1.add(right)

# Members List
scBr = Scrollbar(right)
scBr.pack(side=RIGHT, fill=Y)

x = Listbox(right, selectmode=SINGLE, font='Times 15 bold', yscrollcommand=scBr.set)
scBr.configure(command=x.yview)
x.insert(END, *e)
x.pack(fill=BOTH)
Button(right, text='send', command=lambda: stuff()).pack()

# Member information
bottom = Frame(m1)
m1.add(bottom)


def quitHandler():
    root.destroy()
    Stp()


def stuff():
    select = x.curselection()
    index = list(select)[0]
    print(index, e[index])


barcodeEntry.focus_force()
root.protocol("WM_DELETE_WINDOW", quitHandler)
root.mainloop()
