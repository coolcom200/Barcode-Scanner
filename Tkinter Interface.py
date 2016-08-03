from tkinter import *
import threading
import time
from tkinter import messagebox
from sys import exit as Stp
from tkinter.ttk import *
from openpyxl import load_workbook
from openpyxl.cell import column_index_from_string as ColToInt
from time import strftime
from matplotlib import pyplot as plt

date = strftime('%B %d %Y')  # Gets date in the from of MONTH DATE YEAR
saveFile = 'ATT.xlsx'  # Sets the save file
wb = load_workbook(saveFile)
ws = wb.active


def save():  # saves file
    wb.save(saveFile)


def find(findValue, row=None, column=None):
    if isinstance(findValue, str):
        findValue = findValue.lower()
    if row is None and column is None:
        for row in ws.iter_rows:
            for cell in row:
                if cell.value is findValue:
                    return cell.row, ColToInt(cell.column)

    elif row is not None and column is None:  # scanning through a row
        index = 1
        while True:
            cell = ws.cell(row=row, column=index)
            if cell.value is findValue:
                return cell.row, ColToInt(cell.column)
            index += 1

    elif row is None and column is not None:  # scanning through a column
        index = 2
        lastCell = ws.cell(row=index - 1, column=column)
        while True:
            cell = ws.cell(row=index, column=column)
            if lastCell.value is None and cell.value is None:
                return None, None
            if cell.value == findValue:
                return cell.row, ColToInt(cell.column)
            index += 1
            lastCell = cell


def create_attendance_date(row, column):  # Creates meeting date
    column = ColToInt(column)  # Converts letter column to number
    ws.cell(row=row, column=column, value=date)  # Adds date column
    save()  # Saves file


def find_empty_date_column():
    global attCol
    index = 2  # starts column count at 2('B')
    lastCell = ws['A1']  # first cell is set as 'A1'
    while True:  # runs loop to find empty cell in row
        cell = ws.cell(row=1, column=index)
        if lastCell.value is None and cell.value is None:
            # [date][None][None]...[None] finds 2 empty cells and the
            # uses the previous cell as the point at which to add the
            # date stamp
            create_attendance_date(lastCell.row, lastCell.column)
            attCol = lastCell.column
            break

        elif cell.value == date:

            # Makes attCol known to the whole program for attendance registration
            attCol = cell.column
            break

        # elif cell.value is None and lastCell.value is not None:
        index += 1
        lastCell = cell


def emptyRow():
    return find(None, None, 1)[0]


def number_check_in(num):
    timeIn = strftime("%I:%M %p")
    row, col = find(num, None, 1)
    if row is None and col is None:
        add_member(num)
    elif ws.cell(row=row, column=ColToInt(attCol)).value is None:
        ws.cell(row=row, column=ColToInt(attCol), value=timeIn)
        add_name_listbox(row)
    save()


def name_check_in(name):
    row, col = find(name, None, 2)
    if row is None and col is None:
        add_member(name, True)
    elif ws.cell(row=row, column=ColToInt(attCol)).value is None:
        timeIn = strftime("%I:%M %p")
        ws.cell(row=row, column=ColToInt(attCol), value=timeIn)
        add_name_listbox(row)
    save()


def check_in(identifier, UseName):
    find_empty_date_column()
    if UseName:
        # Name based check in
        popup.destroy()
        name_check_in(identifier.lower())
    else:
        # Number based check in
        number_check_in(identifier)


def approve_payment(identifier):
    if isinstance(identifier, str):
        identifier = identifier.lower()
        coll = 2
    else:
        coll = 1
    row, col = find(identifier, None, coll)
    if ws.cell(row=row, column=3).value is None:
        ws.cell(row=row, column=3, value='PAID')
    save()


def gather_att_data():
    indices = []
    names = []
    attendData = []
    colTup = ws.columns
    for column in colTup:
        present = 0
        for cell in column:
            if cell.value is not None:
                present += 1

        attendData.append(present - 1)
    attendData = attendData[3:]
    for i in range(len(attendData)):
        indices.append(i)
        if i % 5 == 0:
            names.append('Mt #' + str(i))
        else:
            names.append('')
    return attendData, indices, names


def show_graph():
    data, indices, label = gather_att_data()
    plt.bar(indices, data, 0.6)
    indices = [x + 0.3 for x in indices]
    print(indices)
    plt.xticks(indices, label)
    plt.show()


def add_member(identifier, strr=False):
    choice = messagebox.askquestion('Not In Database',
                                    'Sorry you are not in the database.\nWould you like to become a member?')

    def create_mem_row(num, name):
        num = int(num)
        row = emptyRow()
        ws.cell(row=row, column=1, value=num)
        ws.cell(row=row, column=2, value=name)
        check_in(num, False)
        save()

    if choice == 'yes':
        FLname = None
        StuNumber = None
        questionF = Toplevel()
        if strr:
            FLname = identifier
            Label(questionF, text='Enter Student Number:').pack()
            num = Entry(questionF)
            num.pack()

            def num_ret():
                StuNumber = num.get()
                questionF.destroy()
                create_mem_row(StuNumber, FLname)

            Button(questionF, text='Submit', command=num_ret).pack()

        else:
            StuNumber = identifier
            Label(questionF, text='Enter First Name: ').pack()
            Fn = Entry(questionF)
            Fn.pack()
            Label(questionF, text='Enter Last Name: ').pack()
            Ln = Entry(questionF)
            Ln.pack()

            def name_ret():
                FLname = Fn.get() + ' ' + Ln.get()
                FLname = FLname.lower()
                create_mem_row(StuNumber, FLname)
                questionF.destroy()

            Button(questionF, text='Submit', command=name_ret).pack()


def check_eligible(ID):
    row = find(ID, None, 2)[0]
    days = ws.iter_rows('D' + str(row) + ':' + attCol + str(row))
    notPres = 0
    attend = 0
    for i in list(days)[0]:
        if i.value is None:
            notPres += 1
        else:
            attend += 1
    perc = round(attend / (attend + notPres) * 100, 0)

    # must add math to count attendance to classes
    if perc > 60:
        return 'Yes'
    else:
        return 'No'


signInList = []

root = Tk()
root.resizable(width=False, height=False)
Style().configure("TButton", relief="flat", padding=5, font='Times 14 bold')


def name_check_in_GUI():
    global popup, FN, LN
    popup = Toplevel()
    popup.tkraise(root)
    Label(popup, text='Enter First Name').grid(row=1, column=1, sticky=W)
    FN = Entry(popup, font='Times 15 bold')
    FN.grid(row=1, column=2, sticky=W)
    Label(popup, text='Enter Last Name').grid(row=2, column=1, sticky=W)
    LN = Entry(popup, font='Times 15 bold')
    LN.grid(row=2, column=2, sticky=W)
    Button(popup, text='Check In', command=lambda: retrieve_name_entry()).grid(row=3, column=1, sticky=W)


def retrieve_name_entry():
    f = FN.get()
    l = LN.get()
    if f.strip() == '' or l.strip() == '':
        FN.delete(0, END)
        LN.delete(0, END)
    else:
        name = f + ' ' + l
        check_in(name, True)


menubar = Menu(root)
menuB = Menu(menubar, tearoff=0)
menuB2 = Menu(menubar, tearoff=0)
menuB3 = Menu(menubar, tearoff=1)
menuB.add_command(label="Exit", command=lambda: quit_handler())
menuB2.add_command(label="Attendance Graph", command=show_graph)
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

barcodeEntry = Entry(top, font='Times 24 bold', width=44)
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


def retrieve_entry(name=False):
    identifier = barcodeEntry.get()
    if name is False:
        identifier = int(identifier)

    barcodeEntry.delete(0, END)
    check_in(identifier, name)


clearEntry = Button(top, text='Clear', command=lambda: barcodeEntry.delete(0, END)).pack(side=LEFT)
nameEntry = Button(top, text='Name Check In', command=name_check_in_GUI).pack(side=LEFT)

smallMid = Frame(m1)
m1.add(smallMid)

Label(smallMid, text='Members that have checked in', font='Times 15 bold').grid(row=1, column=1)
NumCheckIn = Label(smallMid, text="Checked In: ", font='Times 15 bold')
smallMid.grid_columnconfigure(2, minsize=1300)
NumCheckIn.grid(row=1, column=2)

middle = Frame(m1)
m1.add(middle)
# Members List
scBr = Scrollbar(middle)
scBr.pack(side=RIGHT, fill=Y)

dataList = Listbox(middle, selectmode=SINGLE, font='Times 15 bold', yscrollcommand=scBr.set)
scBr.configure(command=dataList.yview)

dataList.pack(fill=BOTH)


def add_name_listbox(row):
    name = ws.cell(row=row, column=2).value
    name = name.split(' ')
    name = name[0][0].upper() + name[0][1:] + ' ' + name[1][0].upper() + name[1][1:]
    number = ws.cell(row=row, column=1).value
    paid = ws.cell(row=row, column=3).value
    if paid is None:
        paid = 'NO'
    dataList.insert(0, name)
    signInList.insert(0, [name, number, paid])
    NumCheckIn.config(text='Checked In: ' + str(len(signInList)))


# Member information
bottom = Frame(m1)
m1.add(bottom)


def display_member_info(memberIndex):
    def member_info_template(name='Not Applicable', number='Not Applicable', paid='NO'):
        if name == paid == number and name == 'None':
            color1, color2, color3, color4 = None, None, None, None
        else:
            color1, color2, color3, color4 = 'white', 'lightgreen', 'lightgreen', 'red'
        if memberIndex != '':
            eligible = check_eligible(signInList[memberIndex][0])
        eligible = "NO"

        if name == 'Not Applicable':
            color1 = 'Red'

        if paid == 'NO':
            color3 = 'Red'

        if number == 'Not Applicable':
            color2 = 'red'

        if eligible == 'Yes':
            color4 = 'lightgreen'

        nameLab = Label(bottom, text='Name: ' + name, font='Times 15 bold', background=color1)
        nameLab.grid(row=1, column=1, sticky=W)
        numLab = Label(bottom, text='Number: ' + str(number), font='Times 15 bold', background=color2)
        numLab.grid(row=2, column=1, sticky=W)
        payLab = Label(bottom, text='Paid: ' + paid, font='Times 15 bold', background=color3)
        payLab.grid(row=3, column=1, sticky=W)
        eligLab = Label(bottom, text='Eligible for Credit: ' + eligible, font='Times 15 bold', background=color4)
        eligLab.grid(row=4, column=1, sticky=W)

        return [nameLab, numLab, payLab, eligLab]

    memberIndex = (str(memberIndex))
    memberIndex = memberIndex.replace('(', '')
    memberIndex = memberIndex.replace(')', '')
    memberIndex = memberIndex.replace(',', '')
    if memberIndex is '':
        labelIDs = member_info_template('None', 'None', 'None')

    else:
        memberIndex = int(memberIndex)  # must incoperate excel doc into display process
        labelIDs = member_info_template(signInList[memberIndex][0], signInList[memberIndex][1],
                                        signInList[memberIndex][2])
    return labelIDs


def clean_mem_info_panel(labelIDList):
    if labelIDList is None:
        pass
    else:
        l1, l2, l3, l4 = labelIDList[0], labelIDList[1], labelIDList[2], labelIDList[3]
        l1.destroy()
        l2.destroy()
        l3.destroy()
        l4.destroy()


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
