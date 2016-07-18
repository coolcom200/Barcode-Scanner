from openpyxl import load_workbook
from openpyxl.cell import column_index_from_string as ColToInt
from time import strftime

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
    elif row is not None and column is None:
        index = 1
        while True:
            cell = ws.cell(row=row, column=index)
            if cell.value is findValue:
                return cell.row, ColToInt(cell.column)
            index += 1

    elif row is None and column is not None:
        index = 1
        while True:
            cell = ws.cell(row=index, column=column)
            if cell.value == findValue:
                return cell.row, ColToInt(cell.column)
            index += 1


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
            attCol = (cell.column)
            break

        elif cell.value == date:

            # Makes attCol known to the whole program for attendance registration
            attCol = (cell.column)
            break

        # elif cell.value is None and lastCell.value is not None:
        index += 1
        lastCell = cell


def number_check_in(num):
    timeIn = strftime("%I:%M %p")
    row, col = find(num, None, 1)
    if ws.cell(row=row, column=col).value is None:
        ws.cell(row=row, column=ColToInt(attCol), value=timeIn)
    save()


def name_check_in(name):
    timeIn = strftime("%I:%M %p")
    row, col = find(name, None, 2)
    if ws.cell(row=row, column=col).value is None:
        ws.cell(row=row, column=ColToInt(attCol), value=timeIn)
    save()


def check_in(number=None, name=None):
    if number is not None:
        # Name based check in
        number_check_in(number)
    else:
        # Number based check in
        name_check_in(name.lower())


find_empty_date_column()
name_check_in('LF')
number_check_in(2)
