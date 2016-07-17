from openpyxl import load_workbook
from openpyxl.cell import column_index_from_string as ColToInt
from time import strftime

date = strftime('%B %d %Y')  # Gets date in the from of MONTH DATE YEAR
saveFile = 'ATT.xlsx'  # Sets the save file
wb = load_workbook(saveFile)
ws = wb.active


def save():  # saves file
    wb.save(saveFile)


def create_attendance_date(row, column):  # Creates meeting date
    column = ColToInt(column)  # Converts letter column to number
    ws.cell(row=row, column=column, value=date)  # Adds date column
    save()  # Saves file


def find_empty_date_column():
    index = 2  # starts column count at 2('B')
    lastCell = ws['A1']  # first cell is set as 'A1'
    while True:  # runs loop to find empty cell in row
        cell = ws.cell(row=1, column=index)
        if lastCell.value is None and cell.value is None:
            # [date][None][None]...[None] finds 2 empty cells and the
            # uses the previous cell as the point at which to add the
            # date stamp
            create_attendance_date(lastCell.row, lastCell.column)
            break

        elif cell.value == date:
            global attCol
            # Makes attCol known to the whole program for attendance registration
            attCol = (cell.column, ColToInt(cell.column))

        # elif cell.value is None and lastCell.value is not None:
        index += 1
        lastCell = cell


find_empty_date_column()
