import cx_Freeze
import sys
import os
from tkinter import *
import threading
import time
from tkinter import messagebox
from sys import exit as Stp
from openpyxl import load_workbook
from openpyxl.cell import column_index_from_string as ColToInt
from matplotlib import pyplot as plt

os.environ['TCL_LIBRARY'] = "C:\\Users\\Advait\\AppData\\Local\\Programs\\Python\\Python35\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\Users\\Advait\\AppData\\Local\\Programs\\Python\\Python35\\tcl\\tk8.6"

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

execu = [cx_Freeze.Executable("Barcode Scanner Interface.py", base=base)]
opt = {'build_exe': {"packages": ["tkinter", "threading", "time", "openpyxl", "matplotlib"], "include_files":["icon.ico", "Attendance.xlsx", "C:\\Users\\Advait\\AppData\\Local\\Programs\\Python\\Python35\\DLLs\\tcl86t.dll", "C:\\Users\\Advait\\AppData\\Local\\Programs\\Python\\Python35\\DLLs\\tk86t.dll"]}}

cx_Freeze.setup(
    name="Barcode Scanner Interface",
    options=opt,
    version="1.0",
    description="Tool to track attendance using barcodes from student cards.",
    executables=execu
)

