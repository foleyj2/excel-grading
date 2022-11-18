#!/usr/bin/python3
## Trigger VBA code on multiple files
## By Joseph T. Foley <foley AT RU.IS>
## https://towardsdatascience.com/how-to-supercharge-excel-with-python-726b0f8e22c2
## 1. Install WinPython/Anaconda
## 2. xlwings addin install
## 3. Enable User Defined Functions for xlwings (in Excel Add-ins)
## 4. Enable Trust acces to VBA project
##     File > Options > Trust Center > Trust Center Settings > Macro Settings
##     Check "Trust access to the VBA project object model"

from pathlib import PurePath##https://docs.python.org/3/library/pathlib.html#module-pathlib
import xlwings as xw
## TODO:  Learn how to deal with these crazy windows paths
wb_path = PurePath("notebook-eval.xls")
wb_macros = xw.Book("grading-macros.xlsm")
macro_call = wb_macros.macro("ExportToPDFs")
macro_call()#put arguments here if needed
