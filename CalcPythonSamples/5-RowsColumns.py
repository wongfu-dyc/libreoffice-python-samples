# coding: utf-8
from __future__ import unicode_literals
import time
import random


def RowColumnsProperties(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    for i in range(5):
        for j in range(5):
            cell = sheet.getCellByPosition(i,j)
            cell.setValue((i+1)*(j+1))

    CalcRange = sheet.getCellRangeByName("A1:D15")

    TheRows = CalcRange.Rows
    TheCols = CalcRange.Columns
    NbL = CalcRange.Rows.Count
    NbC = CalcRange.Columns.Count
    TheRow = CalcRange.Rows[1]
    TheCol = CalcRange.Columns[1]

    TheRows.OptimalHeight = True
    TheCol.OptimalWidth = True
    TheCol.IsVisible = False

def InsertRowsColumns(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    rows = sheet.getCellRangeByPosition(0,0,3,3).Rows
    columns = sheet.getCellRangeByPosition(0,0,3,3).Columns
    
    for i in range(20):
        rows.insertByIndex(1, 1)
        cell = sheet.getCellByPosition(0,1)
        cell.setValue(2**i)
    columns.insertByIndex(0, 1)

def DeleteRowsColumns(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    for i in range(500):
        cell = sheet.getCellByPosition(1,i)
        cell.setValue(i)

    rows = sheet.getCellRangeByPosition(0,0,9,9).Rows
    columns = sheet.getCellRangeByPosition(0,0,9,9).Columns    
    for i in range(493):
        rows.removeByIndex(2, 1)
    columns.removeByIndex(0, 1)
