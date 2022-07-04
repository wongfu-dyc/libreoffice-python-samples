# coding: utf-8
from __future__ import unicode_literals
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table.CellContentType import EMPTY,VALUE,TEXT,FORMULA


def AccessToCells():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet
    
    # By cell default notation
    cell = sheet.getCellRangeByName("A4")
    # By coordinates X and Y 
    cell = sheet.getCellByPosition(0,3)

    cell.setString("a")


def AccessToActiveCell():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    if doc.CurrentController.getSelection().supportsService("com.sun.star.sheet.SheetCell"):
        # It's a cell
        Activecell =doc.CurrentController.getSelection()
        ActiveCel.CellBackColor=120*(256**2) +255*(256**1) +120*(256**0)
        ActiveCel.setString("Active")


def SelectCell():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet
    cell = sheet.getCellByPosition(0,3)
    doc.CurrentController.select(cel)


def AccessCellContents():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet
    
    cell = sheet.getCellByPosition(0,0)
    cell.setFormula("=ROUND(POWER(1.2345+0.55;2);2)")

    MyText = cell.getString()
    aNumber = cell.getValue()
    TheFormula = cell.getFormula()
    heType = cell.Type
    cells = sheet.getCellRangeByPosition(0,1,1,4)
    cells[3,1].setString(heType.value)
    
    cells[0,0].setString("Hello !")
    cells[0,1].setString(MyText)
    
    cells[1,0].setValue(1.234)
    cells[1,1].setString(aNumber)
    
    cells[2,0].setFormula('=AND(A1="YES";A2="OK")') 
    cells[2,1].setString(TheFormula)
    
    return


def TraverseCells():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    Cells= sheet.getCellRangeByPosition(0,0,30,120)

    for i in range(11): sheet.clearContents(2**i)
    for i in range(Cells.getRows().Count):
        for j in range(Cells.getColumns().Count):
            Cells[i,j].setValue(i+j*(120+1))