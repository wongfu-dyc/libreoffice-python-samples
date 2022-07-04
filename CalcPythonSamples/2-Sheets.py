# coding: utf-8
from __future__ import unicode_literals
import uno
from com.sun.star.beans import PropertyValue


def ModifySheets():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()

    # add sheet
    for i in range(1,4):
        Name = "Sheet{0:02d}".format(i)
        doc.Sheets.insertNewByName(Name,0)

    #delete sheet
    doc.Sheets.removeByName("Sheet02")
    #Duplicate a sheet
    doc.Sheets.copyByName("Sheet01","Sheet06", 0)  
    #Move sheet
    doc.Sheets.moveByName("Sheet06", 2)

def AccessToSheets():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.ActiveSheet

    # Sheet list 
    AllSheets = doc.Sheets
    # Number of sheets 
    NumberSheets = doc.Sheets.Count
    # Sheet object (by index [base 0])
    index=0 
    sheet = doc.Sheets[index]
    # Sheet object (by name) 
    SheetName ="Sheet01"
    sheet = doc.Sheets.getByName(SheetName)
    # Check existence (name) 
    Exist = doc.Sheets.hasByName(SheetName)
    # Sheet index 
    Index = sheet.RangeAddress.Sheet
    

def ProtectSheets():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    index=1
    sheet = doc.Sheets[index]
    # Activate sheet 
    doc.CurrentController.ActiveSheet = sheet
    # Protect sheet
    password="123"
    sheet.protect(password)
    # Unprotect sheet 
    sheet.unprotect(password)
    # Tab color 
    sheet.TabColor= 255*(256**2) +255*(256**1) +0*(256**0)


from com.sun.star.sheet.CellFlags import VALUE,DATETIME,STRING,ANNOTATION,FORMULA,HARDATTR,STYLES,OBJECTS,EDITATTR,FORMATTED 
def EraseSheet():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    MyRange = sheet.getCellRangeByName("A2:D15")
    
    for EraseMode in (VALUE,DATETIME,STRING,ANNOTATION,FORMULA,HARDATTR,STYLES,OBJECTS,EDITATTR,FORMATTED):
        sheet.clearContents(EraseMode)
    '''
    for i in range(11):
        EraseMode = 2**i
        sheet.clearContents(EraseMode)
    '''