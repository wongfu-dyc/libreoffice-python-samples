# coding: utf-8
from __future__ import unicode_literals
import uno
from com.sun.star.beans import PropertyValue

def AccessToRanges():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet
    
    #By cell default notation 
    Ran = sheet.getCellRangeByName("C2:G14")
    
    # By name 
    # Ran = sheet.getCellRangeByName("RangeName")
    
    # By coordinates (X1, Y1, X2, Y2) 
    Ran = sheet.getCellRangeByPosition(2, 1, 6, 13)
    
    doc.CurrentController.select(Ran)

def AccessToActiveRange():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    if doc.CurrentController.getSelection().supportsService("com.sun.star.sheet.SheetCellRange"):
        # It's a Range
        ActiveRan =doc.CurrentController.getSelection()
        ActiveRan.setString("Active")

def RangeSelection():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    Ran = sheet.getCellRangeByName("C2:G14")
    
    doc.CurrentController.select(Ran)

def RangeCoordinates():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    CoordAddress = MyRange.RangeAddress
    # Sheet index (Integer) 
    # Ran = MyRange.RangeAddress.Sheet
    # Column rank (Long)    top/left corner
    NumCHG = MyRange.RangeAddress.StartColumn
    # Row rank (Long) top/left corner
    NumLHG = MyRange.RangeAddress.StartRow
    # Column rank (Long) bottom/right corner
    NumCBD = MyRange.RangeAddress.EndColumn
    # Row rank (Long) bottom/right corner
    NumLBD = MyRow.RangeAddress.EndRow
    #Sheet container object 
    sheet = MyRange.Spreadsheet
    # Absolute coordinates (String) 
    Coord = MyRange.AbsoluteNamen

def NamedRanges():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    CellRef = sheet.getCellRangeByName("A6")
    Coord = "A6:C12"
    TheRanges = doc.NamedRanges
    
    # remove range name
    if TheRanges.hasByName("RangeName"):
        TheRanges.removeByName("RangeName")
    
    # new range name
    TheRanges.addNewByName("Rangename", Coord, CellRef.CellAddress, 0)
    Nb = TheRanges.Count
    
    # Check existence 
    Exist = TheRanges.hasByName("RangeName")
    # Get range
    MyRange = TheRanges[0]
    MyRange = TheRanges.getByName("RangeName")

    cell = sheet.getCellByPosition(0,0)
    cell.setValue(Nb)

from com.sun.star.sheet.CellFlags import VALUE,DATETIME,STRING,ANNOTATION,FORMULA,HARDATTR,STYLES,OBJECTS,EDITATTR,FORMATTED 
def EraseRange():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet =  doc.CurrentController.ActiveSheet

    MyRange = sheet.getCellRangeByName("A2:D15")
    
    for EraseMode in (ANNOTATION,DATETIME,FORMULA,STRING,VALUE):
        MyRange.clearContents(EraseMode)
        # sheet.clearContents(EraseMode)
    '''
    for i in range(11):
        EraseMode = 2**i
        MyRange.clearContents(EraseMode)
    '''
    
def CopyCellContents():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    
    sheet = doc.CurrentController.ActiveSheet
    MyRange = sheet.getCellRangeByName("A1:D4")
    MyRange[1,1].setValue(12)
    MyRange[2,2].setValue(15)
    MyRange[3,3].setValue(888)
    Ran = sheet.getCellRangeByName("A5:D8")

    MyTable = MyRange.DataArray # MyTable takes the range dimensions
    #(give values to the tabel elements)
    
    Ran.DataArray = MyTable 
    cell= sheet.getCellByPosition(0,13)
    cell.setString(str(type(MyTable)))


def UseDataArray():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.ActiveSheet

    Ran = sheet.getCellRangeByName("A1:B1")
    Ran.DataArray=((4,5),)

    Ran = sheet.getCellRangeByName("A2:A4")
    Ran.DataArray=((13,),(12,),(11,))

    import random
    Data = [(x,) for x in range(1,15)]
    random.shuffle(Data)
    Data = tuple(Data)
    Ran = sheet.getCellRangeByName("A4:A17")
    Ran.DataArray=Data




import random
from apso_utils import mri
def TraverseCellsInRange():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    
    sheet = doc.CurrentController.ActiveSheet
    MyRange = sheet.getCellRangeByName("A1:D4")
    MyRange.DataArray = ((0,0,0,0),(0,0,0,0),(0,0,0,0),(0,0,0,0))

    MyRanges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")
    MyRanges.insertByName("", MyRange)
    LEnum = MyRanges.Cells.createEnumeration()

    x = 1 
    while LEnum.hasMoreElements():
        x += 1
        MyCell = LEnum.nextElement()
        MyCell.setValue(x)
        # apply instructions to object cell 
