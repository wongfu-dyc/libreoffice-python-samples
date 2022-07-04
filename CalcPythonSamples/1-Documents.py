# coding: utf-8
from __future__ import unicode_literals
import uno
from com.sun.star.beans import PropertyValue


def SaveCalcDocumentDemo01():
    Path = r"C:\Path\To\CalcFile.ods" 
    # ConvertToURL
    Pathdoc = uno.systemPathToFileUrl(Path)
    
    # set the properties
    '''
    Props = (PropertyValue(),)
    Props[0].Name  = "Hidden"  #the document will open hidden"
    Props[0].Value = True
    '''
    Props=()

    # loadComponentFromURL
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    doc = desktop.loadComponentFromURL(PathDoc, "_blank", 0, Props)

    save_path = r"C:\Path\To\CalcFile.xls" 
    
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()

    Props = (PropertyValue('FilterName',0,'MS Excel 97',0),)
    #*.xlsx "Calc MS Excel 2007 XML"
    #*.xls "MS Excel 97"
    #*.csv "Text - txt - csv (StarCalc)"
    #*.html "HTML (StarCalc)"
    url = uno.systemPathToFileUrl(save_path)
    doc.storeAsURL(url,Props) 

    # Close Document
    doc.close(True)

def OpenDocumentVisibleMode():
    Path = r"C:\Path\To\CalcFile.ods" 

    # ConvertToURL
    Pathdoc = uno.systemPathToFileUrl(Path)

    # loadComponentFromURL
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop=smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    doc = desktop.loadComponentFromURL(PathDoc,"_blank", 0, ())

def OpenDocumentInvisibleMode():
    Path = r"C:\Path\To\CalcFile.ods" 
    
    # ConvertToURL
    Pathdoc = uno.systemPathToFileUrl(Path)

    # set the properties
    Props = (PropertyValue(),)
    Props[0].Name  = "Hidden"  #the document will open hidden"
    Props[0].Value = True
    
    # loadComponentFromURL
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    doc = desktop.loadComponentFromURL(PathDoc, "_blank", 0, Props)

def CreateCalcDocument():
    Model = "private:factory/scalc"
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    
    doc = desktop.loadComponentFromURL(Model,"_blank", 0, ())

def SaveCalcDocument():
    save_path = r"C:\Path\To\CalcFile.xls" 

    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()

    Props = (PropertyValue('FilterName',0,'MS Excel 97',0),)
    url = uno.systemPathToFileUrl(save_path)
    doc.storeAsURL(url,Props) 

def CloseCalcDocument():
    # Use the method close from the document object
    desktop = XSCRIPTCONTEXT.getDesktop()
    desktop.getCurrentComponent().close(True)

    
def Automatic_calculation():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    MySheet = doc.CurrentController.ActiveSheet

    Auto = doc.isAutomaticCalculationEnabled
    doc.enableAutomaticCalculation(False)
    doc.enableAutomaticCalculation(True)
    doc.calculate() # only for formulas not updated
    doc.calculateAll() #all formulas 