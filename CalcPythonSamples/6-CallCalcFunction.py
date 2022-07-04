# coding: utf-8
from __future__ import unicode_literals

def CallCalcFunction1(*args):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.ActiveSheet

    FunAccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess",ctx)
    Params = [sheet.getCellRangeByName("A1:A14")] # =SUM(A1:A14)
    Results = FunAccess.callFunction("SUM", Params)
    sheet.getCellByPosition(1,0).setValue(Results)

def CallCalcFunction2(*args):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.ActiveSheet
    
    FunAccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess",ctx)
    Params=[1.2345,1] # =ROUND(1.2345,1)
    Results = FunAccess.callFunction("ROUND", Params)
    sheet.getCellByPosition(0,0).setValue(Results)