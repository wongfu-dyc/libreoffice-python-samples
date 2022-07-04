# coding: utf-8
from platform import system
import uno

def _get_macro_directory():
    '''
    https://wiki.documentfoundation.org/Macros/Python_Guide/Introduction
    https://help.libreoffice.org/latest/sq/text/sbasic/python/python_locations.html

    '''
    platform = system()
    location=["My Macros:","","LibreOffice Macros:",""]
    if(platform == "Windows"):
        location[1] = r"%APPDATA%\LibreOffice\4\user\Scripts\python"
        location[3] = r"C:\Program Files\LibreOffice\share\Scripts\python"
    elif platform == "Linux":
        location[1] = "~/.config/libreoffice/4/user/Scripts/python"
        location[3] = "/usr/lib/libreoffice/share/Scripts/python"
    elif platform == "OS X":
        location[1]= "~/Library/Application Support/LibreOffice/4/user/Scripts/python/"
    else :
        location[1] = uno.fileUrlToSystemPath(__file__)
    return location

def show_macro_directory():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    sheet = doc.Sheets[0]

    location = _get_macro_directory()
    for i in range(4):
        cell = sheet.getCellRangeByName(f'A{i+1}')
        cell.setString(location[i])


g_exportedScripts = (show_macro_directory,)

