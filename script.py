import os,sys
#import comtypes
from comtypes.client import CreateObject
#import win32api
#import win32con
from win32com.client import Dispatch

#key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
 #                           "Software\\Microsoft\\Office\\16.0\\Excel"
  #                          + "\\Security", 0, win32con.KEY_ALL_ACCESS)
#win32api.RegSetValueEx(key, "AccessVBOM", 0, win32con.REG_DWORD, 2)


#get the excel file

Excelfile = sys.argv[1]


#vba code
strcode = \
'''
Sub Auto_Open()

    Dim eicarPart1 As String
    Dim eicarPart2 As String
    eicarPart1 = "X5O!P%@AP[4\PZX54(P^)7C"
    eicarPart2 = "C)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"

    MsgBox eicarPart1 + eicarPart2
End Sub

'''
print(Excelfile)
cwd = os.getcwd()
com_instance = CreateObject("Excel.Application",dynamic=True)
#com_instance = Dispatch("Excel.Application")
com_instance.Visible = True
com_instance.DisplayAlerts = False
Excelfile = os.path.join(cwd,Excelfile)
objworkbook = com_instance.Workbooks.Open(Excelfile)
xlmodule = objworkbook.VBProject.VBComponents.Add(1)
xlmodule.CodeModule.AddFromString(strcode.strip())
objworkbook.SaveAs(os.path.join(cwd,"result.xlsm"))

com_instance.Quit()

