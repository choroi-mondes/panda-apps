Attribute VB_Name = "funcSys"
Option Explicit

Function startProc()
    Application.ScreenUpdating = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    Application.Calculation = xlCalculationManual
'    Application.Cursor = xlWait
End Function

Function closeProc()
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.Calculation = xlCalculationAutomatic
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    Application.ScreenUpdating = True
End Function
