Attribute VB_Name = "Module1"
Sub clearback()
Attribute clearback.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clearback Macro
'

'
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayOutline = False
        .DisplayZeros = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
    End With
    Application.ShowStartupDialog = False
    Application.DisplayFullScreen = True
End Sub
Sub showback()
Attribute showback.VB_ProcData.VB_Invoke_Func = " \n14"
'
' showback Macro
'

'
    Application.DisplayFullScreen = False
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayOutline = True
        .DisplayZeros = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
End Sub
