Attribute VB_Name = "showback"

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
ActiveWindow.View = xlNormalView

End Sub
