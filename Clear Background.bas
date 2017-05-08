Attribute VB_Name = "clearback"
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
    Application.CommandBars("Full Screen").Visible = False
    ActiveWindow.View = xlPageBreakPreview
End Sub


