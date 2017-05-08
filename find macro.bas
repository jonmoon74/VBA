Attribute VB_Name = "Module5"
Sub find()
Attribute find.VB_Description = "Macro recorded 02/07/2008 by Jon Moon"
Attribute find.VB_ProcData.VB_Invoke_Func = " \n14"
'
' find Macro
' Macro recorded 02/07/2008 by Jon Moon
'

'
    Cells.find(What:="debra", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
End Sub
