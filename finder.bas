Attribute VB_Name = "Module1"
Sub finder1()
Attribute finder1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' finder1 Macro
'

'
    Cells.Find(What:="Bronzefield", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.FindNext(After:=ActiveCell).Activate
End Sub

Sub finderbz()

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer


e = 0

    Cells.Find(What:="Bronzefield", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
     Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
    Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
    Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
    MsgBox e, vbOKOnly, "Bronzefield"
    
    
        
End Sub

Sub finderpb()
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer


e = 0

    Cells.Find(What:="Peterborough", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
     Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
    Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
    Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
    MsgBox e, vbOKOnly, "Peterborough"
End Sub

Sub finderFB()
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer


e = 0

    Cells.Find(What:="Forest Bank", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
     Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
        
    Cells.FindNext(After:=ActiveCell).Activate
        a = ActiveCell.Row
        b = ActiveCell.Column
        c = b + 1
        d = Cells(a, c).Value
        e = e + d
'    Cells.FindNext(After:=ActiveCell).Activate
'        a = ActiveCell.Row
'        b = ActiveCell.Column
'        c = b + 1
'        d = Cells(a, c).Value
'        e = e + d
        
    MsgBox e, vbOKOnly, "Forest Bank"
End Sub
