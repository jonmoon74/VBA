VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newentry 
   Caption         =   "New Entry"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   OleObjectBlob   =   "newentry2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub okbutton_Click()
Dim x As Integer, fn As String, sn As String, dept As String, gl As String, gldate As Date

If Cells(3, 1) <> "" Then
ActiveSheet.Range("a2").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentry.fn.Value
ActiveSheet.Cells(x, 2).Value = newentry.sn.Value
ActiveSheet.Cells(x, 3).Value = newentry.dept.Value
ActiveSheet.Cells(x, 4).Value = newentry.gl.Value
ActiveSheet.Cells(x, 5).Value = newentry.gldate.Value
ActiveSheet.Cells(x, 5) = Format(ActiveSheet.Cells(x, 5), "dd/mm/yy")

If ActiveSheet.Cells(x, 4).Value = "N" Then ActiveSheet.Cells(x, 5).Interior.ColorIndex = 15
If Cells(x, 4).Value = "Y" And Cells(x, 5) = "" Then Cells(x, 5).Interior.ColorIndex = 6
If Cells(x, 4).Value = "Y" Then
Cells(x, 6).Value = newentry.gldate.Value
Cells(x, 6) = Format(Cells(x, 6), "dd/mm/yy")
Cells(x, 7).Value = newentry.gldate.Value
Cells(x, 7) = Format(Cells(x, 7), "dd/mm/yy")
Cells(x, 8).Value = newentry.gldate.Value
Cells(x, 8) = Format(Cells(x, 8), "dd/mm/yy")
Cells(x, 9).Value = newentry.gldate.Value
Cells(x, 9) = Format(Cells(x, 9), "dd/mm/yy")
Cells(x, 6).Interior.ColorIndex = 15
Cells(x, 7).Interior.ColorIndex = 15
Cells(x, 8).Interior.ColorIndex = 15
Cells(x, 9).Interior.ColorIndex = 15
End If

Else

ActiveSheet.Cells(3, 1).Value = newentry.fn.Value
ActiveSheet.Cells(3, 2).Value = newentry.sn.Value
ActiveSheet.Cells(3, 3).Value = newentry.dept.Value
ActiveSheet.Cells(3, 4).Value = newentry.gl.Value
ActiveSheet.Cells(3, 5).Value = newentry.gldate.Value
ActiveSheet.Cells(3, 5) = Format(ActiveSheet.Cells(3, 5), "dd/mm/yy")

If ActiveSheet.Cells(3, 4).Value = "N" Then ActiveSheet.Cells(3, 5).Interior.ColorIndex = 15
If Cells(3, 4).Value = "Y" And Cells(3, 5) = "" Then Cells(3, 5).Interior.ColorIndex = 6
If Cells(3, 4).Value = "Y" Then
Cells(3, 6).Value = newentry.gldate.Value
Cells(3, 6) = Format(Cells(3, 6), "dd/mm/yy")
Cells(3, 7).Value = newentry.gldate.Value
Cells(3, 7) = Format(Cells(3, 7), "dd/mm/yy")
Cells(3, 8).Value = newentry.gldate.Value
Cells(3, 8) = Format(Cells(3, 8), "dd/mm/yy")
Cells(3, 9).Value = newentry.gldate.Value
Cells(3, 9) = Format(Cells(3, 9), "dd/mm/yy")
Cells(3, 6).Interior.ColorIndex = 15
Cells(3, 7).Interior.ColorIndex = 15
Cells(3, 8).Interior.ColorIndex = 15
Cells(3, 9).Interior.ColorIndex = 15
End If

End If

Unload Me
End Sub

Private Sub submitbutton_Click()
Dim x As Integer, fn As String, sn As String, dept As String, gl As String, gldate As Date

If ActiveSheet.Cells(3, 1) <> "" Then

ActiveSheet.Range("a2").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentry.fn.Value
ActiveSheet.Cells(x, 2).Value = newentry.sn.Value
ActiveSheet.Cells(x, 3).Value = newentry.dept.Value
ActiveSheet.Cells(x, 4).Value = newentry.gl.Value
ActiveSheet.Cells(x, 5).Value = newentry.gldate.Value
ActiveSheet.Cells(x, 5) = Format(ActiveSheet.Cells(x, 5), "dd/mm/yy")

If ActiveSheet.Cells(x, 4).Value = "N" Then ActiveSheet.Cells(x, 5).Interior.ColorIndex = 15
If Cells(x, 4).Value = "Y" And Cells(x, 5) = "" Then Cells(x, 5).Interior.ColorIndex = 6
If Cells(x, 4).Value = "Y" Then
Cells(x, 6).Value = newentry.gldate.Value
Cells(x, 6) = Format(Cells(x, 6), "dd/mm/yy")
Cells(x, 7).Value = newentry.gldate.Value
Cells(x, 7) = Format(Cells(x, 7), "dd/mm/yy")
Cells(x, 8).Value = newentry.gldate.Value
Cells(x, 8) = Format(Cells(x, 8), "dd/mm/yy")
Cells(x, 9).Value = newentry.gldate.Value
Cells(x, 9) = Format(Cells(x, 9), "dd/mm/yy")
Cells(x, 6).Interior.ColorIndex = 15
Cells(x, 7).Interior.ColorIndex = 15
Cells(x, 8).Interior.ColorIndex = 15
Cells(x, 9).Interior.ColorIndex = 15
End If

newentry.fn.Value = ""
newentry.sn.Value = ""
newentry.dept.Value = ""
newentry.gl.Value = ""
newentry.gldate.Value = ""

Else

ActiveSheet.Cells(3, 1).Value = newentry.fn.Value
ActiveSheet.Cells(3, 2).Value = newentry.sn.Value
ActiveSheet.Cells(3, 3).Value = newentry.dept.Value
ActiveSheet.Cells(3, 4).Value = newentry.gl.Value
ActiveSheet.Cells(3, 5).Value = newentry.gldate.Value
ActiveSheet.Cells(3, 5) = Format(ActiveSheet.Cells(3, 5), "dd/mm/yy")

If ActiveSheet.Cells(3, 4).Value = "N" Then ActiveSheet.Cells(3, 5).Interior.ColorIndex = 15
If Cells(3, 4).Value = "Y" And Cells(3, 5) = "" Then Cells(3, 5).Interior.ColorIndex = 6
If Cells(3, 4).Value = "Y" Then
Cells(3, 6).Value = newentry.gldate.Value
Cells(3, 6) = Format(Cells(3, 6), "dd/mm/yy")
Cells(3, 7).Value = newentry.gldate.Value
Cells(3, 7) = Format(Cells(3, 7), "dd/mm/yy")
Cells(3, 8).Value = newentry.gldate.Value
Cells(3, 8) = Format(Cells(3, 8), "dd/mm/yy")
Cells(3, 9).Value = newentry.gldate.Value
Cells(3, 9) = Format(Cells(3, 9), "dd/mm/yy")
Cells(3, 6).Interior.ColorIndex = 15
Cells(3, 7).Interior.ColorIndex = 15
Cells(3, 8).Interior.ColorIndex = 15
Cells(3, 9).Interior.ColorIndex = 15
End If

newentry.fn.Value = ""
newentry.sn.Value = ""
newentry.dept.Value = ""
newentry.gl.Value = ""
newentry.gldate.Value = ""

End If

End Sub

Private Sub Cancelbutton_Click()
Unload Me
End Sub


