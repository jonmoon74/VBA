VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newentryform 
   Caption         =   "New Entry"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   OleObjectBlob   =   "newentryform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newentryform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub okbutton_Click()
Dim x As Integer, fn As String, sn As String, role As String, stdate As Date

If Cells(3, 1) <> "" Then

ActiveSheet.Range("a2").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentryform.fn.Value
ActiveSheet.Cells(x, 2).Value = newentryform.sn.Value
ActiveSheet.Cells(x, 3).Value = newentryform.role.Value
ActiveSheet.Cells(x, 4).Value = newentryform.stdate.Value
ActiveSheet.Cells(x, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

Else

ActiveSheet.Cells(3, 1).Value = newentryform.fn.Value
ActiveSheet.Cells(3, 2).Value = newentryform.sn.Value
ActiveSheet.Cells(3, 3).Value = newentryform.role.Value
ActiveSheet.Cells(3, 4).Value = newentryform.stdate.Value
ActiveSheet.Cells(3, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

End If

Unload Me
End Sub

Private Sub submit_Click()
Dim x As Integer, fn As String, sn As String, role As String, stdate As Date


If Cells(3, 1) <> "" Then

ActiveSheet.Range("a2").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentryform.fn.Value
ActiveSheet.Cells(x, 2).Value = newentryform.sn.Value
ActiveSheet.Cells(x, 3).Value = newentryform.role.Value
ActiveSheet.Cells(x, 4).Value = newentryform.stdate.Value
ActiveSheet.Cells(x, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

Else

ActiveSheet.Cells(3, 1).Value = newentryform.fn.Value
ActiveSheet.Cells(3, 2).Value = newentryform.sn.Value
ActiveSheet.Cells(3, 3).Value = newentryform.role.Value
ActiveSheet.Cells(3, 4).Value = newentryform.stdate.Value
ActiveSheet.Cells(3, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

End If

newentryform.fn.Value = ""
newentryform.sn.Value = ""
newentryform.role.Value = ""
newentryform.stdate.Value = ""
End Sub

Private Sub cancelbutton_Click()
Unload Me
End Sub


