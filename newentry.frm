VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newentry 
   Caption         =   "New Entry"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   OleObjectBlob   =   "newentry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okbutton_Click()
Dim x As Integer, fn As String, sn As String, role As String, stdate As Date


ActiveSheet.Range("a4").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentry.fn.Value
ActiveSheet.Cells(x, 2).Value = newentry.sn.Value
ActiveSheet.Cells(x, 3).Value = newentry.role.Value
ActiveSheet.Cells(x, 4).Value = newentry.stdate.Value
ActiveSheet.Cells(x, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

Unload Me
End Sub

Private Sub submit_Click()
Dim x As Integer, fn As String, sn As String, role As String, stdate As Date


ActiveSheet.Range("a4").End(xlDown).Select
x = ActiveCell.Row + 1
ActiveSheet.Cells(x, 1).Value = newentry.fn.Value
ActiveSheet.Cells(x, 2).Value = newentry.sn.Value
ActiveSheet.Cells(x, 3).Value = newentry.role.Value
ActiveSheet.Cells(x, 4).Value = newentry.stdate.Value
ActiveSheet.Cells(x, 4) = Format(ActiveSheet.Cells(x, 4), "dd/mm/yy")

newentry.fn.Value = ""
newentry.sn.Value = ""
newentry.role.Value = ""
newentry.stdate.Value = ""
End Sub

Private Sub cancelbutton_Click()
Unload Me
End Sub


