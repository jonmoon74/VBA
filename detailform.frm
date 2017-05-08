VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} detailform 
   Caption         =   "Data"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   OleObjectBlob   =   "detailform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "detailform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelbutton_Click()
Unload Me
End Sub


Private Sub cs1a_Click()
Dim b As Integer
b = ActiveCell.Row

Cells(b, 6).Value = detailform.cs1box.Value
Cells(b, 6) = Format(Cells(b, 6), "dd/mm/yy")
detailform.cs1box.Value = Cells(b, 6).Value
detailform.cs1box = Format(Cells(b, 6), "dd/mm/yy")

End Sub

Private Sub cs1c_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 6).Value = ""
detailform.cs1box.Value = ""

End Sub

Private Sub cs2a_Click()
Dim b As Integer

b = ActiveCell.Row

Cells(b, 7).Value = detailform.cs2box.Value
Cells(b, 7) = Format(Cells(b, 7), "dd/mm/yy")
detailform.cs2box.Value = Cells(b, 7).Value
detailform.cs2box = Format(Cells(b, 7), "dd/mm/yy")

End Sub

Private Sub cs2c_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 7).Value = ""
detailform.cs2box.Value = ""
End Sub

Private Sub cs3a_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 8).Value = detailform.cs3box.Value
Cells(b, 8) = Format(Cells(b, 8), "dd/mm/yy")
detailform.cs3box.Value = Cells(b, 8).Value
detailform.cs3box = Format(Cells(b, 8), "dd/mm/yy")

End Sub

Private Sub cs3c_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 8).Value = ""
detailform.cs3box.Value = ""
End Sub

Private Sub cs4a_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 9).Value = detailform.cs4box.Value
Cells(b, 9) = Format(Cells(b, 9), "dd/mm/yy")
detailform.cs4box.Value = Cells(b, 9).Value
detailform.cs4box = Format(Cells(b, 9), "dd/mm/yy")

End Sub

Private Sub cs4c_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 9).Value = ""
detailform.cs4box.Value = ""
End Sub

Private Sub lua_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 10).Value = detailform.lubox.Value
Cells(b, 10) = Format(Cells(b, 10), "dd/mm/yy")
detailform.lubox.Value = Cells(b, 10).Value
detailform.lubox = Format(Cells(b, 10), "dd/mm/yy")

End Sub

Private Sub luc_Click()
Dim b As Integer

b = ActiveCell.Row
Cells(b, 10).Value = ""
detailform.lubox.Value = ""
End Sub

Private Sub okbutton_Click()
Dim b As Integer
b = ActiveCell.Row

Cells(b, 1).Value = detailform.fnbox.Value
Cells(b, 2).Value = detailform.snbox.Value
Cells(b, 3).Value = detailform.deptbox.Value
If detailform.yesbutton = True Then Cells(b, 4).Value = "Y"
If detailform.nobutton = True Then Cells(b, 4).Value = "N"
Cells(b, 5).Value = detailform.gldatebox.Value
Cells(b, 5) = Format(Cells(b, 5), "dd/mm/yy")
Cells(b, 6).Value = detailform.cs1box.Value
Cells(b, 6) = Format(Cells(b, 6), "dd/mm/yy")
Cells(b, 7).Value = detailform.cs2box.Value
Cells(b, 7) = Format(Cells(b, 7), "dd/mm/yy")
Cells(b, 8).Value = detailform.cs3box.Value
Cells(b, 8) = Format(Cells(b, 8), "dd/mm/yy")
Cells(b, 9).Value = detailform.cs4box.Value
Cells(b, 9) = Format(Cells(b, 9), "dd/mm/yy")
Cells(b, 10).Value = detailform.lubox.Value
Cells(b, 10) = Format(Cells(b, 10), "dd/mm/yy")

If detailform.yesbutton.Value = True Then
detailform.cs1box.BackColor = &H80000000
detailform.cs2box.BackColor = &H80000000
detailform.cs3box.BackColor = &H80000000
detailform.cs4box.BackColor = &H80000000
Cells(b, 6).Interior.ColorIndex = 15
Cells(b, 7).Interior.ColorIndex = 15
Cells(b, 8).Interior.ColorIndex = 15
Cells(b, 9).Interior.ColorIndex = 15
End If
If detailform.yesbutton.Value = True And detailform.gldatebox.Value = "" Then Cells(b, 5).Interior.ColorIndex = 6

Unload Me

End Sub

Private Sub printbutton_Click()
detailform.PrintForm
End Sub

Private Sub searchbutton_Click()
Dim a As Integer

On Error Resume Next

ActiveSheet.Cells.find(what:=detailform.findbox.Value, after:=ActiveCell, LookIn:=xlValues, lookat:=xlPart, _
searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate

a = ActiveCell.Row

detailform.fnbox.Value = Cells(a, 1).Value
detailform.snbox.Value = Cells(a, 2).Value
detailform.deptbox.Value = Cells(a, 3).Value
If Cells(a, 4).Value = "Y" Then detailform.yesbutton = True
If Cells(a, 4).Value = "N" Then detailform.nobutton = True
detailform.gldatebox.Value = Cells(a, 5).Value
detailform.gldatebox = Format(Cells(a, 5), "dd/mm/yy")
detailform.cs1box.Value = Cells(a, 6).Value
detailform.cs1box = Format(Cells(a, 6), "dd/mm/yy")
detailform.cs2box.Value = Cells(a, 7).Value
detailform.cs2box = Format(Cells(a, 7), "dd/mm/yy")
detailform.cs3box.Value = Cells(a, 8).Value
detailform.cs3box = Format(Cells(a, 8), "dd/mm/yy")
detailform.cs4box.Value = Cells(a, 9).Value
detailform.cs4box = Format(Cells(a, 9), "dd/mm/yy")
detailform.lubox.Value = Cells(a, 10).Value
detailform.lubox = Format(Cells(a, 10), "dd/mm/yy")

If detailform.yesbutton.Value = True And detailform.gldatebox.Value = "" Then
detailform.gldatebox.BackColor = &HFFFF&
detailform.cs1box.BackColor = &H80000000
detailform.cs2box.BackColor = &H80000000
detailform.cs3box.BackColor = &H80000000
detailform.cs4box.BackColor = &H80000000
End If


End Sub


Private Sub SpinButton1_SpinDown()
Dim g As Integer, a As Integer

g = ActiveCell.Row
a = g + 1
Cells(a, 1).Activate

detailform.fnbox.Value = Cells(a, 1).Value
detailform.snbox.Value = Cells(a, 2).Value
detailform.deptbox.Value = Cells(a, 3).Value
If Cells(a, 4).Value = "Y" Then detailform.yesbutton = True
If Cells(a, 4).Value = "N" Then detailform.nobutton = True
detailform.gldatebox.Value = Cells(a, 5).Value
detailform.gldatebox = Format(Cells(a, 5), "dd/mm/yy")
detailform.cs1box.Value = Cells(a, 6).Value
detailform.cs1box = Format(Cells(a, 6), "dd/mm/yy")
detailform.cs2box.Value = Cells(a, 7).Value
detailform.cs2box = Format(Cells(a, 7), "dd/mm/yy")
detailform.cs3box.Value = Cells(a, 8).Value
detailform.cs3box = Format(Cells(a, 8), "dd/mm/yy")
detailform.cs4box.Value = Cells(a, 9).Value
detailform.cs4box = Format(Cells(a, 9), "dd/mm/yy")
detailform.lubox.Value = Cells(a, 10).Value
detailform.lubox = Format(Cells(a, 10), "dd/mm/yy")

If detailform.yesbutton.Value = True And detailform.gldatebox.Value = "" Then
detailform.gldatebox.BackColor = &HFFFF&
detailform.cs1box.BackColor = &H80000000
detailform.cs2box.BackColor = &H80000000
detailform.cs3box.BackColor = &H80000000
detailform.cs4box.BackColor = &H80000000
End If


End Sub

Private Sub SpinButton1_SpinUp()
Dim a As Integer, g As Integer



g = ActiveCell.Row
If g < 3 Then
a = 3
Else
a = g - 1
End If
Cells(a, 1).Activate
detailform.fnbox.Value = Cells(a, 1).Value
detailform.snbox.Value = Cells(a, 2).Value
detailform.deptbox.Value = Cells(a, 3).Value
If Cells(a, 4).Value = "Y" Then detailform.yesbutton = True
If Cells(a, 4).Value = "N" Then detailform.nobutton = True
detailform.gldatebox.Value = Cells(a, 5).Value
detailform.gldatebox = Format(Cells(a, 5), "dd/mm/yy")
detailform.cs1box.Value = Cells(a, 6).Value
detailform.cs1box = Format(Cells(a, 6), "dd/mm/yy")
detailform.cs2box.Value = Cells(a, 7).Value
detailform.cs2box = Format(Cells(a, 7), "dd/mm/yy")
detailform.cs3box.Value = Cells(a, 8).Value
detailform.cs3box = Format(Cells(a, 8), "dd/mm/yy")
detailform.cs4box.Value = Cells(a, 9).Value
detailform.cs4box = Format(Cells(a, 9), "dd/mm/yy")
detailform.lubox.Value = Cells(a, 10).Value
detailform.lubox = Format(Cells(a, 10), "dd/mm/yy")

If detailform.yesbutton.Value = True And detailform.gldatebox.Value = "" Then
detailform.gldatebox.BackColor = &HFFFF&
detailform.cs1box.BackColor = &H80000000
detailform.cs2box.BackColor = &H80000000
detailform.cs3box.BackColor = &H80000000
detailform.cs4box.BackColor = &H80000000
End If


End Sub

Private Sub submitbutton_Click()
Dim b As Integer
b = ActiveCell.Row

Cells(b, 1).Value = detailform.fnbox.Value
Cells(b, 2).Value = detailform.snbox.Value
Cells(b, 3).Value = detailform.deptbox.Value
If detailform.yesbutton = True Then Cells(b, 4).Value = "Y"
If detailform.nobutton = True Then Cells(b, 4).Value = "N"
Cells(b, 5).Value = detailform.gldatebox.Value
Cells(b, 5) = Format(Cells(b, 5), "dd/mm/yy")
Cells(b, 6).Value = detailform.cs1box.Value
Cells(b, 6) = Format(Cells(b, 6), "dd/mm/yy")
Cells(b, 7).Value = detailform.cs2box.Value
Cells(b, 7) = Format(Cells(b, 7), "dd/mm/yy")
Cells(b, 8).Value = detailform.cs3box.Value
Cells(b, 8) = Format(Cells(b, 8), "dd/mm/yy")
Cells(b, 9).Value = detailform.cs4box.Value
Cells(b, 9) = Format(Cells(b, 9), "dd/mm/yy")
Cells(b, 10).Value = detailform.lubox.Value
Cells(b, 10) = Format(Cells(b, 10), "dd/mm/yy")

If detailform.yesbutton.Value = True Then
detailform.cs1box.BackColor = &H80000000
detailform.cs2box.BackColor = &H80000000
detailform.cs3box.BackColor = &H80000000
detailform.cs4box.BackColor = &H80000000
Cells(b, 6).Interior.ColorIndex = 15
Cells(b, 7).Interior.ColorIndex = 15
Cells(b, 8).Interior.ColorIndex = 15
Cells(b, 9).Interior.ColorIndex = 15
End If
If detailform.yesbutton.Value = True And detailform.gldatebox.Value = "" Then Cells(b, 5).Interior.ColorIndex = 6
End Sub


