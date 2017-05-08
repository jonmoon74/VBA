VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dataform 
   Caption         =   "Candidate Data"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   OleObjectBlob   =   "dataform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dataform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbutton_Click()
Unload Me
End Sub

Private Sub delbutton_Click()
Dim a As Integer
On Error Resume Next

a = ActiveCell.Row
Rows(a).Select
    Selection.Delete Shift:=xlUp
dataform.OptionButton1 = False
dataform.OptionButton2 = False
dataform.OptionButton3 = False
dataform.OptionButton4 = False
dataform.OptionButton5 = False
dataform.OptionButton6 = False
dataform.OptionButton7 = False
dataform.OptionButton8 = False
dataform.OptionButton9 = False
dataform.OptionButton10 = False
dataform.OptionButton11 = False
dataform.OptionButton12 = False
dataform.OptionButton13 = False
dataform.OptionButton14 = False
dataform.OptionButton15 = False
dataform.OptionButton16 = False
dataform.OptionButton17 = False
dataform.OptionButton18 = False
dataform.OptionButton19 = False
dataform.OptionButton20 = False
dataform.OptionButton21 = False
dataform.OptionButton22 = False
dataform.OptionButton23 = False
dataform.OptionButton24 = False
dataform.OptionButton25 = False
dataform.OptionButton26 = False
dataform.OptionButton27 = False
dataform.OptionButton28 = False
dataform.OptionButton29 = False
dataform.OptionButton30 = False
dataform.OptionButton31 = False
dataform.OptionButton32 = False
dataform.OptionButton33 = False
dataform.OptionButton34 = False
dataform.OptionButton35 = False
dataform.OptionButton36 = False

dataform.empbox.Value = Cells(a, 1).Value
dataform.fname.Value = Cells(a, 2).Value
dataform.sname.Value = Cells(a, 3).Value
dataform.posbox.Value = Cells(a, 4).Value
dataform.manbox.Value = Cells(a, 5).Value
dataform.segbox.Value = Cells(a, 6).Value
dataform.deptbox.Value = Cells(a, 7).Value
dataform.geobox.Value = Cells(a, 8).Value
dataform.lc1.Value = Cells(a, 9).Value
dataform.lc2.Value = Cells(a, 10).Value
dataform.lc3.Value = Cells(a, 11).Value
dataform.cc1.Value = Cells(a, 12).Value
dataform.cc2.Value = Cells(a, 13).Value
dataform.cc3.Value = Cells(a, 14).Value
dataform.o1.Value = Cells(a, 15).Value
dataform.o2.Value = Cells(a, 16).Value
dataform.o3.Value = Cells(a, 17).Value
If Cells(a, 9).Interior.ColorIndex = 3 Then dataform.OptionButton1 = True
If Cells(a, 9).Interior.ColorIndex = 6 Then dataform.OptionButton2 = True
If Cells(a, 9).Interior.ColorIndex = 4 Then dataform.OptionButton3 = True
If Cells(a, 9).Interior.ColorIndex = 33 Then dataform.OptionButton4 = True
If Cells(a, 10).Interior.ColorIndex = 3 Then dataform.OptionButton5 = True
If Cells(a, 10).Interior.ColorIndex = 6 Then dataform.OptionButton6 = True
If Cells(a, 10).Interior.ColorIndex = 4 Then dataform.OptionButton7 = True
If Cells(a, 10).Interior.ColorIndex = 33 Then dataform.OptionButton8 = True
If Cells(a, 11).Interior.ColorIndex = 3 Then dataform.OptionButton9 = True
If Cells(a, 11).Interior.ColorIndex = 6 Then dataform.OptionButton10 = True
If Cells(a, 11).Interior.ColorIndex = 4 Then dataform.OptionButton11 = True
If Cells(a, 11).Interior.ColorIndex = 33 Then dataform.OptionButton12 = True
If Cells(a, 12).Interior.ColorIndex = 3 Then dataform.OptionButton13 = True
If Cells(a, 12).Interior.ColorIndex = 6 Then dataform.OptionButton14 = True
If Cells(a, 12).Interior.ColorIndex = 4 Then dataform.OptionButton15 = True
If Cells(a, 12).Interior.ColorIndex = 33 Then dataform.OptionButton16 = True
If Cells(a, 13).Interior.ColorIndex = 3 Then dataform.OptionButton17 = True
If Cells(a, 13).Interior.ColorIndex = 6 Then dataform.OptionButton18 = True
If Cells(a, 13).Interior.ColorIndex = 4 Then dataform.OptionButton19 = True
If Cells(a, 13).Interior.ColorIndex = 33 Then dataform.OptionButton20 = True
If Cells(a, 14).Interior.ColorIndex = 3 Then dataform.OptionButton21 = True
If Cells(a, 14).Interior.ColorIndex = 6 Then dataform.OptionButton22 = True
If Cells(a, 14).Interior.ColorIndex = 4 Then dataform.OptionButton23 = True
If Cells(a, 14).Interior.ColorIndex = 33 Then dataform.OptionButton24 = True
If Cells(a, 15).Interior.ColorIndex = 3 Then dataform.OptionButton25 = True
If Cells(a, 15).Interior.ColorIndex = 6 Then dataform.OptionButton26 = True
If Cells(a, 15).Interior.ColorIndex = 4 Then dataform.OptionButton27 = True
If Cells(a, 15).Interior.ColorIndex = 33 Then dataform.OptionButton28 = True
If Cells(a, 16).Interior.ColorIndex = 3 Then dataform.OptionButton29 = True
If Cells(a, 16).Interior.ColorIndex = 6 Then dataform.OptionButton30 = True
If Cells(a, 16).Interior.ColorIndex = 4 Then dataform.OptionButton31 = True
If Cells(a, 16).Interior.ColorIndex = 33 Then dataform.OptionButton32 = True
If Cells(a, 17).Interior.ColorIndex = 3 Then dataform.OptionButton33 = True
If Cells(a, 17).Interior.ColorIndex = 6 Then dataform.OptionButton34 = True
If Cells(a, 17).Interior.ColorIndex = 4 Then dataform.OptionButton35 = True
If Cells(a, 17).Interior.ColorIndex = 33 Then dataform.OptionButton36 = True
End Sub

Private Sub newbutton_Click()
Dim x As Integer

If Cells(3, 2) <> "" Then
ActiveSheet.Range("a2").End(xlDown).Select
x = ActiveCell.Row + 1

Cells(x, 1).Value = dataform.empbox.Value
Cells(x, 2).Value = dataform.fname.Value
Cells(x, 3).Value = dataform.sname.Value
Cells(x, 4).Value = dataform.posbox.Value
Cells(x, 5).Value = dataform.manbox.Value
Cells(x, 6).Value = dataform.segbox.Value
Cells(x, 7).Value = dataform.deptbox.Value
Cells(x, 8).Value = dataform.geobox.Value
Cells(x, 9).Value = dataform.lc1.Value
Cells(x, 10).Value = dataform.lc2.Value
Cells(x, 11).Value = dataform.lc3.Value
Cells(x, 12).Value = dataform.cc1.Value
Cells(x, 13).Value = dataform.cc2.Value
Cells(x, 14).Value = dataform.cc3.Value
Cells(x, 15).Value = dataform.o1.Value
Cells(x, 16).Value = dataform.o2.Value
Cells(x, 17).Value = dataform.o3.Value
If dataform.OptionButton1 = True Then Cells(x, 9).Interior.ColorIndex = 3
If dataform.OptionButton2 = True Then Cells(x, 9).Interior.ColorIndex = 6
If dataform.OptionButton3 = True Then Cells(x, 9).Interior.ColorIndex = 4
If dataform.OptionButton4 = True Then Cells(x, 9).Interior.ColorIndex = 33
If dataform.OptionButton5 = True Then Cells(x, 10).Interior.ColorIndex = 3
If dataform.OptionButton6 = True Then Cells(x, 10).Interior.ColorIndex = 6
If dataform.OptionButton7 = True Then Cells(x, 10).Interior.ColorIndex = 4
If dataform.OptionButton8 = True Then Cells(x, 10).Interior.ColorIndex = 33
If dataform.OptionButton9 = True Then Cells(x, 11).Interior.ColorIndex = 3
If dataform.OptionButton10 = True Then Cells(x, 11).Interior.ColorIndex = 6
If dataform.OptionButton11 = True Then Cells(x, 11).Interior.ColorIndex = 4
If dataform.OptionButton12 = True Then Cells(x, 11).Interior.ColorIndex = 33
If dataform.OptionButton13 = True Then Cells(x, 12).Interior.ColorIndex = 3
If dataform.OptionButton14 = True Then Cells(x, 12).Interior.ColorIndex = 6
If dataform.OptionButton15 = True Then Cells(x, 12).Interior.ColorIndex = 4
If dataform.OptionButton16 = True Then Cells(x, 12).Interior.ColorIndex = 33
If dataform.OptionButton17 = True Then Cells(x, 13).Interior.ColorIndex = 3
If dataform.OptionButton18 = True Then Cells(x, 13).Interior.ColorIndex = 6
If dataform.OptionButton19 = True Then Cells(x, 13).Interior.ColorIndex = 4
If dataform.OptionButton20 = True Then Cells(x, 13).Interior.ColorIndex = 33
If dataform.OptionButton21 = True Then Cells(x, 14).Interior.ColorIndex = 3
If dataform.OptionButton22 = True Then Cells(x, 14).Interior.ColorIndex = 6
If dataform.OptionButton23 = True Then Cells(x, 14).Interior.ColorIndex = 4
If dataform.OptionButton24 = True Then Cells(x, 14).Interior.ColorIndex = 33
If dataform.OptionButton25 = True Then Cells(x, 15).Interior.ColorIndex = 3
If dataform.OptionButton26 = True Then Cells(x, 15).Interior.ColorIndex = 6
If dataform.OptionButton27 = True Then Cells(x, 15).Interior.ColorIndex = 4
If dataform.OptionButton28 = True Then Cells(x, 15).Interior.ColorIndex = 33
If dataform.OptionButton29 = True Then Cells(x, 16).Interior.ColorIndex = 3
If dataform.OptionButton30 = True Then Cells(x, 16).Interior.ColorIndex = 6
If dataform.OptionButton31 = True Then Cells(x, 16).Interior.ColorIndex = 4
If dataform.OptionButton32 = True Then Cells(x, 16).Interior.ColorIndex = 33
If dataform.OptionButton33 = True Then Cells(x, 17).Interior.ColorIndex = 3
If dataform.OptionButton34 = True Then Cells(x, 17).Interior.ColorIndex = 6
If dataform.OptionButton35 = True Then Cells(x, 17).Interior.ColorIndex = 4
If dataform.OptionButton36 = True Then Cells(x, 17).Interior.ColorIndex = 33

Else

Cells(3, 1).Value = dataform.empbox.Value
Cells(3, 2).Value = dataform.fname.Value
Cells(3, 3).Value = dataform.sname.Value
Cells(3, 4).Value = dataform.posbox.Value
Cells(3, 5).Value = dataform.manbox.Value
Cells(3, 6).Value = dataform.segbox.Value
Cells(3, 7).Value = dataform.deptbox.Value
Cells(3, 8).Value = dataform.geobox.Value
Cells(3, 9).Value = dataform.lc1.Value
Cells(3, 10).Value = dataform.lc2.Value
Cells(3, 11).Value = dataform.lc3.Value
Cells(3, 12).Value = dataform.cc1.Value
Cells(3, 13).Value = dataform.cc2.Value
Cells(3, 14).Value = dataform.cc3.Value
Cells(3, 15).Value = dataform.o1.Value
Cells(3, 16).Value = dataform.o2.Value
Cells(3, 17).Value = dataform.o3.Value
If dataform.OptionButton1 = True Then Cells(3, 9).Interior.ColorIndex = 3
If dataform.OptionButton2 = True Then Cells(3, 9).Interior.ColorIndex = 6
If dataform.OptionButton3 = True Then Cells(3, 9).Interior.ColorIndex = 4
If dataform.OptionButton4 = True Then Cells(3, 9).Interior.ColorIndex = 33
If dataform.OptionButton5 = True Then Cells(3, 10).Interior.ColorIndex = 3
If dataform.OptionButton6 = True Then Cells(3, 10).Interior.ColorIndex = 6
If dataform.OptionButton7 = True Then Cells(3, 10).Interior.ColorIndex = 4
If dataform.OptionButton8 = True Then Cells(3, 10).Interior.ColorIndex = 33
If dataform.OptionButton9 = True Then Cells(3, 11).Interior.ColorIndex = 3
If dataform.OptionButton10 = True Then Cells(3, 11).Interior.ColorIndex = 6
If dataform.OptionButton11 = True Then Cells(3, 11).Interior.ColorIndex = 4
If dataform.OptionButton12 = True Then Cells(3, 11).Interior.ColorIndex = 33
If dataform.OptionButton13 = True Then Cells(3, 12).Interior.ColorIndex = 3
If dataform.OptionButton14 = True Then Cells(3, 12).Interior.ColorIndex = 6
If dataform.OptionButton15 = True Then Cells(3, 12).Interior.ColorIndex = 4
If dataform.OptionButton16 = True Then Cells(3, 12).Interior.ColorIndex = 33
If dataform.OptionButton17 = True Then Cells(3, 13).Interior.ColorIndex = 3
If dataform.OptionButton18 = True Then Cells(3, 13).Interior.ColorIndex = 6
If dataform.OptionButton19 = True Then Cells(3, 13).Interior.ColorIndex = 4
If dataform.OptionButton20 = True Then Cells(3, 13).Interior.ColorIndex = 33
If dataform.OptionButton21 = True Then Cells(3, 14).Interior.ColorIndex = 3
If dataform.OptionButton22 = True Then Cells(3, 14).Interior.ColorIndex = 6
If dataform.OptionButton23 = True Then Cells(3, 14).Interior.ColorIndex = 4
If dataform.OptionButton24 = True Then Cells(3, 14).Interior.ColorIndex = 33
If dataform.OptionButton25 = True Then Cells(3, 15).Interior.ColorIndex = 3
If dataform.OptionButton26 = True Then Cells(3, 15).Interior.ColorIndex = 6
If dataform.OptionButton27 = True Then Cells(3, 15).Interior.ColorIndex = 4
If dataform.OptionButton28 = True Then Cells(3, 15).Interior.ColorIndex = 33
If dataform.OptionButton29 = True Then Cells(3, 16).Interior.ColorIndex = 3
If dataform.OptionButton30 = True Then Cells(3, 16).Interior.ColorIndex = 6
If dataform.OptionButton31 = True Then Cells(3, 16).Interior.ColorIndex = 4
If dataform.OptionButton32 = True Then Cells(3, 16).Interior.ColorIndex = 33
If dataform.OptionButton33 = True Then Cells(3, 17).Interior.ColorIndex = 3
If dataform.OptionButton34 = True Then Cells(3, 17).Interior.ColorIndex = 6
If dataform.OptionButton35 = True Then Cells(3, 17).Interior.ColorIndex = 4
If dataform.OptionButton36 = True Then Cells(3, 17).Interior.ColorIndex = 33
End If
End Sub

Private Sub okbutton_Click()
Dim a As Integer

a = ActiveCell.Row

Cells(a, 1).Value = dataform.empbox.Value
Cells(a, 2).Value = dataform.fname.Value
Cells(a, 3).Value = dataform.sname.Value
Cells(a, 4).Value = dataform.posbox.Value
Cells(a, 5).Value = dataform.manbox.Value
Cells(a, 6).Value = dataform.segbox.Value
Cells(a, 7).Value = dataform.deptbox.Value
Cells(a, 8).Value = dataform.geobox.Value
Cells(a, 9).Value = dataform.lc1.Value
Cells(a, 10).Value = dataform.lc2.Value
Cells(a, 11).Value = dataform.lc3.Value
Cells(a, 12).Value = dataform.cc1.Value
Cells(a, 13).Value = dataform.cc2.Value
Cells(a, 14).Value = dataform.cc3.Value
Cells(a, 15).Value = dataform.o1.Value
Cells(a, 16).Value = dataform.o2.Value
Cells(a, 17).Value = dataform.o3.Value
If dataform.OptionButton1 = True Then Cells(a, 9).Interior.ColorIndex = 3
If dataform.OptionButton2 = True Then Cells(a, 9).Interior.ColorIndex = 6
If dataform.OptionButton3 = True Then Cells(a, 9).Interior.ColorIndex = 4
If dataform.OptionButton4 = True Then Cells(a, 9).Interior.ColorIndex = 33
If dataform.OptionButton5 = True Then Cells(a, 10).Interior.ColorIndex = 3
If dataform.OptionButton6 = True Then Cells(a, 10).Interior.ColorIndex = 6
If dataform.OptionButton7 = True Then Cells(a, 10).Interior.ColorIndex = 4
If dataform.OptionButton8 = True Then Cells(a, 10).Interior.ColorIndex = 33
If dataform.OptionButton9 = True Then Cells(a, 11).Interior.ColorIndex = 3
If dataform.OptionButton10 = True Then Cells(a, 11).Interior.ColorIndex = 6
If dataform.OptionButton11 = True Then Cells(a, 11).Interior.ColorIndex = 4
If dataform.OptionButton12 = True Then Cells(a, 11).Interior.ColorIndex = 33
If dataform.OptionButton13 = True Then Cells(a, 12).Interior.ColorIndex = 3
If dataform.OptionButton14 = True Then Cells(a, 12).Interior.ColorIndex = 6
If dataform.OptionButton15 = True Then Cells(a, 12).Interior.ColorIndex = 4
If dataform.OptionButton16 = True Then Cells(a, 12).Interior.ColorIndex = 33
If dataform.OptionButton17 = True Then Cells(a, 13).Interior.ColorIndex = 3
If dataform.OptionButton18 = True Then Cells(a, 13).Interior.ColorIndex = 6
If dataform.OptionButton19 = True Then Cells(a, 13).Interior.ColorIndex = 4
If dataform.OptionButton20 = True Then Cells(a, 13).Interior.ColorIndex = 33
If dataform.OptionButton21 = True Then Cells(a, 14).Interior.ColorIndex = 3
If dataform.OptionButton22 = True Then Cells(a, 14).Interior.ColorIndex = 6
If dataform.OptionButton23 = True Then Cells(a, 14).Interior.ColorIndex = 4
If dataform.OptionButton24 = True Then Cells(a, 14).Interior.ColorIndex = 33
If dataform.OptionButton25 = True Then Cells(a, 15).Interior.ColorIndex = 3
If dataform.OptionButton26 = True Then Cells(a, 15).Interior.ColorIndex = 6
If dataform.OptionButton27 = True Then Cells(a, 15).Interior.ColorIndex = 4
If dataform.OptionButton28 = True Then Cells(a, 15).Interior.ColorIndex = 33
If dataform.OptionButton29 = True Then Cells(a, 16).Interior.ColorIndex = 3
If dataform.OptionButton30 = True Then Cells(a, 16).Interior.ColorIndex = 6
If dataform.OptionButton31 = True Then Cells(a, 16).Interior.ColorIndex = 4
If dataform.OptionButton32 = True Then Cells(a, 16).Interior.ColorIndex = 33
If dataform.OptionButton33 = True Then Cells(a, 17).Interior.ColorIndex = 3
If dataform.OptionButton34 = True Then Cells(a, 17).Interior.ColorIndex = 6
If dataform.OptionButton35 = True Then Cells(a, 17).Interior.ColorIndex = 4
If dataform.OptionButton36 = True Then Cells(a, 17).Interior.ColorIndex = 33

Unload Me
End Sub

Private Sub printbutton_Click()
dataform.PrintForm
End Sub

Private Sub searchbutton_Click()
Dim a As Integer

dataform.OptionButton1 = False
dataform.OptionButton2 = False
dataform.OptionButton3 = False
dataform.OptionButton4 = False
dataform.OptionButton5 = False
dataform.OptionButton6 = False
dataform.OptionButton7 = False
dataform.OptionButton8 = False
dataform.OptionButton9 = False
dataform.OptionButton10 = False
dataform.OptionButton11 = False
dataform.OptionButton12 = False
dataform.OptionButton13 = False
dataform.OptionButton14 = False
dataform.OptionButton15 = False
dataform.OptionButton16 = False
dataform.OptionButton17 = False
dataform.OptionButton18 = False
dataform.OptionButton19 = False
dataform.OptionButton20 = False
dataform.OptionButton21 = False
dataform.OptionButton22 = False
dataform.OptionButton23 = False
dataform.OptionButton24 = False
dataform.OptionButton25 = False
dataform.OptionButton26 = False
dataform.OptionButton27 = False
dataform.OptionButton28 = False
dataform.OptionButton29 = False
dataform.OptionButton30 = False
dataform.OptionButton31 = False
dataform.OptionButton32 = False
dataform.OptionButton33 = False
dataform.OptionButton34 = False
dataform.OptionButton35 = False
dataform.OptionButton36 = False

On Error Resume Next
ActiveSheet.Cells.Find(what:=dataform.searchbox.Value, after:=ActiveCell, LookIn:=xlValues, lookat:=xlPart, _
searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate

a = ActiveCell.Row

dataform.empbox.Value = Cells(a, 1).Value
dataform.fname.Value = Cells(a, 2).Value
dataform.sname.Value = Cells(a, 3).Value
dataform.posbox.Value = Cells(a, 4).Value
dataform.manbox.Value = Cells(a, 5).Value
dataform.segbox.Value = Cells(a, 6).Value
dataform.deptbox.Value = Cells(a, 7).Value
dataform.geobox.Value = Cells(a, 8).Value
dataform.lc1.Value = Cells(a, 9).Value
dataform.lc2.Value = Cells(a, 10).Value
dataform.lc3.Value = Cells(a, 11).Value
dataform.cc1.Value = Cells(a, 12).Value
dataform.cc2.Value = Cells(a, 13).Value
dataform.cc3.Value = Cells(a, 14).Value
dataform.o1.Value = Cells(a, 15).Value
dataform.o2.Value = Cells(a, 16).Value
dataform.o3.Value = Cells(a, 17).Value
If Cells(a, 9).Interior.ColorIndex = 3 Then dataform.OptionButton1 = True
If Cells(a, 9).Interior.ColorIndex = 6 Then dataform.OptionButton2 = True
If Cells(a, 9).Interior.ColorIndex = 4 Then dataform.OptionButton3 = True
If Cells(a, 9).Interior.ColorIndex = 33 Then dataform.OptionButton4 = True
If Cells(a, 10).Interior.ColorIndex = 3 Then dataform.OptionButton5 = True
If Cells(a, 10).Interior.ColorIndex = 6 Then dataform.OptionButton6 = True
If Cells(a, 10).Interior.ColorIndex = 4 Then dataform.OptionButton7 = True
If Cells(a, 10).Interior.ColorIndex = 33 Then dataform.OptionButton8 = True
If Cells(a, 11).Interior.ColorIndex = 3 Then dataform.OptionButton9 = True
If Cells(a, 11).Interior.ColorIndex = 6 Then dataform.OptionButton10 = True
If Cells(a, 11).Interior.ColorIndex = 4 Then dataform.OptionButton11 = True
If Cells(a, 11).Interior.ColorIndex = 33 Then dataform.OptionButton12 = True
If Cells(a, 12).Interior.ColorIndex = 3 Then dataform.OptionButton13 = True
If Cells(a, 12).Interior.ColorIndex = 6 Then dataform.OptionButton14 = True
If Cells(a, 12).Interior.ColorIndex = 4 Then dataform.OptionButton15 = True
If Cells(a, 12).Interior.ColorIndex = 33 Then dataform.OptionButton16 = True
If Cells(a, 13).Interior.ColorIndex = 3 Then dataform.OptionButton17 = True
If Cells(a, 13).Interior.ColorIndex = 6 Then dataform.OptionButton18 = True
If Cells(a, 13).Interior.ColorIndex = 4 Then dataform.OptionButton19 = True
If Cells(a, 13).Interior.ColorIndex = 33 Then dataform.OptionButton20 = True
If Cells(a, 14).Interior.ColorIndex = 3 Then dataform.OptionButton21 = True
If Cells(a, 14).Interior.ColorIndex = 6 Then dataform.OptionButton22 = True
If Cells(a, 14).Interior.ColorIndex = 4 Then dataform.OptionButton23 = True
If Cells(a, 14).Interior.ColorIndex = 33 Then dataform.OptionButton24 = True
If Cells(a, 15).Interior.ColorIndex = 3 Then dataform.OptionButton25 = True
If Cells(a, 15).Interior.ColorIndex = 6 Then dataform.OptionButton26 = True
If Cells(a, 15).Interior.ColorIndex = 4 Then dataform.OptionButton27 = True
If Cells(a, 15).Interior.ColorIndex = 33 Then dataform.OptionButton28 = True
If Cells(a, 16).Interior.ColorIndex = 3 Then dataform.OptionButton29 = True
If Cells(a, 16).Interior.ColorIndex = 6 Then dataform.OptionButton30 = True
If Cells(a, 16).Interior.ColorIndex = 4 Then dataform.OptionButton31 = True
If Cells(a, 16).Interior.ColorIndex = 33 Then dataform.OptionButton32 = True
If Cells(a, 17).Interior.ColorIndex = 3 Then dataform.OptionButton33 = True
If Cells(a, 17).Interior.ColorIndex = 6 Then dataform.OptionButton34 = True
If Cells(a, 17).Interior.ColorIndex = 4 Then dataform.OptionButton35 = True
If Cells(a, 17).Interior.ColorIndex = 33 Then dataform.OptionButton36 = True
End Sub



Private Sub SpinButton1_SpinDown()
Dim y As Integer

dataform.OptionButton1 = False
dataform.OptionButton2 = False
dataform.OptionButton3 = False
dataform.OptionButton4 = False
dataform.OptionButton5 = False
dataform.OptionButton6 = False
dataform.OptionButton7 = False
dataform.OptionButton8 = False
dataform.OptionButton9 = False
dataform.OptionButton10 = False
dataform.OptionButton11 = False
dataform.OptionButton12 = False
dataform.OptionButton13 = False
dataform.OptionButton14 = False
dataform.OptionButton15 = False
dataform.OptionButton16 = False
dataform.OptionButton17 = False
dataform.OptionButton18 = False
dataform.OptionButton19 = False
dataform.OptionButton20 = False
dataform.OptionButton21 = False
dataform.OptionButton22 = False
dataform.OptionButton23 = False
dataform.OptionButton24 = False
dataform.OptionButton25 = False
dataform.OptionButton26 = False
dataform.OptionButton27 = False
dataform.OptionButton28 = False
dataform.OptionButton29 = False
dataform.OptionButton30 = False
dataform.OptionButton31 = False
dataform.OptionButton32 = False
dataform.OptionButton33 = False
dataform.OptionButton34 = False
dataform.OptionButton35 = False
dataform.OptionButton36 = False

y = ActiveCell.Row
y = y + 1
Cells(y, 1).Activate

dataform.empbox.Value = Cells(y, 1).Value
dataform.fname.Value = Cells(y, 2).Value
dataform.sname.Value = Cells(y, 3).Value
dataform.posbox.Value = Cells(y, 4).Value
dataform.manbox.Value = Cells(y, 5).Value
dataform.segbox.Value = Cells(y, 6).Value
dataform.deptbox.Value = Cells(y, 7).Value
dataform.geobox.Value = Cells(y, 8).Value
dataform.lc1.Value = Cells(y, 9).Value
dataform.lc2.Value = Cells(y, 10).Value
dataform.lc3.Value = Cells(y, 11).Value
dataform.cc1.Value = Cells(y, 12).Value
dataform.cc2.Value = Cells(y, 13).Value
dataform.cc3.Value = Cells(y, 14).Value
dataform.o1.Value = Cells(y, 15).Value
dataform.o2.Value = Cells(y, 16).Value
dataform.o3.Value = Cells(y, 17).Value
If Cells(y, 9).Interior.ColorIndex = 3 Then dataform.OptionButton1 = True
If Cells(y, 9).Interior.ColorIndex = 6 Then dataform.OptionButton2 = True
If Cells(y, 9).Interior.ColorIndex = 4 Then dataform.OptionButton3 = True
If Cells(y, 9).Interior.ColorIndex = 33 Then dataform.OptionButton4 = True
If Cells(y, 10).Interior.ColorIndex = 3 Then dataform.OptionButton5 = True
If Cells(y, 10).Interior.ColorIndex = 6 Then dataform.OptionButton6 = True
If Cells(y, 10).Interior.ColorIndex = 4 Then dataform.OptionButton7 = True
If Cells(y, 10).Interior.ColorIndex = 33 Then dataform.OptionButton8 = True
If Cells(y, 11).Interior.ColorIndex = 3 Then dataform.OptionButton9 = True
If Cells(y, 11).Interior.ColorIndex = 6 Then dataform.OptionButton10 = True
If Cells(y, 11).Interior.ColorIndex = 4 Then dataform.OptionButton11 = True
If Cells(y, 11).Interior.ColorIndex = 33 Then dataform.OptionButton12 = True
If Cells(y, 12).Interior.ColorIndex = 3 Then dataform.OptionButton13 = True
If Cells(y, 12).Interior.ColorIndex = 6 Then dataform.OptionButton14 = True
If Cells(y, 12).Interior.ColorIndex = 4 Then dataform.OptionButton15 = True
If Cells(y, 12).Interior.ColorIndex = 33 Then dataform.OptionButton16 = True
If Cells(y, 13).Interior.ColorIndex = 3 Then dataform.OptionButton17 = True
If Cells(y, 13).Interior.ColorIndex = 6 Then dataform.OptionButton18 = True
If Cells(y, 13).Interior.ColorIndex = 4 Then dataform.OptionButton19 = True
If Cells(y, 13).Interior.ColorIndex = 33 Then dataform.OptionButton20 = True
If Cells(y, 14).Interior.ColorIndex = 3 Then dataform.OptionButton21 = True
If Cells(y, 14).Interior.ColorIndex = 6 Then dataform.OptionButton22 = True
If Cells(y, 14).Interior.ColorIndex = 4 Then dataform.OptionButton23 = True
If Cells(y, 14).Interior.ColorIndex = 33 Then dataform.OptionButton24 = True
If Cells(y, 15).Interior.ColorIndex = 3 Then dataform.OptionButton25 = True
If Cells(y, 15).Interior.ColorIndex = 6 Then dataform.OptionButton26 = True
If Cells(y, 15).Interior.ColorIndex = 4 Then dataform.OptionButton27 = True
If Cells(y, 15).Interior.ColorIndex = 33 Then dataform.OptionButton28 = True
If Cells(y, 16).Interior.ColorIndex = 3 Then dataform.OptionButton29 = True
If Cells(y, 16).Interior.ColorIndex = 6 Then dataform.OptionButton30 = True
If Cells(y, 16).Interior.ColorIndex = 4 Then dataform.OptionButton31 = True
If Cells(y, 16).Interior.ColorIndex = 33 Then dataform.OptionButton32 = True
If Cells(y, 17).Interior.ColorIndex = 3 Then dataform.OptionButton33 = True
If Cells(y, 17).Interior.ColorIndex = 6 Then dataform.OptionButton34 = True
If Cells(y, 17).Interior.ColorIndex = 4 Then dataform.OptionButton35 = True
If Cells(y, 17).Interior.ColorIndex = 33 Then dataform.OptionButton36 = True
End Sub

Private Sub SpinButton1_SpinUp()
Dim y As Integer

dataform.OptionButton1 = False
dataform.OptionButton2 = False
dataform.OptionButton3 = False
dataform.OptionButton4 = False
dataform.OptionButton5 = False
dataform.OptionButton6 = False
dataform.OptionButton7 = False
dataform.OptionButton8 = False
dataform.OptionButton9 = False
dataform.OptionButton10 = False
dataform.OptionButton11 = False
dataform.OptionButton12 = False
dataform.OptionButton13 = False
dataform.OptionButton14 = False
dataform.OptionButton15 = False
dataform.OptionButton16 = False
dataform.OptionButton17 = False
dataform.OptionButton18 = False
dataform.OptionButton19 = False
dataform.OptionButton20 = False
dataform.OptionButton21 = False
dataform.OptionButton22 = False
dataform.OptionButton23 = False
dataform.OptionButton24 = False
dataform.OptionButton25 = False
dataform.OptionButton26 = False
dataform.OptionButton27 = False
dataform.OptionButton28 = False
dataform.OptionButton29 = False
dataform.OptionButton30 = False
dataform.OptionButton31 = False
dataform.OptionButton32 = False
dataform.OptionButton33 = False
dataform.OptionButton34 = False
dataform.OptionButton35 = False
dataform.OptionButton36 = False


y = ActiveCell.Row
y = y - 1
If y < 3 Then y = "3"
Cells(y, 1).Activate

dataform.empbox.Value = Cells(y, 1).Value
dataform.fname.Value = Cells(y, 2).Value
dataform.sname.Value = Cells(y, 3).Value
dataform.posbox.Value = Cells(y, 4).Value
dataform.manbox.Value = Cells(y, 5).Value
dataform.segbox.Value = Cells(y, 6).Value
dataform.deptbox.Value = Cells(y, 7).Value
dataform.geobox.Value = Cells(y, 8).Value
dataform.lc1.Value = Cells(y, 9).Value
dataform.lc2.Value = Cells(y, 10).Value
dataform.lc3.Value = Cells(y, 11).Value
dataform.cc1.Value = Cells(y, 12).Value
dataform.cc2.Value = Cells(y, 13).Value
dataform.cc3.Value = Cells(y, 14).Value
dataform.o1.Value = Cells(y, 15).Value
dataform.o2.Value = Cells(y, 16).Value
dataform.o3.Value = Cells(y, 17).Value
If Cells(y, 9).Interior.ColorIndex = 3 Then dataform.OptionButton1 = True
If Cells(y, 9).Interior.ColorIndex = 6 Then dataform.OptionButton2 = True
If Cells(y, 9).Interior.ColorIndex = 4 Then dataform.OptionButton3 = True
If Cells(y, 9).Interior.ColorIndex = 33 Then dataform.OptionButton4 = True
If Cells(y, 10).Interior.ColorIndex = 3 Then dataform.OptionButton5 = True
If Cells(y, 10).Interior.ColorIndex = 6 Then dataform.OptionButton6 = True
If Cells(y, 10).Interior.ColorIndex = 4 Then dataform.OptionButton7 = True
If Cells(y, 10).Interior.ColorIndex = 33 Then dataform.OptionButton8 = True
If Cells(y, 11).Interior.ColorIndex = 3 Then dataform.OptionButton9 = True
If Cells(y, 11).Interior.ColorIndex = 6 Then dataform.OptionButton10 = True
If Cells(y, 11).Interior.ColorIndex = 4 Then dataform.OptionButton11 = True
If Cells(y, 11).Interior.ColorIndex = 33 Then dataform.OptionButton12 = True
If Cells(y, 12).Interior.ColorIndex = 3 Then dataform.OptionButton13 = True
If Cells(y, 12).Interior.ColorIndex = 6 Then dataform.OptionButton14 = True
If Cells(y, 12).Interior.ColorIndex = 4 Then dataform.OptionButton15 = True
If Cells(y, 12).Interior.ColorIndex = 33 Then dataform.OptionButton16 = True
If Cells(y, 13).Interior.ColorIndex = 3 Then dataform.OptionButton17 = True
If Cells(y, 13).Interior.ColorIndex = 6 Then dataform.OptionButton18 = True
If Cells(y, 13).Interior.ColorIndex = 4 Then dataform.OptionButton19 = True
If Cells(y, 13).Interior.ColorIndex = 33 Then dataform.OptionButton20 = True
If Cells(y, 14).Interior.ColorIndex = 3 Then dataform.OptionButton21 = True
If Cells(y, 14).Interior.ColorIndex = 6 Then dataform.OptionButton22 = True
If Cells(y, 14).Interior.ColorIndex = 4 Then dataform.OptionButton23 = True
If Cells(y, 14).Interior.ColorIndex = 33 Then dataform.OptionButton24 = True
If Cells(y, 15).Interior.ColorIndex = 3 Then dataform.OptionButton25 = True
If Cells(y, 15).Interior.ColorIndex = 6 Then dataform.OptionButton26 = True
If Cells(y, 15).Interior.ColorIndex = 4 Then dataform.OptionButton27 = True
If Cells(y, 15).Interior.ColorIndex = 33 Then dataform.OptionButton28 = True
If Cells(y, 16).Interior.ColorIndex = 3 Then dataform.OptionButton29 = True
If Cells(y, 16).Interior.ColorIndex = 6 Then dataform.OptionButton30 = True
If Cells(y, 16).Interior.ColorIndex = 4 Then dataform.OptionButton31 = True
If Cells(y, 16).Interior.ColorIndex = 33 Then dataform.OptionButton32 = True
If Cells(y, 17).Interior.ColorIndex = 3 Then dataform.OptionButton33 = True
If Cells(y, 17).Interior.ColorIndex = 6 Then dataform.OptionButton34 = True
If Cells(y, 17).Interior.ColorIndex = 4 Then dataform.OptionButton35 = True
If Cells(y, 17).Interior.ColorIndex = 33 Then dataform.OptionButton36 = True
End Sub

Private Sub updatebutton_Click()
Dim a As Integer

a = ActiveCell.Row

Cells(a, 1).Value = dataform.empbox.Value
Cells(a, 2).Value = dataform.fname.Value
Cells(a, 3).Value = dataform.sname.Value
Cells(a, 4).Value = dataform.posbox.Value
Cells(a, 5).Value = dataform.manbox.Value
Cells(a, 6).Value = dataform.segbox.Value
Cells(a, 7).Value = dataform.deptbox.Value
Cells(a, 8).Value = dataform.geobox.Value
Cells(a, 9).Value = dataform.lc1.Value
Cells(a, 10).Value = dataform.lc2.Value
Cells(a, 11).Value = dataform.lc3.Value
Cells(a, 12).Value = dataform.cc1.Value
Cells(a, 13).Value = dataform.cc2.Value
Cells(a, 14).Value = dataform.cc3.Value
Cells(a, 15).Value = dataform.o1.Value
Cells(a, 16).Value = dataform.o2.Value
Cells(a, 17).Value = dataform.o3.Value
If dataform.OptionButton1 = True Then Cells(a, 9).Interior.ColorIndex = 3
If dataform.OptionButton2 = True Then Cells(a, 9).Interior.ColorIndex = 6
If dataform.OptionButton3 = True Then Cells(a, 9).Interior.ColorIndex = 4
If dataform.OptionButton4 = True Then Cells(a, 9).Interior.ColorIndex = 33
If dataform.OptionButton5 = True Then Cells(a, 10).Interior.ColorIndex = 3
If dataform.OptionButton6 = True Then Cells(a, 10).Interior.ColorIndex = 6
If dataform.OptionButton7 = True Then Cells(a, 10).Interior.ColorIndex = 4
If dataform.OptionButton8 = True Then Cells(a, 10).Interior.ColorIndex = 33
If dataform.OptionButton9 = True Then Cells(a, 11).Interior.ColorIndex = 3
If dataform.OptionButton10 = True Then Cells(a, 11).Interior.ColorIndex = 6
If dataform.OptionButton11 = True Then Cells(a, 11).Interior.ColorIndex = 4
If dataform.OptionButton12 = True Then Cells(a, 11).Interior.ColorIndex = 33
If dataform.OptionButton13 = True Then Cells(a, 12).Interior.ColorIndex = 3
If dataform.OptionButton14 = True Then Cells(a, 12).Interior.ColorIndex = 6
If dataform.OptionButton15 = True Then Cells(a, 12).Interior.ColorIndex = 4
If dataform.OptionButton16 = True Then Cells(a, 12).Interior.ColorIndex = 33
If dataform.OptionButton17 = True Then Cells(a, 13).Interior.ColorIndex = 3
If dataform.OptionButton18 = True Then Cells(a, 13).Interior.ColorIndex = 6
If dataform.OptionButton19 = True Then Cells(a, 13).Interior.ColorIndex = 4
If dataform.OptionButton20 = True Then Cells(a, 13).Interior.ColorIndex = 33
If dataform.OptionButton21 = True Then Cells(a, 14).Interior.ColorIndex = 3
If dataform.OptionButton22 = True Then Cells(a, 14).Interior.ColorIndex = 6
If dataform.OptionButton23 = True Then Cells(a, 14).Interior.ColorIndex = 4
If dataform.OptionButton24 = True Then Cells(a, 14).Interior.ColorIndex = 33
If dataform.OptionButton25 = True Then Cells(a, 15).Interior.ColorIndex = 3
If dataform.OptionButton26 = True Then Cells(a, 15).Interior.ColorIndex = 6
If dataform.OptionButton27 = True Then Cells(a, 15).Interior.ColorIndex = 4
If dataform.OptionButton28 = True Then Cells(a, 15).Interior.ColorIndex = 33
If dataform.OptionButton29 = True Then Cells(a, 16).Interior.ColorIndex = 3
If dataform.OptionButton30 = True Then Cells(a, 16).Interior.ColorIndex = 6
If dataform.OptionButton31 = True Then Cells(a, 16).Interior.ColorIndex = 4
If dataform.OptionButton32 = True Then Cells(a, 16).Interior.ColorIndex = 33
If dataform.OptionButton33 = True Then Cells(a, 17).Interior.ColorIndex = 3
If dataform.OptionButton34 = True Then Cells(a, 17).Interior.ColorIndex = 6
If dataform.OptionButton35 = True Then Cells(a, 17).Interior.ColorIndex = 4
If dataform.OptionButton36 = True Then Cells(a, 17).Interior.ColorIndex = 33
End Sub


Private Sub wipe1_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 9).Delete
Cells(a, 9).Interior.ColorIndex = xlNone
dataform.lc1.Value = Cells(a, 9).Value
dataform.OptionButton1 = False
dataform.OptionButton2 = False
dataform.OptionButton3 = False
dataform.OptionButton4 = False
End Sub

Private Sub wipe2_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 10).ClearContents
Cells(a, 10).Interior.ColorIndex = xlNone
dataform.lc2.Value = Cells(a, 10).Value
dataform.OptionButton5 = False
dataform.OptionButton6 = False
dataform.OptionButton7 = False
dataform.OptionButton8 = False
End Sub

Private Sub wipe3_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 11).ClearContents
Cells(a, 11).Interior.ColorIndex = xlNone
dataform.lc3.Value = Cells(a, 11).Value
dataform.OptionButton9 = False
dataform.OptionButton10 = False
dataform.OptionButton11 = False
dataform.OptionButton12 = False
End Sub

Private Sub wipe4_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 12).ClearContents
Cells(a, 12).Interior.ColorIndex = xlNone
dataform.cc1.Value = Cells(a, 12).Value
dataform.OptionButton13 = False
dataform.OptionButton14 = False
dataform.OptionButton15 = False
dataform.OptionButton16 = False
End Sub

Private Sub wipe5_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 13).ClearContents
Cells(a, 13).Interior.ColorIndex = xlNone
dataform.cc2.Value = Cells(a, 13).Value
dataform.OptionButton17 = False
dataform.OptionButton18 = False
dataform.OptionButton19 = False
dataform.OptionButton20 = False
End Sub

Private Sub wipe6_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 14).ClearContents
Cells(a, 14).Interior.ColorIndex = xlNone
dataform.cc3.Value = Cells(a, 14).Value
dataform.OptionButton21 = False
dataform.OptionButton22 = False
dataform.OptionButton23 = False
dataform.OptionButton24 = False
End Sub

Private Sub wipe7_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 15).ClearContents
Cells(a, 15).Interior.ColorIndex = xlNone
dataform.o1.Value = Cells(a, 15).Value
dataform.OptionButton25 = False
dataform.OptionButton26 = False
dataform.OptionButton27 = False
dataform.OptionButton28 = False
End Sub

Private Sub wipe8_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 16).ClearContents
Cells(a, 16).Interior.ColorIndex = xlNone
dataform.o2.Value = Cells(a, 16).Value
dataform.OptionButton29 = False
dataform.OptionButton30 = False
dataform.OptionButton31 = False
dataform.OptionButton32 = False
End Sub

Private Sub wipe9_Click()
Dim a As Integer
a = ActiveCell.Row
On Error Resume Next
Cells(a, 17).ClearContents
Cells(a, 17).Interior.ColorIndex = xlNone
dataform.o3.Value = Cells(a, 17).Value
dataform.OptionButton33 = False
dataform.OptionButton34 = False
dataform.OptionButton35 = False
dataform.OptionButton36 = False
End Sub
