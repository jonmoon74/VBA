VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updateform 
   Caption         =   "Candidate Update Form"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   OleObjectBlob   =   "updateform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "updateform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelbutton1_Click()
Cells(1, 3).Value = ""
Unload Me
adminform.Show False
End Sub

Private Sub findbutton_Click()
Dim a As Integer, fname As String, sname As String, dptmt As String, ssdate As Date, comdate As Date
Dim role As Boolean, effect As Boolean, team As Boolean, assert As Boolean
Dim time As Boolean, accid As Boolean, attend As Boolean, recruit As Boolean, trainthe As Boolean
Dim assess As Boolean
Dim talent As Boolean, pdrs As Boolean


On Error Resume Next

ActiveSheet.Cells.Find(what:=updateform.findfield.Value, after:=ActiveCell, LookIn:=xlValues, lookat:=xlPart, _
searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate

a = ActiveCell.Row
Cells(1, 3).Value = a

updateform.fname.Value = Cells(a, 1).Value
updateform.sname.Value = Cells(a, 2).Value
updateform.dptmt.Value = Cells(a, 3).Value
updateform.ssdate.Value = Cells(a, 4).Value
updateform.ssdate = Format(Cells(a, 4), "dd/mm/yy")
updateform.comdate.Value = Cells(a, 17).Value
updateform.comdate = Format(Cells(a, 17), "dd/mm/yy")
If Cells(a, 5) = "1" Then updateform.role.Value = True
If Cells(a, 5) <> "1" Then updateform.role.Value = False
    If Cells(a, 6) = "1" Then updateform.effect.Value = True
    If Cells(a, 6) <> "1" Then updateform.effect.Value = False
If Cells(a, 7) = "1" Then updateform.team.Value = True
If Cells(a, 7) <> "1" Then updateform.team.Value = False
    If Cells(a, 8) = "1" Then updateform.assert.Value = True
    If Cells(a, 8) <> "1" Then updateform.assert.Value = False
If Cells(a, 9) = "1" Then updateform.time.Value = True
If Cells(a, 9) <> "1" Then updateform.time.Value = False
    If Cells(a, 10) = "1" Then updateform.accid.Value = True
    If Cells(a, 10) <> "1" Then updateform.accid.Value = False
If Cells(a, 11) = "1" Then updateform.attend.Value = True
If Cells(a, 11) <> "1" Then updateform.attend.Value = False
    If Cells(a, 12) = "1" Then updateform.recruit.Value = True
    If Cells(a, 12) <> "1" Then updateform.recruit.Value = False
If Cells(a, 13) = "1" Then updateform.trainthe.Value = True
If Cells(a, 13) <> "1" Then updateform.trainthe.Value = False
    If Cells(a, 14) = "1" Then updateform.assess.Value = True
    If Cells(a, 14) <> "1" Then updateform.assess.Value = False
If Cells(a, 15) = "1" Then updateform.talent.Value = True
If Cells(a, 15) <> "1" Then updateform.talent.Value = False
    If Cells(a, 16) = "1" Then updateform.pdrs.Value = True
    If Cells(a, 16) <> "1" Then updateform.pdrs.Value = False

End Sub

Private Sub okbutton1_Click()
Dim a As Integer, fname As String, sname As String, dptmt As String, ssdate As Date, comdate As Date
Dim role As Boolean, effect As Boolean, team As Boolean, assert As Boolean
Dim time As Boolean, accid As Boolean, attend As Boolean, recruit As Boolean, trainthe As Boolean
Dim assess As Boolean
Dim talent As Boolean, pdrs As Boolean

a = ActiveCell.Row

Cells(a, 1).Value = updateform.fname.Value
Cells(a, 2).Value = updateform.sname.Value
Cells(a, 3).Value = updateform.dptmt.Value
Cells(a, 4).Value = updateform.ssdate.Value
Cells(a, 4) = Format(Cells(a, 4), "dd/mm/yy")
Cells(a, 17).Value = updateform.comdate.Value
Cells(a, 17) = Format(Cells(a, 17), "dd/mm/yy")
If updateform.role.Value = True Then Cells(a, 5).Value = "1"
If updateform.role.Value = False Then Cells(a, 5).Value = ""
    If updateform.effect.Value = True Then Cells(a, 6).Value = "1"
    If updateform.effect.Value = False Then Cells(a, 6).Value = ""
If updateform.team.Value = True Then Cells(a, 7).Value = "1"
If updateform.team.Value = False Then Cells(a, 7).Value = ""
    If updateform.assert.Value = True Then Cells(a, 8).Value = "1"
    If updateform.assert.Value = False Then Cells(a, 8).Value = ""
If updateform.time.Value = True Then Cells(a, 9).Value = "1"
If updateform.time.Value = False Then Cells(a, 9).Value = ""
    If updateform.accid.Value = True Then Cells(a, 10).Value = "1"
    If updateform.accid.Value = False Then Cells(a, 10).Value = ""
If updateform.attend.Value = True Then Cells(a, 11).Value = "1"
If updateform.attend.Value = False Then Cells(a, 11).Value = ""
    If updateform.recruit.Value = True Then Cells(a, 12).Value = "1"
    If updateform.recruit.Value = False Then Cells(a, 12).Value = ""
If updateform.trainthe.Value = True Then Cells(a, 13).Value = "1"
If updateform.trainthe.Value = False Then Cells(a, 13).Value = ""
    If updateform.assess.Value = True Then Cells(a, 14).Value = "1"
    If updateform.assess.Value = False Then Cells(a, 14).Value = ""
If updateform.talent.Value = True Then Cells(a, 15).Value = "1"
If updateform.talent.Value = False Then Cells(a, 15).Value = ""
    If updateform.pdrs.Value = True Then Cells(a, 16).Value = "1"
    If updateform.pdrs.Value = False Then Cells(a, 16).Value = ""


Cells(1, 3).Value = ""
Unload Me
adminform.Show False
End Sub

Private Sub pbutt_Click()
updateform.PrintForm
End Sub

Private Sub SpinButton1_SpinDown()
Dim x As Integer, fname As String, sname As String, dptmt As String, ssdate As Date, comdate As Date
Dim role As Boolean, effect As Boolean, team As Boolean, assert As Boolean
Dim time As Boolean, accid As Boolean, attend As Boolean, recruit As Boolean, trainthe As Boolean
Dim assess As Boolean
Dim talent As Boolean, pdrs As Boolean

Cells(1, 3).Value = Cells(1, 3).Value + 1

x = Cells(1, 3).Value
Cells(x, 1).Activate

updateform.fname.Value = Cells(x, 1).Value
updateform.sname.Value = Cells(x, 2).Value
updateform.dptmt.Value = Cells(x, 3).Value
updateform.ssdate.Value = Cells(x, 4).Value
updateform.ssdate = Format(Cells(x, 4), "dd/mm/yy")
updateform.comdate.Value = Cells(x, 17).Value
updateform.comdate = Format(Cells(x, 17), "dd/mm/yy")
If Cells(x, 5) = "1" Then updateform.role.Value = True
If Cells(x, 5) <> "1" Then updateform.role.Value = False
    If Cells(x, 6) = "1" Then updateform.effect.Value = True
    If Cells(x, 6) <> "1" Then updateform.effect.Value = False
If Cells(x, 7) = "1" Then updateform.team.Value = True
If Cells(x, 7) <> "1" Then updateform.team.Value = False
    If Cells(x, 8) = "1" Then updateform.assert.Value = True
    If Cells(x, 8) <> "1" Then updateform.assert.Value = False
If Cells(x, 9) = "1" Then updateform.time.Value = True
If Cells(x, 9) <> "1" Then updateform.time.Value = False
    If Cells(x, 10) = "1" Then updateform.accid.Value = True
    If Cells(x, 10) <> "1" Then updateform.accid.Value = False
If Cells(x, 11) = "1" Then updateform.attend.Value = True
If Cells(x, 11) <> "1" Then updateform.attend.Value = False
    If Cells(x, 12) = "1" Then updateform.recruit.Value = True
    If Cells(x, 12) <> "1" Then updateform.recruit.Value = False
If Cells(x, 13) = "1" Then updateform.trainthe.Value = True
If Cells(x, 13) <> "1" Then updateform.trainthe.Value = False
    If Cells(x, 14) = "1" Then updateform.assess.Value = True
    If Cells(x, 14) <> "1" Then updateform.assess.Value = False
If Cells(x, 15) = "1" Then updateform.talent.Value = True
If Cells(x, 15) <> "1" Then updateform.talent.Value = False
    If Cells(x, 16) = "1" Then updateform.pdrs.Value = True
    If Cells(x, 16) <> "1" Then updateform.pdrs.Value = False
End Sub

Private Sub SpinButton1_SpinUp()
Dim x As Integer, fname As String, sname As String, dptmt As String, ssdate As Date, comdate As Date
Dim role As Boolean, effect As Boolean, team As Boolean, assert As Boolean
Dim time As Boolean, accid As Boolean, attend As Boolean, recruit As Boolean, trainthe As Boolean
Dim assess As Boolean
Dim talent As Boolean, pdrs As Boolean
Cells(1, 3).Value = Cells(1, 3).Value - 1
If Cells(1, 3).Value < 3 Then Cells(1, 3).Value = "3"

x = Cells(1, 3).Value
Cells(x, 1).Activate

updateform.fname.Value = Cells(x, 1).Value
updateform.sname.Value = Cells(x, 2).Value
updateform.dptmt.Value = Cells(x, 3).Value
updateform.ssdate.Value = Cells(x, 4).Value
updateform.ssdate = Format(Cells(x, 4), "dd/mm/yy")
updateform.comdate.Value = Cells(x, 17).Value
updateform.comdate = Format(Cells(x, 17), "dd/mm/yy")
If Cells(x, 5) = "1" Then updateform.role.Value = True
If Cells(x, 5) <> "1" Then updateform.role.Value = False
    If Cells(x, 6) = "1" Then updateform.effect.Value = True
    If Cells(x, 6) <> "1" Then updateform.effect.Value = False
If Cells(x, 7) = "1" Then updateform.team.Value = True
If Cells(x, 7) <> "1" Then updateform.team.Value = False
    If Cells(x, 8) = "1" Then updateform.assert.Value = True
    If Cells(x, 8) <> "1" Then updateform.assert.Value = False
If Cells(x, 9) = "1" Then updateform.time.Value = True
If Cells(x, 9) <> "1" Then updateform.time.Value = False
    If Cells(x, 10) = "1" Then updateform.accid.Value = True
    If Cells(x, 10) <> "1" Then updateform.accid.Value = False
If Cells(x, 11) = "1" Then updateform.attend.Value = True
If Cells(x, 11) <> "1" Then updateform.attend.Value = False
    If Cells(x, 12) = "1" Then updateform.recruit.Value = True
    If Cells(x, 12) <> "1" Then updateform.recruit.Value = False
If Cells(x, 13) = "1" Then updateform.trainthe.Value = True
If Cells(x, 13) <> "1" Then updateform.trainthe.Value = False
    If Cells(x, 14) = "1" Then updateform.assess.Value = True
    If Cells(x, 14) <> "1" Then updateform.assess.Value = False
If Cells(x, 15) = "1" Then updateform.talent.Value = True
If Cells(x, 15) <> "1" Then updateform.talent.Value = False
    If Cells(x, 16) = "1" Then updateform.pdrs.Value = True
    If Cells(x, 16) <> "1" Then updateform.pdrs.Value = False
End Sub

Private Sub submitbutton1_Click()
Dim a As Integer, fname As String, sname As String, dptmt As String, ssdate As Date, comdate As Date
Dim role As Boolean, effect As Boolean, team As Boolean, assert As Boolean
Dim time As Boolean, accid As Boolean, attend As Boolean, recruit As Boolean, trainthe As Boolean
Dim assess As Boolean
Dim talent As Boolean, pdrs As Boolean

a = ActiveCell.Row

Cells(a, 1).Value = updateform.fname.Value
Cells(a, 2).Value = updateform.sname.Value
Cells(a, 3).Value = updateform.dptmt.Value
Cells(a, 4).Value = updateform.ssdate.Value
Cells(a, 4) = Format(Cells(a, 4), "dd/mm/yy")
Cells(a, 17).Value = updateform.comdate.Value
Cells(a, 17) = Format(Cells(a, 17), "dd/mm/yy")
If updateform.role.Value = True Then Cells(a, 5).Value = "1"
If updateform.role.Value = False Then Cells(a, 5).Value = ""
    If updateform.effect.Value = True Then Cells(a, 6).Value = "1"
    If updateform.effect.Value = False Then Cells(a, 6).Value = ""
If updateform.team.Value = True Then Cells(a, 7).Value = "1"
If updateform.team.Value = False Then Cells(a, 7).Value = ""
    If updateform.assert.Value = True Then Cells(a, 8).Value = "1"
    If updateform.assert.Value = False Then Cells(a, 8).Value = ""
If updateform.time.Value = True Then Cells(a, 9).Value = "1"
If updateform.time.Value = False Then Cells(a, 9).Value = ""
    If updateform.accid.Value = True Then Cells(a, 10).Value = "1"
    If updateform.accid.Value = False Then Cells(a, 10).Value = ""
If updateform.attend.Value = True Then Cells(a, 11).Value = "1"
If updateform.attend.Value = False Then Cells(a, 11).Value = ""
    If updateform.recruit.Value = True Then Cells(a, 12).Value = "1"
    If updateform.recruit.Value = False Then Cells(a, 12).Value = ""
If updateform.trainthe.Value = True Then Cells(a, 13).Value = "1"
If updateform.trainthe.Value = False Then Cells(a, 13).Value = ""
    If updateform.assess.Value = True Then Cells(a, 14).Value = "1"
    If updateform.assess.Value = False Then Cells(a, 14).Value = ""
If updateform.talent.Value = True Then Cells(a, 15).Value = "1"
If updateform.talent.Value = False Then Cells(a, 15).Value = ""
    If updateform.pdrs.Value = True Then Cells(a, 16).Value = "1"
    If updateform.pdrs.Value = False Then Cells(a, 16).Value = ""


End Sub

Private Sub UserForm_Initialize()
Cells(1, 3).Value = "2"
End Sub
