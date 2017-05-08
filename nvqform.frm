VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nvqform 
   Caption         =   "NVQ Data"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14145
   OleObjectBlob   =   "nvqform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "nvqform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancelbutton_Click()
Unload Me
End Sub

Private Sub searchbutton_Click()
FormatDateTime (Date = [vbShortDate])


Dim a As Integer, fname As String, sname As String, role As String, sdate As Date
Dim dept As String, ltwo As String, bytwo As String, datetwo As Date
Dim lthree As String, bythree As String, datethree As Date, course As String
Dim level As Integer, started As Date, leveltwo As String, datefour As Date
Dim levelthree As String, datefive As Date

On Error Resume Next

ActiveSheet.Cells.Find(what:=dataform.searchbox.Value, after:=ActiveCell, LookIn:=xlValues, lookat:=xlPart, _
searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate

a = ActiveCell.Row

dataform.fname.Value = Cells(a, 1).Value
dataform.sname.Value = Cells(a, 2).Value
dataform.role.Value = Cells(a, 3).Value
dataform.sdate.Value = Cells(a, 4).Value
dataform.sdate = Format(Cells(a, 4), "dd/mm/yy")
dataform.dept.Value = Cells(a, 5).Value
dataform.ltwo.Value = Cells(a, 6).Value
dataform.bytwo.Value = Cells(a, 7).Value
dataform.datetwo.Value = Cells(a, 8).Value
dataform.datetwo = Format(Cells(a, 8), "dd/mm/yy")
dataform.lthree.Value = Cells(a, 9).Value
dataform.bythree.Value = Cells(a, 10).Value
dataform.datethree.Value = Cells(a, 11).Value
dataform.datethree = Format(Cells(a, 11), "dd/mm/yy")
dataform.course.Value = Cells(a, 12).Value
dataform.level.Value = Cells(a, 13).Value
dataform.started.Value = Cells(a, 14).Value
dataform.started = Format(Cells(a, 14), "dd/mm/yy")
dataform.leveltwo.Value = Cells(a, 15).Value
dataform.datefour.Value = Cells(a, 16).Value
dataform.datefour = Format(Cells(a, 16), "dd/mm/yy")
dataform.levelthree.Value = Cells(a, 17).Value
dataform.datefive.Value = Cells(a, 18).Value
dataform.datefive = Format(Cells(a, 18), "dd/mm/yy")

End Sub

