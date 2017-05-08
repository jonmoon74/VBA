VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} statsform 
   Caption         =   "Statistics Summary"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   OleObjectBlob   =   "statsform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "statsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcbuttone_Click()
Dim a As Integer, b As Integer, r2 As Integer, y2 As Integer, g2 As Integer, b2 As Integer
Dim c As Integer, d As Integer, tot2 As Integer, geogr As String, checker As String

If geog.Value = "" Or geog.Value = "total" Then

Application.ScreenUpdating = False
ActiveSheet.Range("a3").End(xlDown).Select
c = ActiveCell.Row
r2 = 0
y2 = 0
g2 = 0
b2 = 0

For a = 3 To c
For b = 9 To 17
Cells(a, b).Select
    With Selection.Interior
        If Cells(a, b).Interior.ColorIndex = 3 Then r2 = r2 + 1
        If Cells(a, b).Interior.ColorIndex = 6 Then y2 = y2 + 1
        If Cells(a, b).Interior.ColorIndex = 4 Then g2 = g2 + 1
        If Cells(a, b).Interior.ColorIndex = 33 Then b2 = b2 + 1
    End With
Next b
Next a

bookno.Value = r2
tobookno.Value = y2
planno.Value = g2
unplanno.Value = b2
tot2 = r2 + y2 + g2 + b2
totnum.Value = tot2
bookper.Value = Round((r2 / tot2) * 100, 1)
tobookper.Value = Round((y2 / tot2) * 100, 1)
planper.Value = Round((g2 / tot2) * 100, 1)
unplanper.Value = Round((b2 / tot2) * 100, 1)

Else

Application.ScreenUpdating = False
ActiveSheet.Range("a3").End(xlDown).Select
c = ActiveCell.Row
r2 = 0
y2 = 0
g2 = 0
b2 = 0
geogr = geog.Value

For a = 3 To c
For b = 9 To 17
checker = Cells(a, 8).Value
    If checker = geogr Then
    Cells(a, b).Select
        With Selection.Interior
            If Cells(a, b).Interior.ColorIndex = 3 Then r2 = r2 + 1
            If Cells(a, b).Interior.ColorIndex = 6 Then y2 = y2 + 1
            If Cells(a, b).Interior.ColorIndex = 4 Then g2 = g2 + 1
            If Cells(a, b).Interior.ColorIndex = 33 Then b2 = b2 + 1
        End With
    End If
Next b
Next a
 On Error Resume Next
bookno.Value = r2
tobookno.Value = y2
planno.Value = g2
unplanno.Value = b2
tot2 = r2 + y2 + g2 + b2
totnum.Value = tot2
bookper.Value = Round((r2 / tot2) * 100, 1)
tobookper.Value = Round((y2 / tot2) * 100, 1)
planper.Value = Round((g2 / tot2) * 100, 1)
unplanper.Value = Round((b2 / tot2) * 100, 1)
Application.ScreenUpdating = True
End If
End Sub

Private Sub okbutton_Click()
Unload Me
End Sub

Private Sub printbutton_Click()
statsform.PrintForm
End Sub

Private Sub UserForm_Initialize()
Dim a As Integer, b As Integer, red As Integer, yellow As Integer, green As Integer, blue As Integer
Dim c As Integer, d As Integer, totno As Integer

Application.ScreenUpdating = False
ActiveSheet.Range("a3").End(xlDown).Select
c = ActiveCell.Row
red = 0
yellow = 0
green = 0
blue = 0

For a = 3 To c
For b = 9 To 17
Cells(a, b).Select
    With Selection.Interior
        If Cells(a, b).Interior.ColorIndex = 3 Then red = red + 1
        If Cells(a, b).Interior.ColorIndex = 6 Then yellow = yellow + 1
        If Cells(a, b).Interior.ColorIndex = 4 Then green = green + 1
        If Cells(a, b).Interior.ColorIndex = 33 Then blue = blue + 1
    End With
Next b
Next a

bookno.Value = red
tobookno.Value = yellow
planno.Value = green
unplanno.Value = blue
totno = red + yellow + green + blue
totnum.Value = totno
bookper.Value = Round((red / totno) * 100, 1)
tobookper.Value = Round((yellow / totno) * 100, 1)
planper.Value = Round((green / totno) * 100, 1)
unplanper.Value = Round((blue / totno) * 100, 1)

Application.ScreenUpdating = True

End Sub
