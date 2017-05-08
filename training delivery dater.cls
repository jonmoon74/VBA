VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim a As Integer

Application.ScreenUpdating = False

For a = 5 To 10000
If Cells(a, 3) <> "" Then
Cells(a, 2) = Cells(a, 7)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 8)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 9)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 10)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 11)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 12)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 13)
Cells(a, 2) = Format(Cells(a, 2), "mmm")
Cells(a, 1) = Cells(a - 1, 1) + 1
End If
Next a

Application.ScreenUpdating = True

End Sub

Private Sub Workbook_Open()
Sheet5.Visible = xlSheetHidden
Sheet1.Activate

End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
Dim a As Integer

Application.ScreenUpdating = False

For a = 5 To 10000
If Cells(a, 3) <> "" Then
Cells(a, 2) = Cells(a, 7)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 8)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 9)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 10)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 11)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 12)
    If Cells(a, 2) = "" Then Cells(a, 2) = Cells(a, 13)
Cells(a, 2) = Format(Cells(a, 2), "mmm")
Cells(a, 1) = Cells(a - 1, 1) + 1
End If
Next a

Application.ScreenUpdating = True
End Sub
