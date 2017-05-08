Attribute VB_Name = "Module1"
Sub namesort()
Dim a As Integer, fullname As String, pos As Integer, cn As String
Dim sn As String
For a = 4 To 6000
    If Cells(a, 3) <> "" Then
    fullname = Cells(a, 3).Value
    pos = InStr(fullname, " ")
    cn = Left(fullname, pos - 1)
    sn = Mid(fullname, pos + 1)
    Cells(a, 2) = sn
    Cells(a, 1) = cn
    End If
Next a
End Sub
