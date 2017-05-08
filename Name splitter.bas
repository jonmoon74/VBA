Attribute VB_Name = "Module1"
Sub name_splitter()
Dim fullname As String, pos As Integer, fn As String, sn As String, a As Integer

For a = 2 To 600
    fullname = Cells(a, 1).Value
    pos = InStr(fullname, " ")
    fn = Left(fullname, pos - 1)
    sn = Mid(fullname, pos + 1)
    Cells(a, 2) = fn
    Cells(a, 3) = sn
Next a
End Sub
