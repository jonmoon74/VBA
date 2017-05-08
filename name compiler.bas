Attribute VB_Name = "Module1"
Sub name_compiler()
Dim a As Integer
Dim fn As String, sn As String, fullname As String

For a = 2 To 600
    fn = Cells(a, 1).Value
    sn = Cells(a, 2).Value
    fullname = fn & " " & sn
    Cells(a, 3) = fullname
Next a

End Sub
