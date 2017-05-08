Attribute VB_Name = "Module12"
Sub name_compiler()
Dim a As Integer
Dim fn As String, sn As String, fullname As String

For a = 2 To 70
On Error Resume Next
    fn = Cells(a, 4).Value
    sn = Cells(a, 5).Value
    fullname = fn & "." & sn & "@sodexojusticeservices.com"
    Cells(a, 6) = fullname
Next a

End Sub
