Attribute VB_Name = "Module1"
Sub name_compiler()
Dim a As Integer
Dim fn As String, sn As String, fullname As String
Dim site As String

For a = 3 To 31
    site = Cells(a, 3).Value
    If site <> "NU" Then
        fn = Cells(a, 4).Value
        sn = Cells(a, 5).Value
        fullname = fn & "." & sn & "@sodexojusticeservices.com"
        Cells(a, 6) = fullname
    Else
        fn = Cells(a, 4).Value
        sn = Cells(a, 5).Value
        fullname = fn & "." & sn & "@hmps.gsi.gov.uk"
        Cells(a, 6) = fullname
    End If
Next a

End Sub
