Attribute VB_Name = "Module1"
Sub servicelength()
Dim calcdate As Date, startdate As Date, a As Integer
Dim servy As Integer


calcdate = "01 April 2013"

For a = 2 To 50
    startdate = Cells(a, 5).Value
    servy = DateDiff("yyyy", startdate, calcdate)
    Cells(a, 8).Value = servy
Next a

    
End Sub
Sub holiday_entitlement()

Dim a As Integer, daysent As Integer, hworked As Variant, holent As Integer, servicelength As Integer, ans As Variant

Dim myrange As Range, daysworked As Integer
Dim arg2 As Integer, arg3 As Boolean

Set myrange = Sheet1.Range("B7:I81")


For a = 2 To 50
    hworked = Sheet2.Cells(a, 9).Value
    daysworked = Sheet2.Cells(a, 12).Value
    servicelength = Sheet2.Cells(a, 8).Value
    holent = 7
    If servicelength < 5 Then holent = 6
    If servicelength > 9 Then holent = 8
    
    arg3 = False

    ans = Application.VLookup(hworked, myrange, holent, arg3)
    
    daysent = daysworked * ans
    Sheet2.Cells(a, 10) = daysent

Next a

End Sub
