Attribute VB_Name = "Module1"
Sub facilitatorlist()
Attribute facilitatorlist.VB_Description = "collect code for range copy and remove duplicates"
Attribute facilitatorlist.VB_ProcData.VB_Invoke_Func = " \n14"

'   Collects facilitator names from drop down selected list and removes duplicates
Dim a As Integer, b As Integer, c As Integer

    Sheet1.Select
    Sheet1.Range("I2:I6000").Select
    Selection.Copy
    Sheet2.Select
    Sheet2.Range("D10").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Sheet2.Range("D10:D6010").RemoveDuplicates Columns:=1, Header:=xlNo
        
        Application.ScreenUpdating = False
        For a = 10 To 6010
        If Sheet2.Cells(a, 4) = "" Then
        Sheet2.Cells(a, 4).Select
        Selection.Delete shift:=xlUp
        End If
        Next a
        Application.ScreenUpdating = True
            
    Sheet1.Select
    Sheet1.Range("K2:K6000").Select
    Selection.Copy
    Sheet2.Select
    Sheet2.Range("E10").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Sheet2.Range("E10:E6010").RemoveDuplicates Columns:=1, Header:=xlNo
    
        Application.ScreenUpdating = False
        For b = 10 To 6010
        If Sheet2.Cells(b, 5) = "" Then
        Sheet2.Cells(b, 5).Select
        Selection.Delete shift:=xlUp
        End If
        Next b
       
    Application.ScreenUpdating = False
    Sheet2.Range("D10:D30").Select
    Selection.Copy
    Sheet2.Range("F10").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Sheet2.Range("E10:E40").Select
    Selection.Copy
    Sheet2.Range("F30").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    
    Sheet2.Range("F10:F60").Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
        For c = 10 To 60
        If Sheet2.Cells(c, 6) = "" Then
        Sheet2.Cells(c, 6).Select
        Selection.Delete shift:=xlUp
        End If
        Next c
        Application.ScreenUpdating = True

    Sheet2.Range("f10").Select
    
End Sub

