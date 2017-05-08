Attribute VB_Name = "Module1"

Public Sub All()                                'shortcut key shift+ctrl+A
Attribute All.VB_Description = "Make All Sheets and Update buttons visible"
Attribute All.VB_ProcData.VB_Invoke_Func = "A\n14"
Application.ScreenUpdating = False
Sheet1.Visible = xlSheetVisible
Sheet2.Visible = xlSheetVisible
Sheet3.Visible = xlSheetVisible
Sheet4.Visible = xlSheetVisible
Sheet5.Visible = xlSheetVisible
Sheet6.Visible = xlSheetVisible
Sheet7.Visible = xlSheetVisible
Sheet8.Visible = xlSheetVisible
Sheet9.Visible = xlSheetVisible
Sheet10.Visible = xlSheetVisible
Sheet11.Visible = xlSheetVisible
Sheet12.Visible = xlSheetVisible
Sheet1.janupdate.Visible = True
Sheet2.febupdate.Visible = True
Sheet3.marupdate.Visible = True
Sheet4.aprupdate.Visible = True
Sheet5.mayupdate.Visible = True
Sheet6.junupdate.Visible = True
Sheet7.julupdate.Visible = True
Sheet8.augupdate.Visible = True
Sheet9.septupdate.Visible = True
Sheet10.octupdate.Visible = True
Sheet11.novupdate.Visible = True
Sheet12.decupdate.Visible = True
Sheet1.jancleanse.Visible = True
Sheet2.febcleanse.Visible = True
Sheet3.marcleanse.Visible = True
Sheet4.aprcleanse.Visible = True
Sheet5.maycleanse.Visible = True
Sheet6.juncleanse.Visible = True
Sheet7.julcleanse.Visible = True
Sheet8.augcleanse.Visible = True
Sheet9.septcleanse.Visible = True
Sheet10.octcleanse.Visible = True
Sheet11.novcleanse.Visible = True
Sheet12.decleanse.Visible = True
Application.ScreenUpdating = True

End Sub



Public Sub HighlightChange()                                'shortcut key shift+ctrl+H
Attribute HighlightChange.VB_Description = "Highlight changes"
Attribute HighlightChange.VB_ProcData.VB_Invoke_Func = "H\n14"
Dim a As Integer, b As Integer, c As Integer, d As Integer

Application.ScreenUpdating = False

For a = 4 To 43
If ActiveSheet.Cells(a, 5) <> "" And ActiveSheet.Cells(a, 5).Value < ActiveSheet.Cells(a, 4).Value Then
ActiveSheet.Cells(a, 3).Font.Color = vbRed
End If
If ActiveSheet.Cells(a, 5) <> "" And ActiveSheet.Cells(a, 5).Value > ActiveSheet.Cells(a, 4).Value Then
ActiveSheet.Cells(a, 3).Font.Color = vbGreen
End If
If ActiveSheet.Cells(a, 5).Value = ActiveSheet.Cells(a, 4).Value Then
ActiveSheet.Cells(a, 3).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(a, 8) <> "" And ActiveSheet.Cells(a, 8).Value < ActiveSheet.Cells(a, 7).Value Then
    ActiveSheet.Cells(a, 6).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(a, 8) <> "" And ActiveSheet.Cells(a, 8).Value > ActiveSheet.Cells(a, 7).Value Then
    ActiveSheet.Cells(a, 6).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(a, 8).Value = ActiveSheet.Cells(a, 7).Value Then
    ActiveSheet.Cells(a, 6).Font.ColorIndex = xlColorIndexAutomatic
    End If
If ActiveSheet.Cells(a, 11) <> "" And ActiveSheet.Cells(a, 11).Value < ActiveSheet.Cells(a, 10).Value Then
ActiveSheet.Cells(a, 9).Font.Color = vbRed
End If
If ActiveSheet.Cells(a, 11) <> "" And ActiveSheet.Cells(a, 11).Value > ActiveSheet.Cells(a, 10).Value Then
ActiveSheet.Cells(a, 9).Font.Color = vbGreen
End If
If ActiveSheet.Cells(a, 11).Value = ActiveSheet.Cells(a, 10).Value Then
ActiveSheet.Cells(a, 9).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(a, 14) <> "" And ActiveSheet.Cells(a, 14).Value < ActiveSheet.Cells(a, 13).Value Then
    ActiveSheet.Cells(a, 12).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(a, 14) <> "" And ActiveSheet.Cells(a, 14).Value > ActiveSheet.Cells(a, 13).Value Then
    ActiveSheet.Cells(a, 12).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(a, 14).Value = ActiveSheet.Cells(a, 13).Value Then
    ActiveSheet.Cells(a, 12).Font.ColorIndex = xlColorIndexAutomatic
    End If
Next a

Application.ScreenUpdating = False

For b = 46 To 81
If ActiveSheet.Cells(b, 5) <> "" And ActiveSheet.Cells(b, 5).Value < ActiveSheet.Cells(b, 4).Value Then
ActiveSheet.Cells(b, 3).Font.Color = vbRed
End If
If ActiveSheet.Cells(b, 5) <> "" And ActiveSheet.Cells(b, 5).Value > ActiveSheet.Cells(b, 4).Value Then
ActiveSheet.Cells(b, 3).Font.Color = vbGreen
End If
If ActiveSheet.Cells(b, 5).Value = ActiveSheet.Cells(b, 4).Value Then
ActiveSheet.Cells(b, 3).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(b, 8) <> "" And ActiveSheet.Cells(b, 8).Value < ActiveSheet.Cells(a, 7).Value Then
    ActiveSheet.Cells(b, 6).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(b, 8) <> "" And ActiveSheet.Cells(b, 8).Value > ActiveSheet.Cells(b, 7).Value Then
    ActiveSheet.Cells(b, 6).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(b, 8).Value = ActiveSheet.Cells(b, 7).Value Then
    ActiveSheet.Cells(b, 6).Font.ColorIndex = xlColorIndexAutomatic
    End If
If ActiveSheet.Cells(b, 11) <> "" And ActiveSheet.Cells(b, 11).Value < ActiveSheet.Cells(b, 10).Value Then
ActiveSheet.Cells(b, 9).Font.Color = vbRed
End If
If ActiveSheet.Cells(b, 11) <> "" And ActiveSheet.Cells(b, 11).Value > ActiveSheet.Cells(b, 10).Value Then
ActiveSheet.Cells(b, 9).Font.Color = vbGreen
End If
If ActiveSheet.Cells(b, 11).Value = ActiveSheet.Cells(b, 10).Value Then
ActiveSheet.Cells(b, 9).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(b, 14) <> "" And ActiveSheet.Cells(b, 14).Value < ActiveSheet.Cells(b, 13).Value Then
    ActiveSheet.Cells(b, 12).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(b, 14) <> "" And ActiveSheet.Cells(b, 14).Value > ActiveSheet.Cells(b, 13).Value Then
    ActiveSheet.Cells(b, 12).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(b, 14).Value = ActiveSheet.Cells(b, 13).Value Then
    ActiveSheet.Cells(b, 12).Font.ColorIndex = xlColorIndexAutomatic
    End If
Next b

Application.ScreenUpdating = False

For c = 84 To 123
If ActiveSheet.Cells(c, 5) <> "" And ActiveSheet.Cells(c, 5).Value < ActiveSheet.Cells(c, 4).Value Then
ActiveSheet.Cells(c, 3).Font.Color = vbRed
End If
If ActiveSheet.Cells(c, 5) <> "" And ActiveSheet.Cells(c, 5).Value > ActiveSheet.Cells(c, 4).Value Then
ActiveSheet.Cells(c, 3).Font.Color = vbGreen
End If
If ActiveSheet.Cells(c, 5).Value = ActiveSheet.Cells(c, 4).Value Then
ActiveSheet.Cells(c, 3).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(c, 8) <> "" And ActiveSheet.Cells(c, 8).Value < ActiveSheet.Cells(c, 7).Value Then
    ActiveSheet.Cells(c, 6).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(c, 8) <> "" And ActiveSheet.Cells(c, 8).Value > ActiveSheet.Cells(c, 7).Value Then
    ActiveSheet.Cells(c, 6).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(c, 8).Value = ActiveSheet.Cells(c, 7).Value Then
    ActiveSheet.Cells(c, 6).Font.ColorIndex = xlColorIndexAutomatic
    End If
If ActiveSheet.Cells(c, 11) <> "" And ActiveSheet.Cells(c, 11).Value < ActiveSheet.Cells(c, 10).Value Then
ActiveSheet.Cells(c, 9).Font.Color = vbRed
End If
If ActiveSheet.Cells(c, 11) <> "" And ActiveSheet.Cells(c, 11).Value > ActiveSheet.Cells(c, 10).Value Then
ActiveSheet.Cells(c, 9).Font.Color = vbGreen
End If
If ActiveSheet.Cells(c, 11).Value = ActiveSheet.Cells(c, 10).Value Then
ActiveSheet.Cells(c, 9).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(c, 14) <> "" And ActiveSheet.Cells(c, 14).Value < ActiveSheet.Cells(c, 13).Value Then
    ActiveSheet.Cells(c, 12).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(c, 14) <> "" And ActiveSheet.Cells(c, 14).Value > ActiveSheet.Cells(c, 13).Value Then
    ActiveSheet.Cells(c, 12).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(c, 14).Value = ActiveSheet.Cells(c, 13).Value Then
    ActiveSheet.Cells(c, 12).Font.ColorIndex = xlColorIndexAutomatic
    End If
Next c

Application.ScreenUpdating = False

For d = 126 To 163
If ActiveSheet.Cells(d, 5) <> "" And ActiveSheet.Cells(d, 5).Value < ActiveSheet.Cells(d, 4).Value Then
ActiveSheet.Cells(d, 3).Font.Color = vbRed
End If
If ActiveSheet.Cells(d, 5) <> "" And ActiveSheet.Cells(d, 5).Value > ActiveSheet.Cells(d, 4).Value Then
ActiveSheet.Cells(d, 3).Font.Color = vbGreen
End If
If ActiveSheet.Cells(d, 5).Value = ActiveSheet.Cells(d, 4).Value Then
ActiveSheet.Cells(d, 3).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(d, 8) <> "" And ActiveSheet.Cells(d, 8).Value < ActiveSheet.Cells(d, 7).Value Then
    ActiveSheet.Cells(d, 6).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(d, 8) <> "" And ActiveSheet.Cells(d, 8).Value > ActiveSheet.Cells(d, 7).Value Then
    ActiveSheet.Cells(d, 6).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(d, 8).Value = ActiveSheet.Cells(d, 7).Value Then
    ActiveSheet.Cells(d, 6).Font.ColorIndex = xlColorIndexAutomatic
    End If
If ActiveSheet.Cells(d, 11) <> "" And ActiveSheet.Cells(d, 11).Value < ActiveSheet.Cells(d, 10).Value Then
ActiveSheet.Cells(d, 9).Font.Color = vbRed
End If
If ActiveSheet.Cells(d, 11) <> "" And ActiveSheet.Cells(d, 11).Value > ActiveSheet.Cells(d, 10).Value Then
ActiveSheet.Cells(d, 9).Font.Color = vbGreen
End If
If ActiveSheet.Cells(d, 11).Value = ActiveSheet.Cells(d, 10).Value Then
ActiveSheet.Cells(d, 9).Font.ColorIndex = xlColorIndexAutomatic
End If
    If ActiveSheet.Cells(d, 14) <> "" And ActiveSheet.Cells(d, 14).Value < ActiveSheet.Cells(d, 13).Value Then
    ActiveSheet.Cells(d, 12).Font.Color = vbRed
    End If
    If ActiveSheet.Cells(d, 14) <> "" And ActiveSheet.Cells(d, 14).Value > ActiveSheet.Cells(d, 13).Value Then
    ActiveSheet.Cells(d, 12).Font.Color = vbGreen
    End If
    If ActiveSheet.Cells(d, 14).Value = ActiveSheet.Cells(d, 13).Value Then
    ActiveSheet.Cells(d, 12).Font.ColorIndex = xlColorIndexAutomatic
    End If
Next d
Application.ScreenUpdating = True
End Sub


Public Sub RemoveHighlights()                               'shortcut key = shift+ctrl=R
Attribute RemoveHighlights.VB_ProcData.VB_Invoke_Func = "R\n14"
ActiveSheet.UsedRange.Font.ColorIndex = xlColorIndexAutomatic
End Sub
