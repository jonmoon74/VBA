Attribute VB_Name = "Module2"
Sub sorter()
Attribute sorter.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim x As Integer

ActiveSheet.Range("a3").End(xlDown).Select
x = ActiveCell.Row

ActiveSheet.Range(Cells(3, 1), Cells(x, 92)).Sort _
key1:=Range("B3"), key2:=Range("a3")

End Sub
