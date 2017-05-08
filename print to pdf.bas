Attribute VB_Name = "Module1"
Sub printpdf()
Attribute printpdf.VB_Description = "Macro recorded 05/11/2009 by Jon Moon"
Attribute printpdf.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ActivePrinter = "PrimoPDF on Ne00:"
    ActiveWindow.SelectedSheets.PrintOut From:=2, to:=2, Copies:=1, _
        ActivePrinter:="PrimoPDF on Ne00:", Collate:=True
End Sub
