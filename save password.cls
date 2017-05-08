VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim pwd As String
pwd = InputBox("Enter Password", "Security Violation!")
        If pwd = "Pa55word" Or pwd = "5au5age5" Then
        Cancel = False
        Else
        MsgBox "You Do Not Have The Necessary Permission To Save This File", vbOKOnly, "Security Violation!"
        Cancel = True
        End If
End Sub

