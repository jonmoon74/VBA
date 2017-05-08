VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} passform 
   Caption         =   "Password Required"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3555
   OleObjectBlob   =   "passform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "passform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbutton_Click()
Unload Me
End Sub

Private Sub okbutton_Click()
Dim pwd As String
pwd = passwordbox.Value
Unload Me
If pwd <> "" Then
    If pwd = "Pa55word" Or pwd = "5au5age5" Then
    mgrtna.Visible = xlSheetVisible
    mgrtna.Activate
    Else
    MsgBox "You Do Not Have Permission To View This Sheet!", vbOKOnly, "Error!"
    End If
Else
Exit Sub
End If
End Sub
