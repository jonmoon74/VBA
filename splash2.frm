VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} splash 
   Caption         =   "Splash"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   OleObjectBlob   =   "splash2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub userform_Initialize()
Application.OnTime Now + TimeValue("00:00:02"), "splashclose"

End Sub
