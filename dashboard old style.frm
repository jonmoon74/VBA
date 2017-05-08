VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dashboard 
   Caption         =   "Dashboard"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   OleObjectBlob   =   "dashboard old style.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pth As String, nme As String, fnme As String, ofnme As String
Dim dbd As String, pos As Integer, slash As String, pos2 As Integer, pth2 As String

Private Sub accredbutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "Accredited Training.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme
End Sub

Private Sub caresbutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "CARES Status.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme
End Sub

Private Sub cateringbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Catering Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

Private Sub closebutton_Click()
Unload Me

Application.DisplayFullScreen = False
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayOutline = True
        .DisplayZeros = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
ActiveWindow.View = xlNormalView

ActiveWorkbook.Close savechanges:=False
Application.Quit
Application.Quit

End Sub

Private Sub cmbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "cleaning matters.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

Private Sub compliancebutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "Training Status.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub



Private Sub helpdeskbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Helpdesk Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

'Private Sub hotelservicesbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Hotel Services Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme
'
'End Sub

Private Sub ioshbutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "iosh.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub

Private Sub linenbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Linen Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

Private Sub logo_Click()
Unload Me

Application.DisplayFullScreen = False
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayOutline = True
        .DisplayZeros = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
    End With
ActiveWindow.View = xlNormalView
End Sub

Private Sub monitoringbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Monitoring Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

Private Sub nvqbutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "NVQ.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub

Private Sub portersbutton_Click()
'Unload Me
'Application.DisplayFullScreen = False
'pth = ActiveWorkbook.Path
'fnme = "Porters Skills.xls"
'slash = "\Databases\"
'ofnme = pth & slash & fnme
'
'Workbooks.Open Filename:=ofnme

End Sub

Private Sub sdpbutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "Supervisor Development Program Tracking.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub
Private Sub tnabutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "TNA.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub

Private Sub UserForm_Initialize()
With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayOutline = False
        .DisplayZeros = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
    End With
    Application.ShowStartupDialog = False
    Application.DisplayFullScreen = True
    Application.CommandBars("Full Screen").Visible = False
    ActiveWindow.View = xlPageBreakPreview
End Sub
Private Sub machinebutton_Click()
Unload Me
Application.DisplayFullScreen = False
pth = ActiveWorkbook.Path
fnme = "Machine Skills.xls"
slash = "\Databases\"
ofnme = pth & slash & fnme

Workbooks.Open Filename:=ofnme

End Sub
