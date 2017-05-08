Attribute VB_Name = "TNAModule"
Sub dataform_show()
dataform.Show

End Sub

Sub statsform_show()
statsform.Show
End Sub

Sub toolbar()
Dim mytoolbar As CommandBar
Set mytoolbar = CommandBars.Add("Data Management")
With mytoolbar
    .Position = msoBarFloating
    .Visible = True
End With

Dim src As CommandBarControl, stats As CommandBarControl

Set src = mytoolbar.Controls.Add(Type:=msoControlButton)
With src
    .FaceId = 1849
    .OnAction = "dataform_show"
    .TooltipText = "Search"
End With

Set stats = mytoolbar.Controls.Add(Type:=msoControlButton)
With stats
    .FaceId = 3736
    .OnAction = "statsform_show"
    .TooltipText = "Statistics"
End With
End Sub
Sub toolbarclose()
Application.CommandBars("Data Management").Delete
End Sub
