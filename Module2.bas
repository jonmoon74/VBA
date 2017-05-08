Attribute VB_Name = "Module2"
Sub toolbar()
Dim mytoolbar As CommandBar                               'create custom toolbar
Set mytoolbar = CommandBars.Add("Compliance")
With mytoolbar
    .Position = msoBarFloating
    .Visible = True
End With
                                                         'add buttons & Macros
Dim hs As CommandBarControl, fire As CommandBarControl
Dim cust As CommandBarControl, infect As CommandBarControl
Dim food As CommandBarControl, hsp As CommandBarControl, bfsp As CommandBarControl
Dim te As CommandBarControl, snf As CommandBarControl



Set te = mytoolbar.Controls.Add(Type:=msoControlButton)
With te
    .FaceId = 1849
    .OnAction = "showform"
    .TooltipText = "Search"
End With

Set snf = mytoolbar.Controls.Add(Type:=msoControlButton)
With snf
    .FaceId = 505
    .OnAction = "shownewform"
    .TooltipText = "New Entry"
End With

Set hs = mytoolbar.Controls.Add(Type:=msoControlButton)
With hs
    .FaceId = 463
    .OnAction = "HealthSafety"
    .TooltipText = "Health & Safety"
End With

Set fire = mytoolbar.Controls.Add(Type:=msoControlButton)
With fire
    .FaceId = 2151
    .OnAction = "fire"
    .TooltipText = "Fire"
End With

Set cust = mytoolbar.Controls.Add(Type:=msoControlButton)
With cust
    .FaceId = 2148
    .OnAction = "CustomerService"
    .TooltipText = "Customer Service"
End With

Set food = mytoolbar.Controls.Add(Type:=msoControlButton)
With food
    .FaceId = 3205
    .OnAction = "FoodHygiene"
    .TooltipText = "Food Hygiene"
End With

Set infect = mytoolbar.Controls.Add(Type:=msoControlButton)
With infect
    .FaceId = 3202
    .OnAction = "InfectionControl"
    .TooltipText = "Infection Control"
End With

Set hsp = mytoolbar.Controls.Add(Type:=msoControlButton)
With hsp
    .FaceId = 565
    .OnAction = "HSP"
    .TooltipText = "Health & Safety Passport"
End With

Set bfsp = mytoolbar.Controls.Add(Type:=msoControlButton)
With bfsp
    .FaceId = 1672
    .OnAction = "BFSP"
    .TooltipText = "Food Passport"
End With

End Sub

Sub toolbarclose()

Application.CommandBars("Compliance").Delete

End Sub

