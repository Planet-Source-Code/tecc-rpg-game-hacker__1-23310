Attribute VB_Name = "CollisionDetectionEngine"

Public Function TestForCollision(Direction1 As String)

' the game area will be the section of the game
'you are in later in game development
'for now, gamearea is always 0
Gamearea = 0
Select Case Gamearea

Case 0

Select Case LCase(Direction1)

Case "up"

'collision Detection for running up (label3-Barrels)
If Form1.MainFigure.Top <= Form1.label3.Top + Form1.label3.Height And Form1.MainFigure.Top >= Form1.label3.Top Then
If Form1.MainFigure.Left + (Form1.MainFigure.Width / 2) >= Form1.label3.Left And Form1.MainFigure.Left <= Form1.label3.Left + Form1.label3.Width - (Form1.MainFigure.Width / 2) Then
'put the character beside the collided object
Form1.MainFigure.Top = Form1.label3.Top + Form1.label3.Height
'stop the character from moving!
Form1.RunUp.Enabled = False
'place the correct stand-still picture in the
'mainfigure picture
Form1.MainFigure.Picture = Form1.Stillback.Picture
'call the dialog function and tell it to
'make a dialog box (same in all directions)(optional)
'Arrange_dialog (Rnd * 1000), (Rnd * 1000), 1000, 300, "Some Barrels"
PlayMidi (App.Path & "\data\sfx\" & "hitmetal1.wav")
Else
End If
Else
End If

If Form1.MainFigure.Top <= Form1.Puddle.Top + Form1.Puddle.Height And Form1.MainFigure.Top >= Form1.Puddle.Top Then
If Form1.MainFigure.Left + (Form1.MainFigure.Width / 2) >= Form1.Puddle.Left And Form1.MainFigure.Left <= Form1.Puddle.Left + Form1.Puddle.Width - (Form1.MainFigure.Width / 2) Then
WalkOnWater = True
Else
WalkOnWater = False
End If
Else
WalkOnWater = False
End If

'collision Detection for running up (label2-Top warehouse)
If Form1.MainFigure.Top <= Form1.Label2.Top + Form1.Label2.Height Then
Form1.MainFigure.Picture = Form1.Stillback.Picture
Form1.MainFigure.Top = Form1.Label2.Top + Form1.Label2.Height
Form1.RunUp.Enabled = False
P = 1
Else
End If


Case "down"

'collision Detection for running down
If Form1.MainFigure.Top + Form1.MainFigure.Height >= Form1.label3.Top And Form1.MainFigure.Top <= Form1.label3.Top + Form1.label3.Height Then
If Form1.MainFigure.Left + (Form1.MainFigure.Width / 2) >= Form1.label3.Left And Form1.MainFigure.Left <= Form1.label3.Left + Form1.label3.Width - (Form1.MainFigure.Width / 2) Then
Form1.MainFigure.Top = Form1.label3.Top - Form1.MainFigure.Height
Form1.Rundown.Enabled = False
Form1.MainFigure.Picture = Form1.Stillforward.Picture
PlayMidi (App.Path & "\data\sfx\" & "hitmetal1.wav")
Else
End If
Else
End If

If Form1.MainFigure.Top + Form1.MainFigure.Height >= Form1.Puddle.Top And Form1.MainFigure.Top <= Form1.Puddle.Top + Form1.Puddle.Height Then
If Form1.MainFigure.Left + (Form1.MainFigure.Width / 2) >= Form1.Puddle.Left And Form1.MainFigure.Left <= Form1.Puddle.Left + Form1.Puddle.Width - (Form1.MainFigure.Width / 2) Then
WalkOnWater = True
Else
WalkOnWater = False
End If
Else
WalkOnWater = False
End If

If Form1.MainFigure.Top >= Form1.Label1.Top - Form1.MainFigure.Height Then
Form1.MainFigure.Picture = Form1.Stillforward.Picture
Form1.MainFigure.Top = Form1.Label1.Top - Form1.MainFigure.Height
Form1.Rundown.Enabled = False
Arrange_dialog 500, 250, 1000, 400, "Cannot exit out of this area (yet)", False, "", ""
AA = 1
Else
End If

Case "left"

'collision Detection for running left
If Form1.MainFigure.Left + Form1.MainFigure.Width >= Form1.label3.Left And Form1.MainFigure.Left <= Form1.label3.Left + Form1.label3.Width Then
If Form1.MainFigure.Top + (Form1.MainFigure.Height / 2) >= Form1.label3.Top And Form1.MainFigure.Top <= Form1.label3.Top + Form1.label3.Height - (Form1.MainFigure.Height / 2) Then
Form1.MainFigure.Left = Form1.label3.Left + Form1.label3.Width
Form1.Runleft.Enabled = False
Form1.MainFigure.Picture = Form1.Stillleft.Picture
PlayMidi (App.Path & "\data\sfx\" & "hitmetal1.wav")
Else
End If
Else
End If

If Form1.MainFigure.Left + Form1.MainFigure.Width >= Form1.Puddle.Left And Form1.MainFigure.Left <= Form1.Puddle.Left + Form1.Puddle.Width Then
If Form1.MainFigure.Top + (Form1.MainFigure.Height / 2) >= Form1.Puddle.Top And Form1.MainFigure.Top <= Form1.Puddle.Top + Form1.Puddle.Height - (Form1.MainFigure.Height / 2) Then
WalkOnWater = True
Else
WalkOnWater = False
End If
Else
WalkOnWater = False
End If

If Form1.MainFigure.Left <= 0 Then
Form1.MainFigure.Left = 0
Form1.Runleft.Enabled = False
Form1.MainFigure.Picture = Form1.Stillleft.Picture
Else
End If

Case "right"

'collision Detection for running right
If Form1.MainFigure.Left + Form1.MainFigure.Width >= Form1.label3.Left And Form1.MainFigure.Left <= Form1.label3.Left + Form1.label3.Width Then
If Form1.MainFigure.Top + (Form1.MainFigure.Height / 2) >= Form1.label3.Top And Form1.MainFigure.Top <= Form1.label3.Top + Form1.label3.Height - (Form1.MainFigure.Height / 2) Then
Form1.MainFigure.Left = Form1.label3.Left - Form1.MainFigure.Width
Form1.Runright.Enabled = False
Form1.MainFigure.Picture = Form1.Stillright.Picture
PlayMidi (App.Path & "\data\sfx\" & "hitmetal1.wav")
Else
End If
Else
End If

If Form1.MainFigure.Left + Form1.MainFigure.Width >= Form1.Puddle.Left And Form1.MainFigure.Left <= Form1.Puddle.Left + Form1.Puddle.Width Then
If Form1.MainFigure.Top + (Form1.MainFigure.Height / 2) >= Form1.Puddle.Top And Form1.MainFigure.Top <= Form1.Puddle.Top + Form1.Puddle.Height - (Form1.MainFigure.Height / 2) Then
WalkOnWater = True
Else
WalkOnWater = False
End If
Else
WalkOnWater = False
End If

If Form1.MainFigure.Left >= (Form1.ScaleWidth - Form1.MainFigure.Width) Then
Form1.MainFigure.Left = Form1.ScaleWidth - Form1.MainFigure.Width
Form1.Runright.Enabled = False
Form1.MainFigure.Picture = Form1.Stillright.Picture
Else
End If

End Select
End Select
End Function
