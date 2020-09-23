Attribute VB_Name = "Module1"
Public I As Integer
Public E1 As Integer
Public P As Integer
Public c As Integer
Public AA As Integer
Public DWn As Integer
Public DWn1 As Integer
Public ChoiceReturn As Integer
Public Gamearea As Integer
Public Gameexit As Boolean
Public speed As Long
Public e102e As Integer
Public e100e As Integer
Public e104e As Integer
Public e98e As Integer
Public MenuSoundsEnabled As Boolean
Public Sndd1 As String, Sndd2 As String
Public WalkOnWater As Boolean
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Public Function Arrange_dialog(l1 As Long, t1 As Long, long1 As Long, hight1 As Long, Dialog As String, ShowChoice As Boolean, Choice1text As String, Choice2text As String)
On Error Resume Next
Form1.DialogBox(0).Visible = False
Form1.DialogBox(1).Visible = False
Form1.DialogBox(2).Visible = False
Form1.DialogBox(3).Visible = False
Form1.DialogBox(4).Visible = False
Form1.DialogBox(5).Visible = False
Form1.DialogBox(6).Visible = False
Form1.DialogBox(7).Visible = False
Form1.DialogBox(8).Visible = False
Form1.dialog1.Visible = False

Form1.DialogBox(0).Move l1, t1
Form1.DialogBox(1).Move l1 + long1, t1
Form1.DialogBox(2).Move l1 + long1, t1 + hight1 + Form1.DialogBox(2).Height
Form1.DialogBox(3).Move l1, t1 + hight1 + Form1.DialogBox(3).Height
Form1.DialogBox(4).Move Form1.DialogBox(0).Left + 60, Form1.DialogBox(0).Top + 90, long1 + Form1.DialogBox(0).Width - 120 + 15, hight1 + Form1.DialogBox(0).Height + 160 + 15
Form1.DialogBox(5).Move l1, t1 + Form1.DialogBox(0).Height, Form1.DialogBox(5).Width, hight1 + Form1.DialogBox(0).Height - 120
Form1.DialogBox(6).Move l1 + long1 + Form1.DialogBox(0).Width - Form1.DialogBox(6).Width, Form1.DialogBox(0).Top + 140, Form1.DialogBox(6).Width, hight1 + Form1.DialogBox(0).Height
Form1.DialogBox(7).Move l1 + Form1.DialogBox(0).Width, Form1.DialogBox(0).Top + Form1.DialogBox(0).Height + Form1.DialogBox(5).Height + 30, long1 - 120
Form1.DialogBox(8).Move l1 + Form1.DialogBox(0).Width, t1 + 30, Form1.DialogBox(7).Width
Form1.dialog1.Move Form1.DialogBox(4).Left + 120, Form1.DialogBox(4).Top + 120, Form1.DialogBox(4).Width - 240, Form1.DialogBox(4).Height - 240
If ShowChoice = False Then
Form1.DialogChoice(0).Visible = False
Form1.DialogChoice(1).Visible = False

Else
Form1.DialogChoice(1).Caption = Choice2text
Form1.DialogChoice(0).Caption = Choice1text
Form1.DialogChoice(0).Visible = True
Form1.DialogChoice(1).Visible = True
Form1.DialogChoice(0).Move Form1.dialog1.Left, Form1.dialog1.Top + Form1.dialog1.Height - Form1.DialogChoice(0).Height
Form1.DialogChoice(1).Move Form1.DialogChoice(0).Left + Form1.DialogChoice(0).Width, Form1.DialogChoice(0).Top
End If
Form1.dialog1.Caption = Dialog

Form1.DialogBox(0).Visible = True
Form1.DialogBox(1).Visible = True
Form1.DialogBox(2).Visible = True
Form1.DialogBox(3).Visible = True
Form1.DialogBox(4).Visible = True
Form1.DialogBox(5).Visible = True
Form1.DialogBox(6).Visible = True
Form1.DialogBox(7).Visible = True


Form1.DialogBox(8).Visible = True
Form1.dialog1.Visible = True

End Function


Public Function PlayMidi(Soundfile As String)

wFlags% = SND_ASYNC Or SND_NODEFAULT
X% = sndPlaySound(Soundfile, wFlags%)

End Function
