VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "STARZ"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar scrl 
      Height          =   4335
      LargeChange     =   5
      Left            =   120
      Max             =   1000
      Min             =   1
      TabIndex        =   2
      Top             =   840
      Value           =   100
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00000000&
      Caption         =   "Color"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00000000&
      Caption         =   "Clear"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer prosses 
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stars As Integer, all() As Single, r As Double, size As Single _
, center_x As Integer, center_y As Integer, i As Integer, c As Integer
Private Sub chk_Click(Index As Integer)
Cls
generate_stars
End Sub
Private Sub Form_Load()
Cls
r = 3.1415926 / 180 '1 digree
direction = 1
If stars < 1 Then stars = 100 '# of stars
ReDim all(stars, 5) 'star,angle,distance,angle delta,dist delta,color
center_x = Form1.ScaleWidth / 2: center_y = Form1.ScaleHeight / 2 'center
If center_x > center_y Then size = center_y Else size = center_x
generate_stars 'self explanitory
End Sub
Private Sub generate_stars()
For i = 0 To stars - 1
all(i, 0) = Rnd * 360 'begining angle
all(i, 1) = Rnd * size 'begining distance from orgin (0,0)
all(i, 2) = (Rnd * 4) - 2 'angle delta
all(i, 3) = (Rnd * 6) - 3 'distance delta
If chk(1).Value = 1 Then all(i, 4) = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255)) Else _
c = Int(Rnd * 255): all(i, 4) = RGB(c, c, c)
Next i
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And chk(0).Visible = False Then mnu_show Else mnu_hide
End Sub
Private Sub mnu_show()
For i = 0 To 1
chk(i).Visible = True
Next i
scrl.Visible = True
End Sub
Private Sub mnu_hide()
For i = 0 To 1
chk(i).Visible = False
Next i
scrl.Visible = False
End Sub
Private Sub Form_Resize()
Cls
Form_Load
End Sub
Private Sub prosses_Timer()
If chk(0).Value = 1 Then Cls
For i = 0 To stars - 1
all(i, 0) = all(i, 0) + all(i, 2): If all(i, 0) > 360 Then all(i, 0) = 0
all(i, 1) = all(i, 1) + (all(i, 3)): If all(i, 1) > size Or all(i, 1) < (-1 * size) _
Then all(i, 3) = all(i, 3) * -1
X = Sin(all(i, 0) * r) * all(i, 1)
Y = Cos(all(i, 0) * r) * all(i, 1)
Circle (X + center_x, Y + center_y), 1, all(i, 4)
Next i
End Sub
Private Sub scrl_Change()
stars = scrl.Value
Form_Load
End Sub
Private Sub scrl_Scroll()
stars = scrl.Value
Form_Load
End Sub
