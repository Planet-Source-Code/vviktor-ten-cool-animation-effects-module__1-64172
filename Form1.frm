VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEffect 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Please click on the PictureBox under this text to start the presentation of effects"
      Top             =   0
      Width           =   8295
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   3240
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   6
      Top             =   0
      Width           =   8295
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Index           =   3
      Left            =   3720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2760
      Index           =   2
      Left            =   1320
      Picture         =   "Form1.frx":42D4
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":7367
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":A1C2
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   267
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   4560
      Index           =   4
      Left            =   3840
      Picture         =   "Form1.frx":DD55
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Height          =   4560
      Index           =   5
      Left            =   3960
      Picture         =   "Form1.frx":10E1A
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurPic As Integer
Dim CurEffect As Integer
Dim Animation As Boolean


Private Sub Form_Load()
    Set modEffects.Buffer = picBuffer
End Sub

Private Sub Form_Resize()
    picMain.Width = ScaleWidth
    picMain.Height = ScaleHeight
    txtEffect.Width = ScaleWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Animation Then Cancel = 1
End Sub

Private Sub picMain_Click()
    If Animation Then Exit Sub
    Animation = True
    Randomize Timer
    Select Case CurEffect
        Case 0
            txtEffect.Text = "BrickLayer"
            modEffects.BrickLayer picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 1
            txtEffect.Text = "BlackBox"
            modEffects.BlackBox picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 2
            txtEffect.Text = "BlackCircle"
            modEffects.BlackCircle picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 3
            txtEffect.Text = "Laser"
            modEffects.Laser picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 4
            txtEffect.Text = "Checker"
            modEffects.Checker picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain, 5
        Case 5
            txtEffect.Text = "Enlarge (center)"
            modEffects.Enlarge picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 6
            txtEffect.Text = "Enlarge (vertical)"
            modEffects.Enlarge picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain, VerticalEnlarge
        Case 7
            txtEffect.Text = "Enlarge (horizontal)"
            modEffects.Enlarge picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain, HorizontalEnlarge
        Case 8
            txtEffect.Text = "Slash"
            modEffects.Slash picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
        Case 9
            txtEffect.Text = "Emerge"
            modEffects.Emerge picMain.ScaleWidth / 2, picMain.ScaleHeight / 2, picSample(CurPic), picMain
            
    End Select
    CurPic = CurPic + 1
    If CurPic = picSample.Count Then CurPic = 0
    CurEffect = CurEffect + 1
    If CurEffect = 10 Then CurEffect = 0
    Animation = False
End Sub

Private Sub txtEffect_GotFocus()
    picMain.SetFocus
End Sub
