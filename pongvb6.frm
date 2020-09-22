VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Pong - - - by Richard Banks"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   FillColor       =   &H0000C000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8400
      TabIndex        =   7
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTER WINS THE GAME :("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER 1 WINS THE GAME :)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER 1 WINS!WELL DONE! CLICK TO PLAY AGAIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTER WINS! HARD LUCK.CLICK TO CARRY ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3120
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1 Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Shape shpPlayer1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape shpPlayer2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape shpBall 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape shpWallBottom 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   5760
      Width           =   10215
   End
   Begin VB.Shape shpWallTop 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'============//GAME VARIABLES//==================
Dim yspeed As Integer
Dim xspeed As Integer
Dim paddleSpeed As Integer
Dim origPaddleLoc As Integer
Dim origBallLocY As Integer
Dim origBallLocX As Integer
Dim playerscore As Integer
Dim computerscore As Integer
 

'===========================================

'============//PADDLE MOVEMENT//=================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 38 Then
    paddleSpeed = -300
ElseIf KeyCode = 40 Then
    paddleSpeed = 300
End If
 
'============================================

'=============//RESET GAME//=====================
If KeyCode = 13 Then
If shpBall.Left + shpBall.Width < 0 Then
Timer1.Enabled = True
shpBall.Left = origBallLocX
    shpBall.Top = origBallLocY
    shpPlayer1.Top = origPaddleLoc
    shpPlayer2.Top = origPaddleLoc
    xspeed = xspeed - 50
    yspeed = 0
    computerscore = computerscore + 100
    Label2.Caption = "Computer Score" & ":"
    Label8.Caption = computerscore
    Label8.ForeColor = vbYellow
    Label3.Visible = False
End If
If shpBall.Left > Form1.Width Then
shpBall.Left = origBallLocX
    shpBall.Top = origBallLocY
    shpPlayer1.Top = origPaddleLoc
    shpPlayer2.Top = origPaddleLoc
    xspeed = xspeed + 50
    yspeed = 0
    Label4.Visible = False
    Timer1.Enabled = True
    playerscore = playerscore + 100
    Label1.Caption = "Playerscore" & ":"
    Label7.Caption = playerscore
    Label7.ForeColor = vbYellow
    End If
    End If
End Sub
'============================================

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
paddleSpeed = 0    'stop the paddle from moving
End Sub

Private Sub Form_Load()
xspeed = -250
yspeed = 0
Label3.Visible = False
Label4.Visible = False
App.TaskVisible = False


'==========//ORIGINAL START LOCATIONS//===========
origPaddleLoc = shpPlayer1.Top
origBallLocX = shpBall.Left
origBallLocY = shpBall.Top
 
'==========================================
End Sub

 

Private Sub Timer1_Timer()
'=============//MOVE BALL//====================
shpBall.Top = shpBall.Top + yspeed
shpBall.Left = shpBall.Left + xspeed
'==========================================

'==============//WALL COLLISION//===============
If shpBall.Top + shpBall.Height >= shpWallBottom.Top Then
    shpBall.Top = shpWallBottom.Top - shpBall.Height
    yspeed = -yspeed
    Beep
ElseIf shpBall.Top <= shpWallTop.Top + shpWallTop.Height Then
    shpBall.Top = shpWallTop.Top + shpWallTop.Height
    yspeed = -yspeed
    Beep
End If
'==========================================

 
If paddleSpeed <> 0 Then
    shpPlayer1.Top = shpPlayer1.Top + paddleSpeed
End If

'============//PADDLE HIT WALL//=================
If shpPlayer1.Top <= shpWallTop.Top + shpWallTop.Height Then
    shpPlayer1.Top = shpWallTop.Top + shpWallTop.Height
ElseIf shpPlayer1.Top + shpPlayer1.Height >= shpWallBottom.Top Then
    shpPlayer1.Top = shpWallBottom.Top - shpPlayer1.Height
End If

If shpPlayer2.Top <= shpWallTop.Top + shpWallTop.Height Then
    shpPlayer2.Top = shpWallTop.Top + shpWallTop.Height
ElseIf shpPlayer2.Top + shpPlayer2.Height >= shpWallBottom.Top Then
    shpPlayer2.Top = shpWallBottom.Top - shpPlayer2.Height
End If
'================================================

'===========//MOVE COMPUTER PADDLE//===================
If shpBall.Top < shpPlayer2.Top Then
    shpPlayer2.Top = shpPlayer2.Top - 300
ElseIf shpBall.Top > shpPlayer2.Top + shpPlayer2.Height Then
    shpPlayer2.Top = shpPlayer2.Top + 300
End If
'=================================================


'===============//PLAYER COLLISION//====================
If shpBall.Left <= shpPlayer1.Left + shpPlayer1.Width And shpBall.Left >= shpPlayer1.Left - shpPlayer1.Width Then
    If shpBall.Top + shpBall.Height >= shpPlayer1.Top And shpBall.Top <= shpPlayer1.Top + shpPlayer1.Height Then
        'calculate the angle it's deflecting off at
        tmp = ((shpPlayer1.Top + (shpPlayer1.Height / 2)) - (shpBall.Top + (shpBall.Height / 2))) * 0.55
        yspeed = yspeed + -tmp
        Beep
        shpBall.Left = shpPlayer1.Left + shpPlayer1.Width
        'deflect the ball
        xspeed = -xspeed
        Label7.ForeColor = vbWhite
    End If
    
    If xspeed >= 800 Then
    Timer1.Enabled = False
    If Val(Label7.Caption) > Val(Label8.Caption) Then
    Label5.Visible = True
    ElseIf Val(Label8.Caption) > Val(Label7.Caption) Then
    Label6.Visible = True
    End If
    
    End If
    
End If
'=================================================

'=============//COMPUTER COLLISION//====================
If shpBall.Left + shpBall.Width >= shpPlayer2.Left And shpBall.Left <= shpPlayer2.Left + shpPlayer2.Width Then
    If shpBall.Top + shpBall.Height >= shpPlayer2.Top And shpBall.Top <= shpPlayer2.Top + shpPlayer2.Height Then
    tmp = ((shpPlayer2.Top + (shpPlayer2.Height / 2)) - (shpBall.Top + (shpBall.Height / 2))) * 0.55
        ypseed = yspeed + -tmp
        Beep
        shpBall.Left = shpPlayer2.Left - shpBall.Width
         xspeed = -xspeed
          Label8.ForeColor = vbWhite
    End If
End If

'=================================================

 '=============//WINNER//============================
If shpBall.Left + shpBall.Width < 0 Then
    Label3.Visible = True
    Timer1.Enabled = False
     
    

ElseIf shpBall.Left > Form1.Width Then
Label4.Visible = True
Timer1.Enabled = False
 End If
End Sub
'=================================================

