VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   6030
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   840
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3960
      Top             =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Richard Banks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 
 

 

Private Sub Form_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Me.ZOrder (1)
App.TaskVisible = False
End Sub

 

 
