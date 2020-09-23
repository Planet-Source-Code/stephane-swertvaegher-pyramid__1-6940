VERSION 5.00
Begin VB.Form NextLevel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   360
      Top             =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Prepare yourself"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   465
      Left            =   495
      TabIndex        =   1
      Top             =   990
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   495
      TabIndex        =   0
      Top             =   270
      Width           =   3570
   End
End
Attribute VB_Name = "NextLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If L <> 21 Then
Label1.Caption = "Level " & L
Label2.Caption = "Prepare yourself"
Else
Label1.Caption = "Congratulations"
Label2.Caption = "You made it !"
End If
Call PLAY(8, 9)
Timer1.Enabled = True
Pyramid.Label6.Caption = "Your score:" & vbCr & Format(Score, "0000000000")
End Sub

Private Sub Form_Load()
ColBox NextLevel, 0, 0, NextLevel.ScaleWidth, NextLevel.ScaleHeight, 7, 128, 128, 0, 164, 255, 64
End Sub

Private Sub Timer1_Timer()
Pyramid.Enabled = True
Pyramid.Label6.Caption = "Your score:" & vbCr & Format(Score, "0000000000")
RC = sndPlaySound("", 1)
NextLevel.Hide
Timer1.Enabled = False
LoadLevel
If L <> 21 Then
Pyramid.Timer1.Enabled = True
End If
TxCount = TxCount + 1
If TxCount = 10 Then TxCount = 0
End Sub
