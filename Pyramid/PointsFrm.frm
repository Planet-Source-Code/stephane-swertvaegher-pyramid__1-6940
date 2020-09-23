VERSION 5.00
Begin VB.Form PointsFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4590
      Top             =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4095
      Top             =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   3465
      TabIndex        =   2
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   405
      TabIndex        =   1
      Top             =   180
      Width           =   4680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00 X 50 points ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   765
      TabIndex        =   0
      Top             =   1440
      Width           =   2760
   End
End
Attribute VB_Name = "PointsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newtim

Private Sub Form_Activate()
Pyramid.Enabled = False
Label2.Caption = "00000000"
If 50 - Tim >= 30 Then Label3.Caption = "Excellent !"
If 50 - Tim < 30 Then Label3.Caption = "Not bad at all..."
If 50 - Tim < 20 Then Label3.Caption = "Mmm... could be better !"
If 50 - Tim < 10 Then Label3.Caption = "Bad ! Absolutely bad !"
newtim = Int((50 - Tim) / 3)
Label3.Caption = Label3.Caption & vbCr & "You scored " & newtim * 50 & " points..."
Label1.Caption = Format(newtim, "00") & " units X 50 points"
Points = 0
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Form_Load()
ColBox PointsFrm, 0, 0, PointsFrm.ScaleWidth, PointsFrm.ScaleHeight, 7, 128, 128, 0, 164, 255, 64
End Sub

Private Sub Timer1_Timer()
Points = Points + 50
Label2.Caption = Format(Points, "00000000")
newtim = newtim - 1
DoEvents
Call PLAY(7, 0)
If newtim < 1 Then
Score = Score + Points
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
PointsFrm.Hide
DoEvents
NextLevel.Show 1
Timer2.Enabled = False
End Sub
