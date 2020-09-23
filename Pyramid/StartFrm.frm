VERSION 5.00
Begin VB.Form StartFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to W. A. Mozart (+) for this nice tune..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   330
      Left            =   270
      TabIndex        =   8
      Top             =   3870
      Width           =   7035
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606000&
      Height          =   555
      Left            =   3465
      TabIndex        =   6
      Top             =   3195
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto next level"
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
      Height          =   360
      Index           =   3
      Left            =   4500
      TabIndex        =   5
      Top             =   2025
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   3735
      Picture         =   "StartFrm.frx":0000
      Top             =   1980
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset level"
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
      Height          =   360
      Index           =   2
      Left            =   4500
      TabIndex        =   4
      Top             =   2610
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3735
      Picture         =   "StartFrm.frx":030A
      Top             =   2565
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change color"
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
      Height          =   360
      Index           =   1
      Left            =   1305
      TabIndex        =   3
      Top             =   2610
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   540
      Picture         =   "StartFrm.frx":0BD4
      Top             =   2565
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Balls"
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
      Height          =   360
      Index           =   0
      Left            =   1305
      TabIndex        =   2
      Top             =   2025
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   540
      Picture         =   "StartFrm.frx":149E
      Top             =   1980
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Stephan Swertvaegher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Left            =   1260
      TabIndex        =   1
      Top             =   1215
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pyramid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1230
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   45
      Width           =   6540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pyramid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1230
      Index           =   1
      Left            =   540
      TabIndex        =   7
      Top             =   90
      Width           =   6540
   End
End
Attribute VB_Name = "StartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For xx = 0 To StartFrm.ScaleHeight Step 7
StartFrm.Line (0, xx)-(StartFrm.ScaleWidth, xx), RGB(0, 64, 128)
Next xx
ColBox StartFrm, 0, 0, StartFrm.ScaleWidth, StartFrm.ScaleHeight, 7, 64, 128, 0, 128, 255, 0
MidFile = App.Path & "\sounds\B.mid"
PlayMIDI (MidFile)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H606000
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H606000
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H606000
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H606000
End Sub

Private Sub Label4_Click()
StartFrm.Hide
StopMIDI (MidFile)
NextLevel.Show 1
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HC0C000
End Sub
