VERSION 5.00
Begin VB.Form NameFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   540
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2655
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You reached Level 050"
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
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Top             =   1350
      Width           =   4110
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Give up your name, stranger..."
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
      Left            =   270
      TabIndex        =   1
      Top             =   2115
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pyramid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   90
      Width           =   3885
   End
End
Attribute VB_Name = "NameFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Label3.Caption = "You reached level " & Format(Str(Val(L) - 1), "000")
End Sub

Private Sub Form_Load()
ColBox NameFrm, 0, 0, NameFrm.ScaleWidth, NameFrm.ScaleHeight, 11, 128, 64, 0, 255, 128, 0
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text = "" Then Exit Sub
PlName = Text1.Text
PlName = LCase(PlName)
PlName = UCase(Left(PlName, 1)) & Right(PlName, Len(PlName) - 1)
HScore(10, 0) = Left(PlName, 10)
HScore(10, 1) = Right("00000000" + Trim(Str(Score)), 8)
NameFrm.Hide
HSfrm.Show 1
End If
End Sub
