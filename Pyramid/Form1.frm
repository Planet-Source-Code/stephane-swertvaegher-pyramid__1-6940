VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Pyramid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "« Pyramid »"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   Begin VB.PictureBox Texture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   4905
      ScaleHeight     =   630
      ScaleWidth      =   1185
      TabIndex        =   10
      Top             =   225
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click here if you are trapped"
      Top             =   1710
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2565
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   6
      Top             =   3915
      Width           =   3255
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No more time left..."
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
         Left            =   315
         TabIndex        =   7
         Top             =   225
         Width           =   2685
      End
   End
   Begin VB.PictureBox PB1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   270
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   4
      Top             =   1755
      Width           =   3750
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   90
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3240
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2686
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5EB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2610
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":678E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7362
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":96DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AE86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BA5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      Height          =   510
      Left            =   90
      Top             =   7470
      Width           =   3570
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   3015
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   345
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "The working color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   270
      TabIndex        =   12
      Top             =   7515
      Width           =   2715
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "New life !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   6075
      TabIndex        =   11
      Top             =   1665
      Width           =   1770
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   780
      Left            =   6930
      TabIndex        =   9
      Top             =   135
      Width           =   2490
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   8
      Left            =   8910
      Picture         =   "Form1.frx":D18A
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   7
      Left            =   8415
      Picture         =   "Form1.frx":DA54
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   6
      Left            =   7920
      Picture         =   "Form1.frx":E31E
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   5
      Left            =   7425
      Picture         =   "Form1.frx":EBE8
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   4
      Left            =   6930
      Picture         =   "Form1.frx":F4B2
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   3
      Left            =   6435
      Picture         =   "Form1.frx":FD7C
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   2
      Left            =   5940
      Picture         =   "Form1.frx":10646
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   5445
      Picture         =   "Form1.frx":10F10
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lives:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   4005
      TabIndex        =   5
      Top             =   1260
      Width           =   825
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   4950
      Picture         =   "Form1.frx":117DA
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   5
      Left            =   1305
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   4
      Left            =   1665
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   3
      Left            =   2025
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   2
      Left            =   2385
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   1
      Left            =   2700
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
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
      Height          =   330
      Left            =   4005
      TabIndex        =   3
      Top             =   1665
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colors:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   225
      TabIndex        =   2
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   0
      Left            =   3060
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   135
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1035
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   6750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00007000&
      Height          =   435
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   120
   End
End
Attribute VB_Name = "Pyramid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
PB1.SetFocus
DoEvents
If Trapped = 1 Then Exit Sub
Call PLAY(5, 0)
ResetLevel
Trapped = 1
End Sub

Private Sub Form_Activate()
'LoadLevel
'DrawBack
End Sub

Private Sub Form_Load()
LoadHighScore
TxCount = 0
PSetup
Timer1.Enabled = False
L = "001"
Pyramid.Label6.Caption = "Your score:" & vbCr & Format(Score, "0000000000")
Pyramid.Show
DrawBack
Label7.Visible = False
StartFrm.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Temp = MsgBox("Quit Pyramid ?", vbYesNo + vbQuestion, "Pyramid")
If Temp = vbYes Then End
End Sub

Private Sub Image1_Click(Index As Integer)
If Pause = 1 Then Exit Sub
CurIdx = Index
Idx1 = Int(Index / 19)
Idx0 = Index - (Idx1 * 19)
clidx1 = Int(ClickIdx / 19)
clidx0 = ClickIdx - (clidx1 * 19)
'first click in range 171 - 189
If Index > 170 And First = 0 Then GoTo OK
If Index < 171 And First = 0 Then Call PLAY(2, 1): Exit Sub
'No 0-value
If Lev(Idx0, Idx1) = 0 Then Call PLAY(2, 1): Exit Sub
'has to be appending image
If Idx0 > clidx0 + 1 Then Call PLAY(2, 1): Exit Sub
If Idx0 < clidx0 - 1 Then Call PLAY(2, 1): Exit Sub
If Idx1 > clidx1 + 1 Then Call PLAY(2, 1): Exit Sub
If Idx1 < clidx1 - 1 Then Call PLAY(2, 1): Exit Sub
OK:
'Has to be same color or colorchanger or nextlevel
If Val(Mid(Seq, SeqCnt, 1)) <> Lev(Idx0, Idx1) And Lev(Idx0, Idx1) <> 8 And Lev(Idx0, Idx1) <> 9 Then Call PLAY(2, 1): Exit Sub

'OK! Go on
Image1(Index).Picture = ImageList2.ListImages(Val(Mid(Seq, SeqCnt, 1))).Picture
Sound = 0

'If Reset Level
If Idx0 = 0 And Idx1 = 0 Then GoTo Further1 'skip if coordinates are 0
For t = 0 To 4
If Idx0 = Cor1(t) And Idx1 = Cor2(t) Then
Image1(Index).Picture = ImageList2.ListImages(9).Picture
DoEvents
Sound = 5
GoTo Further
End If
Next t

Further1:
'encounter color change
If Lev(Idx0, Idx1) = 8 Then
Image1(Index).Picture = ImageList2.ListImages(8).Picture
SeqCnt = SeqCnt + 1
If SeqCnt > 14 Then SeqCnt = 14
For t = 0 To 5
Image2(t).Picture = ImageList1.ListImages(Val(Mid(Seq, SeqCnt + t, 1))).Picture
Next t
Image4.Picture = Image2(0).Picture
Sound = 1
End If

'encounter level end
If Lev(Idx0, Idx1) = 9 Then
SeqCnt = SeqCnt + 1
For t = 0 To 5
Image2(t).Picture = ImageList1.ListImages(Val(Mid(Seq, SeqCnt + t, 1))).Picture
Next t
Sound = 4
Timer1.Enabled = False
'goto next level
L = Format(Str(Val(L) + 1), "000")
If L = "021" Then GoTo Further2
End If

Further:
Lev(Idx0, Idx1) = 0
ClickIdx = Index
First = 1
If Sound = 5 Then
Call PLAY(Sound, 0)
ResetLevel
Exit Sub
End If
If Sound = 4 Then
Call PLAY(Sound, 0)
PB1.Cls
Label3.Caption = ""
PointsFrm.Show 1
Exit Sub
End If
Call PLAY(Sound, 1)
Trapped = 0
Exit Sub
Further2:
Timer1.Enabled = False
Label3.Caption = ""
PB1.Cls
Call PLAY(4, 0)
PointsFrm.Show 1
NameFrm.Show 1
End Sub

Private Sub Label6_DblClick()
Timer1.Enabled = False
L = Format(Str(Val(L) + 1), "000")
If L = "020" Then L = "019"
LoadLevel
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Tim = Tim + 1
Label3.Caption = Format(50 - Tim, "00") & " units"
If Tim = 50 Then
Timer1.Enabled = False
Lives = Lives - 1
SetLives
Picture1.Visible = True
Pause = 1
DoEvents
Call PLAY(5, 0)
Picture1.Visible = False
ResetLevel
PB1.Cls
Tim = 0
Timer1.Enabled = True
End If
If Lives = 0 Then
Timer1.Enabled = False
GoTo timer_2
End If
If Tim < 35 Then PB1.Line (Tim * 5, 0)-(Tim * 5 + 3, 10), RGB(255, 128, 0), B
If Tim > 34 Then PB1.Line (Tim * 5, 0)-(Tim * 5 + 3, 10), RGB(255, 255, 0), B
If Tim > 44 Then PB1.Line (Tim * 5, 0)-(Tim * 5 + 3, 10), RGB(255, 0, 0), B
Pause = 0
Exit Sub
timer_2:
NameFrm.Show 1
End Sub
