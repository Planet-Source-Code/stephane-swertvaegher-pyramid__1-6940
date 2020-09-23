VERSION 5.00
Begin VB.Form Designer 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pyramid - Leveldesigner"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "New level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      TabIndex        =   47
      Top             =   5760
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   6570
      TabIndex        =   45
      Top             =   4230
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5085
      TabIndex        =   44
      Top             =   5265
      Width           =   1140
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   8595
      TabIndex        =   40
      Top             =   3645
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   3
      Left            =   1080
      TabIndex        =   17
      Top             =   5985
      Width           =   3660
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Top             =   5625
      Width           =   3660
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   5265
      Width           =   3660
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   11
      Top             =   4905
      Width           =   3660
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click to load level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   465
      Left            =   8640
      TabIndex        =   48
      Top             =   3150
      Width           =   1230
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level001.lev"
      Height          =   285
      Left            =   4995
      TabIndex        =   46
      Top             =   4860
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   8595
      TabIndex        =   43
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   555
      Left            =   9000
      TabIndex        =   42
      Top             =   2430
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   41
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   19
      Left            =   6210
      TabIndex        =   39
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   18
      Left            =   5940
      TabIndex        =   38
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   17
      Left            =   5670
      TabIndex        =   37
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   16
      Left            =   5400
      TabIndex        =   36
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   15
      Left            =   5130
      TabIndex        =   35
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   14
      Left            =   4860
      TabIndex        =   34
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   13
      Left            =   4590
      TabIndex        =   33
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   12
      Left            =   4320
      TabIndex        =   32
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   11
      Left            =   4050
      TabIndex        =   31
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   10
      Left            =   3780
      TabIndex        =   30
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   9
      Left            =   3510
      TabIndex        =   29
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   8
      Left            =   3240
      TabIndex        =   28
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   7
      Left            =   2970
      TabIndex        =   27
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   6
      Left            =   2700
      TabIndex        =   26
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   5
      Left            =   2430
      TabIndex        =   25
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   4
      Left            =   2160
      TabIndex        =   24
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   3
      Left            =   1890
      TabIndex        =   23
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   2
      Left            =   1620
      TabIndex        =   22
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   1
      Left            =   1350
      TabIndex        =   21
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   0
      Left            =   1080
      TabIndex        =   20
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Colorseq."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   19
      Top             =   4455
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "Timing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   18
      Top             =   6030
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   5670
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   5310
      Width           =   780
   End
   Begin VB.Label Label3 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   4950
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8595
      TabIndex        =   10
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "col"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   9540
      TabIndex        =   9
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Index           =   10
      Left            =   9090
      TabIndex        =   8
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   9045
      TabIndex        =   7
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   8595
      TabIndex        =   6
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   9540
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   9045
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   8595
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   9540
      TabIndex        =   2
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   9045
      TabIndex        =   1
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   375
   End
End
Attribute VB_Name = "Designer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx%, yy%, BcCol&, Lev$(9), Sequence$, Cor1%(4), Cor2%(4), ff%, Temp$
Dim Level$, Counter%

Private Sub SetScreen()
For xx = 1 To 189
Load Label1(xx)
Label1(xx).Visible = True
Next xx
For xx = 0 To 9
For yy = 0 To 18
Label1(yy + (xx * 19)).Move (yy * 27) + 10, (xx * 27) + 5
Next yy
Next xx
Text1(1).Text = "Stephan Swertvaegher"
Text1(3).Text = "750"
BcCol = Label6.BackColor
End Sub

Private Sub Command1_Click() 'savelevel
MakeLev
ff = FreeFile
Open App.Path & "\levels\" & Level For Output As #ff
For xx = 0 To 3
Print #ff, Text1(xx).Text
Next xx
For xx = 0 To 9
Print #ff, Lev(xx)
Next xx
Print #ff, Sequence
For xx = 0 To 4
Print #ff, Cor1(xx)
Print #ff, Cor2(xx)
Next xx
Print #ff, "--- [" & Mid(Level, 6, 3) & "] ---"
Print #ff, "[Legend]"
Print #ff, "1 = Blue"
Print #ff, "2 = Red"
Print #ff, "3 = Green"
Print #ff, "4 = Purple"
Print #ff, "5 = Yellow"
Print #ff, "6 = Cyan"
Print #ff, "7 = Grey"
Print #ff, "8 = ColorChange"
Print #ff, "9 = Next Level"
Close #ff
File1.Refresh
End Sub

Private Sub Command2_Click() 'new level
NewLevel
NextLevel
End Sub

Private Sub NewLevel()
Label6.BackColor = Label2(1).BackColor
Label6.Caption = Label2(1).Caption
Label6.ForeColor = Label2(1).ForeColor
For xx = 0 To 189
Label1(xx).BackColor = Label2(0).BackColor
Label1(xx).Caption = Label2(0).Caption
Label1(xx).ForeColor = Label2(0).ForeColor
Next xx
For xx = 0 To 3
Text1(xx).Text = ""
Next xx
For xx = 0 To 19
Label4(xx).BackColor = Label6.BackColor
Next xx
End Sub
Private Sub File1_DblClick() 'loadlevel
Level = File1.List(File1.ListIndex)
Label7.Caption = Level
NewLevel
ff = FreeFile
Open App.Path & "\levels\" & Level For Input As #ff
For xx = 0 To 3
Line Input #ff, Temp
Text1(xx).Text = Temp
Next xx
For xx = 0 To 9
Line Input #ff, Lev(xx)
Next xx
Line Input #ff, Sequence
For xx = 0 To 4
Input #ff, Cor1(xx)
Input #ff, Cor2(xx)
Next xx
Close #ff
For xx = 0 To 9
For yy = 0 To 18
z = Val(Mid(Lev(xx), yy + 1, 1))
Label1(yy + (xx * 19)).BackColor = Label2(z).BackColor
Label1(yy + (xx * 19)).Caption = Label2(z).Caption
Label1(yy + (xx * 19)).ForeColor = Label2(z).ForeColor
Next yy
Next xx
For xx = 0 To 4
If Cor1(xx) <> 0 And Cor2(xx) <> 0 Then
Label1(Cor1(xx) + (Cor2(xx) * 19)).Caption = "X"
Label1(Cor1(xx) + (Cor2(xx) * 19)).ForeColor = 0
End If
Next xx
For yy = 0 To 18
z = Val(Mid(Sequence, yy + 1, 1))
Label4(yy).BackColor = Label2(z).BackColor
Next yy



End Sub

Private Sub Form_Load()
SetScreen
File1.Path = App.Path & "\levels"
NextLevel
End Sub

Private Sub Label1_Click(Index As Integer)
Label1(Index).BackColor = Label6.BackColor
Label1(Index).Caption = Label6.Caption
Label1(Index).ForeColor = Label6.ForeColor
End Sub

Private Sub Label2_Click(Index As Integer)
If Index <> 10 Then Label6.BackColor = Label2(Index).BackColor
Label6.Caption = Label2(Index).Caption
Label6.ForeColor = Label2(Index).ForeColor
End Sub

Private Sub Label4_Click(Index As Integer)
If Label6.Caption <> "" Then Exit Sub
Label4(Index).BackColor = Label6.BackColor
End Sub

Private Sub MakeLev()
Dim z%
z = 0
For xx = 0 To 9
Lev(xx) = ""
Next xx

For xx = 0 To 9
For yy = 0 To 18
    For z = 0 To 9
    If Label1(yy + (xx * 19)).BackColor = Label2(z).BackColor Then Lev(xx) = Lev(xx) & Trim(Str(z))
    Next z
Next yy
Next xx
Sequence = ""
For xx = 0 To 19
    For z = 0 To 9
    If Label4(xx).BackColor = Label2(z).BackColor Then Sequence = Sequence & Trim(Str(z))
    Next z
Next xx


For xx = 0 To 4
Cor1(xx) = 0: Cor2(xx) = 0
Next xx
z = 0
For xx = 0 To 9
For yy = 0 To 18
    If Label1(yy + (xx * 19)).Caption = "X" Then
    Cor2(z) = xx: Cor1(z) = yy
    z = z + 1
    End If
If z = 5 Then GoTo Further
Next yy
Next xx

Further:
List1.Clear
For xx = 0 To 9
List1.AddItem Lev(xx)
Next xx
List1.AddItem Sequence
List1.AddItem Cor1(0) & "     " & Cor2(0)
List1.AddItem Cor1(1) & "     " & Cor2(1)
List1.AddItem Cor1(2) & "     " & Cor2(2)
List1.AddItem Cor1(3) & "     " & Cor2(3)
List1.AddItem Cor1(4) & "     " & Cor2(4)
End Sub

Private Sub NextLevel()
Level = File1.List(File1.ListCount - 1)
Counter = Val(Mid(Level, 6, 3))
Counter = Counter + 1
Level = "Level" & Format(Str(Counter), "000") & ".lev"
Label7.Caption = Level
End Sub

