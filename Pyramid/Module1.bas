Attribute VB_Name = "Module1"
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum

Public xx%, yy%, ff%, First%, SeqCnt%, Lev%(18, 9), Lev2%(18, 9), ClickIdx%, Tim%, Sound%
Public Points&, Score&
Public q%, p%, TxCount%
Public t&
Public L$, Diff$, PLine$(9), Title$, dummy$, Seq$, Progress$
Public Idx0%, Idx1%, clidx0%, clidx1%
Public Cor1%(4), Cor2%(4), Lives%, Pause%, CurIdx%, Trapped%
Public HScore$(10, 1), PlName$, Temp$
Public RC, MidFile
Public SNDbestand As String
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const snd_sync = &H0
Public Const snd_async = &H1
Public Const snd_loop = &H8
Public WavNam(25) As String
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6 ' masks
Public Const SRCPAINT = &HEE0086 ' paint over masks
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Function StopMIDI(MIDIFile$)
    Dim SafeFile$
    SafeFile = Dir(MIDIFile)
    If SafeFile <> "" Then
        Call mciSendString("stop " & MIDIFile, 0&, 0, 0)
    End If
End Function


Public Function PlayMIDI(MIDIFile$)
    Dim SafeFile$
    SafeFile = Dir(MIDIFile)
    If SafeFile <> "" Then
        Call mciSendString("play " & MIDIFile, 0&, 0, 0)
    End If
End Function

Public Sub PSetup()
With Pyramid
Pyramid.Move (Screen.Width / 2) - (Pyramid.Width / 2), 0
Pause = 0
ColBox .Picture1, 0, 0, .Picture1.ScaleWidth, .Picture1.ScaleHeight, 7, 64, 128, 0, 128, 255, 0
ColBox Pyramid, 462, 5, 628, 61, 5, 64, 128, 0, 128, 255, 0
.Picture1.Left = (Pyramid.ScaleWidth / 2) - (.Picture1.ScaleWidth / 2)
.Picture1.Visible = False
For xx = 1 To 189
Load .Image1(xx)
.Image1(xx).Visible = True
Next xx
For yy = 0 To 9
For xx = 0 To 18
.Image1(xx + (yy * 19)).Move (xx * 33) + 10, (yy * 33) + 160
.Image1(xx + (yy * 19)).Visible = False
Next xx
Next yy
ColBox Pyramid, 4, 67, 641, 150, 7, 64, 128, 0, 128, 255, 0
WavNam(0) = App.Path & "\Sounds\boing.wav"
WavNam(1) = App.Path & "\Sounds\effect2.wav"
WavNam(2) = App.Path & "\Sounds\Voltage.wav"
WavNam(3) = App.Path & "\Sounds\ufo.wav"
WavNam(4) = App.Path & "\Sounds\funky.wav"
WavNam(5) = App.Path & "\Sounds\Reset.wav"
WavNam(6) = App.Path & "\Sounds\NewLife.wav"
WavNam(7) = App.Path & "\Sounds\Pon.wav"
WavNam(8) = App.Path & "\Sounds\Start.wav"
Lives = 5
SetLives
Trapped = 0
End With
End Sub

Public Sub DrawBack()
Pyramid.Texture.Picture = LoadPicture(App.Path & "\texture\texture" & Trim(Str(TxCount)) & ".jpg")
For xx = 0 To 7
For yy = 0 To 7
BitBlt Pyramid.hDC, (xx * 79) + 6, (yy * 42) + 157, 79, 42, Pyramid.Texture.hDC, 0, 0, SRCCOPY
Next yy
Next xx
Pyramid.Line (4, 155)-(640, 495), RGB(64, 128, 10), B
Pyramid.Line (5, 156)-(639, 494), RGB(64, 248, 10), B
Pyramid.Line (6, 157)-(638, 493), RGB(64, 128, 10), B
'Pyramid.Refresh
End Sub
Public Sub LoadLevel()
If L = 21 Then Exit Sub
DrawBack
With Pyramid

Dim PCol%
For xx = 0 To 189
.Image1(xx).Visible = False
Next xx
DoEvents
ff = FreeFile
Open App.Path & "\Levels\Level" & L & ".lev" For Input As #ff
Line Input #ff, Title
Input #ff, dummy
Input #ff, Diff
Input #ff, Progress
.Timer1.Interval = Val(Progress)
For t = 0 To 9
Line Input #ff, PLine(t)
Next t
Line Input #ff, Seq
For t = 0 To 4
Input #ff, Cor1(t)
Input #ff, Cor2(t)
Next t
Close #ff
'SetLevel
.Label1(0).Caption = "Level " & L & ": " & Title & vbCr & "Difficulty: " & Diff
.Label1(1).Caption = "Level " & L & ": " & Title & vbCr & "Difficulty: " & Diff
SeqCnt = 1
For t = 0 To 5
.Image2(t).Picture = .ImageList1.ListImages(Val(Mid(Seq, SeqCnt + t, 1))).Picture
Next t
.Image4.Picture = .Image2(0).Picture
DoEvents
For yy = 0 To 9
For xx = 0 To 18
Lev(xx, yy) = Val(Mid(PLine(yy), xx + 1, 1))
Lev2(xx, yy) = Val(Mid(PLine(yy), xx + 1, 1))
If Mid(PLine(yy), xx + 1, 1) <> "0" Then
PCol = Val(Mid(PLine(yy), xx + 1, 1))
.Image1(xx + (yy * 19)).Visible = True
.Image1(xx + (yy * 19)).Picture = .ImageList1.ListImages(PCol).Picture
End If
DoEvents
For t = 0 To 1000
Next t
Next xx
Next yy
First = 0
ClickIdx = 0
.PB1.Cls
Tim = 0
.Label3.Caption = "50 units"
If Val(L) / 5 = Int(Val(L) / 5) Then ' + 1 life every 5 levels
Lives = Lives + 1
If Lives > 8 Then Lives = 8 'max 8 lives
SetLives
.Label7.Visible = True
DoEvents
Call PLAY(6, 0)
.Label7.Visible = False
End If
Trapped = 1
End With
End Sub

Sub PLAY(sndnr As Integer, snd As Integer)
RC = sndPlaySound(WavNam(sndnr), snd)
End Sub

Public Sub ResetLevel()
With Pyramid
For yy = 0 To 9
For xx = 0 To 18
Lev(xx, yy) = Lev2(xx, yy)
If Lev(xx, yy) <> "0" Then
PCol = Lev2(xx, yy)
.Image1(xx + (yy * 19)).Visible = True
.Image1(xx + (yy * 19)).Picture = .ImageList1.ListImages(PCol).Picture
End If
DoEvents
Next xx
Next yy
First = 0
ClickIdx = 0
SeqCnt = 1
For t = 0 To 5
.Image2(t).Picture = .ImageList1.ListImages(Val(Mid(Seq, SeqCnt + t, 1))).Picture
Next t
End With
End Sub

Public Sub SetLives()
For xx = 0 To 8
Pyramid.Image3(xx).Visible = False
If xx <= Lives - 1 Then Pyramid.Image3(xx).Visible = True
Next xx
End Sub

Public Sub ColBar(Obj As Object, St%, h%, R%, G%, B%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(h / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - G) / H3
IvB = Int(BE - B) / H3
Do While h >= H3
Obj.Line (0, St + H2)-(Obj.ScaleWidth, St + H2), RGB(R, G, B)
Obj.Line (0, St + h)-(Obj.ScaleWidth, St + h), RGB(R, G, B)
h = h - 1
H2 = H2 + 1
R = R + IvR
G = G + IvG
B = B + IvB
Loop
End Sub

Public Sub ColBox(Obj As Object, BX%, BY%, EX%, EY%, h%, R%, G%, B%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(h / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - G) / H3
IvB = Int(BE - B) / H3
Do While h >= H3
Obj.Line (BX + H2, BY + H2)-(EX - H2, EY - H2), RGB(R, G, B), B
Obj.Line (BX + h, BY + h)-(EX - h, EY - h), RGB(R, G, B), B
h = h - 1
H2 = H2 + 1
R = R + IvR
G = G + IvG
B = B + IvB
Loop
End Sub

Public Sub LoadHighScore()
Dim ff%
On Error GoTo LHS
ff = FreeFile
Open App.Path & "\HighScore\HighScore.hsc" For Input As #ff
For xx = 0 To 9
Line Input #ff, Temp
HScore(xx, 0) = Left(Temp, 10)
HScore(xx, 1) = Right(Temp, 8)
Next xx
LHS:
Close #ff
End Sub
Public Sub SaveHighScore()
Dim ff%
On Error GoTo SHS
ff = FreeFile
Open App.Path & "\HighScore\HighScore.hsc" For Output As #ff
For xx = 0 To 9
Temp = Left(HScore(xx, 0) + "          ", 10) & HScore(xx, 1)
Print #ff, Temp
Next xx
SHS:
Close #ff
End Sub

Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim T3Dxx%
On Error Resume Next

Obj.Borderstyle = 0 'no border

If IsMissing(Style3D) Then Style3D = 0

If Style3D > 4 Then Style3D = 3

If Style3D = 0 Then 'RaiseRaise
R1 = 240: R2 = 128: R3 = 240: R4 = 128
End If
If Style3D = 1 Then 'RaiseInset
R1 = 240: R2 = 128: R4 = 240: R3 = 128
End If
If Style3D = 2 Then 'InsetRaise
R2 = 240: R1 = 128: R3 = 240: R4 = 128
End If
If Style3D = 3 Then 'InsetInset
R2 = 240: R1 = 128: R4 = 240: R3 = 128
End If
If Style3D = 4 Then 'No Border
R1 = 192: R2 = 192: R3 = 192: R4 = 192
End If

G1 = 0
B1 = R1
G2 = 0
B2 = R2
G3 = 0
B3 = R3
G4 = 0
B4 = R4
Bev = Bev + 1
T3Dxx = Bev
'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
End Function
