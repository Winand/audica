VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   3  'Windows Default
   Begin WAB.ucSlider Pos 
      Height          =   120
      Left            =   120
      Top             =   480
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   212
      BackColor       =   14737632
      SliderImage     =   "Form1.frx":0000
      Orientation     =   0
      RailImage       =   "Form1.frx":0132
      Max             =   1
   End
   Begin VB.Timer PlayTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   720
   End
   Begin WAB.ucButton StopB 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "[]"
      BackColor       =   16777215
      BackOver        =   12632256
      ShowFocusRect   =   0   'False
      MyColorType     =   1
      ButtonType      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WAB.ucButton PlayB 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ">"
      BackColor       =   16777215
      BackOver        =   12632256
      ShowFocusRect   =   0   'False
      MyColorType     =   1
      ButtonType      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WAB.ucButton ucButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "+"
      BackColor       =   16777215
      BackOver        =   12632256
      ShowFocusRect   =   0   'False
      MyColorType     =   1
      ButtonType      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LVContainer 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
   Begin WAB.epCmDlg Comd1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      MultiSelect     =   -1  'True
   End
   Begin VB.Image PIC_Play 
      Height          =   210
      Left            =   3720
      Picture         =   "Form1.frx":014E
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label ProgVal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0:00:00 / 0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents PlayList As cListView
Attribute PlayList.VB_VarHelpID = -1
Public CurrentStream As Long

Private Sub Form_Load()
CreateListView '///
LoadBass
PlayList.ImgLst_AddIcon PIC_Play.Picture
End Sub

Private Sub CreateListView()
Set PlayList = New cListView
Call PlayList.Create(LVContainer.hwnd, LVS_ICON + LVS_NOCOLUMNHEADER + LVS_REPORT + LVS_SHOWSELALWAYS, 0, 0, LVContainer.Width, LVContainer.Height, , WS_EX_STATICEDGE)
Call PlayList.SetStyleEx(LVS_EX_FLATSB + LVS_EX_FULLROWSELECT + LVS_EX_GRIDLINES + LVS_EX_HEADERDRAGDROP)
PlayList.HeaderButtons = False
PlayList.AddColumn 1, , 150
PlayList.AddColumn 2, , 45
PlayList.AddColumn 3, , 200
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call BASS_Free
FreeBass
End Sub

Private Sub Form_Resize()
On Error Resume Next
LVContainer.Move LVContainer.Left, LVContainer.Top, ScaleWidth, ScaleHeight - LVContainer.Top
End Sub

Private Sub LVContainer_Resize()
On Error Resume Next
PlayList.Move 0, 0, LVContainer.Width, LVContainer.Height
End Sub

Private Sub PlayB_Click()
If CurrentStream = 0 Then Exit Sub
'If (PlayList.ListIndex >= 0) Then
    If (BASS_ChannelPlay(CurrentStream, BASSFALSE) = 0) Then _
        Call Error_("Can't play stream"): Exit Sub
    Pos.Max = BASS_ChannelBytes2Seconds(CurrentStream, BASS_ChannelGetLength(CurrentStream))
    PlayTimer.Enabled = True
'End If
End Sub

Private Sub PlayList_DblClick(ByVal iItem As Long, ByVal Button As MouseButtonConstants)
BASS_ChannelStop CurrentStream
BASS_ChannelSetPosition CurrentStream, 0

CurrentStream = PlayList.ItemData(iItem)
'PlayList.ItemIconIndex(PlayList.ListIndex) = 0
if
PlayB_Click
End Sub

Private Sub PlayList_VScroll(ByVal cPos As Long, flag As gbLVVScrollEnum)
PlayList.TopItem
End Sub

Private Sub PlayTimer_Timer()
'If BASS_ChannelIsActive(CurrentStream) <> 1 Then Exit Sub
Pos.Value = BASS_ChannelBytes2Seconds(CurrentStream, BASS_ChannelGetPosition(CurrentStream))
End Sub

Private Sub Pos_Change()
ProgVal.Caption = TimeSerial(0, 0, Pos.Value) & " / " & _
                  Round(Pos.Value / (Pos.Max / 100), 2) & "%"
End Sub

Private Sub Pos_MouseDown(Shift As Integer)
'If CurrentStream = 0 Then Exit Sub
PlayTimer.Enabled = False
End Sub

Private Sub Pos_MouseUp(Shift As Integer)
If CurrentStream = 0 Then Exit Sub
BASS_ChannelSetPosition CurrentStream, BASS_ChannelSeconds2Bytes(CurrentStream, Pos.Value)
PlayTimer.Enabled = True
End Sub

Private Sub StopB_Click()
If CurrentStream = 0 Then Exit Sub
PlayTimer.Enabled = False
BASS_ChannelStop CurrentStream
BASS_ChannelSetPosition CurrentStream, 0
Pos.Value = 0
End Sub

Private Sub ucButton1_Click()
'*.wav;*.aif;*.mp3;*.mp2;*.mp1;*.ogg;*.mo3;*.xm;*.mod;*.s3m;*.it;*.mtm;*.umx
Comd1.Filter = "Все поддерживаемые форматы|*.wav;*.aif;*.mp3;*.mp2;*.mp1;*.ogg;*.mo3;*.xm;*.mod;*.s3m;*.it;*.mtm;*.umx|" & _
               "Streamable Files (*.wav;*.aif;*.mp3;*.mp2;*.mp1;*.ogg)|*.wav;*.aif;*.mp3;*.mp2;*.mp1;*.ogg|" & _
               "MOD Music Files (*.mo3;*.xm;*.mod;*.s3m;*.it;*.mtm;*.umx)|*.mo3;*.xm;*.mod;*.s3m;*.it;*.mtm;*.umx|" & _
               "Все файлы (*.*)|*.*"
Comd1.ShowOpen

Dim i As Long, StreamHandle As Long, FileStr As String, TMP As Long
For i = 1 To Comd1.cFileName.Count
    FileStr = Comd1.cFileName(i)
    
    StreamHandle = BASS_StreamCreateFile(BASSFALSE, FileStr, 0, 0, 0)
'    BASS_MusicLoad(BASSFALSE, FileStr, 0, 0, 0)
    If StreamHandle = 0 Then
        Call Error_("Can't open stream")
    Else
        PlayList.AddItem filename(FileStr), , 1
        PlayList.ItemText(2, PlayList.Count - 1) = FileStr
        TMP = BASS_ChannelBytes2Seconds(StreamHandle, BASS_ChannelGetLength(StreamHandle)) 'Длина в секундах
        PlayList.ItemText(1, PlayList.Count - 1) = TMP \ 60 & ":" & Format(TMP Mod 60, "00")
        PlayList.ItemData(PlayList.Count - 1) = StreamHandle
    End If
Next i

If PlayList.ListIndex < 0 Then
    PlayList.ListIndex = PlayList.Count - 1
    CurrentStream = StreamHandle
End If
End Sub

Private Function PlayinItemVisible() As Long

End Function
