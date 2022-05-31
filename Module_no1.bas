Attribute VB_Name = "Module_no1"
Option Explicit

Rem Filter #############################################
Const StreamFiles As String = "wav|aif|mp3|mp2|mp1|ogg"
Const MusicFiles As String = "mo3|.xm|mod|s3m|.it|mtm|umx"
Rem Filter #############################################

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private BassLib As Long 'Handle of bass.dll

Public Function LoadBass() As Long
Dim AppPath As String
If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" Else AppPath = App.Path
BassLib = LoadLibrary(AppPath & "bass.dll")
LoadBass = BassLib

' check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
    End
End If

' Initialize output - default device, 44100hz, stereo, 16 bits
If BASS_Init(-1, 44100, 0, Form1.hwnd, 0) = BASSFALSE Then
    Call Error_("Can't initialize digital sound system")
    End
End If
End Function

Public Function FreeBass() As Long
If BassLib = 0 Then Exit Function
Call BASS_Free
FreeBass = FreeLibrary(BassLib)
End Function

'display error messages
Public Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")", vbExclamation, "Error")
End Sub
'get file name from file path
Public Function fileNaME(ByVal Path As String) As String
    fileNaME = Mid(Path, InStrRev(Path, "\") + 1)
End Function

Public Function LoadSoundFile(ByVal FileNam As String) As String
    If InStr(1, StreamFiles, Left(FileNam, 3), vbTextCompare) Then
    
    ElseIf InStr(1, StreamFiles, Left(FileNam, 3), vbTextCompare) Then
    
    End If
    LoadSoundFile = Mid(Path, InStrRev(Path, "\") + 1)
End Function

