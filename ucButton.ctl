VERSION 5.00
Begin VB.UserControl ucButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   795
   ScaleWidth      =   2085
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ucButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

'Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Type POINTAPI
    X           As Long
    Y           As Long
End Type
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Public Enum gbHWBtnTypes
    [Netscape 6] = &H0
    [Simple Flat] = &H1
    [Flat Highlight] = &H2
    [Office XP] = &H3
    [KDE 2] = &H4
    [Frame Flat] = &H5
    [Frame Std] = &H6
End Enum
Public Enum gbHWBtnClrTypes
    [Use Windows] = &H0
    Custom = &H1
    [Force Standart] = &H2
    [Use Container] = &H3
End Enum
Public Enum gbHWBtnPicPos
    cbLeft = &H0
    cbRight = &H1
    cbTop = &H2
    cbBottom = &H3
    cbBackground = &H4
End Enum

Private Const COLOR_HIGHLIGHT       As Long = &HD
Private Const COLOR_BTNFACE         As Long = &HF
Private Const COLOR_BTNSHADOW       As Long = &H10
Private Const COLOR_BTNTEXT         As Long = &H12
Private Const COLOR_BTNHIGHLIGHT    As Long = &H14
Private Const COLOR_BTNDKSHADOW     As Long = &H15
Private Const COLOR_BTNLIGHT        As Long = &H16
Private Const PS_SOLID              As Long = &H0
Private Const DT_CENTER             As Long = &H1 Or &H4 Or &H20

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseEnter()
Public Event MouseExit()

'variables
Private MyButtonType        As gbHWBtnTypes, _
        MyColorType         As gbHWBtnClrTypes

Private m_Caption As String, _
        m_BackColor         As Long, _
        m_BackOver          As Long, _
        m_ForeColor         As Long, _
        m_ForeOver          As Long, _
        m_ShowFocusRect     As Boolean, _
        m_PictureNormal     As StdPicture, _
        m_PictureOver       As StdPicture, _
        m_PicturePosition   As gbHWBtnPicPos, _
        m_UseGreyscale      As Boolean, _
        m_CheckBoxBehaviour As Boolean, _
        m_Value             As Boolean, _
        m_BorderStyleEx     As Boolean

Private rcCtl               As RECT, _
        rcRich              As RECT
        
Private picPT               As POINTAPI, _
        picSZ               As POINTAPI  ' picture Position & Size

Private LastButton          As Byte, _
        LastKeyDown         As Byte, _
        LastCaption         As String, _
        LastStat            As Byte, _
        isShown             As Boolean, _
        lTxtFlags           As Long

Private cFace               As Long, _
        cLight              As Long, _
        cHighLight          As Long, _
        cShadow             As Long, _
        cDarkShadow         As Long, _
        cRich               As Long, _
        cRichOver           As Long, _
        cFaceOver           As Long, _
        OXPb                As Long, _
        OXPf                As Long

Private HasFocus As Boolean
Private isOver As Boolean

Private Sub OverTimer_Timer()
    If isMouseOver Then Exit Sub

    OverTimer.Enabled = False: isOver = False
    Call Redraw(&H0, True)

    RaiseEvent MouseExit
End Sub
Private Function isMouseOver() As Boolean
    Dim pt As POINTAPI

    Call GetCursorPos(pt)
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = hWnd)
End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    LastButton = &H1
    Call UserControl_Click
End Sub

Private Sub UserControl_Click()
    If Not LastButton = vbLeftButton Or Not UserControl.Enabled Then Exit Sub

    If m_CheckBoxBehaviour Then m_Value = Not m_Value
    Call Redraw(&H0, True)
    Call UserControl.Refresh

    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If Not LastButton = vbLeftButton Then Exit Sub
    Call UserControl_MouseDown(&H1, &H0, &H0, &H0)
End Sub

Private Sub UserControl_GotFocus()
    HasFocus = True
    Call Redraw(LastStat, True)
End Sub

Private Sub UserControl_Hide()
    isShown = False
End Sub

Private Sub UserControl_Initialize()
    isShown = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)

    LastKeyDown = KeyCode
    Select Case KeyCode
        Case 32         ' spacebar pressed
            Call Redraw(2, False)
        Case 39, 40     ' right and down arrows
            SendKeys "{Tab}"
        Case 37, 38     ' left and up arrows
            SendKeys "+{Tab}"
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)

    If Not KeyCode = 32 Or Not LastKeyDown = 32 Then Exit Sub

    If m_CheckBoxBehaviour Then m_Value = Not m_Value
    Call Redraw(&H0, False)
    Call UserControl.Refresh

    RaiseEvent Click
End Sub

Private Sub UserControl_LostFocus()
    HasFocus = False
    Call Redraw(LastStat, True)
End Sub

Private Sub UserControl_InitProperties()
    'm_ShowFocusRect = True
    MyButtonType = [Simple Flat]

    m_Caption = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    
    m_BackColor = vbButtonFace
    m_BackOver = vbButtonFace
    m_ForeColor = vbButtonText
    m_ForeOver = vbButtonText
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    Call SetCapture(hWnd)
    If Not Button = vbRightButton Then Call Redraw(&H2, False)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button > vbLeftButton Then Exit Sub

    If Not isMouseOver Then         ' Åñëè ìûøêà âíå êíîïêè
        Call Redraw(&H0, False)
    Else                            ' Åñëè íàä êíîïêîé
        If Button = &H0 And Not isOver Then
            OverTimer.Enabled = True
            isOver = True
            Call Redraw(&H0, True)
            RaiseEvent MouseEnter
        ElseIf Button = vbLeftButton Then
            isOver = True
            Call Redraw(&H2, False)
            isOver = False
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not Button = vbRightButton Then Call Redraw(&H0, False)
End Sub

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Öâåò, çàäíåãî âèäà êíîïêè"
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal inColor As OLE_COLOR)
    m_BackColor = inColor

    Call SetColors
    Call Redraw(LastStat, True)
End Property

Public Property Get BackOver() As OLE_COLOR
    BackOver = m_BackOver
End Property
Public Property Let BackOver(ByVal inColor As OLE_COLOR)
    m_BackOver = inColor

    Call SetColors
    Call Redraw(LastStat, True)
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Öâåò íàäïèñè çàãîëîâêà"
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal inColor As OLE_COLOR)
    m_ForeColor = inColor

    Call SetColors
    Call Redraw(LastStat, True)
End Property

Public Property Get ForeOver() As OLE_COLOR
    ForeOver = m_ForeOver
End Property
Public Property Let ForeOver(ByVal inColor As OLE_COLOR)
    m_ForeOver = inColor

    Call SetColors
    Call Redraw(LastStat, True)
End Property

Public Property Get ButtonType() As gbHWBtnTypes
    ButtonType = MyButtonType
End Property
Public Property Let ButtonType(ByVal newValue As gbHWBtnTypes)
    MyButtonType = newValue

    Call UserControl_Resize
    Call Redraw(LastStat, True)
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Çàãîëîâîê êíîïêè (íàäïèñü íà êíîïêå)"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal newValue As String)
    m_Caption = newValue

    Call SetAccessKeys
    Call Redraw(&H0, True)
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Ðàáîòàåò ëè êíîïêà, ëèáî äîñòóïà ê íåé íåò"
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue

    Call Redraw(&H0, True)
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Âîçâðàùàåò îáúåêò Øðèôò"
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef newFont As Font)
    Set UserControl.Font = newFont

    Call Redraw(&H0, True)
End Property

Public Property Get ColorScheme() As gbHWBtnClrTypes
    ColorScheme = MyColorType
End Property
Public Property Let ColorScheme(ByVal newValue As gbHWBtnClrTypes)
    MyColorType = newValue
    
    Call SetColors
    Call Redraw(&H0, True)
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property
Public Property Let ShowFocusRect(ByVal newValue As Boolean)
    m_ShowFocusRect = newValue
    
    Call Redraw(LastStat, True)
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Âûáîð èñïîëüçóåìîãî êóðñîðà íà êíîïêå"
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
    UserControl.MousePointer = newPointer
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Èñïîëüçîâàíèå ñâîåãî êóðñîðà íà êíîïêå"
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewIcon As StdPicture)
    On Local Error Resume Next
    Set UserControl.MouseIcon = NewIcon
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get PictureNormal() As StdPicture
    Set PictureNormal = m_PictureNormal
End Property
Public Property Set PictureNormal(ByVal newPic As StdPicture)
    Set m_PictureNormal = newPic
    
    Call CalcPicSize
    Call Redraw(LastStat, True)
End Property

Public Property Get PictureOver() As StdPicture
    Set PictureOver = m_PictureOver
End Property
Public Property Set PictureOver(ByVal newPic As StdPicture)
    Set m_PictureOver = newPic
    If isOver Then Call Redraw(LastStat, True)
End Property

Public Property Get PicturePosition() As gbHWBtnPicPos
    PicturePosition = m_PicturePosition
End Property
Public Property Let PicturePosition(ByVal newPicPos As gbHWBtnPicPos)
    m_PicturePosition = newPicPos
    
    Call Redraw(LastStat, True)
End Property

Public Property Get UseGreyscale() As Boolean
    UseGreyscale = m_UseGreyscale
End Property
Public Property Let UseGreyscale(ByVal newValue As Boolean)
    m_UseGreyscale = newValue
    
    If Not m_PictureNormal Is Nothing Then Call Redraw(LastStat, True)
End Property

Public Property Get CheckBoxBehaviour() As Boolean
    CheckBoxBehaviour = m_CheckBoxBehaviour
End Property
Public Property Let CheckBoxBehaviour(ByVal newValue As Boolean)
    m_CheckBoxBehaviour = newValue
    
    Call Redraw(LastStat, True)
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property
Public Property Let Value(ByVal newValue As Boolean)
    m_Value = newValue
    
    If m_CheckBoxBehaviour Then Call Redraw(&H0, True)
End Property

Public Property Get BorderStyleEx() As Boolean
    BorderStyleEx = m_BorderStyleEx
End Property
Public Property Let BorderStyleEx(ByVal newValue As Boolean)
    m_BorderStyleEx = newValue
    
    Call Redraw(&H0, True)
End Property

'########## END OF PROPERTIES ##########

Private Sub UserControl_Resize()
    Call GetClientRect(UserControl.hWnd, rcCtl)
    Call Redraw(LastStat, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_BackColor = .ReadProperty("BackColor", vbButtonFace)
        m_BackOver = .ReadProperty("BackOver", vbButtonFace)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_ForeOver = .ReadProperty("ForeOver", vbButtonText)
        m_ShowFocusRect = .ReadProperty("ShowFocusRect", True)
        Set m_PictureNormal = .ReadProperty("PictureNormal", Nothing)
        Set m_PictureOver = .ReadProperty("PictureOver", Nothing)
        m_PicturePosition = .ReadProperty("PicturePosition", cbLeft)
        m_UseGreyscale = .ReadProperty("UseGreyscale", False)
        m_CheckBoxBehaviour = .ReadProperty("CheckBoxBehaviour", False)
        m_Value = .ReadProperty("Value", False)
        m_BorderStyleEx = .ReadProperty("BorderStyleEx", False)
        MyColorType = .ReadProperty("MyColorType", [Use Windows])
        MyButtonType = .ReadProperty("ButtonType", [Simple Flat])
        
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        UserControl.MousePointer = .ReadProperty("MousePointer", &H0)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
    End With

    Call CalcPicSize
    Call SetAccessKeys
End Sub

Private Sub UserControl_Show()
    Call SetColors
    isShown = True
    Call Redraw(&H0, True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("BackColor", m_BackColor, vbButtonFace)
        Call .WriteProperty("BackOver", m_BackOver, vbButtonFace)
        Call .WriteProperty("ForeColor", m_ForeColor, vbButtonText)
        Call .WriteProperty("ForeOver", m_ForeOver, vbButtonText)
        Call .WriteProperty("ShowFocusRect", m_ShowFocusRect, True)
        Call .WriteProperty("PictureNormal", m_PictureNormal, Nothing)
        Call .WriteProperty("PictureOver", m_PictureOver, Nothing)
        Call .WriteProperty("PicturePosition", m_PicturePosition, cbLeft)
        Call .WriteProperty("UseGreyscale", m_UseGreyscale, False)
        Call .WriteProperty("CheckBoxBehaviour", m_CheckBoxBehaviour, False)
        Call .WriteProperty("Value", m_Value, False)
        Call .WriteProperty("BorderStyleEx", m_BorderStyleEx, False)
        Call .WriteProperty("MyColorType", MyColorType, [Use Windows])
        Call .WriteProperty("ButtonType", MyButtonType, [Simple Flat])
        
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, &H0)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
    End With
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Ðèñîâàíèå ñàìîé êíîïêè §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
    
    If m_CheckBoxBehaviour And m_Value Then curStat = &H2
    
    If Not isShown Then Exit Sub
    If Not Force Then If curStat = LastStat And LastCaption = m_Caption Then Exit Sub
    
    Dim tempCol As Long
    
    LastStat = curStat
    LastCaption = m_Caption
    
    If isOver And MyColorType = Custom Then tempCol = m_BackColor: m_BackColor = m_BackOver: SetColors
    
    With UserControl
        Call .Cls
        Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, cFace)
        
        If UserControl.Enabled Then
            Select Case MyButtonType
                
                Case Is = [Simple Flat], [Flat Highlight]
                    If curStat = &H0 Then
                        If MyButtonType = [Simple Flat] Then Call DrawFrame(cHighLight, cShadow, &H0, &H0, False, True)
                        
                        If isOver Then
                            If MyButtonType = [Flat Highlight] Then
                                If Not m_BorderStyleEx Then Call DrawFrame(cHighLight, cShadow, &H0, &H0, False, True)
                            End If
                        End If
                    Else
                        If Not MyButtonType = [Flat Highlight] Then Call DrawFrame(cShadow, cHighLight, &H0, &H0, False, True)
                    End If
                    
                Case Is = [Office XP]
                    If curStat = &H0 Then
                        If isOver Then Call DrawRectangle(&H1, &H1, rcCtl.Right, rcCtl.Bottom, OXPf): _
                                       If Not m_BorderStyleEx Then Call DrawRectangle(&H0, &H0, rcCtl.Right - 1, rcCtl.Bottom - 1, OXPb, True)
                    Else
                        If isOver Then Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, Abs(Not MyColorType = Custom) * ShiftColor(OXPf, -&H20) + Abs(MyColorType = Custom) * ShiftColorOXP(OXPb, &H80))
                        Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, OXPb, True)
                    End If
                    
                Case Is = [Netscape 6]
                    If curStat = &H0 Then
                        Call DrawFrame(ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False)
                    Else
                        Call DrawFrame(cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), False)
                    End If
                    
                Case Is = [Frame Flat]
                    If m_BorderStyleEx And isOver Then
                        Call DrawFrame(cDarkShadow, cDarkShadow, &H0, &H0, False, True)
                    Else
                        Call DrawFrame(cShadow, cShadow, &H0, &H0, False, True)
                    End If
                    
                Case Is = [Frame Std]
                    If curStat = &H0 Then
                        Call DrawFrame(cShadow, cHighLight, cHighLight, cShadow, False)
                    Else
                        Call DrawFrame(cShadow, cHighLight, &H0, &H0, False, True)
                    End If
                    
                Case Is = [KDE 2]
                    If curStat = &H0 Then
                        Dim stepXP As Single, usi As Long
                        
                        If Not isOver Then
                            stepXP = 58 / rcCtl.Bottom
                            For usi = &H1 To rcCtl.Bottom
                                Call DrawLine(0, usi, rcCtl.Right, usi, ShiftColor(cHighLight, -stepXP * usi))
                            Next
                        Else
                            Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, cLight)
                        End If
                        
                        Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, ShiftColor(cShadow, -&H32), True)
                        Call DrawRectangle(&H1, &H1, rcCtl.Right - &H2, rcCtl.Bottom - &H2, ShiftColor(cFace, -&H9), True)
                        Call DrawRectangle(&H2, &H2, rcCtl.Right - &H4, 2, cHighLight)
                        Call DrawRectangle(&H2, &H4, &H2, rcCtl.Bottom - &H6, cHighLight)
                    Else
                        Call DrawRectangle(&H1, &H1, rcCtl.Right, rcCtl.Bottom, ShiftColor(cFace, -&H9))
                        Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, ShiftColor(cShadow, -&H30), True)
                        Call DrawLine(&H2, rcCtl.Bottom - &H2, rcCtl.Right - &H2, rcCtl.Bottom - &H2, cHighLight)
                        Call DrawLine(rcCtl.Right - &H2, &H2, rcCtl.Right - &H2, rcCtl.Bottom - &H1, cHighLight)
                    End If
            End Select
            If curStat = &H0 Then
                Call DrawPicture(Abs(MyButtonType = [Office XP]))
                Call DrawCaption(Abs(isOver))
            Else
                Call DrawPicture(Abs(MyButtonType = [Office XP]) + &H2)
                Call DrawCaption(&H2)
            End If
            Call DrawFocusR
        Else
            Select Case MyButtonType
                Case Is = [Netscape 6]
                    Call DrawFrame(ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False)
                    
                Case Is = [Flat Highlight], [Simple Flat]
                    If MyButtonType = [Simple Flat] Then Call DrawFrame(cHighLight, cShadow, 0, 0, False, True)
                    
                Case Is = [KDE 2]
                    stepXP = 58 / rcCtl.Bottom
                    For usi = &H1 To rcCtl.Bottom
                        DrawLine 0, usi, rcCtl.Right, usi, ShiftColor(cHighLight, -stepXP * usi)
                    Next
                    Call DrawRectangle(&H0, &H0, rcCtl.Right, rcCtl.Bottom, ShiftColor(cShadow, -&H32), True)
                    Call DrawRectangle(&H1, &H1, rcCtl.Right - &H2, rcCtl.Bottom - &H2, ShiftColor(cFace, -&H9), True)
                    Call DrawRectangle(&H2, &H2, rcCtl.Right - &H4, &H2, cHighLight)
                    Call DrawRectangle(&H2, &H4, &H2, rcCtl.Bottom - &H6, cHighLight)
                    
                Case Is = [Frame Flat]
                    Call DrawFrame(cShadow, cShadow, &H0, &H0, False, True)
                    
                Case Is = [Frame Std]
                    Call DrawFrame(cShadow, cHighLight, cHighLight, cShadow, False)
            End Select
            If MyButtonType = [KDE 2] Or MyButtonType = [Office XP] Or MyButtonType = [Frame Flat] Then
                Call DrawPicture(&H5)
                Call DrawCaption(&H4)
            Else
                Call DrawPicture(&H4)
                Call DrawCaption(&H3)
            End If
        End If
    End With
    If isOver And MyColorType = Custom Then m_BackColor = tempCol: SetColors
    UserControl.Refresh
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Ðèñîâàíèå êâàäðàòèêîâ, ôðàéìîâ, ôîêóñîâ è ëèíèé §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub DrawRectangle(ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal Width As Long, _
                          ByVal Height As Long, _
                          ByVal Color As Long, _
                 Optional OnlyBorder As Boolean = False)
    Dim bRECT As RECT, hBrush As Long

    bRECT.Left = X: bRECT.Right = X + Width
    bRECT.Top = Y:  bRECT.Bottom = Y + Height

    hBrush = CreateSolidBrush(Color)
    If OnlyBorder Then Call FrameRect(UserControl.hdc, bRECT, hBrush) Else Call FillRect(UserControl.hdc, bRECT, hBrush)

    Call DeleteObject(hBrush)
End Sub

Private Sub DrawLine(ByVal X1 As Long, _
                     ByVal Y1 As Long, _
                     ByVal X2 As Long, _
                     ByVal Y2 As Long, _
                     ByVal Color As Long)
    Dim pt As POINTAPI
    Dim oldPen As Long, hPen As Long

    With UserControl
        hPen = CreatePen(PS_SOLID, &H1, Color)
        oldPen = SelectObject(.hdc, hPen)

        Call MoveToEx(.hdc, X1, Y1, pt)
        Call LineTo(.hdc, X2, Y2)

        Call SelectObject(.hdc, oldPen)
        Call DeleteObject(hPen)
    End With
End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, _
                      ByVal ColLight As Long, ByVal ColShadow As Long, _
                      ByVal ExtraOffset As Boolean, Optional ByVal Flat As Boolean = False)
    Dim pt As POINTAPI
    Dim frHe As Long, frWi As Long, frXtra As Long

    frHe = rcCtl.Bottom - &H1 + ExtraOffset
    frWi = rcCtl.Right - &H1 + ExtraOffset
    frXtra = Abs(ExtraOffset)

    With UserControl
        Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, &H1, ColHigh)))

        Call MoveToEx(.hdc, frXtra, frHe - 1, pt)
        Call LineTo(.hdc, frXtra, frXtra)
        Call LineTo(.hdc, frWi - 1, frXtra)

        Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, &H1, ColDark)))

        Call LineTo(.hdc, frWi - 1, frHe - 1)
        LineTo .hdc, frXtra - &H1, frHe - 1
        Call MoveToEx(.hdc, frXtra + &H1, frHe - &H1, pt)
        If Flat Then Exit Sub

        Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, &H1, ColLight)))

        Call LineTo(.hdc, frXtra + &H1, frXtra + &H1)
        Call LineTo(.hdc, frWi - &H1, frXtra + &H1)

        Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, &H1, ColShadow)))

        Call LineTo(.hdc, frWi - &H1, frHe - &H1)
        Call LineTo(.hdc, frXtra, frHe - &H1)
    End With
End Sub

Private Sub DrawFocusR()
    If Not m_ShowFocusRect Or Not HasFocus Then Exit Sub
    
    Dim iVal As Integer
    iVal = IIf(MyButtonType = [Netscape 6] Or MyButtonType = [Frame Std], 5, 4)
    Call InflateRect(rcCtl, -iVal, -iVal)
    
    Call SetTextColor(UserControl.hdc, cRich)
    Call DrawFocusRect(UserControl.hdc, rcCtl)
    
    Call InflateRect(rcCtl, iVal, iVal)
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Óñòàíîâêà êíîïî÷åê íà êîíòðîë §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub SetAccessKeys()
    Dim ampersandPos As Long
    UserControl.AccessKeys = vbNullString

    If Len(m_Caption) = &H0 Then Exit Sub
    ampersandPos = InStr(&H1, m_Caption, "&")
    If ampersandPos = &H0 Then Exit Sub

    If Not Mid$(m_Caption, ampersandPos + &H1, &H1) = "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
        UserControl.AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + &H1, &H1))
    Else
        ampersandPos = InStr(ampersandPos + &H2, m_Caption, "&")
        If Not Mid$(m_Caption, ampersandPos + &H1, &H1) = "&" Then UserControl.AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + &H1, &H1))
    End If
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Âíåøíÿÿ ôèãíÿ, óñêîðÿþùàÿ ðàáîòó... §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Public Sub DisableRefresh() 'Äëÿ óñêîðåíèÿ îïåðàöèé
    isShown = False
End Sub

Public Sub Refresh()
    Call SetColors
    isShown = True
    Call Redraw(LastStat, True)
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Ðèñîâàíèå çàãîëîâêà êíîïêè è êàðòèíêè íà íåé §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub DrawCaption(ByVal State As Byte)
    With UserControl
        Select Case State
            Case Is = &H0                       ' normal caption
                Call SetTextColor(.hdc, cRich)
                
            Case Is = &H1                       ' hover caption
                Call SetTextColor(.hdc, cRichOver)
                
            Case Is = &H2                       ' down caption
                Call SetTextColor(.hdc, cRichOver)
                Call OffsetRect(rcRich, &H1, &H1)
                Call DrawText(.hdc, m_Caption, Len(m_Caption), rcRich, lTxtFlags)
                Call OffsetRect(rcRich, &HFFFF, &HFFFF)
                
            Case Is = &H3                       ' disabled embossed caption
                Call SetTextColor(.hdc, cHighLight)
                Call OffsetRect(rcRich, &H1, &H1)
                Call DrawText(.hdc, m_Caption, Len(m_Caption), rcRich, lTxtFlags)
                Call SetTextColor(.hdc, cShadow)
                Call OffsetRect(rcRich, &HFFFF, &HFFFF)
                
            Case Is = &H4                       ' disabled grey caption
                Call SetTextColor(.hdc, cShadow)
        End Select

        If Not State = &H2 Then Call DrawText(.hdc, m_Caption, Len(m_Caption), rcRich, lTxtFlags)
    End With
End Sub

Private Sub DrawPicture(ByVal State As Byte)
    Dim hBr As Long, lFlagN As Long, lFlagO As Long
    
    Call GetClientRect(UserControl.hWnd, rcRich)
    Call CalcPicRichPos
    
    If Not m_PictureOver Is Nothing Then lFlagO = IIf(m_PictureOver.Type = &H1, &H4, &H3)
    If Not m_PictureNormal Is Nothing Then lFlagN = IIf(m_PictureNormal.Type = &H1, &H4, &H3)
    
    hBr = CreateSolidBrush(cShadow)
    
    Select Case State
        Case Is = &H0       ' Íîðìàëüíûå áàòîíû
            If isOver And Not m_PictureOver Is Nothing Then
                Call DrawState(UserControl.hdc, &H0, &H0, m_PictureOver.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagO)
            ElseIf Not m_PictureNormal Is Nothing Then
                Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN)
            End If
        
        Case Is = &H1       ' OfficeXP
            If Not m_PictureNormal Is Nothing Then
                If isOver And Not m_PictureNormal Is Nothing Then
                    Call DrawState(UserControl.hdc, hBr, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN Or &H80)
                    Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X - &H1, picPT.Y - &H1, picSZ.X, picSZ.Y, lFlagN)
                ElseIf Not m_PictureNormal Is Nothing Then
                    Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN)
                End If
            End If
        Case Is = &H2
            If isOver And Not m_PictureOver Is Nothing Then
                Call DrawState(UserControl.hdc, &H0, &H0, m_PictureOver.Handle, &H0, picPT.X + &H1, picPT.Y + &H1, picSZ.X, picSZ.Y, lFlagO)
            ElseIf Not m_PictureNormal Is Nothing Then
                Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X + &H1, picPT.Y + &H1, picSZ.X, picSZ.Y, lFlagN)
            End If
        Case Is = &H3
            If Not m_PictureNormal Is Nothing Then Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN)
        Case Is = &H4
            If Not m_PictureNormal Is Nothing Then Call DrawState(UserControl.hdc, &H0, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN Or &H20)
        Case Is = &H5
            If Not m_PictureNormal Is Nothing Then Call DrawState(UserControl.hdc, hBr, &H0, m_PictureNormal.Handle, &H0, picPT.X, picPT.Y, picSZ.X, picSZ.Y, lFlagN Or &H80)
    End Select
    Call DeleteObject(hBr)
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Ðàñ÷åòû ðàçìåðîâ êàðòèíêè è ïîçèöèè. Âûñòàâëååíèå RECT'a äëÿ òåêñòà §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub CalcPicSize()
    If Not m_PictureNormal Is Nothing Then
        picSZ.X = UserControl.ScaleX(m_PictureNormal.Width, &H8, vbPixels)
        picSZ.Y = UserControl.ScaleY(m_PictureNormal.Height, &H8, vbPixels)
    Else
        picSZ.X = &H0: picSZ.Y = &H0
    End If
End Sub

Private Sub CalcPicRichPos()
    If m_PictureNormal Is Nothing And m_PictureOver Is Nothing Then GoTo none

    If Len(Trim$(m_Caption)) = &H0 Or m_PicturePosition = cbBackground Then GoTo none
    Select Case m_PicturePosition
        Case Is = cbLeft
            picPT.X = (rcCtl.Right - TextWidth(m_Caption) \ Screen.TwipsPerPixelX - &H4) \ &H2 - picSZ.X
            picPT.Y = (rcCtl.Bottom - picSZ.Y) \ &H2
            If picPT.X < &H4 Then picPT.X = &H4
            
            rcRich.Left = picPT.X + picSZ.X + &H4
            lTxtFlags = &H4 Or &H20
        Case Is = cbRight
            picPT.X = rcCtl.Right - (rcCtl.Right - TextWidth(m_Caption) \ Screen.TwipsPerPixelX - &H4) \ &H2
            picPT.Y = (rcCtl.Bottom - picSZ.Y) \ &H2
            If picPT.X > rcCtl.Right - picSZ.X - &H4 Then picPT.X = rcCtl.Right - picSZ.X - &H4
            
            rcRich.Right = picPT.X - &H4
            lTxtFlags = &H4 Or &H20 Or &H2
        Case Is = cbTop
            picPT.X = (rcCtl.Right - picSZ.X) \ &H2
            picPT.Y = rcCtl.Top + &H2
            
            rcRich.Top = picPT.Y + picSZ.Y
            lTxtFlags = &H1 Or &H20 Or &H4
        Case Is = cbBottom
            picPT.X = (rcCtl.Right - picSZ.X) \ &H2
            picPT.Y = rcCtl.Bottom - picSZ.Y - &H2
            
            rcRich.Bottom = picPT.Y
            lTxtFlags = &H1 Or &H20 Or &H4
    End Select
Exit Sub
' §§§§§§§§§§§§§§§§§§§§§§§§§§ Ëèáî êàðòèíîê íåò, ëèáî çàãîëîâêà, èëè ïðîñòîë êàðòèíêà - ôîíîì §§§§§§§§§§§§§§§§§§§§§§§§§§
none:
    picPT.X = (rcCtl.Right - picSZ.X) \ &H2
    picPT.Y = (rcCtl.Bottom - picSZ.Y) \ &H2

    lTxtFlags = &H1 Or &H20 Or &H4
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Íàñòðîéêè öâåòà §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub SetColors()
    If MyColorType = Custom Then
        cFace = GetFromSysColor(m_BackColor)
        cFaceOver = GetFromSysColor(m_BackOver)
        cRich = GetFromSysColor(m_ForeColor)
        cRichOver = GetFromSysColor(m_ForeOver)
        cShadow = ShiftColor(cFace, -&H40)
        cLight = ShiftColor(cFace, &H1F)
        cHighLight = ShiftColor(cFace, &H2F)
        cDarkShadow = ShiftColor(cFace, -&HC0)
        OXPb = ShiftColor(cFace, -&H80)
        OXPf = cFace
    ElseIf MyColorType = [Use Windows] Then
        cFace = GetSysColor(COLOR_BTNFACE)
        cFaceOver = cFace
        cShadow = GetSysColor(COLOR_BTNSHADOW)
        cLight = GetSysColor(COLOR_BTNLIGHT)
        cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
        cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
        cRich = GetSysColor(COLOR_BTNTEXT)
        cRichOver = cRich
        OXPb = GetSysColor(COLOR_HIGHLIGHT)
        OXPf = ShiftColorOXP(OXPb)
    ElseIf MyColorType = [Force Standart] Then
        cFace = &HC0C0C0
        cFaceOver = cFace
        cShadow = &H808080
        cLight = &HDFDFDF
        cDarkShadow = &H0
        cHighLight = &HFFFFFF
        cRich = &H0
        cRichOver = cRich
        OXPb = &H800000
        OXPf = &HD1ADAD
    ElseIf MyColorType = [Use Container] Then
        cFace = GetBkColor(GetDC(GetParent(hWnd)))
        cFaceOver = cFace
        cRich = GetTextColor(GetDC(GetParent(hWnd)))
        cRichOver = cRich
        cShadow = ShiftColor(cFace, -&H40)
        cLight = ShiftColor(cFace, &H1F)
        cHighLight = ShiftColor(cFace, &H2F)
        cDarkShadow = ShiftColor(cFace, -&HC0)
        OXPb = GetSysColor(COLOR_HIGHLIGHT)
        OXPf = ShiftColorOXP(OXPb)
    End If
End Sub

Public Function GetFromSysColor(ByVal theColor As Long) As Long
    Call OleTranslateColor(theColor, &H0, GetFromSysColor)
End Function

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long
    Dim Red As Long, Blue As Long, Green As Long

    'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
    'If isSoft Then Value = Value \ 2

    Blue = ((Color \ &H10000) Mod &H100) + Value
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value

    If Value > 0 Then
        If Red > 255 Then Red = 255
        If Green > 255 Then Green = 255
        If Blue > 255 Then Blue = 255
    ElseIf Value < 0 Then
        If Red < 0 Then Red = 0
        If Green < 0 Then Green = 0
        If Blue < 0 Then Blue = 0
    End If

    ShiftColor = Red + &H100 * Green + &H10000 * Blue
End Function

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
    Dim Red As Long, Blue As Long, Green As Long
    Dim Delta As Long

    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base

    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF

    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255

    ShiftColorOXP = Red + &H100 * Green + &H10000 * Blue
End Function

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Ðèñîâàíèå ïðîçðà÷íîé êàðòèíêè §§§§§§§§§§§§§§§§§§§§§§§§§§

'Private Sub TransparentBlt(ByVal OutDstDC As Long, _
'                           ByVal DstDC As Long, _
'                           ByVal SrcDC As Long, _
'                           ByVal SrcRect As RECT, _
'                           ByVal DstX As Integer, _
'                           ByVal DstY As Integer, _
'                           ByVal TransColor As Long)
'
'     Dim W As Integer, H As Integer
'     Dim MonoMaskDC As Long, hMonoMask As Long
'     Dim MonoInvDC As Long
'     Dim ResultDstDC As Long
'     Dim ResultSrcDC As Long
'     Dim hPrevMask As Long, hPrevInv As Long
'     Dim hPrevSrc As Long, hPrevDst As Long
'
'     W = SrcRect.Right - SrcRect.Left + &H1
'     H = SrcRect.Bottom - SrcRect.Top + &H1
'
'     MonoMaskDC = CreateCompatibleDC(DstDC)                                                 'create monochrome mask and inverse masks
'     MonoInvDC = CreateCompatibleDC(DstDC)
'     hMonoMask = CreateBitmap(W, H, &H1, &H1, ByVal &H0)
'     hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
'     hPrevInv = SelectObject(MonoInvDC, CreateBitmap(W, H, &H1, &H1, ByVal &H0))
'
'     ResultDstDC = CreateCompatibleDC(DstDC)                                                'create keeper DCs and bitmaps
'     ResultSrcDC = CreateCompatibleDC(DstDC)
'     hPrevDst = SelectObject(ResultDstDC, CreateCompatibleBitmap(DstDC, W, H))
'     hPrevSrc = SelectObject(ResultSrcDC, CreateCompatibleBitmap(DstDC, W, H))
'
'     Dim OldBC As Long                                                                      'copy src to monochrome mask
'     OldBC = SetBkColor(SrcDC, TransColor)
'     Call BitBlt(MonoMaskDC, &H0, &H0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
'     TransColor = SetBkColor(SrcDC, OldBC)
'
'     Call BitBlt(MonoInvDC, &H0, &H0, W, H, MonoMaskDC, &H0, &H0, vbNotSrcCopy)             'create inverse of mask
'     Call BitBlt(ResultDstDC, &H0, &H0, W, H, DstDC, DstX, DstY, vbSrcCopy)                 'get background
'     Call BitBlt(ResultDstDC, &H0, &H0, W, H, MonoMaskDC, &H0, &H0, vbSrcAnd)               'AND with Monochrome mask
'     Call BitBlt(ResultSrcDC, &H0, &H0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)  'get overlapper
'     Call BitBlt(ResultSrcDC, &H0, &H0, W, H, MonoInvDC, &H0, &H0, vbSrcAnd)                'AND with inverse monochrome mask
'     Call BitBlt(ResultDstDC, &H0, &H0, W, H, ResultSrcDC, &H0, &H0, vbSrcInvert)           'XOR these two
'     Call BitBlt(OutDstDC, DstX, DstY, W, H, ResultDstDC, &H0, &H0, vbSrcCopy)              'output results
'
'     hMonoMask = SelectObject(MonoMaskDC, hPrevMask)                                        'clean up
'     Call DeleteObject(hMonoMask)
'
'     Call DeleteObject(SelectObject(MonoInvDC, hPrevInv))
'     Call DeleteObject(SelectObject(ResultDstDC, hPrevDst))
'     Call DeleteObject(SelectObject(ResultSrcDC, hPrevSrc))
'
'     Call DeleteDC(MonoMaskDC)
'     Call DeleteDC(MonoInvDC)
'     Call DeleteDC(ResultDstDC)
'     Call DeleteDC(ResultSrcDC)
'End Sub

