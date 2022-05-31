VERSION 5.00
Begin VB.UserControl ucSlider 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   48
   Begin VB.Image imgSlider 
      Height          =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRail 
      Height          =   255
      Left            =   285
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucSlider.ctl
' Author:        Carles P.V. ©2001-2005
' Dependencies:
' Last revision: 2005.05.29 (Original code date: 2001)
' Version:       1.2.0
'========================================================================================

Option Explicit

'-- API:

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal lEdge As Long, ByVal grfFlags As Long) As Long

Private Const BDR_SUNKEN      As Long = &HA
Private Const BDR_RAISED      As Long = &H5
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BF_RECT         As Long = &HF

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
                         
Private Const HWND_TOP       As Long = 0
Private Const HWND_TOPMOST   As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE     As Long = &H1
Private Const SWP_NOMOVE     As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
                         
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT2) As Long

Private Type RECT2
    x1 As Long
    y1 As Long
    X2 As Long
    Y2 As Long
End Type

'-- Public enums.:
Public Enum sOrientationConstants
    [Horizontal] = 0
    [Vertical]
End Enum
Public Enum sRailStyleConstants
    [Sunken] = 0
    [Raised]
    [SunkenSoft]
    [RaisedSoft]
    [ByPicture] = 99
End Enum

'-- Private types:
Private Type Point
    X As Single
    Y As Single
End Type

'-- Private variables:
Private pv_bSliderHooked As Boolean ' imgSlider hooked
Private pv_uSliderOffset As Point   ' imgSlider anchor point
Private pv_uRailRect     As RECT2   ' Rail rectangle
Private pv_uSliderlRect  As RECT2   ' Slider rectangle
Private pv_lAbsCount     As Long    ' pv_lAbsCount = Max - Min
Private pv_lLastValue    As Long    ' Last slider value
Private pv_lTPPx         As Long    ' TwipsPerPixelX
Private pv_lTPPy         As Long    ' TwipsPerPixelY

'-- Default property values:
Private Const m_def_Enabled      As Boolean = True
Private Const m_def_Orientation  As Long = [Vertical]
Private Const m_def_RailStyle    As Long = [Sunken]
Private Const m_def_ShowValueTip As Boolean = True
Private Const m_def_Min          As Long = 0
Private Const m_def_Max          As Long = 10
Private Const m_def_Value        As Long = 0

'-- Property variables:
Private m_Enabled      As Boolean
Private m_Orientation  As sOrientationConstants
Private m_RailStyle    As sRailStyleConstants
Private m_ShowValueTip As Boolean
Private m_Min          As Long
Private m_Max          As Long
Private m_Value        As Long

'-- Event declarations:
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event Change()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)

'###############################################################################
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
''THE POINTAPI STRUCTURE
Private Type POINTAPI
    X As Long                       ' The POINTAPI structure defines the x- and y-coordinates of a point.
    Y As Long
End Type
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
Public Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
Dim lCFrom As Long
Dim lCTo As Long
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long

lCFrom = GetLngColor(oColorFrom)
lCTo = GetLngColor(oColorTo)

lSrcR = lCFrom And &HFF
lSrcG = (lCFrom And &HFF00&) \ &H100&
lSrcB = (lCFrom And &HFF0000) \ &H10000
lDstR = lCTo And &HFF
lDstG = (lCTo And &HFF00&) \ &H100&
lDstB = (lCTo And &HFF0000) \ &H10000

BlendColor = RGB( _
    ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
    ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
    ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
      
End Function

'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
Public Sub DrawGradient(lEndColor As Long, lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hdc As Long, ByVal SegmentSize As Long, Optional bH As Boolean)
    On Error Resume Next
    
    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim er As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    er = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - er) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)

    For ni = 0 To IIf(bH, X2, Y2)
        If ni Mod (SegmentSize + 1) = 0 Then ni = ni + 1
        If bH Then
            DrawLine X + ni, Y, X + ni, Y2, hdc, RGB(er + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine X, Y + ni, X2, Y + ni, hdc, RGB(er + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
    Next ni
End Sub

'DRAWS A LINE WITH A DEFINED COLOR
Public Sub DrawLine( _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim Pos     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
        MoveToEx cHdc, X, Y, Pos
        LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub
'###############################################################################

'========================================================================================
' Usercontrol initialization/termination
'========================================================================================

Private Sub UserControl_Initialize()
    
    pv_lTPPx = Screen.TwipsPerPixelX
    pv_lTPPy = Screen.TwipsPerPixelY
End Sub

'========================================================================================
' Drawing
'========================================================================================

Private Sub UserControl_Show()
    '-- Draw control
    Call Refresh
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    '-- Resize control
    If (m_RailStyle = 99 And imgRail.Picture.handle <> 0) Then
        '-- Horizontal
        If (imgSlider.Height < imgRail.Height) Then
            Size (imgRail.Width + 4) * pv_lTPPx, imgRail.Height * pv_lTPPx
        Else
            Size (imgRail.Width + 4) * pv_lTPPx, imgSlider.Height * pv_lTPPx
        End If
    Else
        '-- Horizontal
        If (Width = 0) Then Width = imgSlider.Width * pv_lTPPx
        Height = imgSlider.Height * pv_lTPPy
    End If
    
    '-- Update slider position
     '-- Horizontal
            If (imgSlider.Height < imgRail.Height And m_RailStyle = 99 And imgRail <> 0) Then
                imgSlider.Top = (imgRail.Height - imgSlider.Height) \ 2
              Else
                imgSlider.Top = 0
            End If
            imgSlider.Left = (m_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
    
    '-- Define rail rectangle
     '-- Horizontal
            With pv_uRailRect
                .y1 = (imgSlider.Height - 4) \ 2
                .Y2 = .y1 + 4
                .x1 = imgSlider.Width \ 2 - 2
                .X2 = .x1 + ScaleWidth - imgSlider.Width + 4
            End With
    
    '-- Refresh control
    Call Refresh
    
    On Error GoTo 0
End Sub

Private Sub Refresh()
    
    '-- Clear control
    Call Cls
    
    '-- Draw rail...
    On Error Resume Next
    
    DrawGradient &HE0E0E0, &H808080, 0, ScaleHeight \ 2 - 2, imgSlider.Left, ScaleHeight \ 2 + 2, hdc, 2, True
    
    '-- Paint image
    Call PaintPicture(imgSlider, imgSlider.Left, imgSlider.Top)
    
    On Error GoTo 0
End Sub

'========================================================================================
' Scrolling
'========================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Me.Enabled) Then
    
        With imgSlider
            
            '-- Hook slider, get offsets and show tip
            If (Button = vbLeftButton) Then
               
                pv_bSliderHooked = True
                
                '-- Mouse over slider
                If (X >= .Left And X < .Left + .Width And Y >= .Top And Y < .Top + .Height) Then
                   
                    pv_uSliderOffset.X = X - .Left
                    pv_uSliderOffset.Y = Y - .Top
                
                Else
                '-- Mouse over rail
                    pv_uSliderOffset.X = .Width \ 2
                    pv_uSliderOffset.Y = .Height \ 2
                    Call UserControl_MouseMove(Button, Shift, X, Y)
                End If
                
                RaiseEvent MouseDown(Shift)
            End If
        End With
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (pv_bSliderHooked) Then
        
        '-- Check limits
        With imgSlider
        
            Select Case m_Orientation
            
                Case 0 '-- Horizontal
                    If (X - pv_uSliderOffset.X < 0) Then
                        .Left = 0
                      ElseIf (X - pv_uSliderOffset.X > ScaleWidth - .Width) Then
                        .Left = ScaleWidth - .Width
                      Else
                        .Left = X - pv_uSliderOffset.X
                    End If
            
                Case 1 '-- Vertical
                    If (Y - pv_uSliderOffset.Y < 0) Then
                        .Top = 0
                      ElseIf (Y - pv_uSliderOffset.Y > ScaleHeight - .Height) Then
                        .Top = ScaleHeight - .Height
                      Else
                        .Top = Y - pv_uSliderOffset.Y
                    End If
            End Select
        End With
        
        '-- Get value from imgSlider position
        Value = pvGetValue
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '-- Click event (If mouse over control area)
    If (X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight And Button = vbLeftButton) Then
        RaiseEvent Click
    End If
    
    '-- MouseUp event (imgSlider has been hooked)
    If (pv_bSliderHooked) Then
        RaiseEvent MouseUp(Shift)
    End If
    
    '-- Unhook slider and hide value tip
    pv_bSliderHooked = False
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvGetValue() As Long
    
    On Error Resume Next
    
    Select Case m_Orientation
    
        Case 0 '-- Horizontal
            pvGetValue = imgSlider.Left / (ScaleWidth - imgSlider.Width) * pv_lAbsCount + m_Min
            imgSlider.Left = (pvGetValue - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
        
        Case 1 '-- Vertical
            pvGetValue = (ScaleHeight - imgSlider.Height - imgSlider.Top) / (ScaleHeight - imgSlider.Height) * pv_lAbsCount + m_Min
            imgSlider.Top = ScaleHeight - imgSlider.Height - (pvGetValue - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
    End Select
    
    On Error GoTo 0
End Function

Private Sub pvResetSlider()

    Select Case m_Orientation
        
        Case 0 '-- Horizontal
            Call imgSlider.Move(0, 0)
             
        Case 1 '-- Vertical
            Call imgSlider.Move(0, ScaleHeight - imgSlider.Height)
    End Select
End Sub

'========================================================================================
' Init/Read/Write properties
'========================================================================================

Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_Orientation = m_def_Orientation
    m_RailStyle = m_def_RailStyle
    m_ShowValueTip = m_def_ShowValueTip
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    
    pv_lAbsCount = 10
    pv_lLastValue = m_Value
    Call pvResetSlider
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
    
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        m_RailStyle = .ReadProperty("RailStyle", m_def_RailStyle)
        m_ShowValueTip = .ReadProperty("ShowValueTip", m_def_ShowValueTip)
        m_Min = .ReadProperty("Min", m_def_Min)
        m_Max = .ReadProperty("Max", m_def_Max)
        m_Value = .ReadProperty("Value", m_def_Value)
        
        Set imgSlider.Picture = .ReadProperty("SliderImage", Nothing)
        Set imgRail = .ReadProperty("RailImage", Nothing)
        
        '-- Get absolute count and set imgSlider position
        pv_lAbsCount = m_Max - m_Min
        pv_lLastValue = m_Value
        imgSlider.Left = (m_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
        imgSlider.Top = (ScaleHeight - imgSlider.Height) - (m_Value - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
        Call .WriteProperty("SliderImage", imgSlider.Picture, Nothing)
        Call .WriteProperty("Orientation", m_Orientation, m_def_Orientation)
        Call .WriteProperty("RailImage", imgRail, Nothing)
        Call .WriteProperty("RailStyle", m_RailStyle, m_def_RailStyle)
        Call .WriteProperty("ShowValueTip", m_ShowValueTip, m_def_ShowValueTip)
        Call .WriteProperty("Min", m_Min, m_def_Min)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Value", m_Value, m_def_Value)
    End With
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call Refresh
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Long)
    If (New_Max <= m_Min) Then Call Err.Raise(380)
    m_Max = New_Max
    pv_lAbsCount = m_Max - m_Min
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Long)
    If (New_Min >= m_Max) Then Err.Raise 380
    m_Min = New_Min
    Value = New_Min
    pv_lAbsCount = m_Max - m_Min
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Long)

    If (New_Value < m_Min Or New_Value > m_Max) Then Call Err.Raise(380)
    
    m_Value = New_Value
        
    If (m_Value <> pv_lLastValue) Then
        
        If (Not pv_bSliderHooked) Then
                   
            Select Case m_Orientation

                Case 0 '-- Horizontal
                    imgSlider.Left = (New_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
                
                Case 1 '-- Vertical
                    imgSlider.Top = ScaleHeight - imgSlider.Height - (New_Value - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
            End Select
        End If
        
        Call Refresh
        pv_lLastValue = m_Value
        
        RaiseEvent Change
        If (m_Value = m_Max) Then RaiseEvent ArrivedLast
        If (m_Value = m_Min) Then RaiseEvent ArrivedFirst
    End If
End Property

Public Property Get Orientation() As sOrientationConstants
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal New_Orientation As sOrientationConstants)
    m_Orientation = New_Orientation
    Call pvResetSlider
    Call UserControl_Resize
End Property

Public Property Get RailStyle() As sRailStyleConstants
    RailStyle = m_RailStyle
End Property
Public Property Let RailStyle(ByVal New_RailStyle As sRailStyleConstants)
    m_RailStyle = New_RailStyle
    Call UserControl_Resize
End Property

Public Property Get SliderImage() As Picture
    Set SliderImage = imgSlider.Picture
End Property
Public Property Set SliderImage(ByVal New_SliderImage As Picture)
    Set imgSlider.Picture = New_SliderImage
    Call UserControl_Resize
End Property

Public Property Get RailImage() As Picture
    Set RailImage = imgRail.Picture
End Property
Public Property Set RailImage(ByVal New_RailImage As Picture)
    Set imgRail.Picture = New_RailImage
    Call UserControl_Resize
End Property

Public Property Get ShowValueTip() As Boolean
    ShowValueTip = m_ShowValueTip
End Property
Public Property Let ShowValueTip(ByVal New_ShowValueTip As Boolean)
    m_ShowValueTip = New_ShowValueTip
End Property
