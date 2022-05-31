VERSION 5.00
Begin VB.UserControl VolSlider 
   AutoRedraw      =   -1  'True
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   ClipControls    =   0   'False
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   98
   Begin VB.Image Slider 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   0
      Picture         =   "VolSlider.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "VolSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Fader.ctl
' Custom control to display custom/skinned slider to control the master volume
' This submission piggybacked from the excellent cpvSlider by Carles P.V.
'**************************************************************************************************
'  Copyright © 2004, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
'  edited by [rm_code]
'  removed mixer handling
'**************************************************************************************************

Option Explicit

'**************************************************************************************************
'  Fader Constants
'**************************************************************************************************
Private Const COL_CNT = 2
Private Const AbsCount = 100

'**************************************************************************************************
'  Fader Structs and Enums
'**************************************************************************************************
Public Enum eOrientation
    [Horizontal]
    [Vertical]
End Enum ' eOrientation

Private Type POINTAPI
    X As Single
    Y As Single
End Type ' POINTAPI

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type ' RECT

'**************************************************************************************************
' Fader Win32 API
'**************************************************************************************************
' drawing api
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long, lpPoint As POINTAPI) As Long

'**************************************************************************************************
' Fader Events
'**************************************************************************************************
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event ValueChanged()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)

'**************************************************************************************************
' Fader Module-Level variables
'**************************************************************************************************
Private SliderHooked As Boolean
Private SliderOffset As POINTAPI
Private LastValue As Long
Private tppX As Long
Private tppY As Long
Attribute tppY.VB_VarHelpID = -1

'**************************************************************************************************
'  Fader Default Control Property Variables
'**************************************************************************************************
Private Const m_def_Enabled = True
Private Const m_def_ForeColor = &HFF00&
Private Const m_def_GradientEndColor = &HFF&
Private Const m_def_GradientMidColor = &HFFFF&
Private Const m_def_GradientStartColor = &HFF00&
Private Const m_def_Max = 100
Private Const m_def_Min = 0
Private Const m_def_Orientation = 0
Private Const m_def_Segmented = True
Private Const m_def_SegmentSize = 3
Private Const m_def_UseGradient = True
Private Const m_def_Value = 0

'**************************************************************************************************
' Fader Property Variables
'**************************************************************************************************
Private m_ForeColor As OLE_COLOR
Private m_Enabled As Boolean
Private m_GradientEndColor As OLE_COLOR
Private m_GradientMidColor As OLE_COLOR
Private m_GradientStartColor As OLE_COLOR
Private m_Orientation As eOrientation
Private m_Segmented As Boolean
Private m_SegmentSize As Long
Private m_UseGradient As Boolean
Private m_Value As Long
'Property Variables:
Dim m_Max As Long



'****************************************************************************************
' Fader Properties Procedures
'****************************************************************************************
Public Property Get BackColor() As OLE_COLOR
     ' Return usercontrol's backcolor
     BackColor = UserControl.BackColor
End Property ' Get BackColor

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
     ' Set usercontrol's backcolor
     UserControl.BackColor() = New_BackColor
     ' Redraw
     Refresh
     ' broadcast change
     PropertyChanged "BackColor"
End Property ' Let BackColor

Public Property Get Enabled() As Boolean
     ' Return property value
     Enabled = m_Enabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     ' Set property variable
     m_Enabled = New_Enabled
     ' broadcast change
     PropertyChanged "Enabled"
End Property ' Let Enabled

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = UserControl.ForeColor()
End Property ' GetForeColor

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
     UserControl.ForeColor() = New_ForeColor
     UserControl.FillColor() = New_ForeColor
     m_ForeColor = New_ForeColor
     Refresh
     PropertyChanged "ForeColor"
End Property ' Let ForeColor

Public Property Get GradientEndColor() As OLE_COLOR
     GradientEndColor = m_GradientEndColor
End Property ' Get GradientEndColor

Public Property Let GradientEndColor(New_GradientEndColor As OLE_COLOR)
     m_GradientEndColor = New_GradientEndColor
     Refresh
     PropertyChanged "GradientEndColor"
End Property ' Let GradientEndColor

Public Property Get GradientMidColor() As OLE_COLOR
     GradientMidColor = m_GradientMidColor
End Property ' Get GradientMidColor

Public Property Let GradientMidColor(New_GradientMidColor As OLE_COLOR)
     m_GradientMidColor = New_GradientMidColor
     Refresh
     PropertyChanged "GradientMidColor"
End Property ' Let GradientMidColor

Public Property Get GradientStartColor() As OLE_COLOR
     GradientStartColor = m_GradientStartColor
End Property ' Get GradientStartColor

Public Property Let GradientStartColor(New_GradientStartColor As OLE_COLOR)
     m_GradientStartColor = New_GradientStartColor
     Refresh
     PropertyChanged "GradientStartColor"
End Property ' Let GradientStartColor

Public Property Get Orientation() As eOrientation
     ' Return property value
     Orientation = m_Orientation
End Property ' Get Orientation

Public Property Let Orientation(ByVal New_Orientation As eOrientation)
     ' Set property variable
     m_Orientation = New_Orientation
     ' Reset position
     ResetSlider
     ' Call resize event
     UserControl_Resize
     ' Broadcast change
     PropertyChanged "Orientation"
End Property ' Let Orientation

Public Property Get Segmented() As Boolean
     Segmented = m_Segmented
End Property ' Get Segmented

Public Property Let Segmented(New_Segmented As Boolean)
     m_Segmented = New_Segmented
     Refresh
     PropertyChanged "Segmented"
End Property ' Let Segmented

Public Property Get SegmentSize() As Long
     SegmentSize = m_SegmentSize
End Property ' Get SegmentSize

Public Property Let SegmentSize(New_SegmentSize As Long)
     ' validation
     If New_SegmentSize > 5 Then New_SegmentSize = 5
     If New_SegmentSize < 2 Then New_SegmentSize = 2
     m_SegmentSize = New_SegmentSize
     Refresh
     PropertyChanged "SegmentSize"
End Property ' Let SegmenetSize

Public Property Get SliderIcon() As Picture
     ' Return property value
     Set SliderIcon = Slider.Picture
End Property ' Get SliderIcon

Public Property Set SliderIcon(ByVal New_SliderIcon As Picture)
     ' Set property variable
     Set Slider.Picture = New_SliderIcon
     ' Call resize event
     UserControl_Resize
     ' Broadcast change
     PropertyChanged "SliderIcon"
End Property ' Set SliderIcon

Public Property Get UseGradient() As Boolean
     UseGradient = m_UseGradient
End Property ' Get UseGradient

Public Property Let UseGradient(New_UseGradient As Boolean)
     m_UseGradient = New_UseGradient
     Refresh
     PropertyChanged "UseGradient"
End Property ' Let UseGradient

Public Property Get Value() As Long
     ' Return property value
     Value = GetValue
End Property ' Get Value

Public Property Let Value(ByVal New_Value As Long)
     ' If New_Value is out of range exit without changes
     If (New_Value < m_def_Min Or New_Value > Max) Then Exit Property
     ' Set property variable
     m_Value = New_Value
     ' If the value has changed
     If (m_Value <> LastValue) Then
          If (Not SliderHooked) Then
               ' Set slider position
               Select Case m_Orientation
                    Case 0 ' Horizontal
                         Slider.Left = (New_Value - m_def_Min) * _
                              (ScaleWidth - Slider.Width) / Max
                    Case 1 ' Vertical
                         Slider.Top = ScaleHeight - Slider.Height - (New_Value - m_def_Min) * _
                              (ScaleHeight - Slider.Height) / Max
               End Select
          End If
          ' Redraw
          Refresh
          ' Update lastvalue variable
          LastValue = m_Value
          ' Raise event
          RaiseEvent ValueChanged
          ' Set tooltip text
          Extender.ToolTipText = CStr(m_Value) + Chr(37)
          ' If arrived at minimum value, raise event
          If (m_Value = Max) Then RaiseEvent ArrivedLast
          ' If arrived at maximum value, raise event
          If (m_Value = m_def_Min) Then RaiseEvent ArrivedFirst
          ' Broadcast change
          PropertyChanged "Value"
    End If
End Property ' Let Value

'****************************************************************************************
' Fader Private Methods
'****************************************************************************************
Private Function ColorDivide(ByVal dblNum As Double, ByVal dblDenom As Double) As Double
     ' Divides dblNum by dblDenom if dblDenom <> 0 to eliminate 'Division By Zero' error.
     If dblDenom = False Then Exit Function
     ColorDivide = dblNum / dblDenom
End Function ' ColorDivide

Private Sub DrawBar()
     Dim lLimit As Long
     Dim lLoop As Long
     Dim lRtn As Long
     Dim lIdx As Long
     Dim lCur As Long
     Dim lSegment As Long
     Dim lRed As Long
     Dim lGreen As Long
     Dim lBlue As Long
     Dim sglRed As Single
     Dim sglGreen As Single
     Dim sglBlue As Single
     Dim lFadeStart As Long
     Dim lFadeMid As Long
     Dim lFadeEnd As Long
     Dim m_level As Long
     Dim m_Colors As Variant
     Dim lCtr As Long
     Dim pt As POINTAPI
     ' convert value to level
     Select Case m_Orientation
          Case 0
               m_level = ScaleWidth * (m_Value / 100)
          Case 1
               m_level = ScaleHeight * (m_Value / 100)
     End Select
     ' set gradient colors
     If m_UseGradient Then
'          If m_Mute Then
'               ' fade the colors
'               lFadeStart = m_GradientStartColor And &H808080
'               lFadeMid = m_GradientMidColor And &H808080
'               lFadeEnd = m_GradientEndColor And &H808080
'               m_Colors = Array(lFadeStart, lFadeMid, lFadeEnd)
'          Else
               m_Colors = Array(m_GradientStartColor, m_GradientMidColor, m_GradientEndColor)
'          End If
     Else
'          If m_Mute Then
'               ' fade the colors
'               lFadeStart = UserControl.FillColor And &H808080
'               lFadeMid = UserControl.FillColor And &H808080
'               lFadeEnd = UserControl.FillColor And &H808080
'               m_Colors = Array(lFadeStart, lFadeMid, lFadeEnd)
'          Else
               m_Colors = Array(UserControl.FillColor, UserControl.FillColor, _
               UserControl.FillColor)
'          End If
     End If
     ' Get our segments sizes for each color
     If m_Orientation = 0 Then
          lLimit = ScaleWidth
     Else
          lLimit = ScaleHeight
     End If
     ' Get our segments sizes for each color
     lSegment = lLimit \ COL_CNT
     ' Dimension segment array and store segments
     If lSegment <= 2 Then
          ' Not enough  real estate to draw a proper gradient
          Exit Sub
     Else
          ' Size segments array to color count and store segment sizes
          ReDim sglSegments(1 To COL_CNT)
          ' Now determine if the color count divides
          ' evenly with the scale height.  If not add
          ' remainder to the first segment
          lRtn = lLimit Mod lSegment
          ' Loop through and add segments to segment array
          For lLoop = 1 To COL_CNT
               If lLoop = 1 Then
                    ' add remainder to first segment
                    sglSegments(lLoop) = lSegment + lRtn
               Else
                    sglSegments(lLoop) = lSegment
               End If
          Next
     End If
     ' Index for ColorArray tracking
     lCur = 1
     ' Dimension color array t
     ReDim lColorArray(1 To lLimit)
     ' Loop and blend the colors stopping at the next to last color
     ' always loop 1 less than color count
    For lLoop = 1 To COL_CNT
          'Extract Red, Blue and Green values from the loop - 1 color
          lRed = (m_Colors(lLoop - 1) And &HFF&)
          lGreen = (m_Colors(lLoop - 1) And &HFF00&) / &H100&
          lBlue = (m_Colors(lLoop - 1) And &HFF0000) / &H10000
          'Find the range of change from one color to another
          sglRed = ColorDivide(CSng((m_Colors(lLoop) And &HFF&) - lRed), _
               sglSegments(lLoop))
          sglGreen = ColorDivide(CSng(((m_Colors(lLoop) And &HFF00&) / &H100&) - lGreen), _
               sglSegments(lLoop))
          sglBlue = ColorDivide(CSng(((m_Colors(lLoop) And &HFF0000) / &H10000) - lBlue), _
               sglSegments(lLoop))
          ' Create the gradients and add colors to array
          For lIdx = 1 To sglSegments(lLoop)
               lColorArray(lCur) = CLng(lRed + (sglRed * lIdx)) + (CLng(lGreen + _
                    (sglGreen * lIdx)) * &H100&) + (CLng(lBlue + (sglBlue * lIdx)) * &H10000)
               lCur = lCur + 1
          Next
     Next     ' clean the canvas
     ' are we horizontal or vertical
     Select Case m_Orientation
          Case 0
               ' Loop through and output gradient stopping at level
               For lIdx = 1 To m_level
                    If m_Segmented Then
                         lCtr = lCtr + 1
                         If lCtr = m_SegmentSize Then
                              lColorArray(lIdx) = UserControl.BackColor
                              lCtr = 0
                         End If
                    End If
                    ' Set the forecolor so the right color line is drawn
                    UserControl.ForeColor = lColorArray(lIdx)
                    ' move the starting point of the line
                    MoveToEx hdc, lIdx, 2, pt
                    ' draw the line
                    LineTo hdc, lIdx, ScaleHeight - 2
               Next
          Case 1
               ' Loop through and output gradient stopping at level
               For lIdx = 1 To m_level 'Step 2
                    If m_Segmented Then
                         lCtr = lCtr + 1
                         If lCtr = m_SegmentSize Then
                              lColorArray(lIdx) = UserControl.BackColor
                              lCtr = 0
                         End If
                    End If
                    ' Set the forecolor so the right color line is drawn
                    UserControl.ForeColor = lColorArray(lIdx)
                    ' Move the starting point of the line
                    MoveToEx hdc, 2, ScaleHeight - lIdx, pt
                    ' draw the line
                    LineTo hdc, ScaleWidth - 2, ScaleHeight - lIdx
               Next
     End Select
End Sub ' DrawBar

Private Function GetValue() As Long
     Dim lValue As Long
     On Error Resume Next
     Select Case m_Orientation
          Case 0 ' Horizontal
               GetValue = Slider.Left / (ScaleWidth - Slider.Width) * AbsCount + m_def_Min
               Slider.Left = (GetValue - m_def_Min) * (ScaleWidth - Slider.Width) / AbsCount
          Case 1 ' Vertical
               GetValue = (ScaleHeight - Slider.Height - Slider.Top) / _
                    (ScaleHeight - Slider.Height) * AbsCount + m_def_Min
               Slider.Top = ScaleHeight - Slider.Height - (GetValue - m_def_Min) * _
                    (ScaleHeight - Slider.Height) / AbsCount
     End Select
     ' convert value
     lValue = &HFFFF& * (GetValue / Max)
End Function ' GetValue

Private Sub Refresh()
     ' Clear control
     Cls
     ' Draw meter
     DrawBar
     ' Paint slider
     PaintPicture Slider, Slider.Left, Slider.Top
End Sub ' Refresh

Private Sub ResetSlider()
     Select Case m_Orientation
          Case 0 ' Horizontal
               Slider.Move 0, 0
          Case 1 ' Vertical
               Slider.Move 0, ScaleHeight - Slider.Height
     End Select
End Sub ' ResetSlider

'****************************************************************************************
' Usercontrol Intrinsic Methods/Events
'****************************************************************************************
Private Sub UserControl_Initialize()
     ' Get twipsperpixel on the x axis
     tppX = Screen.TwipsPerPixelX
     ' Get twipsperpixel on the y axis
     tppY = Screen.TwipsPerPixelY
End Sub ' UserControl_Initialize

Private Sub UserControl_InitProperties()
     ' Set initial property  values
     m_Enabled = m_def_Enabled
     m_GradientEndColor = m_def_GradientEndColor
     m_GradientMidColor = m_def_GradientMidColor
     m_GradientStartColor = m_def_GradientStartColor
     m_Orientation = m_def_Orientation
'     m_Min = m_def_Min
'     m_Max = m_def_Max
     m_SegmentSize = m_def_SegmentSize
     m_Value = m_def_Value
     LastValue = m_Value
     ' Set position
     ResetSlider
    m_Max = m_def_Max
End Sub ' UserControl_InitProperties

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ' if control is active
     If (Me.Enabled) Then
          ' hook and move slider
          With Slider
               ' Hook slider and get offsets
               If (Button = vbLeftButton) Then
                    SliderHooked = True
                    ' Mouse over slider
                    If (X >= .Left And X < .Left + .Width And Y >= .Top And _
                         Y < .Top + .Height) Then
                         ' move slider pic
                         SliderOffset.X = X - .Left
                         SliderOffset.Y = Y - .Top
                    ' Mouse is over control but not over slider pic
                    Else
                         SliderOffset.X = .Width / 2
                         SliderOffset.Y = .Height / 2
                         UserControl_MouseMove Button, Shift, X, Y
                    End If
                    ' Raise the event
                    RaiseEvent MouseDown(Shift)
               Else
                    'If (Button = vbRightButton) Then _
                        Mute = Not (m_Mute)
               End If
          End With
     End If
End Sub ' UserControl_MouseDown

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ' If slider is clicked
     If (SliderHooked) Then
          ' Check min/max limits
          With Slider
               Select Case m_Orientation
                    Case 0 ' Horizontal
                         If (X - SliderOffset.X < 0) Then
                              .Left = 0
                         ElseIf (X - SliderOffset.X > ScaleWidth - .Width) Then
                              .Left = ScaleWidth - .Width
                         Else
                              .Left = X - SliderOffset.X
                         End If
                    Case 1 ' Vertical
                         If (Y - SliderOffset.Y < 0) Then
                              .Top = 0
                         ElseIf (Y - SliderOffset.Y > ScaleHeight - .Height) Then
                              .Top = ScaleHeight - .Height
                         Else
                              .Top = Y - SliderOffset.Y
                         End If
               End Select
          End With
          ' Get value from Slider position
          Value = GetValue
    End If
End Sub ' UserControl_MouseMove

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     ' Click event (If mouse over control area)
     If (X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight And _
          Button = vbLeftButton) Then RaiseEvent Click
     ' MouseUp event (Slider has been hooked)
     If (SliderHooked) Then RaiseEvent MouseUp(Shift)
     ' Unhook slider
     SliderHooked = False
End Sub ' UserControl_MouseUp

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
          UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
          Enabled = .ReadProperty("Enabled", m_def_Enabled)
          GradientEndColor = .ReadProperty("GradientEndColor", m_def_GradientEndColor)
          GradientMidColor = .ReadProperty("GradientMidColor", m_def_GradientMidColor)
          GradientStartColor = .ReadProperty("GradientStartColor", m_def_GradientStartColor)
          Orientation = .ReadProperty("Orientation", m_def_Orientation)
          Value = .ReadProperty("Value", m_def_Value)
          Segmented = .ReadProperty("Segmented", m_def_Segmented)
          SegmentSize = .ReadProperty("SegmentSize", m_def_SegmentSize)
          UseGradient = .ReadProperty("UseGradient", m_def_UseGradient)
          Set Slider.Picture = .ReadProperty("SliderIcon", Nothing)
          ' Set lastvalue = to value
          LastValue = m_Value
          ' set slider position
          Slider.Left = (m_Value - m_def_Min) * (ScaleWidth - Slider.Width) / AbsCount
          Slider.Top = (ScaleHeight - Slider.Height) - (m_Value - m_def_Min) * _
               (ScaleHeight - Slider.Height) / AbsCount
     End With
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     On Error Resume Next
     ' Resize control
     Select Case m_Orientation
          Case 0 ' Horizontal
               If (Width = 0) Then Width = (Slider.Width * tppX)
               Height = Slider.Height * tppY
               Slider.Top = 0
               Slider.Left = (m_Value - m_def_Min) * (ScaleWidth - Slider.Width) \ AbsCount
          Case 1 ' Vertical
               If (Height = 0) Then Height = Slider.Height * tppY
               Width = (Slider.Width) * tppX
               Slider.Left = 0
               Slider.Top = ScaleHeight - Slider.Height - (m_Value - m_def_Min) * _
                    (ScaleHeight - Slider.Height) \ AbsCount
     End Select
     ' Refresh control
     Refresh
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     ' Draw control
     Refresh
End Sub ' UserControl_Show

Private Sub UserControl_Terminate()
     '
End Sub ' UserControl_Terminate

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
          .WriteProperty "Enabled", m_Enabled, m_def_Enabled
          .WriteProperty "GradientEndColor", m_GradientEndColor, m_def_GradientEndColor
          .WriteProperty "GradientMidColor", m_GradientMidColor, m_def_GradientMidColor
          .WriteProperty "GradientStartColor", m_GradientStartColor, m_def_GradientStartColor
          .WriteProperty "Orientation", m_Orientation, m_def_Orientation
          .WriteProperty "Segmented", m_Segmented, m_def_Segmented
          .WriteProperty "SegmentSize", m_SegmentSize, m_def_SegmentSize
          .WriteProperty "SliderIcon", Slider.Picture, Nothing
          .WriteProperty "UseGradient", m_UseGradient, m_def_UseGradient
          .WriteProperty "Value", m_Value, m_def_Value
     End With
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
End Sub ' UserControl_WriteProperties
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Возвращает/Устанавливает цвет, используемый, чтобы заполнить формы, круги, и рамки."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

