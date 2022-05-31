VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
'*            Íàïèñàíî 16.08.2004 ãîäó (Team HomeWork)              *
'*                   e-mail: sne_pro@mail.ru                        *
'********************************************************************
'*       Î ñêîëüêî íàì îòêðûòèé ÷óäíûõ, ãîòîâèò þçåðà Âèíäà!        *
'********************************************************************

'********************************************************************
'                                                                   '
' Æóäêî èçâèíÿþñü ÷òî áåç êîìåíòàðèåâ, íó äà êòî çàõî÷åò, âñåãäà    '
' ðàçáåðåòñÿ :)                                                     '
' Òåì-áîëåå ÷òî ïî íàïèñàííîìó ýòî ñäåëàòü êóäà ïðîùå :)            '
'                                                                   '
'********************************************************************

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ÀÏÈ §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hmenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
'Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, ByVal hicon As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hicon As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Long
Private Declare Function ImageList_RemoveAll Lib "comctl32.dll" (ByVal himl As Long) As Long

Public Enum gbLVBorderStyle
    WS_EX_STATICEDGE = &H20000
    WS_EX_CLIENTEDGE = &H200&
End Enum
Public Enum gbLVStyle
    LVS_ALIGNLEFT = &H800               ' Items are left-aligned in icon and small icon view
    LVS_ALIGNTOP = &H0                  ' Items are aligned with the top of the list-view control in icon and small icon view
    LVS_AUTOARRANGE = &H100             ' Icons are automatically kept arranged in icon and small icon view
    LVS_EDITLABELS = &H200              ' Item text can be edited in place. The parent window must process the LVN_ENDLABELEDIT notification message
    LVS_ICON = &H0                      ' This style specifies icon view
    LVS_LIST = &H3                      ' This style specifies list view
    LVS_NOCOLUMNHEADER = &H4000         ' Column headers are not displayed in report view. By default, columns have headers in report view
    LVS_NOLABELWRAP = &H80              ' Item text is displayed on a single line in icon view. By default, item text may wrap in icon view
    LVS_NOSCROLL = &H2000               ' Scrolling is disabled
    LVS_NOSORTHEADER = &H8000           ' Column headers do not work like buttons. This style can be used if clicking a column header in report view does not carry out an action, such as sorting
    LVS_OWNERDATA = &H1000              ' This style specifies a virtual list-view control. For more information about this list control style, see List-View Controls
    LVS_OWNERDRAWFIXED = &H400          ' The owner window can paint items in report view. The list-view control sends a WM_DRAWITEM message to paint each item; it does not send separate messages for each subitem. The itemData member of the DRAWITEMSTRUCT structure contains the item data for the specified list-view item
    LVS_REPORT = &H1                    ' This style specifies report view
    LVS_SHAREIMAGELISTS = &H40          ' The image list will not be deleted when the control is destroyed. This style enables the use of the same image lists with multiple list-view controls
    LVS_SHOWSELALWAYS = &H8             ' The selection, if any, is always shown, even if the control does not have the focus
    LVS_SINGLESEL = &H4                 ' Only one item at a time can be selected. By default, multiple items may be selected
    LVS_SMALLICON = &H2                 ' This style specifies small icon view
    LVS_SORTASCENDING = &H10            ' Item indexes are sorted based on item text in ascending order
    LVS_SORTDESCENDING = &H20           ' Item indexes are sorted based on item text in descending order
End Enum
Public Enum gbLVStyleEx
    LVS_EX_CHECKBOXES = &H4             ' Enables check boxes for items in a list-view control
    LVS_EX_FLATSB = &H100               ' Enables flat scroll bars in the list view
    LVS_EX_FULLROWSELECT = &H20         ' When an item is selected, the item and all its subitems are highlighted
    LVS_EX_GRIDLINES = &H1              ' Displays gridlines around items and subitems
    LVS_EX_HEADERDRAGDROP = &H10        ' Enables drag-and-drop reordering of columns in a list-view control
    LVS_EX_LABELTIP = &H4000            ' When a list-view control uses the LVS_EX_INFOTIP style, the LVN_GETINFOTIP notification message is sent to the parent window before displaying an item's ToolTip
    LVS_EX_INFOTIP = &H400              ' If a partially hidden label in any list-view mode lacks ToolTip text, the list-view control will unfold the label
    LVS_EX_MULTIWORKAREAS = &H2000      ' If the list-view control has the LVS_AUTOARRANGE style, the control will not autoarrange its icons until one or more work areas are defined (see LVM_SETWORKAREAS)
    LVS_EX_ONECLICKACTIVATE = &H40      ' The list-view control sends an LVN_ITEMACTIVATE notification message to the parent window when the user clicks an item
    LVS_EX_REGIONAL = &H200             ' Sets the list-view window region to include only the item icons and text using SetWindowRgn
    LVS_EX_SUBITEMIMAGES = &H2          ' Allows images to be displayed for subitems
    LVS_EX_TRACKSELECT = &H8            ' Enables hot-track selection in a list-view control
    LVS_EX_TWOCLICKACTIVATE = &H80      ' The list-view control sends an LVN_ITEMACTIVATE notification message to the parent window when the user double-clicks an item
    LVS_EX_UNDERLINECOLD = &H1000       ' Causes those non-hot items that may be activated to be displayed with underlined text
    LVS_EX_UNDERLINEHOT = &H800         ' Causes those hot items that may be activated to be displayed with underlined text
    LVS_EX_BORDERSELECT = &H8000
End Enum
Public Enum gbLVArrange
    LVA_DEFAULT = &H0                   ' Aligns items according to the list-view control's current alignment styles (the default value).
    LVA_ALIGNLEFT = &H1                 ' Aligns items along the left edge of the window
    LVA_ALIGNTOP = &H2                  ' Aligns items along the top edge of the window
    LVA_SNAPTOGRID = &H5                ' Snaps all icons to the nearest grid position
End Enum
Public Enum gbLVGroupAligin
    LVGA_FOOTER_LEFT = &H8
    LVGA_FOOTER_RIGHT = &H20
    LVGA_FOOTER_CENTER = &H10
End Enum
Public Enum gbLVGroupState
    LVGS_NORMAL = &H0
    LVGS_COLLAPSED = &H1
    LVGS_HIDDEN = &H2
End Enum
Public Enum gbLVImgLstType
    LVSIL_NORMAL = &H0                  ' Image list with large icons
    LVSIL_SMALL = &H1                   ' Image list with small icons
    LVSIL_STATE = &H2                   ' Image list with state images
End Enum
Public Enum gbLVGetNextItem
    LVNI_ALL = &H0                      ' Searches for a subsequent item by index (the default value)
    LVNI_FOCUSED = &H1                  ' The item has the LVIS_FOCUSED state flag set
    LVNI_SELECTED = &H2                 ' The item has the LVIS_SELECTED state flag set
    LVNI_CUT = &H4                      ' The item has the LVIS_CUT state flag set
    LVNI_DROPHILITED = &H8              ' The item has the LVIS_DROPHILITED state flag set
    LVNI_ABOVE = &H100                  ' Searches for an item that is above the specified item
    LVNI_BELOW = &H200                  ' Searches for an item that is below the specified item
    LVNI_TOLEFT = &H400                 ' Searches for an item to the left of the specified item
    LVNI_TORIGHT = &H800                ' Searches for an item to the right of the specified item
End Enum
Public Enum gbLVSIValues
    LVIR_BOUNDS = &H0                   ' Returns the bounding rectangle of the entire item, including the icon and label
    LVIR_ICON = &H1                     ' Returns the bounding rectangle of the icon or small icon
    LVIR_LABEL = &H2                    ' Returns the bounding rectangle of the item text
    LVIR_SELECTBOUNDS = &H3             ' Returns the union of the LVIR_ICON and LVIR_LABEL rectangles, but excludes columns in report view
End Enum
Public Enum gbLVICFlag
    LVSICF_NOINVALIDATEALL = &H1        ' The list-view control will not repaint unless affected items are currently in view
    LVSICF_NOSCROLL = &H2               ' The list-view control will not change the scroll position when the item count changes
End Enum

Public Enum gbLVILColourDepth
    ILC_COLOR = &H0
    ILC_MASK = &H1
    ILC_COLOR4 = &H4
    ILC_COLOR8 = &H8
    ILC_COLOR16 = &H10
    ILC_COLOR24 = &H18
    ILC_COLOR32 = &H20
    ILC_PALETTE = &H800
End Enum

Public Enum gbLVHScrollEnum
    SB_LINELEFT = 0
    SB_LINERIGHT = 1
    SB_PAGELEFT = 2
    SB_PAGERIGHT = 3
    SB_THUMBPOSITION = 4
    SB_THUMBTRACK = 5
    SB_LEFT = 6
    SB_RIGHT = 7
    SB_ENDSCROLL = 8
End Enum
Public Enum gbLVVScrollEnum
    SB_LINEUP = 0
    SB_LINEDOWN = 1
    SB_PAGEUP = 2
    SB_PAGEDOWN = 3
    SB_THUMBPOSITION = 4
    SB_THUMBTRACK = 5
    SB_TOP = 6
    SB_BOTTOM = 7
    SB_ENDSCROLL = 8
End Enum

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Òèïû §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Type POINTAPI
    X               As Long
    Y               As Long
End Type
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Type LVITEM
    mask            As Long
    iItem           As Long
    iSubItem        As Long
    State           As Long
    stateMask       As Long
    pszText         As String
    cchTextMax      As Long
    iImage          As Long
    lParam          As Long
    iIndent         As Long
End Type
Private Type LVCOLUMN
   mask             As Long
   fmt              As Long
   cx               As Long
   pszText          As String
   cchTextMax       As Long
   iSubItem         As Long
   iImage           As Long
   iOrder           As Long
End Type
Private Type LVBKIMAGE
    ulFlags         As Long
    hbm             As Long
    pszImage        As String
    cchImageMax     As Long
    xOffsetPercent  As Long
    yOffsetPercent  As Long
End Type
Private Type LVTILEVIEWINFO
    cbSize          As Long
    dwMask          As Long
    dwFlags         As Long
    sizeTile        As POINTAPI         ' Size
    cLines          As Long
    rcLabelMargin   As RECT
End Type
Private Type LVHITTESTINFO
    pt              As POINTAPI
    flags           As Long
    iItem           As Long
    iSubItem        As Long
End Type
Private Type LVGROUP
    cbSize          As Long
    mask            As Long
    pszHeader       As Long
    cchHeader       As Long
    pszFooter       As Long
    cchFooter       As Long
    iGroupId        As Long
    stateMask       As Long
    State           As Long
    uAlign          As Long
End Type


' §§§§§§§§§§§§§§§§§§§§§§§§§§ Êîíñòàíòû §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Const cNull                 As Long = &H0

Private Const WS_VISIBLE            As Long = &H10000000
Private Const WS_CHILD              As Long = &H40000000

Private Const GWL_EXSTYLE           As Long = (-20)
Private Const GWL_STYLE             As Long = (-16)

Private Const WM_NOTIFY             As Long = &H4E

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ImageList §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Const SM_CXICON             As Long = 11
Private Const SM_CYICON             As Long = 12

Private Const SM_CYSMICON           As Long = 50
Private Const SM_CXSMICON           As Long = 49

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ListView §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Const LVIF_TEXT             As Long = &H1
Private Const LVIF_IMAGE            As Long = &H2
Private Const LVIF_PARAM            As Long = &H4
Private Const LVIF_STATE            As Long = &H8
Private Const LVIF_INDENT           As Long = &H10

Private Const LVCF_FMT              As Long = &H1
Private Const LVCF_WIDTH            As Long = &H2
Private Const LVCF_TEXT             As Long = &H4
Private Const LVCF_IMAGE            As Long = &H10
Private Const LVCFMT_IMAGE          As Long = &H800

Private Const LVIS_SELECTED         As Long = &H2
Private Const LVIS_FOCUSED          As Long = &H1
Private Const LVIS_CUT              As Long = &H4
Private Const LVIS_STATEIMAGEMASK   As Long = &HF000&

Private Const LVM_FIRST             As Long = &H1000
Private Const LVM_GETBKCOLOR        As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR        As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST      As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST      As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT      As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEM           As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEM           As Long = (LVM_FIRST + 6)
Private Const LVM_DELETEITEM        As Long = (LVM_FIRST + 8)
Private Const LVM_INSERTITEM        As Long = (LVM_FIRST + 7)
Private Const LVM_DELETEALLITEMS    As Long = (LVM_FIRST + 9)
Private Const LVM_GETNEXTITEM       As Long = (LVM_FIRST + 12)
Private Const LVM_GETITEMRECT       As Long = (LVM_FIRST + 14)
Private Const LVM_GETITEMPOSITION   As Long = (LVM_FIRST + 16)
Private Const LVM_GETSTRINGWIDTH    As Long = (LVM_FIRST + 17)
Private Const LVM_HITTEST           As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE     As Long = (LVM_FIRST + 19)
Private Const LVM_SCROLL            As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS       As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE           As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABEL         As Long = (LVM_FIRST + 23)
Private Const LVM_GETEDITCONTROL    As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMN         As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMN         As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMN      As Long = (LVM_FIRST + 27)
Private Const LVM_DELETECOLUMN      As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH    As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH    As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER         As Long = (LVM_FIRST + 31)
Private Const LVM_GETVIEWRECT       As Long = (LVM_FIRST + 34)
Private Const LVM_GETTEXTCOLOR      As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR      As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR    As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR    As Long = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX       As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE   As Long = (LVM_FIRST + 40)
Private Const LVM_GETORIGIN         As Long = (LVM_FIRST + 41)
Private Const LVM_SETITEMSTATE      As Long = (LVM_FIRST + 43)
Private Const LVM_UPDATE            As Long = (LVM_FIRST + 42)
Private Const LVM_GETITEMSTATE      As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT       As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXT       As Long = (LVM_FIRST + 46)
Private Const LVM_SETITEMCOUNT      As Long = (LVM_FIRST + 47)
Private Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Private Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Private Const LVM_GETITEMSPACING    As Long = (LVM_FIRST + 51)
Private Const LVM_SETICONSPACING    As Long = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETSUBITEMRECT    As Long = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST    As Long = (LVM_FIRST + 57)
Private Const LVM_SETHOTITEM        As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM        As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR      As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR      As Long = (LVM_FIRST + 63)
Private Const LVM_SETSELECTIONMARK  As Long = (LVM_FIRST + 67)
Private Const LVM_GETSELECTIONMARK  As Long = (LVM_FIRST + 66)
Private Const LVM_SETBKIMAGE        As Long = (LVM_FIRST + 68)
'Private Const LVM_GETBKIMAGE        As Long = (LVM_FIRST + 69)
Private Const LVM_SETHOVERTIME      As Long = (LVM_FIRST + 71)
Private Const LVM_GETHOVERTIME      As Long = (LVM_FIRST + 72)
Private Const LVM_GETNUMBEROFWORKAREAS As Long = (LVM_FIRST + 73)
Private Const LVM_GETTOOLTIPS       As Long = (LVM_FIRST + 78)
Private Const LVM_SETTOOLTIPS       As Long = (LVM_FIRST + 74)
Private Const LVM_INSERTGROUP       As Long = (LVM_FIRST + 145)
Private Const LVM_SETGROUPINFO      As Long = (LVM_FIRST + 147)
Private Const LVM_GETGROUPINFO      As Long = (LVM_FIRST + 149)
Private Const LVM_SETTILEVIEWINFO   As Long = (LVM_FIRST + 162)
Private Const LVM_GETTILEVIEWINFO   As Long = (LVM_FIRST + 163)

'Private Const LVBKIF_SOURCE_NONE    As Long = &H0
'Private Const LVBKIF_SOURCE_HBITMAP As Long = &H1
Private Const LVBKIF_SOURCE_URL     As Long = &H2
'Private Const LVBKIF_SOURCE_MASK    As Long = &H3
'Private Const LVBKIF_STYLE_TILE     As Long = &H10

Private Const LVTVIM_COLUMNS        As Long = &H2

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Notifications §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Const NM_FIRST              As Long = 0
Private Const LVN_FIRST             As Long = (-100)
Private Const LVN_BEGINDRAG         As Long = (LVN_FIRST - 9)
Private Const LVN_BEGINLABELEDIT    As Long = (LVN_FIRST - 5)
Private Const LVN_BEGINRDRAG        As Long = (LVN_FIRST - 11)
Private Const LVN_COLUMNCLICK       As Long = (LVN_FIRST - 8)
Private Const LVN_DELETEALLITEMS    As Long = (LVN_FIRST - 4)
Private Const LVN_DELETEITEM        As Long = (LVN_FIRST - 3)
Private Const LVN_ENDLABELEDIT      As Long = (LVN_FIRST - 6)
Private Const LVN_GETDISPINFO       As Long = (LVN_FIRST - 50)
Private Const LVN_GETINFOTIP        As Long = (LVN_FIRST - 57)
Private Const LVN_HOTTRACK          As Long = (LVN_FIRST - 21)
Private Const LVN_INSERTITEM        As Long = (LVN_FIRST - 2)
Private Const LVN_ITEMACTIVATE      As Long = (LVN_FIRST - 14)
Private Const LVN_ITEMCHANGED       As Long = (LVN_FIRST - 1)
Private Const LVN_ITEMCHANGING      As Long = (LVN_FIRST - 0)
Private Const LVN_KEYDOWN           As Long = (LVN_FIRST - 55)

Private Const NM_CLICK              As Long = (NM_FIRST - 2)
Private Const NM_CUSTOMDRAW         As Long = (NM_FIRST - 12)
Private Const NM_DBLCLK             As Long = (NM_FIRST - 3)
Private Const NM_HOVER              As Long = (NM_FIRST - 13)
Private Const NM_KILLFOCUS          As Long = (NM_FIRST - 8)
Private Const NM_RCLICK             As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK            As Long = (NM_FIRST - 6)
Private Const NM_RELEASEDCAPTURE    As Long = (NM_FIRST - 16)
Private Const NM_RETURN             As Long = (NM_FIRST - 4)
Private Const NM_SETFOCUS           As Long = (NM_FIRST - 7)

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Header §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Const HDM_FIRST             As Long = &H1200
Private Const HDM_GETITEMCOUNT      As Long = (HDM_FIRST + 0)

Private Const HDS_BUTTONS           As Long = &H2
Private Const HDS_HOTTRACK          As Long = &H4

Private Const LVGF_HEADER           As Long = &H1
Private Const LVGF_STATE            As Long = &H4
Private Const LVGF_ALIGN            As Long = &H8
Private Const LVGF_GROUPID          As Long = &H10

Private Const LVGA_HEADER_LEFT      As Long = &H1

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Ñàáêëàññèðîâùèê §§§§§§§§§§§§§§§§§§§§§§§§§§

Const AsmMain As String = "558BEC83C4FC8D45FC50FF7514FF7510FF750CFF75086800000000B800000000FFD08B45FCC9C21000"

Private OldCtl      As Long
Private OldPrnt     As Long

Private ASM_Ctl()   As Byte
Private ASM_Prnt()  As Byte

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Events §§§§§§§§§§§§§§§§§§§§§§§§§§

Public Event Click(ByVal iItem As Long, ByVal Button As MouseButtonConstants)
Public Event DblClick(ByVal iItem As Long, ByVal Button As MouseButtonConstants)
Public Event Hover()
Public Event GotFocus()
Public Event LostFocus()
Public Event ReturnKey()
Public Event ReleaseCapture()
Public Event DeleteAllItems()
Public Event DeleteItem(ByVal iIndex As Long)
Public Event CustomDraw(ByVal lParam As Long)
Public Event ItemInsert(ByVal iIndex As Long)
Public Event ItemChanged(ByVal iIndex As Long)
Public Event ItemActivate(ByVal iIndex As Long)
Public Event ColumnClick(ByVal iColumn As Long)
Public Event GetInfoTip(ByVal iItem As Long, ByRef sText As String)
Public Event KeyDown(ByVal KeyCode As Long, ByVal Shift As Integer)
Public Event ItemChanging(ByVal iIndex As Long, ByRef Cancel As Long)
Public Event LabelStartEdit(ByVal iItem As Long, ByRef Cancel As Long, ByVal lParam As Long)
Public Event HotTrackSelect(ByVal iItem As Long, ByVal iSubItem As Long, ByRef Cancel As Long)
Public Event DragBegin(ByVal iItem As Long, ByVal iSubItem As Long, ByVal Button As MouseButtonConstants)
Public Event LabelEndEdit(ByVal iItem As Long, ByRef sNewText As String, ByRef Cancel As Long, ByVal lParam As Long)

Public Event HScroll(ByVal cPos As Long, flag As gbLVHScrollEnum)
Public Event VScroll(ByVal cPos As Long, flag As gbLVVScrollEnum)

Public Event MouseDown(ByVal Button As MouseButtonConstants, ByVal X As Long, ByVal Y As Long, ByVal Shift As Integer)
Public Event MouseMove(ByVal X As Long, ByVal Y As Long, ByVal Shift As Integer)
Public Event MouseUp(ByVal Button As MouseButtonConstants, ByVal X As Long, ByVal Y As Long, ByVal Shift As Integer)

Public Event KeyUp(ByVal KeyCode As Long, ByVal Shift As Integer)
Public Event KeyPress(KeyAscii As Long)

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Varisants §§§§§§§§§§§§§§§§§§§§§§§§§§

Private hListView   As Long     ' Õýíäë êîíòðîëà
Private hLVParent   As Long     ' Õýíäë ðîäèòåëüñêîãî îêíà êîíòðîëà
Private m_AtIntIL   As Boolean  ' Èñïîëüçóåòñÿ-ëè àëüòåðíàòèâíûé ÈìýéäæËèñò

Public Function zLVCallBack(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case &H200 To &H209
            Dim TMP As POINTAPI

            Call GetCursorPos(TMP)
            Call ScreenToClient(hListView, TMP)
    End Select

    Select Case uMsg
        Case Is = &H114: RaiseEvent HScroll(HiWord(wParam), LoWord(wParam)) ' WM_HSCROLL
        Case Is = &H115: RaiseEvent VScroll(HiWord(wParam), LoWord(wParam)) ' WM_VSCROLL

'        Case Is = &H100: RaiseEvent KeyDown(wParam, giGetShiftState)        ' WM_KEYDOWN
        Case Is = &H101: RaiseEvent KeyUp(wParam, giGetShiftState)          ' WM_KEYUP
        Case Is = &H102: RaiseEvent KeyPress(wParam)                        ' WM_CHAR

        Case Is = &H200: RaiseEvent MouseMove(TMP.X, TMP.Y, giGetShiftState)

        Case Is = &H201: RaiseEvent MouseDown(vbLeftButton, TMP.X, TMP.Y, giGetShiftState)
        Case Is = &H204: RaiseEvent MouseDown(vbRightButton, TMP.X, TMP.Y, giGetShiftState)
        Case Is = &H207: RaiseEvent MouseDown(vbMiddleButton, TMP.X, TMP.Y, giGetShiftState)

        Case Is = &H202: RaiseEvent MouseUp(vbLeftButton, TMP.X, TMP.Y, giGetShiftState)
        Case Is = &H205: RaiseEvent MouseUp(vbRightButton, TMP.X, TMP.Y, giGetShiftState)
        Case Is = &H208: RaiseEvent MouseUp(vbMiddleButton, TMP.X, TMP.Y, giGetShiftState)
    End Select

    zLVCallBack = CallWindowProc(OldCtl, hwnd, uMsg, wParam, lParam)
End Function

Public Function zLVpCallBack(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lng As Long, lng1 As Long, lng2 As Long, pt As POINTAPI, _
            tStr As String

    If uMsg = WM_NOTIFY Then
        Call CopyMemory(lng, ByVal lParam + 8&, 4&)             ' Âûäåëåíèå èç NMHDR ïîëÿ code (lParam - óêàçàòåëü íà ñòðóêòóðó)

        Select Case lng
            Case Is = LVN_BEGINDRAG, LVN_BEGINRDRAG
                Call CopyMemory(lng1, ByVal lParam + 12, 4&)    ' iItem
                Call CopyMemory(lng2, ByVal lParam + 16, 4&)    ' iSubItem
                RaiseEvent DragBegin(lng1, lng2, IIf(lng = LVN_BEGINDRAG, vbLeftButton, vbRightButton))

            Case Is = LVN_BEGINLABELEDIT
                Call CopyMemory(lng1, ByVal lParam + 16&, 4&)   ' iItem
                Call CopyMemory(lng2, ByVal lParam + 44&, 4&)   ' lParam
                RaiseEvent LabelStartEdit(lng1, zLVpCallBack, lng2)

            Case Is = LVN_COLUMNCLICK
                Call CopyMemory(lng1, ByVal lParam + 16&, 4&)
                RaiseEvent ColumnClick(lng1)

            Case Is = LVN_DELETEALLITEMS
                RaiseEvent DeleteAllItems

            Case Is = LVN_DELETEITEM
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)
                RaiseEvent DeleteItem(lng1)

            Case Is = LVN_ENDLABELEDIT
                'LabelEndEdit
                Call CopyMemory(lng1, ByVal lParam + 16&, 4&)   ' iItem (ïîëó÷åíèå èç ñòðóêòóðû)
                Call CopyMemory(lng2, ByVal lParam + 32&, 4&)   ' pszText
                Call CopyMemory(pt.X, ByVal lParam + 36&, 4&)   ' cchTextMax
                Call CopyMemory(pt.Y, ByVal lParam + 44&, 4&)   ' lParam

                If Not lng2 = cNull Then                        ' Âðîäå êàê íàñèëüñòâåííîå èçìåíåíèå òåêñòà
                    tStr = String$(pt.X, cNull)
                    Call CopyMemory(ByVal tStr, ByVal lng2, Len(tStr))
                End If

                RaiseEvent LabelEndEdit(lng1, tStr, zLVpCallBack, pt.Y)

                If Not lng2 = cNull Then                        ' Âðîäå êàê íàñèëüñòâåííîå èçìåíåíèå òåêñòà
                    tStr = tStr & vbNullChar
                    Call CopyMemory(ByVal lng2, ByVal tStr, Len(tStr))
                    Call CopyMemory(ByVal lParam + 36&, Len(tStr), 4&)
                End If
                zLVpCallBack = (zLVpCallBack = cNull)

            Case Is = LVN_GETDISPINFO
            Case Is = LVN_GETINFOTIP
                Call CopyMemory(lng1, ByVal lParam + 24&, 4&)   ' iItem
                Call CopyMemory(lng2, ByVal lParam + 16&, 4&)   ' pszText
                Call CopyMemory(pt.X, ByVal lParam + 20&, 4&)   ' cchTextMax

                If Not lng2 = cNull Then
                    tStr = String$(pt.X, cNull)
                    Call CopyMemory(ByVal tStr, ByVal lng2, pt.X)
                End If

                RaiseEvent GetInfoTip(lng1, tStr)

                If Not lng2 = cNull Then
                    tStr = tStr & vbNullChar
                    Call CopyMemory(ByVal lng2, ByVal tStr, Len(tStr))
                    Call CopyMemory(ByVal lParam + 20&, Len(tStr), 4&)
                End If

            Case Is = LVN_HOTTRACK
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)   ' iItem
                Call CopyMemory(lng2, ByVal lParam + 16&, 4&)   ' iSubItem

                RaiseEvent HotTrackSelect(lng1, lng2, zLVpCallBack)

            Case Is = LVN_INSERTITEM
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)   ' iItem
                RaiseEvent ItemInsert(lng1)

            Case Is = LVN_ITEMACTIVATE
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)   ' iItem
                RaiseEvent ItemActivate(lng1)

            Case Is = LVN_ITEMCHANGED
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)   ' iItem
                RaiseEvent ItemChanged(lng1)

            Case Is = LVN_ITEMCHANGING
                Call CopyMemory(lng1, ByVal lParam + 12&, 4&)   ' iItem
                RaiseEvent ItemChanging(lng1, zLVpCallBack)

            Case Is = LVN_KEYDOWN
                Call CopyMemory(lng1, ByVal lParam + 12&, 2&)   ' wKeyCode
                RaiseEvent KeyDown(lng1, giGetShiftState)
'                RaiseEvent KeyDown(wParam, giGetShiftState) 'Ýòîò wParam âðîäå âñåãäà 0 ðàâåí!!

            Case Is = NM_CLICK, NM_RCLICK
                Call CopyMemory(lng1, ByVal lParam + 12&, 2&)   ' iItem
                RaiseEvent Click(lng1, IIf(lng = NM_CLICK, vbLeftButton, vbRightButton))

            Case Is = NM_CUSTOMDRAW
                RaiseEvent CustomDraw(lParam)                   ' Êîìó íàäî, ñàìè âñå èçó÷àò...

            Case Is = NM_DBLCLK, NM_RDBLCLK
                Call CopyMemory(lng1, ByVal lParam + 12&, 2&)   ' iItem
                RaiseEvent DblClick(lng1, IIf(lng = NM_DBLCLK, vbLeftButton, vbRightButton))

            Case Is = NM_HOVER
                RaiseEvent Hover

            Case Is = NM_KILLFOCUS
                RaiseEvent LostFocus

            Case Is = NM_RELEASEDCAPTURE
                RaiseEvent ReleaseCapture

            Case Is = NM_RETURN
                RaiseEvent ReturnKey

            Case Is = NM_SETFOCUS
                RaiseEvent GotFocus

        End Select

        Exit Function
    End If
    
    zLVpCallBack = CallWindowProc(OldPrnt, hwnd, uMsg, wParam, lParam)
Exit Function

cb: Call CallWindowProc(OldPrnt, hwnd, uMsg, wParam, lParam)
End Function

Public Function Create(ByVal hParent As Long, _
                       ByVal dwStyle As gbLVStyle, _
                       ByVal nLeft As Long, _
                       ByVal nTop As Long, _
                       ByVal nWidth As Long, _
                       ByVal nHeight As Long, _
              Optional ByVal bInitImageList As Boolean = True, _
              Optional ByVal Border As gbLVBorderStyle) As Long     ' Ñîáñòâåííî ñîçäàíèå...

    Call InitCommonControls

    hLVParent = CreateWindowEx(cNull, "STATIC", "(c) Team HW (sne)", WS_VISIBLE Or WS_CHILD, nLeft, nTop, nWidth, nHeight, hParent, 0&, App.hInstance, ByVal 0&)
    hListView = CreateWindowEx(Border, "SysListView32", vbNullString, WS_VISIBLE Or WS_CHILD Or dwStyle, 0&, 0&, nWidth, nHeight, hLVParent, cNull, App.hInstance, ByVal cNull)

    Create = hListView

    If bInitImageList Then Call ImgLst_Init                      ' Èíèöèàëèçàöèÿ ëèñòà
    m_AtIntIL = bInitImageList

    Call StartSubclass(ASM_Ctl, hListView, OldCtl, cNull)           ' Íà÷àëî ñàáêëàññèíãà êîíòðîëà
    Call StartSubclass(ASM_Prnt, hLVParent, OldPrnt, vbNull)        ' Íà÷àëî ñàáêëàññèíãà ðîäèòåëüñêîãî îêíà
End Function

Public Sub Destroy()                                                ' Îòêëþ÷åíèå è óäàëåíèå âñåãî è âñÿ...
    If OldCtl Then Call StopSabclass(hListView, OldCtl):    OldCtl = cNull
    If OldPrnt Then Call StopSabclass(hLVParent, OldPrnt):  OldPrnt = cNull

    If hListView Then Call DestroyWindow(hListView):        hListView = cNull
    If hLVParent Then Call DestroyWindow(hLVParent):        hLVParent = cNull

    If m_AtIntIL Then Call ImgLst_Destroy
End Sub

Public Sub AddItem(ByVal sText As String, _
          Optional ByVal lIndex As Long = &HFFFF, _
          Optional ByVal iIcon As Long = cNull, _
          Optional ByVal iIndent As Long = cNull, _
          Optional ByVal lItemData As Long = cNull)

    Dim tLV As LVITEM

    If lIndex = &HFFFF Then lIndex = Count

    tLV.pszText = sText & vbNullChar
    tLV.cchTextMax = Len(tLV.pszText)
    tLV.iImage = iIcon
    tLV.iIndent = iIndent
    tLV.lParam = lItemData
    tLV.iItem = lIndex
    tLV.mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_INDENT Or IIf(iIcon = &HFFFF, cNull, LVIF_IMAGE)

    Call SendMessage(hListView, LVM_INSERTITEM, 0, tLV)
End Sub

Public Sub AddColumn(ByVal Index As Long, _
            Optional ByVal sText As String = vbNullString, _
            Optional ByVal lWidth As Long = 100&)

    Dim lvc As LVCOLUMN

    With lvc
        .mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT

        .pszText = sText & vbNullChar
        .cchTextMax = Len(.pszText)

        .cx = lWidth
    End With

    Call SendMessage(hListView, LVM_INSERTCOLUMN, Index, lvc)
    Call Refresh
End Sub

Public Function AddGroup(ByVal sText As String, _
                Optional ByVal lIndex As Long = &HFFFF, _
                Optional ByVal lItemData As Long = cNull)

    Dim tLVG As LVGROUP
    Dim lR As Long
    Dim hMem As Long
    Dim lPtrMem As Long
    Dim b() As Byte

    tLVG.cbSize = Len(tLVG)
    tLVG.iGroupId = lItemData
    tLVG.mask = LVGF_ALIGN Or LVGF_STATE Or LVGF_HEADER Or LVGF_GROUPID
    tLVG.State = LVGS_NORMAL
    tLVG.uAlign = LVGA_HEADER_LEFT
    tLVG.cchHeader = Len(sText)
    tLVG.pszHeader = StrPtr(sText)

    AddGroup = Not (SendMessage(hListView, LVM_INSERTGROUP, lIndex, tLVG) = &HFFFF)
End Function

' §§§§§§§§§§§§§§§§§§§§§§§§§§ /* ImageList §§§§§§§§§§§§§§§§§§§§§§§§§§

Public Property Let ImageList(ByVal ilType As gbLVImgLstType, ByVal hImgLst As Long)
    Call SendMessage(hListView, LVM_SETIMAGELIST, ilType, ByVal hImgLst)
End Property
Public Property Get ImageList(ByVal ilType As gbLVImgLstType) As Long
    ImageList = SendMessage(hListView, LVM_GETIMAGELIST, ilType, ByVal cNull)
End Property

Public Sub ImgLst_Init(Optional ByVal lColor As gbLVILColourDepth = ILC_COLOR24 Or ILC_MASK)
    Dim hLarge As Long, _
        hSmall As Long ', _
        hState As Long

    ' Create the full-sized icon image lists.
    hLarge = ImageList_Create(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), lColor, vbNull, vbNull)
    hSmall = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), lColor, vbNull, vbNull)
    'hState = ImageList_Create(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), ILC_MASK, vbNull, vbNull)

    ' Assign the image lists to the list-view control.
    ImageList(LVSIL_NORMAL) = hLarge
    ImageList(LVSIL_SMALL) = hSmall
    'ImageList(LVSIL_STATE) = hState
End Sub

Public Sub ImgLst_Destroy()
    Call ImageList_Destroy(ImageList(LVSIL_NORMAL))
    Call ImageList_Destroy(ImageList(LVSIL_SMALL))
End Sub

Public Sub ImgLst_AddIcon(Optional ByVal hIconLarge As Long = cNull, _
                          Optional ByVal hIconSmall As Long = cNull)

    Call ImgLst_ReplaceIcon(, hIconLarge, hIconSmall)
End Sub

Public Sub ImgLst_ReplaceIcon(Optional ByVal lIndex As Long = &HFFFF, _
                              Optional ByVal hIconLarge As Long = cNull, _
                              Optional ByVal hIconSmall As Long = cNull)

    If Not slVer(hIconLarge, hIconSmall) Then Exit Sub

    Call ImageList_ReplaceIcon(ImageList(LVSIL_NORMAL), lIndex, hIconLarge)
    Call ImageList_ReplaceIcon(ImageList(LVSIL_SMALL), lIndex, hIconSmall)
End Sub

Public Sub ImageListAddMasked(ByVal hImgLarge As Long, ByVal cMskLarge As Long, _
                              ByVal hImgSmall As Long, ByVal cMskSmall As Long)
    
    If Not slVer(hImgLarge, hImgSmall) Then Exit Sub

    Call ImageList_AddMasked(ImageList(LVSIL_NORMAL), hImgLarge, cMskLarge)
    Call ImageList_AddMasked(ImageList(LVSIL_SMALL), hImgSmall, cMskSmall)
End Sub

Public Property Get ImgLst_Count(ByVal iType As gbLVImgLstType) As Long
    ImgLst_Count = ImageList_GetImageCount(ImageList(iType))
End Property

Public Sub ImgLst_Remove(ByVal lIndex As Long)
    Call ImageList_Remove(ImageList(LVSIL_NORMAL), lIndex)
    Call ImageList_Remove(ImageList(LVSIL_SMALL), lIndex)
End Sub

Public Sub ImgLst_Clear()
    Call ImageList_RemoveAll(ImageList(LVSIL_NORMAL))
    Call ImageList_RemoveAll(ImageList(LVSIL_SMALL))
End Sub

Private Function slVer(ByRef hIconLarge As Long, ByRef hIconSmall As Long) As Boolean
    If hIconLarge = cNull And hIconSmall = cNull Then Exit Function

    If hIconLarge = cNull Then hIconLarge = hIconSmall
    If hIconSmall = cNull Then hIconSmall = hIconLarge

    slVer = True
End Function

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ImageList */ §§§§§§§§§§§§§§§§§§§§§§§§§§

Public Property Get HeaderhWnd() As Long
    HeaderhWnd = SendMessage(hListView, LVM_GETHEADER, cNull, ByVal cNull)
End Property

Public Property Get ToolTipCtl() As Long
    ToolTipCtl = SendMessage(hListView, LVM_GETTOOLTIPS, cNull, ByVal cNull)
End Property
Public Property Let ToolTipCtl(ByVal inValue As Long)
    Call SendMessage(hListView, LVM_SETTOOLTIPS, cNull, ByVal inValue)
End Property

Public Property Get ColumnsCount() As Long
    ColumnsCount = SendMessage(HeaderhWnd, HDM_GETITEMCOUNT, cNull, ByVal cNull)
End Property

Public Property Get ItemCountPerPage() As Long
   ItemCountPerPage = SendMessage(hListView, LVM_GETCOUNTPERPAGE, cNull, ByVal cNull)
End Property

Public Property Get ListItemsPerPage() As Long
   ListItemsPerPage = SendMessage(hListView, LVM_GETCOUNTPERPAGE, cNull, ByVal cNull)
End Property

Public Property Get WorkAreasCount() As Long
    Call SendMessage(hListView, LVM_GETNUMBEROFWORKAREAS, cNull, WorkAreasCount)
End Property

Public Property Get SelectedCount() As Long
    SelectedCount = SendMessage(hListView, LVM_GETSELECTEDCOUNT, cNull, ByVal cNull)
End Property

Public Property Let ColumnWidth(ByVal iColumn As Long, ByVal lValue As Long)
    Call SendMessage(hListView, LVM_SETCOLUMNWIDTH, iColumn, ByVal lValue)
End Property
Public Property Get ColumnWidth(ByVal iColumn As Long) As Long
    ColumnWidth = SendMessage(hListView, LVM_GETCOLUMNWIDTH, iColumn, ByVal cNull)
End Property

Public Property Get OriginX() As Long
    Dim tP As POINTAPI

    Call SendMessage(hListView, LVM_GETORIGIN, cNull, tP)
    OriginX = tP.X
End Property
Public Property Get OriginY() As Long
    Dim tP As POINTAPI

    Call SendMessage(hListView, LVM_GETORIGIN, cNull, tP)
    OriginY = tP.Y
End Property

Public Property Get ItemCaption(ByVal lIndex As Long) As String
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_TEXT, LVI)
    ItemCaption = LVI.pszText
End Property

Public Property Let ItemCaption(ByVal lIndex As Long, ByVal sCaption As String)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_TEXT, LVI)
    Call SetLVIStyle(lIndex, LVIF_TEXT, LVI, sCaption)
End Property

Public Property Get SubItemCaption(ByVal lIndex As Long, ByVal lSubItemIndex As Long) As String
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_TEXT, LVI, lSubItemIndex)
    SubItemCaption = LVI.pszText
End Property

Public Property Let SubItemCaption(ByVal lIndex As Long, ByVal lSubItemIndex As Long, ByVal sCaption As String)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_TEXT, LVI, lSubItemIndex)
    Call SetLVIStyle(lIndex, LVIF_TEXT, LVI, sCaption, lSubItemIndex)
End Property

Public Property Get SubItemIconIndex(ByVal lIndex As Long, ByVal lSubItemIndex As Long) As Long
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_IMAGE, LVI, lSubItemIndex)
    SubItemIconIndex = LVI.iImage
End Property

Public Property Let SubItemIconIndex(ByVal lIndex As Long, ByVal lSubItemIndex As Long, ByVal lImage As Long)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_IMAGE, LVI, lSubItemIndex)

    LVI.iImage = lImage
    Call SetLVIStyle(lIndex, LVIF_IMAGE, LVI, , lSubItemIndex)
End Property

Public Property Get ItemIconIndex(ByVal lIndex As Long) As Long
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_IMAGE, LVI)
    ItemIconIndex = LVI.iImage
End Property

Public Property Let ItemIconIndex(ByVal lIndex As Long, ByVal lIconIndex As Long)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_IMAGE, LVI)

    LVI.iImage = lIconIndex
    Call SetLVIStyle(lIndex, LVIF_IMAGE, LVI)
End Property

Public Property Get ItemIndent(ByVal lIndex As Long) As Long
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_INDENT, LVI)
    ItemIndent = LVI.iIndent
End Property

Public Property Let ItemIndent(ByVal lIndex As Long, ByVal lIndent As Long)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_INDENT, LVI)
   
    LVI.iIndent = lIndent
    Call SetLVIStyle(lIndex, LVIF_INDENT, LVI)
End Property

Public Property Get ItemData(ByVal lIndex As Long) As Long
    Dim LVI As LVITEM
   
    Call GetLVIStyle(lIndex, LVIF_PARAM, LVI)
    ItemData = LVI.lParam
End Property

Public Property Let ItemData(ByVal lIndex As Long, ByVal lItemData As Long)
    Dim LVI As LVITEM

    Call GetLVIStyle(lIndex, LVIF_PARAM, LVI)
   
    LVI.lParam = lItemData
    Call SetLVIStyle(lIndex, LVIF_PARAM, LVI)
End Property

Public Property Get ItemCut(ByVal lIndex As Long) As Boolean
    ItemCut = IsState(lIndex, LVIS_CUT)
End Property
Public Property Let ItemCut(ByVal lIndex As Long, ByVal bState As Boolean)
    Call SetState(lIndex, LVIS_CUT, bState)
End Property

Public Property Get ItemHot(ByVal lIndex As Long) As Boolean
    ItemHot = (lIndex = SendMessage(hListView, LVM_GETHOTITEM, cNull, ByVal cNull))
End Property
Public Property Let ItemHot(ByVal lIndex As Long, ByVal bState As Boolean)
    If bState Then Call SendMessage(hListView, LVM_SETHOTITEM, lIndex, ByVal cNull)
End Property

Public Property Get ItemFocused(ByVal lIndex As Long) As Boolean
    ItemFocused = IsState(lIndex, LVIS_FOCUSED)
End Property

Public Property Get ItemSelected(ByVal lIndex As Long) As Boolean
    ItemSelected = IsState(lIndex, LVIS_SELECTED)
End Property
Public Property Let ItemSelected(ByVal lIndex As Long, ByVal bState As Boolean)
    Call SetState(lIndex, LVIS_SELECTED, bState)
End Property

Public Property Get ItemChecked(ByVal lIndex As Long) As Boolean
    ItemChecked = ((SendMessage(hListView, LVM_GETITEMSTATE, lIndex, ByVal LVIS_STATEIMAGEMASK) And &H2000&) = &H2000&)
End Property
Public Property Let ItemChecked(ByVal lIndex As Long, ByVal bState As Boolean)
    Dim LVI As LVITEM

    LVI.iItem = lIndex
    LVI.mask = LVIF_STATE
    LVI.stateMask = &H3000&

    LVI.State = IIf(bState, &H2000&, &H1000&)
    '                       check,   uncheck
    Call SendMessage(hListView, LVM_SETITEMSTATE, lIndex, LVI)
End Property

Public Property Get ItemText(ByVal iSubItem As Long, ByVal lIndex As Long) As String
    Dim lresult As Long, _
        LVI As LVITEM

    LVI.cchTextMax = 260
    LVI.pszText = String$(LVI.cchTextMax, cNull)
    LVI.iSubItem = iSubItem

    lresult = SendMessage(hListView, LVM_GETITEMTEXT, lIndex, LVI)
    ItemText = Left$(LVI.pszText, lresult)
End Property
Public Property Let ItemText(ByVal iSubItem As Long, ByVal lIndex As Long, ByVal sText As String)
    Dim LVI As LVITEM

    LVI.iSubItem = iSubItem
    LVI.pszText = sText & vbNullChar
    LVI.cchTextMax = Len(LVI.pszText)

    Call SendMessage(hListView, LVM_SETITEMTEXT, lIndex, LVI)
End Property

Public Property Get SelectionMark() As Long
    SelectionMark = SendMessage(hListView, LVM_GETSELECTIONMARK, cNull, ByVal cNull)
End Property
Public Property Let SelectionMark(ByVal inValue As Long)
    Call SendMessage(hListView, LVM_SETSELECTIONMARK, cNull, ByVal inValue)
End Property

Public Property Get TopItem() As Long
   TopItem = SendMessage(hListView, LVM_GETTOPINDEX, cNull, ByVal cNull)
End Property

Friend Property Get SelectedItem() As Long
    Dim lFlags As Long

    lFlags = LVNI_SELECTED
    If GetFocus() = hListView Then lFlags = lFlags Or LVNI_FOCUSED

    SelectedItem = SendMessage(hListView, LVM_GETNEXTITEM, &HFFFF, ByVal lFlags)
End Property

Friend Property Let SelectedItem(ByVal rValue As Long)

    Dim ulvi As LVITEM

    ulvi.stateMask = LVIS_SELECTED
    
    '// Although passing -1 and State=LVIS_SELECTED to LVM_SETITEMSTATE will
    '// select all the items in the control, I opted to have -1 clear all the
    '// selections instead. To select everything on the control, use the
    '// SelectAll() method.
    If rValue < 0 Then
        ulvi.State = 0
    Else
        ulvi.State = LVIS_SELECTED
    End If
    
    Call SendMessage(hListView, LVM_SETITEMSTATE, rValue, ulvi)

End Property

Friend Property Get ListIndex() As Long
ListIndex = SendMessage(hListView, LVM_GETNEXTITEM, -1, ByVal LVIS_FOCUSED)
Rem Ñíà÷àëî áûëî òàê: SendMessage(..., ..., 0, ...)
End Property

Friend Property Let ListIndex(ByVal Index As Long)
Dim ulvi As LVITEM
ulvi.stateMask = LVIS_FOCUSED       'Ôëàã - Ôîêóñ

ulvi.State = LVIS_FOCUSED           'Ñòàâèì ôîêóñ íà ýëåìåíò
Call SendMessage(hListView, LVM_SETITEMSTATE, Index, ulvi)

ulvi.stateMask = LVIS_SELECTED      'Ôëàã - Âûäåëåíèå

ulvi.State = Not (LVIS_SELECTED)    'Óáèðàåì âûäåëåíèå ýëåìåíòîâ
Call SendMessage(hListView, LVM_SETITEMSTATE, -1, ulvi)

ulvi.State = LVIS_SELECTED          'Âûäåëÿåì òåêóùèé ýëåìåíò
Call SendMessage(hListView, LVM_SETITEMSTATE, Index, ulvi)

RaiseEvent Click(Index, 1)
End Property

Public Property Get ColumnImage(ByVal lColumn As Long) As Long
    Dim tHI As LVCOLUMN
   
    tHI.mask = LVCF_FMT
   
    If Not (GetHeaderIInfo(lColumn, tHI)) Then Exit Property
    If Not (tHI.fmt And LVCFMT_IMAGE) = LVCFMT_IMAGE Then Exit Property
         
    tHI.mask = LVCF_IMAGE
    ColumnImage = IIf(GetHeaderIInfo(lColumn, tHI), tHI.iImage, &HFFFF)
End Property

Public Property Let ColumnImage(ByVal lColumn As Long, ByVal lImage As Long)
    Dim tHI As LVCOLUMN

    tHI.mask = LVCF_FMT

    If (ColumnImage(lColumn) = lImage) Then Exit Property
    If Not GetHeaderIInfo(lColumn, tHI) Then Exit Property
         
    '    If (pbValidImage(lImage) < 0) Then
    '        tHI.fmt = tHI.fmt Or LVCFMT_IMAGE
    '        tHI.mask = tHI.mask Or LVCF_IMAGE
    '        tHI.iImage = lImage
    '    Else
    '        tHI.fmt = tHI.fmt And Not LVCFMT_IMAGE
    '    End If

    Call SetHeaderIInfo(lColumn, tHI)
End Property

Public Property Get TileViewItemLines() As Long
    Dim tLVI As LVTILEVIEWINFO
   
    tLVI.cbSize = Len(tLVI)
    tLVI.dwMask = LVTVIM_COLUMNS

    Call SendMessage(hListView, LVM_GETTILEVIEWINFO, cNull, tLVI)

    TileViewItemLines = tLVI.cLines
End Property

Public Property Let TileViewItemLines(ByVal lLines As Long)
    Dim tLVI As LVTILEVIEWINFO
   
    tLVI.cbSize = Len(tLVI)
    tLVI.dwMask = LVTVIM_COLUMNS

    Call SendMessage(hListView, LVM_GETTILEVIEWINFO, cNull, tLVI)
    tLVI.cLines = lLines

    Call SendMessage(hListView, LVM_SETTILEVIEWINFO, cNull, tLVI)
End Property

Public Property Let AutoArrange(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(False, LVS_AUTOARRANGE, cNull)
    Else
        Call SetLVStyle(False, cNull, LVS_AUTOARRANGE)
    End If
End Property

Public Property Let Arrange(ByVal newArrange As gbLVArrange)
    Call SendMessage(hListView, LVM_ARRANGE, newArrange, ByVal cNull)
End Property

Public Property Let LabelEdit(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(False, LVS_EDITLABELS, cNull)
    Else
        Call SetLVStyle(False, cNull, LVS_EDITLABELS)
    End If
End Property

Public Property Get ItemPosX(ByVal iIndex As Long) As Long
    Dim pnt As POINTAPI

    Call SendMessage(hListView, LVM_GETITEMPOSITION, iIndex, pnt)
    ItemPosX = pnt.X
End Property

Public Property Get ItemPosY(ByVal iIndex As Long) As Long
    Dim pnt As POINTAPI

    Call SendMessage(hListView, LVM_GETITEMPOSITION, iIndex, pnt)
    ItemPosY = pnt.Y
End Property

Public Sub SetItemPosition(ByVal iIndex As Long, ByVal X As Long, ByVal Y As Long)
    Dim pnt As POINTAPI
   
    pnt.X = X: pnt.Y = Y
    Call SendMessage(hListView, LVM_SETITEMPOSITION32, iIndex, pnt)
End Sub

Public Sub GetItemRect(ByVal iIndex As Long, flags As gbLVSIValues, _
              Optional ByRef lLeft As Long, _
              Optional ByRef lTop As Long, _
              Optional ByRef lRight As Long, _
              Optional ByVal lBottom As Long)

    Dim rct As RECT
    rct.Left = flags

    Call SendMessage(hListView, LVM_GETITEMRECT, iIndex, rct)

    lLeft = rct.Left:    lTop = rct.Top
    lBottom = rct.Bottom: lRight = rct.Right
End Sub

Public Sub GetSubItemRect(ByVal iIndex As Long, ByVal iSubItem As Long, ByVal flags As gbLVSIValues, _
                 Optional ByRef lLeft As Long, _
                 Optional ByRef lTop As Long, _
                 Optional ByRef lRight As Long, _
                 Optional ByVal lBottom As Long)
    Dim rct As RECT

    rct.Top = iSubItem
    rct.Left = flags

    Call SendMessage(hListView, LVM_GETSUBITEMRECT, iIndex, rct)

    lLeft = rct.Left:    lTop = rct.Top
    lBottom = rct.Bottom: lRight = rct.Right
End Sub

Public Sub GetViewRect(Optional ByRef lLeft As Long, _
                       Optional ByRef lTop As Long, _
                       Optional ByRef lRight As Long, _
                       Optional ByRef lBottom As Long)
                       
    Dim rct As RECT

    Call SendMessage(hListView, LVM_GETVIEWRECT, cNull, rct)

    lLeft = rct.Left:    lTop = rct.Top
    lBottom = rct.Bottom: lRight = rct.Right
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = TranslateColor(SendMessage(hListView, LVM_GETBKCOLOR, cNull, ByVal cNull))
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    If (oColor = &HFFFF) Then
        Call SendMessage(hListView, LVM_SETBKCOLOR, cNull, ByVal oColor)
        Call SendMessage(hListView, LVM_SETTEXTBKCOLOR, cNull, ByVal oColor)
    Else
        Call SendMessage(hListView, LVM_SETBKCOLOR, cNull, ByVal oColor)
        Call SendMessage(hListView, LVM_SETTEXTBKCOLOR, cNull, ByVal oColor)
    End If
End Property

Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = TranslateColor(SendMessage(hListView, LVM_GETTEXTBKCOLOR, cNull, ByVal cNull))
End Property
Public Property Let TextBackColor(ByVal oColor As OLE_COLOR)
    Call SendMessage(hListView, LVM_SETTEXTBKCOLOR, cNull, ByVal oColor)
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TranslateColor(SendMessage(hListView, LVM_GETTEXTCOLOR, cNull, ByVal cNull))
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
    Call SendMessage(hListView, LVM_SETTEXTCOLOR, cNull, ByVal oColor)
End Property

Public Property Let HeaderButtons(ByVal bState As Boolean)
    Dim hwnd As Long, tStyle As Long

    hwnd = HeaderhWnd
    If hwnd = cNull Then Exit Property
    
    tStyle = GetWindowLong(hwnd, GWL_STYLE)
    tStyle = IIf(bState, tStyle Or HDS_BUTTONS, tStyle And Not HDS_BUTTONS)
    Call SetWindowLong(hwnd, GWL_STYLE, tStyle)
End Property

Public Property Let HeaderTrackSelect(ByVal bState As Boolean)
    Dim tStyle As Long
    Dim hwnd As Long

    hwnd = HeaderhWnd
    If hwnd = cNull Then Exit Property

    tStyle = GetWindowLong(hwnd, GWL_STYLE)
    tStyle = IIf(bState, tStyle Or HDS_HOTTRACK, tStyle And Not HDS_HOTTRACK)
    Call SetWindowLong(hwnd, GWL_STYLE, tStyle)
End Property

Public Property Let MultiSelect(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(False, cNull, LVS_SINGLESEL)
    Else
        Call SetLVStyle(False, LVS_SINGLESEL, cNull)
    End If
End Property

Public Property Let HideSelection(ByVal bHide As Boolean)
    If bHide Then
        Call SetLVStyle(False, cNull, LVS_SHOWSELALWAYS)
    Else
        Call SetLVStyle(False, LVS_SHOWSELALWAYS, cNull)
    End If
End Property

Public Property Let SubItemImages(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_SUBITEMIMAGES, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_SUBITEMIMAGES)
    End If
End Property

Public Property Let FlatScrollBar(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_FLATSB, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_FLATSB)
    End If
End Property

Public Property Let GridLines(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_GRIDLINES, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_GRIDLINES)
    End If
End Property

Public Property Let OneClickActivate(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_ONECLICKACTIVATE, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_ONECLICKACTIVATE)
    End If
End Property

Public Property Let ItemBorderSelect(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_BORDERSELECT, cNull)
    Else
       Call SetLVStyle(True, cNull, LVS_EX_BORDERSELECT)
    End If
End Property

Public Property Let InfoTips(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_INFOTIP, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_INFOTIP)
    End If
End Property

Public Property Let LabelTips(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_LABELTIP, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_LABELTIP)
    End If
End Property

Public Property Let BackPicture(ByVal sURL As String)
    Dim tLBI As LVBKIMAGE

    tLBI.pszImage = sURL & vbNullChar
    tLBI.cchImageMax = Len(tLBI.pszImage)
    tLBI.ulFlags = LVBKIF_SOURCE_URL

    Call SendMessage(hListView, LVM_SETBKIMAGE, cNull, tLBI)
    Call SendMessage(hListView, LVM_SETTEXTBKCOLOR, cNull, &HFFFF)
End Property

Public Property Let CheckBoxes(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_CHECKBOXES, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_CHECKBOXES)
    End If
End Property

Public Property Let TrackSelect(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_TRACKSELECT, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_TRACKSELECT)
    End If
End Property

Public Property Let HeaderDragDrop(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_HEADERDRAGDROP, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_HEADERDRAGDROP)
    End If
End Property

Public Property Let FullRowSelect(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(True, LVS_EX_FULLROWSELECT, cNull)
    Else
        Call SetLVStyle(True, cNull, LVS_EX_FULLROWSELECT)
    End If
End Property

Public Property Let NoScrollBar(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(False, LVS_NOSCROLL, cNull)
    Else
        Call SetLVStyle(False, cNull, LVS_NOSCROLL)
    End If
End Property

Public Property Let NoColumnHeaders(ByVal bState As Boolean)
    If bState Then
        Call SetLVStyle(False, LVS_NOCOLUMNHEADER, cNull)
    Else
        Call SetLVStyle(False, cNull, LVS_NOCOLUMNHEADER)
    End If
End Property

Public Property Let GroupHeader(ByVal lPtrGroup As Long, ByVal sHeader As String)
    Dim tLVG As LVGROUP
    
    tLVG.cbSize = Len(tLVG)
    tLVG.mask = LVGF_HEADER
    tLVG.cchHeader = Len(sHeader)
    tLVG.pszHeader = StrPtr(sHeader)

    Call SendMessage(hListView, LVM_SETGROUPINFO, lPtrGroup, tLVG)
End Property

Public Property Get GroupAlign(ByVal lPtrGroup As Long) As gbLVGroupAligin
    Dim tLVG As LVGROUP

    tLVG.cbSize = Len(tLVG)
    tLVG.mask = LVGF_ALIGN

    Call SendMessage(hListView, LVM_GETGROUPINFO, lPtrGroup, tLVG)

    GroupAlign = tLVG.uAlign
End Property
Public Property Let GroupAlign(ByVal lPtrGroup As Long, ByVal eAlign As gbLVGroupAligin)
    Dim tLVG As LVGROUP

    tLVG.cbSize = Len(tLVG)
    tLVG.mask = LVGF_ALIGN
    tLVG.uAlign = eAlign

    Call SendMessage(hListView, LVM_GETGROUPINFO, lPtrGroup, tLVG)
End Property

Public Property Get GroupState(ByVal lPtrGroup As Long) As gbLVGroupState
    Dim tLVG As LVGROUP

    tLVG.cbSize = Len(tLVG)
    tLVG.mask = LVGF_STATE

    Call SendMessage(hListView, LVM_GETGROUPINFO, lPtrGroup, tLVG)

    GroupState = tLVG.State
End Property
Public Property Let GroupState(ByVal lPtrGroup As Long, ByVal eState As gbLVGroupState)
    Dim tLVG As LVGROUP

    tLVG.cbSize = Len(tLVG)
    tLVG.mask = LVGF_STATE
    tLVG.State = eState
   Call SendMessage(hListView, LVM_GETGROUPINFO, lPtrGroup, tLVG)
End Property

Public Property Let HotCursor(ByVal hCursor As Long)
    Call SendMessage(hListView, LVM_SETHOTCURSOR, cNull, ByVal hCursor)
End Property
Public Property Get HotCursor() As Long
    HotCursor = SendMessage(hListView, LVM_GETHOTCURSOR, cNull, ByVal cNull)
End Property

Public Property Get HoverTime() As Long
    HoverTime = SendMessage(hListView, LVM_GETHOVERTIME, cNull, ByVal cNull)
End Property
Public Property Let HoverTime(inValue As Long)
    Call SendMessage(hListView, LVM_SETHOVERTIME, cNull, ByVal inValue)
End Property

Public Property Get GetItemSpasing(ByVal fSmallIconView As Boolean) As Long
    GetItemSpasing = SendMessage(hListView, LVM_GETITEMSPACING, fSmallIconView, ByVal cNull)
End Property

Public Property Get hWnd_Parent() As Long
    hWnd_Parent = hLVParent
End Property

Public Property Get hwnd() As Long
    hwnd = hListView
End Property

Public Property Get hWnd_Edit() As Long
   hWnd_Edit = SendMessage(hListView, LVM_GETEDITCONTROL, cNull, ByVal cNull)
End Property

Public Function Count() As Long
    Count = SendMessage(hListView, LVM_GETITEMCOUNT, cNull, ByVal cNull)
End Function
Public Sub CountSet(ByVal flag As gbLVICFlag, ByVal inValue As Long)
    Call SendMessage(hListView, LVM_SETITEMCOUNT, inValue, ByVal flag)
End Sub

Public Function ItemNext(ByVal lStartIndex As Long, ByVal fParam As gbLVGetNextItem) As Long
    ItemNext = SendMessage(hListView, LVM_GETNEXTITEM, lStartIndex, ByVal fParam)
End Function

Public Sub StartEdit(ByVal lIndex As Long)
    Call SendMessage(hListView, LVM_EDITLABEL, &HFFFF, ByVal cNull)
    Call SendMessage(hListView, LVM_EDITLABEL, lIndex, ByVal cNull)
End Sub

Public Function GetStringWidth(ByVal inString As String) As Long
    GetStringWidth = SendMessage(hListView, LVM_GETSTRINGWIDTH, cNull, ByVal StrPtr(inString))
End Function

Public Sub EnsureVisible(ByVal lIndex As Long)
    Call SendMessage(hListView, LVM_ENSUREVISIBLE, lIndex - vbNull, ByVal cNull)
End Sub

Public Function HitTest(ByVal X As Long, ByVal Y As Long) As Long
    Dim tLVHI As LVHITTESTINFO

    tLVHI.pt.X = X
    tLVHI.pt.Y = Y

    HitTest = SendMessage(hListView, LVM_HITTEST, cNull, tLVHI)
End Function

Public Function SubItemHitTest(ByVal X As Long, ByVal Y As Long) As Long
    Dim tLVHI As LVHITTESTINFO

    tLVHI.pt.X = X
    tLVHI.pt.Y = Y

    SubItemHitTest = SendMessage(hListView, LVM_SUBITEMHITTEST, cNull, tLVHI)
End Function

Public Sub SetIconsPacing(ByVal cx As Integer, ByVal cy As Integer)
    Call SendMessage(hListView, LVM_SETICONSPACING, cNull, MakeDWord(cx, cy))
End Sub

Public Sub ItemPosition(ByVal iIndex As Long, ByRef X As Long, ByRef Y As Long)
    Dim pnt As POINTAPI

    Call SendMessage(hListView, LVM_GETITEMPOSITION, iIndex, pnt)
    X = pnt.X: Y = pnt.Y
End Sub


Public Sub Scroll(ByVal X As Long, ByVal Y As Long)
    Call SendMessage(hListView, LVM_SCROLL, X, ByVal Y)
End Sub

Public Sub Clear()
    Call SendMessage(hListView, LVM_DELETEALLITEMS, cNull, ByVal cNull)
End Sub

Public Sub RemoveItem(ByVal lIndex As Long)
    Call SendMessage(hListView, LVM_DELETEITEM, lIndex, ByVal cNull)
End Sub

Public Sub RemoveColumn(ByVal lIndex As Long)
    Call SendMessage(hListView, LVM_DELETECOLUMN, lIndex, cNull)
End Sub

Public Sub Refresh()
    Call SendMessage(hListView, LVM_UPDATE, cNull, ByVal cNull)
End Sub

Public Sub RedrawItem(ByVal iIndex As Long)
    Call SendMessage(hListView, LVM_REDRAWITEMS, iIndex, ByVal iIndex)
End Sub

Public Sub RedrawAllItems(ByVal iIndex As Long)
    Call SendMessage(hListView, LVM_REDRAWITEMS, 0, ByVal Count)
End Sub

Public Sub SetStyleEx(inValue As gbLVStyleEx)
    Call SetLVStyle(True, inValue, cNull)
End Sub
Public Sub DelStyleEx(inValue As gbLVStyleEx)
    Call SetLVStyle(True, cNull, inValue)
End Sub

Public Sub SetFocus()                                               ' Óñòàíîâêà ôîêóñà
    Call SetFocusAPI(hListView)
End Sub
                                                                    ' Ïåðåäâèæåíèå/èçìåíåíèå ðàçìåðîâ
Public Sub Move(ByVal nLeft As Long, ByVal nTop As Long, _
       Optional ByVal nWidth As Long, Optional ByVal nHeight As Long)
    Dim hRect As RECT

    Call GetClientRect(hLVParent, hRect)

    If nWidth = cNull Then nWidth = hRect.Right
    If nHeight = cNull Then nHeight = hRect.Bottom

    Call MoveWindow(hLVParent, nLeft, nTop, nWidth, nHeight, vbNull)
    Call MoveWindow(hListView, nLeft, nTop, nWidth, nHeight, vbNull)
End Sub

' §§§§§§§§§§§§§§§§§§§§§§§§§§  §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub SetLVStyle(ByVal SetStyleEx As Boolean, _
       Optional ByVal lAddStyle As Long = cNull, _
       Optional ByVal lDelStyle As Long = cNull)

    Dim tStyle As Long

    If SetStyleEx Then
        tStyle = SendMessage(hListView, LVM_SETEXTENDEDLISTVIEWSTYLE, cNull, ByVal cNull)
    Else
        tStyle = GetWindowLong(hListView, IIf(SetStyleEx, GWL_EXSTYLE, GWL_STYLE))
    End If

    tStyle = tStyle And Not lDelStyle
    tStyle = tStyle Or lAddStyle

    If SetStyleEx Then
        Call SendMessage(hListView, LVM_SETEXTENDEDLISTVIEWSTYLE, cNull, ByVal tStyle)
    Else
        Call SetWindowLong(hListView, IIf(SetStyleEx, GWL_EXSTYLE, GWL_STYLE), tStyle)
    End If
End Sub

Private Sub GetLVIStyle(ByVal lIndex As Long, ByVal lMask As Long, _
                        ByRef m_tLV As LVITEM, Optional ByVal lSubItem As Long = cNull)
    Dim lPos As Long
    
    m_tLV.mask = lMask
    
    If (lMask And LVIF_TEXT) = LVIF_TEXT Then
        m_tLV.pszText = String$(261, 0)
        m_tLV.cchTextMax = 260
    End If

    m_tLV.iItem = lIndex
    m_tLV.iSubItem = lSubItem

    Call SendMessage(hListView, LVM_GETITEM, IIf(lSubItem = cNull, cNull, lIndex), m_tLV)

    lPos = InStr(m_tLV.pszText, vbNullChar)
    
    If lPos > 0 Then m_tLV.pszText = Left$(m_tLV.pszText, lPos - vbNull)
    m_tLV.cchTextMax = Len(m_tLV.pszText)
End Sub

Private Sub SetLVIStyle(ByVal lIndex As Long, ByVal lMask As Long, m_tLV As LVITEM, _
               Optional ByVal sText As String, Optional ByVal lSubItem As Long = 0)

    If ((lMask And LVIF_TEXT) = LVIF_TEXT) Then
        m_tLV.pszText = String$(261, cNull)
        m_tLV.cchTextMax = 260
        
        LSet m_tLV.pszText = sText & vbNullChar
        m_tLV.cchTextMax = Len(sText) + vbNull
    End If

    m_tLV.iItem = lIndex
    m_tLV.iSubItem = lSubItem
    m_tLV.mask = lMask

    Call SendMessage(hListView, LVM_SETITEM, IIf(lSubItem = 0, cNull, lIndex), m_tLV)
    Call UpdateWindow(hListView)
End Sub

Private Function IsState(ByVal lIndex As Long, _
                         ByVal lValue As Long) As Long
    Dim m_tLV As LVITEM

    m_tLV.stateMask = lValue
    m_tLV.iItem = lIndex

    Call GetLVIStyle(lIndex, LVIF_STATE, m_tLV)

    IsState = CBool(m_tLV.State And lValue)
End Function

Private Sub SetState(ByVal lIndex As Long, _
                     ByVal lValue As Long, _
                     ByVal bState As Boolean)
    Dim m_tLV As LVITEM

    Call GetLVIStyle(lIndex, LVIF_STATE, m_tLV)
   
    m_tLV.State = IIf(bState, m_tLV.State Or lValue, m_tLV.State And (Not lValue))
    Call SetLVIStyle(lIndex, LVIF_STATE, m_tLV)
End Sub

Private Function GetHeaderIInfo(ByVal lCol As Long, tHI As LVCOLUMN) As Boolean
    lCol = lCol - vbNull
    tHI.iSubItem = lCol

    GetHeaderIInfo = Not (SendMessage(hListView, LVM_GETCOLUMN, lCol, tHI) = cNull)
End Function

Private Function SetHeaderIInfo(ByVal lCol As Long, tHI As LVCOLUMN) As Boolean
    lCol = lCol - vbNull
    tHI.iSubItem = lCol

    SetHeaderIInfo = Not (SendMessage(hListView, LVM_SETCOLUMN, lCol, tHI) = cNull)
End Function

Private Function TranslateColor(ByVal clr As Long, Optional hPal As Long = cNull) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = &HFFFF
End Function

Private Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Private Function LoWord(DWord As Long) As Integer
    LoWord = IIf(DWord And &H8000&, DWord Or &HFFFF0000, DWord And &HFFFF&)
End Function

Private Function HiWord(DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Private Function giGetShiftState() As Integer                       ' Äîâîëüíî ïîëåçíî...
    giGetShiftState = (-vbShiftMask * gbKeyIsPressed(vbKeyShift))
    giGetShiftState = giGetShiftState Or (-vbAltMask * gbKeyIsPressed(vbKeyMenu))
    giGetShiftState = giGetShiftState Or (-vbCtrlMask * gbKeyIsPressed(vbKeyControl))
End Function

Private Function gbKeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean
    If (GetAsyncKeyState(nVirtKeyCode) And &H8000&) = &H8000& Then gbKeyIsPressed = True
End Function

' §§§§§§§§§§§§§§§§§§§§§§§§§§ Âîò îí ïðîãðåññ â VB §§§§§§§§§§§§§§§§§§§§§§§§§§

Private Sub StartSubclass(ByRef ASM() As Byte, _
       ByVal hwnd As Long, _
       ByRef OldWndProc As Long, _
       Optional ByVal ProcNumber As Long)                 ' Ñàáêëàññèíã ñ ïîì. ASM (àâòîðà íå çíàþ...)
    Dim lng As Long, tPtr As Long
    lng = Len(AsmMain) \ 2&

    ReDim ASM(cNull To lng - vbNull)

    For lng = cNull To lng - vbNull
        ASM(lng) = Val("&H" & Mid$(AsmMain, (lng) * 2& + vbNull, 2&))
    Next

    Call CopyMemory(tPtr, ByVal ObjPtr(Me), 4&)
    Call CopyMemory(lng, ByVal tPtr + &H1C + (4& * ProcNumber), 4&)

    Call CopyMemory(ASM(23), ObjPtr(Me), 4&)
    Call CopyMemory(ASM(28), lng, 4&)

    OldWndProc = SetWindowLong(hwnd, &HFFFC, VarPtr(ASM(0)))
End Sub

Private Sub StopSabclass(ByVal hwnd As Long, ByVal OldWndProc As Long)
    Call SetWindowLong(hwnd, &HFFFC, OldWndProc)
End Sub

Private Sub Class_Terminate()
    Call Destroy
End Sub

'HELP
'ListIndex  Èíäåêñ òåêóùåãî ýëåìåíòà. Ïðè ïðèñâîåíèè íîâîãî çíà÷åíèÿ ñíèìàåò
'           âûäåëåíèå ñ äð. ýëåìåíòîâ è ñòàâèò ôîêóñ è âûäåëåíèå íà óêàçàííûé.
