VERSION 5.00
Begin VB.UserControl cToolbar 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ClipControls    =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   3855
   ToolboxBitmap   =   "cToolbar.ctx":0000
   Begin VB.Label lblInfo 
      Caption         =   "'Toolbar control'"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4275
   End
End
Attribute VB_Name = "cToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =========================================================================
' vbAccelerator Toolbar control v3.0
' Copyright © 1998-2000 Steve McMahon (steve@vbaccelerator.com)
'
' This is a complete form toolbar implementation designed
' for hosting in a vbAccelerator ReBar control.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================

Private Type TBADDBITMAP
    hInst As Long
    nID As Long
End Type

Private Type NMTOOLBAR_SHORT
    hdr As NMHDR
    iItem As Long
End Type

Private Type NMTOOLBAR
    hdr As NMHDR
    iItem As Long
    tbBtn As TBBUTTON
    cchText As Long
    lpszString As Long
End Type

Private Type NMTBHOTITEM
   hdr As NMHDR
   idOld As Long
   idNew As Long
   dwFlags As Long           '// HICF_*
End Type

Private Type NMTBCUSTOMDRAW
   nmcd As NMCUSTOMDRAW
   hbrMonoDither As Long
   hbrLines As Long
   hpenLines As Long
   clrText As Long
   clrMark As Long
   clrTextHighlight As Long
   clrBtnFace As Long
   clrBtnHighlight As Long
   clrHighlightHotTrack As Long
   rcText As RECT
   nStringBkMode As Long
   nHLStringBkMode As Long
End Type

' Toolbar button states:
Private Enum ectbButtonStates
   TBSTATE_CHECKED = &H1
   TBSTATE_PRESSED = &H2
   TBSTATE_ENABLED = &H4
   TBSTATE_WRAP = &H20
   TBSTATE_ELLIPSES = &H40
   TBSTATE_INDETERMINATE = &H10
   TBSTATE_HIDDEN = &H8
   TBSTATE_MARKED = &H80
End Enum


' Toolbar messages:

Private Const TB_SETSTATE = (WM_USER + 17)
Private Const TB_GETSTATE = (WM_USER + 18)

Private Const TB_ADDBITMAP = (WM_USER + 19)
Private Const TB_ADDBUTTONS = (WM_USER + 20)
Private Const TB_INSERTBUTTON = (WM_USER + 21)
Private Const TB_DELETEBUTTON = (WM_USER + 22)
Private Const TB_GETBUTTON = (WM_USER + 23)
Private Const TB_COMMANDTOINDEX = (WM_USER + 25)

Private Const TB_SAVERESTOREA = (WM_USER + 26)
Private Const TB_SAVERESTOREW = (WM_USER + 76)
Private Const TB_CUSTOMIZE = (WM_USER + 27)
Private Const TB_ADDSTRING = (WM_USER + 28)

Private Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
Private Const TB_SETBUTTONSIZE = (WM_USER + 31)
Private Const TB_SETBITMAPSIZE = (WM_USER + 32)
Private Const TB_AUTOSIZE = (WM_USER + 33)

Private Const TB_GETTOOLTIPS = (WM_USER + 35)
Private Const TB_SETTOOLTIPS = (WM_USER + 36)
Private Const TB_SETPARENT = (WM_USER + 37)
Private Const TB_SETROWS = (WM_USER + 39)
Private Const TB_GETROWS = (WM_USER + 40)
Private Const TB_SETCMDID = (WM_USER + 42)
Private Const TB_CHANGEBITMAP = (WM_USER + 43)
Private Const TB_GETBITMAP = (WM_USER + 44)
Private Const TB_GETBUTTONTEXTA = (WM_USER + 45)
Private Const TB_GETBUTTONTEXTW = (WM_USER + 75)

'#if (_WIN32_IE >= 0x0300)
Private Const TB_SETINDENT = (WM_USER + 47)
Private Const TB_SETIMAGELIST = (WM_USER + 48)
Private Const TB_GETIMAGELIST = (WM_USER + 49)
Private Const TB_LOADIMAGES = (WM_USER + 50)
Private Const TB_GETRECT = (WM_USER + 51)             '// wParam is the Cmd instead of index
Private Const TB_SETHOTIMAGELIST = (WM_USER + 52)
Private Const TB_GETHOTIMAGELIST = (WM_USER + 53)
Private Const TB_SETDISABLEDIMAGELIST = (WM_USER + 54)
Private Const TB_GETDISABLEDIMAGELIST = (WM_USER + 55)
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const TB_GETBUTTONSIZE = (WM_USER + 58)
Private Const TB_SETBUTTONWIDTH = (WM_USER + 59)
Private Const TB_SETMAXTEXTROWS = (WM_USER + 60)
Private Const TB_GETTEXTROWS = (WM_USER + 61)
'#endif

'#if (_WIN32_IE >= 0x0400)
Private Const TB_GETOBJECT = (WM_USER + 62)            '// wParam == IID, lParam void **ppv
Private Const TB_SETANCHORHIGHLIGHT = (WM_USER + 73)   '// wParam == TRUE/FALSE
Private Const TB_GETANCHORHIGHLIGHT = (WM_USER + 74)
Private Const TB_MAPACCELERATORA = (WM_USER + 78)      '// wParam == ch, lParam int * pidBtn
Private Const TB_MAPACCELERATORW = (WM_USER + 90)      '// wParam == ch,
Private Const TB_MAPACCELERATOR = TB_MAPACCELERATORA

Private Type TBINSERTMARK
    iButton As Long
    dwFlags As Long
End Type
Private Const TBIMHT_AFTER = &H1      '// TRUE = insert After iButton, otherwise before
Private Const TBIMHT_BACKGROUND = &H2 '// TRUE iff missed buttons completely

Private Const TB_GETINSERTMARK = (WM_USER + 79)        '// lParam == LPTBINSERTMARK
Private Const TB_SETINSERTMARK = (WM_USER + 80)        '// lParam == LPTBINSERTMARK
Private Const TB_INSERTMARKHITTEST = (WM_USER + 81)    '// wParam == LPPOINT lParam == LPTBINSERTMARK
Private Const TB_MOVEBUTTON = (WM_USER + 82)

Private Const TB_GETMAXSIZE = (WM_USER + 83)           '// lParam == LPSIZE

' Extended style:
Private Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)    ' // For TBSTYLE_EX_*
Private Const TB_GETEXTENDEDSTYLE = (WM_USER + 85)     '// For TBSTYLE_EX_*
Private Const TB_GETPADDING = (WM_USER + 86)
Private Const TB_SETPADDING = (WM_USER + 87)
Private Const TB_SETINSERTMARKCOLOR = (WM_USER + 88)
Private Const TB_GETINSERTMARKCOLOR = (WM_USER + 89)

Private Const TB_SETCOLORSCHEME = CCM_SETCOLORSCHEME       '// lParam is color scheme
Private Const TB_GETCOLORSCHEME = CCM_GETCOLORSCHEME       '// fills in COLORSCHEME pointed to by lParam
'#endif  // _WIN32_IE >= 0x0400

Private Const TBSTYLE_EX_DRAWDDARROWS = &H1

'//Standard image types:
Private Const IDB_STD_SMALL_COLOR = 0
Private Const IDB_STD_LARGE_COLOR = 1
Private Const IDB_VIEW_SMALL_COLOR = 4
Private Const IDB_VIEW_LARGE_COLOR = 5
Private Const IDB_HIST_SMALL_COLOR = 8
Private Const IDB_HIST_LARGE_COLOR = 9

'// icon indexes for standard bitmap

Private Const STD_CUT = 0
Private Const STD_COPY = 1
Private Const STD_PASTE = 2
Private Const STD_UNDO = 3
Private Const STD_REDOW = 4
Private Const STD_DELETE = 5
Private Const STD_FILENEW = 6
Private Const STD_FILEOPEN = 7
Private Const STD_FILESAVE = 8
Private Const STD_PRINTPRE = 9
Private Const STD_PROPERTIES = 10
Private Const STD_HELP = 11
Private Const STD_FIND = 12
Private Const STD_REPLACE = 13
Private Const STD_PRINT = 14

'// icon indexes for standard view bitmap

Private Const VIEW_LARGEICONS = 0
Private Const VIEW_SMALLICONS = 1
Private Const VIEW_LIST = 2
Private Const VIEW_DETAILS = 3
Private Const VIEW_SORTNAME = 4
Private Const VIEW_SORTSIZE = 5
Private Const VIEW_SORTDATE = 6
Private Const VIEW_SORTTYPE = 7
Private Const VIEW_PARENTFOLDER = 8
Private Const VIEW_NETCONNECT = 9
Private Const VIEW_NETDISCONNECT = 10
Private Const VIEW_NEWFOLDER = 11
'#if (_WIN32_IE >= 0x0400)
Private Const VIEW_VIEWMENU = 12
'#End If

'#if (_WIN32_IE >= 0x0300)
Private Const HIST_BACK = 0
Private Const HIST_FORWARD = 1
Private Const HIST_FAVORITES = 2
Private Const HIST_ADDTOFAVORITES = 3
Private Const HIST_VIEWTREE = 4
'#End If

Private Declare Function CreateToolbarEx Lib "comctl32" (ByVal hwnd As Long, ByVal ws As Long, ByVal wId As Long, ByVal nBitmaps As Long, ByVal hBMInst As Long, ByVal wBMID As Long, ByRef lpButtons As TBBUTTON, ByVal iNumButtons As Long, ByVal dxButton As Long, ByVal dyButton As Long, ByVal dxBitmap As Long, ByVal dyBitmap As Long, ByVal uStructSize As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "Comctl32.dll" ( _
        ByVal hIml As Long _
    ) As Long
Private Declare Function ImageList_GetImageRect Lib "Comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function ImageList_Draw Lib "Comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_PRESERVEALPHA = &H1000&        '// This preserves the alpha channel in dest
Private Const ILD_OVERLAYMASK = 3840

' ==============================================================================
' INTERFACE
' ==============================================================================
' Enumerations:
Public Enum ECTBToolButtonSyle
    CTBNormal = TBSTYLE_BUTTON
    CTBSeparator = TBSTYLE_SEP
    CTBCheck = TBSTYLE_CHECK
    CTBCheckGroup = TBSTYLE_CHECKGROUP
    CTBDropDown = TBSTYLE_DROPDOWN
    CTBAutoSize = TBSTYLE_AUTOSIZE
    CTBDropDownArrow = BTNS_WHOLEDROPDOWN
End Enum
Public Enum ECTBImageListTypes
   CTBImageListNormal = TB_SETIMAGELIST
   CTBImageListHot = TB_SETHOTIMAGELIST
   CTBImageListDisabled = TB_SETDISABLEDIMAGELIST
End Enum
Public Enum ECTBToolbarStyle
    CTBFlat = TBSTYLE_FLAT
    CTBList = TBSTYLE_LIST
    CTBTransparent = -1 ' special - here we remove Toolbar from owner window
End Enum
Public Enum ECTBImageSourceTypes
    CTBResourceBitmap
    CTBLoadFromFile
    CTBExternalImageList
    CTBPicture
    CTBStandardImageSources
End Enum
Public Enum ECTBStandardImageSourceTypes
   CTBHistoryLargeColor = IDB_HIST_LARGE_COLOR
   CTBHistorySmallColor = IDB_HIST_SMALL_COLOR
   CTBStandardLargeColor = IDB_STD_LARGE_COLOR
   CTBStandardSmallColor = IDB_STD_SMALL_COLOR
   CTBViewLargeColor = IDB_VIEW_LARGE_COLOR
   CTBViewSmallColor = IDB_VIEW_SMALL_COLOR
End Enum
Public Enum ECTBStandardImageIndexConstants
   ' History:
   CTBHistAddToFavourites = HIST_ADDTOFAVORITES ' 'Add 'to 'favorites.
   CTBHistBack = HIST_BACK ' 'Move 'back.
   CTBHistFavourites = HIST_FAVORITES ' 'Open 'favorites 'folder.
   CTBHistForward = HIST_FORWARD ' 'Move 'forward.
   CTBHistViewTree = HIST_VIEWTREE ' 'View 'tree.
   'Standard:
   CTBStdCopy = STD_COPY ' 'Copy 'operation.
   CTBStdCut = STD_CUT ' 'Cut 'operation.
   CTBStdDelete = STD_DELETE ' 'Delete 'operation.
   CTBStdFileNew = STD_FILENEW ' 'New 'file 'operation.
   CTBStdFileOpen = STD_FILEOPEN ' 'Open 'file 'operation.
   CTBStdFIleSave = STD_FILESAVE ' 'Save 'file 'operation.
   CTBStdFind = STD_FIND ' 'Find 'operation.
   CTBStdHelp = STD_HELP ' 'Help 'operation.
   CTBStdPaste = STD_PASTE ' 'Paste 'operation.
   CTBStdPrint = STD_PRINT ' 'Print 'operation.
   CTBStdPrintPreview = STD_PRINTPRE ' 'Print 'preview 'operation.
   CTBStdProperties = STD_PROPERTIES ' 'Properties 'operation.
   CTBStdRedo = STD_REDOW ' 'Redo 'operation.
   CTBStdReplace = STD_REPLACE ' 'Replace 'operation.
   CTBStdUndo = STD_UNDO ' 'Undo 'operation.
   'View
   CTBViewDetails = VIEW_DETAILS ' 'Details 'view.
   CTBViewLargeIcons = VIEW_LARGEICONS ' 'Large 'icons 'view.
   CTBViewList = VIEW_LIST ' 'List 'view.
   CTBViewNetConnect = VIEW_NETCONNECT ' 'Connect 'to 'network 'drive.
   CTBViewNetDisconnect = VIEW_NETDISCONNECT ' 'Disconnect 'from 'network 'drive.
   CTBViewNewFolder = VIEW_NEWFOLDER ' 'New 'folder.
   CTBViewParentFolder = VIEW_PARENTFOLDER ' 'Go 'to 'parent 'folder.
   CTBViewSmallIcons = VIEW_SMALLICONS ' 'Small 'icon 'view.
   CTBViewSortDate = VIEW_SORTDATE ' 'Sort 'by 'date.
   CTBViewSortName = VIEW_SORTNAME ' 'Sort 'by 'name.
   CTBViewSortSize = VIEW_SORTSIZE ' 'Sort 'by 'size.
   CTBViewSortType = VIEW_SORTTYPE ' 'Sort 'by 'type.
End Enum
Public Enum ECTBHotItemChangeReasonConstants
   HICF_OTHER = 0
   HICF_MOUSE = 1 '// Triggered by mouse
   HICF_ARROWKEYS = 2 ' // Triggered by arrow keys
   HICF_ACCELERATOR = 4  '// Triggered by accelerator
   HICF_DUPACCEL = 8               '// This accelerator is not unique
   HICF_ENTERING = 10               '// idOld is invalid
   HICF_LEAVING = 20                '// idNew is invalid
   HICF_RESELECT = 40               '// hot item reselected
End Enum
Public Enum ECTBToolbarFromMenuStyle
   CTBMenuStyle
   CTBToolbarStyle
End Enum
Public Enum ECTBDropDownAlign
   CTBDropDownAlignBottom
   CTBDropDownAlignLeft
End Enum
Public Enum ECTBChevronAdditionalButtons
   CTBChevronAdditionalAddorRemove
   CTBChevronAdditionalCustomise
   CTBChevronAdditionalReset
End Enum
Public Enum ECTBToolbarDrawStyle
   CTBDrawStandard
   CTBDrawNoVisualStyles
   CTBDrawOfficeXPStyle
End Enum

' Events:
Public Event ButtonClick(ByVal lButton As Long)
Attribute ButtonClick.VB_Description = "Raised when a toolbar button is clicked."
Public Event DropDownPress(ByVal lButton As Long)
Attribute DropDownPress.VB_Description = "Raised when a drop-down arrow on a drop-down button is pressed (Note: COMCTL32.DLL versions below 4.71 do not display drop-down buttons)"
Public Event HotItemChange(ByVal iNew As Long, ByVal iOld As Long, ByVal eReason As ECTBHotItemChangeReasonConstants)
Attribute HotItemChange.VB_Description = "Raised whenever the hot button changes in a flat toolbar."
Public Event CustomiseBegin()
Public Event CustomiseCanInsertBefore(ByVal lButton As Long, ByRef bCanInsert As Boolean)
Public Event CustomiseCanDelete(ByVal lButton As Long, ByRef bCanDelete As Boolean)
Public Event CustomiseHelpPressed()
Public Event CustomiseResetPressed()

Private Const DROPDOWN_ARROW_WIDTH = 13

' ==============================================================================
' INTERNAL INFORMATION
' ==============================================================================
' Subclassing
Implements ISubclass
Private m_bInSubClass As Boolean

' Classes to turn toolbar into a menu:
Private m_cMenu As cTbarMenu

Private m_bIsMenu As Boolean
Private m_hMenu As Long
Private m_eCreateFromMenuStyle  As ECTBToolbarFromMenuStyle
Private m_bCreateFromMenu2 As Boolean
Private m_lPtrMenu As Long
Private m_eDropDownAlign As ECTBDropDownAlign
Private m_bMenuShown As Boolean
Private m_bMenuLoop As Boolean
Private m_bAltPressed As Boolean

Private m_cNCM As New pcNCMetrics

' Hwnd of tool bar itself:
Private m_hWndToolBar As Long
Private m_hWndChevronToolbar As Long
Private m_hWndParentForm As Long

' Chevron information:
Private m_bChevronAdditionalButton(0 To 2) As Boolean
Private m_sChevronAdditionalButton(0 To 2) As String
Private m_iChevronIDMap() As Long
Private m_iChevronIDMapCount As Long
Private m_cChevronWindow As cChevronWindow

' Where the button images are coming from
Private m_eImageSourceType As ECTBImageSourceTypes
Private m_pic As StdPicture
Private m_sFileName As String
Private m_lResourceID As Long
Private m_hInstance As Long
Private m_hIml As Long
Private m_hImlHot As Long
Private m_hImlDis As Long
Private m_ptrVb6ImageList As Long
Private m_eStandardType As ECTBStandardImageSourceTypes
Private m_lIconWidth As Long
Private m_lIconHeight As Long
Private m_lTransColor As Long
Private m_cMemDC As cAlphaDibSection

' Button size:
Private m_iButtonWidth As Integer
Private m_iButtonHeight As Integer
Private m_lOrigButtonSize As Long

' Style information:
Private m_bWithText As Boolean
Private m_bWrappable As Boolean
Private m_eVisualStyle As ECTBToolbarDrawStyle
Private m_bListStyle  As Boolean

Private m_bVisible As Boolean

' Button information:
' Types:
Private Type ButtonInfoStore
    wId As Integer
    iImage As Integer
    sTipText As String
    iTextIndexNum As Integer
    sCaption As String
    bShowText As Boolean
    idString As Long
    iLarge As Integer
    xWidth As Integer
    xHeight As Integer
    sKey As String
    eStyle As ECTBToolButtonSyle
    hSubMenu As Long
    hWndCapture As Long
    hWndParentOrig As Long
    bStretch As Boolean
    bControl As Boolean
    bDropped As Boolean
End Type
Private m_tBInfo() As ButtonInfoStore
' Last return code from toolbar API or sendmessage call
Private m_lR As Long

' Strings in the toolbar:
Private m_lStringIDCount As Long
Private m_sString() As String
Private m_lStringID() As Long

' Common Controls Version:
Private m_lMajorVer As Long
Private m_lMinorVer As Long
Private m_lBuild As Long

' Whether to keep in focus when showing tool wins
Private m_bTitleBarModifier As Boolean

Private m_tRebarBand As RECT

Private m_sCtlName As String


Public Property Get DrawStyle() As ECTBToolbarDrawStyle
   DrawStyle = m_eVisualStyle
End Property
Public Property Let DrawStyle(ByVal eStyle As ECTBToolbarDrawStyle)
   m_eVisualStyle = eStyle
   If Not (m_hWndToolBar = 0) Then
      If (eStyle = CTBDrawStandard) Then
         ' Allow XP Visual Styles
         On Error Resume Next
         SetWindowTheme m_hWndToolBar, StrPtr("Toolbar"), StrPtr("Toolbar")
         On Error GoTo 0
      Else
         ' No XP Visual Styles
         On Error Resume Next
         SetWindowTheme m_hWndToolBar, StrPtr(" "), StrPtr(" ")
         On Error GoTo 0
      End If
   End If
   PropertyChanged "DrawStyle"
End Property

Public Sub ChevronPress(ByVal x As Long, ByVal y As Long)

Dim lhWndChevronToolBar As Long
Dim dwStyle As Long
Dim dwExStyle As Long
Dim Button As TBBUTTON
Dim lParam As Long
Dim i As Long
Dim tR As RECT
Dim lW As Long, lH As Long
Dim iNotVisibleIndex As Long
Dim lhWndParent As Long
Dim lExStyle As Long
Dim bMenu As Boolean
Dim hMenu As Long
Dim hSubMenu As Long
Dim tPM As TPMPARAMS
Dim lCmd As Long
Dim lR As Long
Dim cT As Object
Dim tP As POINTAPI
Dim cMenu As Object
Dim iPos As Long
Dim tMII As MENUITEMINFO
Dim tMI() As MENUITEMINFO
Dim iMenuItemCount As Long
Dim bButtonStyle As Boolean
Dim lIndex As Long
Dim bCustomOnly As Boolean
Dim lChevronAddition() As Long
Dim sChevronAddition() As String
Dim lChevronAdditionCount As Long
Dim lChevronTop As Long
Dim lMenu As Long
Dim lTopLevelMenu As Long
Dim bNoAdditionalCustomSeparator As Boolean
Dim sKeyBit As String

   bMenu = (Not (m_lPtrMenu = 0))
   lhWndParent = m_hWndParentForm
   If Not (getActiveWindow() = lhWndParent) Then
      On Error Resume Next
      UserControl.Parent.ZOrder
      On Error GoTo 0
   End If
   
   If Not bMenu Then
      ' toolbar
      If Not (m_hWndChevronToolbar = 0) Then
         ShowWindow m_hWndChevronToolbar, SW_HIDE
         SetParent m_hWndChevronToolbar, 0
         DestroyWindow m_hWndChevronToolbar
      End If
      
      ' Create a toolbar to show:
      dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN
      dwStyle = dwStyle Or CCS_NOPARENTALIGN Or CCS_NORESIZE Or CCS_NODIVIDER
      dwStyle = dwStyle Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT
      dwStyle = dwStyle Or TBSTYLE_LIST
      dwStyle = dwStyle Or TBSTYLE_WRAPABLE
      dwStyle = dwStyle Or TBSTYLE_REGISTERDROP
      
      dwExStyle = WS_EX_TOOLWINDOW
      lExStyle = GetWindowLong(lhWndParent, GWL_EXSTYLE)
      lExStyle = lExStyle And (WS_EX_RIGHT Or WS_EX_RTLREADING)
      dwExStyle = dwExStyle Or lExStyle
      lhWndChevronToolBar = CreateWindowEX(dwExStyle, "ToolbarWindow32", "", _
            dwStyle, _
            0, 0, 0, 0, UserControl.hwnd, 0&, App.hInstance, 0&)
      SendMessageLong lhWndChevronToolBar, TB_SETPARENT, UserControl.hwnd, 0
      m_lR = SendMessageLong(lhWndChevronToolBar, TB_BUTTONSTRUCTSIZE, LenB(Button), 0)
      AddBitmapIfRequired lhWndChevronToolBar
      If m_eImageSourceType <> -1 Then
         lParam = m_lOrigButtonSize + (m_lOrigButtonSize * &H10000)
      Else
         lParam = 0
      End If
      m_lR = SendMessageLong(lhWndChevronToolBar, TB_SETBITMAPSIZE, 0, lParam)
      ' Ok, now we have a toolbar to work with, add copies of the
      ' buttons that are currently out of view in the toolbar:
   Else
      ' Create a menu to add items to:
      'hMenu = CreatePopupMenu()
      CopyMemory cT, m_lPtrMenu, 4
      Set cMenu = cT
      CopyMemory cT, 0&, 4
      
   End If
   
   iNotVisibleIndex = findFirstNonVisibleButton()
   m_iChevronIDMapCount = 0
   
   ' Is there anything to do?
   bCustomOnly = (bMenu And (m_bChevronAdditionalButton(CTBChevronAdditionalAddorRemove) Or m_bChevronAdditionalButton(CTBChevronAdditionalCustomise)))
   
   If (iNotVisibleIndex < 0) And Not (bCustomOnly) Then
      If lhWndChevronToolBar Then
         DestroyWindow lhWndChevronToolBar
      End If
      Exit Sub
   End If
      
   If bMenu Then
      
      ' Remove items which can be seen in the toolbar:
      If iNotVisibleIndex < 0 Then
         iNotVisibleIndex = GetMenuItemCount(m_hMenu)
         bNoAdditionalCustomSeparator = True
      End If
      For i = iNotVisibleIndex - 1 To 0 Step -1
         tMII.fMask = MIIM_ID
         tMII.cbSize = Len(tMII)
         GetMenuItemInfo m_hMenu, i, True, tMII
         lIndex = cMenu.ItemForID(tMII.wId)
         ' Debug.Print lIndex, cMenu.Caption(lIndex)
         If cMenu.Visible(lIndex) Then
            iMenuItemCount = iMenuItemCount + 1
            ReDim Preserve tMI(1 To iMenuItemCount) As MENUITEMINFO
            LSet tMI(iMenuItemCount) = tMII
            cMenu.Visible(lIndex) = False
         End If
      Next i
      
      
      lMenu = 0
      For i = 1 To cMenu.count
         If (cMenu.hMenu(i) = m_hMenu) Then
            If cMenu.Visible(i) Then
               lMenu = i
               lTopLevelMenu = cMenu.ItemParentIndex(lMenu)
               Exit For
            End If
            lTopLevelMenu = cMenu.ItemParentIndex(i)
         End If
      Next i
      
      If m_bChevronAdditionalButton(CTBChevronAdditionalAddorRemove) Or m_bChevronAdditionalButton(CTBChevronAdditionalCustomise) Or m_bChevronAdditionalButton(CTBChevronAdditionalReset) Then
         If Not bNoAdditionalCustomSeparator Then
            lChevronAdditionCount = lChevronAdditionCount + 1
            ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
            ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
            sChevronAddition(lChevronAdditionCount) = sKeyBit & ":SEP:1"
            lChevronAddition(lChevronAdditionCount) = cMenu.AddItem("-", , VBALCHEVRONMENUCONST, lTopLevelMenu, , , , sChevronAddition(lChevronAdditionCount))
         End If
      
         sKeyBit = "_VBALCC:" & m_hWndToolBar
         ' add the "Add or Remove Buttons" option:
         lChevronAdditionCount = lChevronAdditionCount + 1
         ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
         ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
         sChevronAddition(lChevronAdditionCount) = sKeyBit & ":AOR"
         lChevronAddition(lChevronAdditionCount) = cMenu.AddItem(m_sChevronAdditionalButton(CTBChevronAdditionalAddorRemove), , VBALCHEVRONMENUCONST, lTopLevelMenu, , , , sChevronAddition(lChevronAdditionCount))
         lChevronTop = lChevronAddition(lChevronAdditionCount)
         If lMenu <= 0 Then
            lMenu = lChevronAddition(lChevronAdditionCount)
         End If
         i = -1
         If (m_bChevronAdditionalButton(CTBChevronAdditionalAddorRemove)) Then
            ' add the add/remove details:
            For i = 0 To ButtonCount - 1
               lChevronAdditionCount = lChevronAdditionCount + 1
               ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
               ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
               sChevronAddition(lChevronAdditionCount) = sKeyBit & ":BTN:" & i & ":" & ButtonKey(i)
               lChevronAddition(lChevronAdditionCount) = cMenu.AddItem( _
                  ButtonCaption(i), , _
                  VBALCHEVRONMENUCONST, _
                  lChevronTop, _
                  m_tBInfo(i).iImage, ButtonVisible(i), , _
                  sChevronAddition(lChevronAdditionCount))
               cMenu.RedisplayMenuOnClick(lChevronAddition(lChevronAdditionCount)) = True
               cMenu.ShowCheckAndIcon(lChevronAddition(lChevronAdditionCount)) = True
            Next i
         End If
         If m_bChevronAdditionalButton(CTBChevronAdditionalReset) Then
            ' add the reset toolbar button:
            If i > -1 Then
               i = -1
               lChevronAdditionCount = lChevronAdditionCount + 1
               ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
               ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
               sChevronAddition(lChevronAdditionCount) = sKeyBit & ":SEP:2"
               lChevronAddition(lChevronAdditionCount) = cMenu.AddItem("-", , VBALCHEVRONMENUCONST, lChevronTop, , , , sChevronAddition(lChevronAdditionCount))
            End If
            lChevronAdditionCount = lChevronAdditionCount + 1
            ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
            ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
            sChevronAddition(lChevronAdditionCount) = sKeyBit & ":RST"
            lChevronAddition(lChevronAdditionCount) = cMenu.AddItem(m_sChevronAdditionalButton(CTBChevronAdditionalReset), , VBALCHEVRONMENUCONST, lChevronTop, , , , sChevronAddition(lChevronAdditionCount))
         End If
         If m_bChevronAdditionalButton(CTBChevronAdditionalCustomise) Then
            ' add the customise button:
            If i > -1 Then
               i = -1
               lChevronAdditionCount = lChevronAdditionCount + 1
               ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
               ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
               sChevronAddition(lChevronAdditionCount) = sKeyBit & ":SEP:3"
               lChevronAddition(lChevronAdditionCount) = cMenu.AddItem("-", , VBALCHEVRONMENUCONST, lChevronTop, , , , sChevronAddition(lChevronAdditionCount))
            End If
            lChevronAdditionCount = lChevronAdditionCount + 1
            ReDim Preserve lChevronAddition(1 To lChevronAdditionCount) As Long
            ReDim Preserve sChevronAddition(1 To lChevronAdditionCount) As String
            sChevronAddition(lChevronAdditionCount) = sKeyBit & ":CST"
            lChevronAddition(lChevronAdditionCount) = cMenu.AddItem(m_sChevronAdditionalButton(CTBChevronAdditionalCustomise), , VBALCHEVRONMENUCONST, lChevronTop, , , , sChevronAddition(lChevronAdditionCount))
         End If
         
      End If
      
   Else
      For i = iNotVisibleIndex To ButtonCount - 1
         If Not m_tBInfo(i).eStyle = CTBSeparator Then
            m_iChevronIDMapCount = m_iChevronIDMapCount + 1
            plAddButton lhWndChevronToolBar, m_tBInfo(i).wId, m_tBInfo(i).sTipText, m_tBInfo(i).iImage, , m_tBInfo(i).iLarge, m_tBInfo(i).sCaption, m_tBInfo(i).eStyle 'And Not CTBAutoSize
            'plAddButton lhWndChevronToolBar, m_tBInfo(i).wId, m_tBInfo(i).sTipText, m_tBInfo(i).iImage, , m_tBInfo(i).iLarge, , m_tBInfo(i).eStyle 'And Not CTBAutoSize
            SendMessageLong lhWndChevronToolBar, TB_ENABLEBUTTON, m_tBInfo(i).wId, Abs(ButtonEnabled(i))
            SendMessageLong lhWndChevronToolBar, TB_CHECKBUTTON, m_tBInfo(i).wId, Abs(ButtonChecked(i))
            ReDim Preserve m_iChevronIDMap(1 To m_iChevronIDMapCount) As Long
            m_iChevronIDMap(m_iChevronIDMapCount) = i
         End If
      Next i
   End If
   
   If bMenu Then
      
      tP.x = x: tP.y = y
      ScreenToClient cMenu.hWndOwner, tP
      
      lIndex = cMenu.ShowPopupMenuAtIndex(tP.x * Screen.TwipsPerPixelX, tP.y * Screen.TwipsPerPixelY, , , , , , lMenu)
      
      ' add menu items back in again:
      For i = iMenuItemCount To 1 Step -1
         lIndex = cMenu.ItemForID(tMI(i).wId)
         cMenu.Visible(lIndex) = True
      Next i
      
      ' remove the chevron items:
      For i = lChevronAdditionCount To 1 Step -1
         cMenu.RemoveItem sChevronAddition(i) 'lChevronAddition(i)
      Next i
      
   Else
      ' Evaluate the size of the chevron bar:
      lW = 0: lH = 0
      For i = 0 To plButtonCount(lhWndChevronToolBar) - 1
         SendMessage lhWndChevronToolBar, TB_GETITEMRECT, i, tR
         If tR.right - tR.left > lW Then
            lW = tR.right - tR.left
         End If
         lH = lH + tR.bottom - tR.top
      Next i
      ' account for borders:
      lW = lW
      lH = lH
      
      If y + lH > Screen.height \ Screen.TwipsPerPixelY - 2 Then
         y = Screen.height \ Screen.TwipsPerPixelY - lH - 2
      End If
      If x + lW > Screen.width \ Screen.TwipsPerPixelX - 2 Then
         x = Screen.width \ Screen.TwipsPerPixelX - lW - 2
      End If
   
      ' Show the chevron window at the appropriate position:
      Set m_cChevronWindow = New cChevronWindow
      
      m_hWndChevronToolbar = lhWndChevronToolBar
      m_cChevronWindow.Show m_hWndParentForm, m_hWndChevronToolbar, x, y, lW, lH
      If Not m_cChevronWindow Is Nothing Then
         m_cChevronWindow.Destroy
      End If
      m_hWndChevronToolbar = 0
         
   End If
   
End Sub

Public Property Get ChevronButton(ByVal eButton As ECTBChevronAdditionalButtons) As Boolean
   ChevronButton = m_bChevronAdditionalButton(eButton)
End Property
Public Property Let ChevronButton(ByVal eButton As ECTBChevronAdditionalButtons, ByVal bState As Boolean)
   m_bChevronAdditionalButton(eButton) = bState
End Property
Public Property Get ChevronButtonCaption(ByVal eButton As ECTBChevronAdditionalButtons) As String
   ChevronButtonCaption = m_sChevronAdditionalButton(eButton)
End Property
Public Property Let ChevronButtonCaption(ByVal eButton As ECTBChevronAdditionalButtons, ByVal sCaption As String)
   m_sChevronAdditionalButton(eButton) = sCaption
End Property
Friend Function InMenuLoop() As Boolean
   If (m_bMenuShown) Then
      m_bMenuLoop = False
   End If
   InMenuLoop = m_bMenuLoop
End Function
Friend Function AltKeyPress(ByVal eKeyCode As KeyCodeConstants, ByVal bKeyUp As Boolean, ByVal bAlt As Boolean, ByVal bShift As Boolean) As Boolean
Dim wId As Long
Dim iKey As Long
Dim iB As Long
Dim i As Long
Dim j As Long
Dim sAccel As String
Dim lR As Long

   If m_hWndToolBar <> 0 Then
      ' Am i a member of an active form?
      If getTheActiveWindow() Then
         If (bAlt) And (eKeyCode <> 18) And Not (bKeyUp) Then
            
            iB = -1
            sAccel = UCase$(Chr$(eKeyCode))
            For i = 0 To ButtonCount - 1
               If psGetAccelerator(m_tBInfo(i).sCaption) = sAccel Then
                  iB = i
                  wId = m_tBInfo(i).wId
                  Exit For
               End If
            Next i
            
            If iB > -1 Then
               If (ButtonVisible(iB)) Then
                  'SetFocusAPI m_hWndToolBar
                  ButtonPressed(iB) = True
                  SendMessageLong m_hWndToolBar, WM_COMMAND, wId, m_hWndToolBar
                  ButtonPressed(iB) = False
                  AltKeyPress = True
               End If
            End If
            
         Else
            
            If (m_eCreateFromMenuStyle = CTBMenuStyle) Then
            
               If Not (m_bMenuShown) Then
                     
                  Dim iFirst As Long
                  Dim iFirstHot As Long
                  
                  iFirstHot = -1
                  iFirst = -1
                  For i = 0 To ButtonCount - 1
                     If ButtonVisible(i) Then
                        If (iFirst = -1) Then
                           iFirst = i
                        End If
                        If (ButtonHot(i)) Then
                           If (iFirstHot = -1) Then
                              iFirstHot = i
                           End If
                        End If
                     End If
                  Next i
                  
                  If Not (m_bMenuLoop) Then
                     ' Not in menu loop:
                     If (eKeyCode = 18) Then
                        If Not bKeyUp Then
                           ' show the accelerators:
                           m_bAltPressed = True
                           showAccelerators True
                        Else
                           ' Highlight the first item:
                           If (iFirstHot < 0) Then
                              ButtonHot(iFirst) = True
                           End If
                           m_bMenuLoop = True
                           m_bAltPressed = False
                        End If
                        AltKeyPress = True
                     End If
                  Else
                     ' Menu Loop:
                     
                     Select Case eKeyCode
                     Case 18
                        Debug.Print "18", m_bMenuLoop
                        If (bKeyUp) Then
                           ' un-highlight the first item in the toolbar
                           If (iFirstHot >= 0) Then
                              'ButtonHot(iFirst) = False
                              ButtonHot(iFirstHot) = False
                           End If
                           m_bMenuLoop = False
                           showAccelerators False
                           AltKeyPress = True
                        End If
                        
                     Case vbKeyLeft
                        If Not (bKeyUp) Then
                           If (iFirstHot = -1) Then
                              ButtonHot(iFirst) = True
                           Else
                              Debug.Print iFirstHot
                              i = iFirstHot - 1
                              Do While j < ButtonCount
                                 If (i < 0) Then
                                    i = ButtonCount - 1
                                 End If
                                 If ButtonVisible(i) Then
                                    If (iFirstHot >= 0) Then
                                       ButtonHot(iFirstHot) = False
                                    End If
                                    ButtonHot(i) = True
                                    Exit Do
                                 End If
                                 j = j + 1
                                 i = i - 1
                              Loop
                           End If
                           
                           AltKeyPress = True
                        End If
                     Case vbKeyRight
                        If Not (bKeyUp) Then
                           If (iFirstHot = -1) Then
                              ButtonHot(iFirst) = True
                           Else
                              Debug.Print iFirstHot
                              i = iFirstHot + 1
                              Do While j < ButtonCount
                                 If (i >= ButtonCount) Then
                                    i = 0
                                 End If
                                 If ButtonVisible(i) Then
                                    If (iFirstHot >= 0) Then
                                       ButtonHot(iFirstHot) = False
                                    End If
                                    ButtonHot(i) = True
                                    Exit Do
                                 End If
                                 j = j + 1
                                 i = i + 1
                              Loop
                           End If
                        End If
                        
                        AltKeyPress = True
                        
                     Case vbKeyDown, vbKeyUp, vbKeyReturn
                        If bKeyUp Then
                           m_bMenuLoop = False
                           ButtonPressed(iFirstHot) = True
                           SendMessageLong m_hWndToolBar, WM_COMMAND, m_tBInfo(iFirstHot).wId, m_hWndToolBar
                           ButtonPressed(iFirstHot) = False
                        End If
                        AltKeyPress = True
                        
                     Case vbKeyEscape
                        ' exit menu loop:
                        m_bMenuLoop = False
                        showAccelerators False
                        If Not bKeyUp Then
                           If (iFirstHot > -1) Then
                              ButtonHot(iFirstHot) = False
                           End If
                        End If
                        AltKeyPress = True
                        
                     Case Else
                        If bKeyUp Then
                           iB = -1
                           sAccel = UCase$(Chr$(eKeyCode))
                           For i = 0 To ButtonCount - 1
                              If psGetAccelerator(m_tBInfo(i).sCaption) = sAccel Then
                                 iB = i
                                 wId = m_tBInfo(i).wId
                                 Exit For
                              End If
                           Next i
                           
                           If (iB > -1) Then
                              If ButtonVisible(iB) Then
                                 ' start menu tracking:
                                 m_bMenuLoop = False
                                 For i = 0 To ButtonCount - 1
                                    ButtonHot(i) = False
                                 Next i
                                 ButtonPressed(iB) = True
                                 SendMessageLong m_hWndToolBar, WM_COMMAND, wId, m_hWndToolBar
                                 ButtonPressed(iB) = False
                                 AltKeyPress = True
                                 Exit Function
                              End If
                           End If
                           
                           ' Not a valid key:
                           Beep
                           m_bMenuLoop = False
                           If (iFirstHot > -1) Then
                              ButtonHot(iFirstHot) = False
                           End If
                           AltKeyPress = True
                        End If
                     End Select
                  End If
               End If
               
            End If
         End If
      End If
   End If
   
End Function
Private Sub showAccelerators(ByVal bState As Boolean)
   ' To do
End Sub

Private Function getTheActiveWindow() As Boolean
Dim lhWnd As Long
   lhWnd = getActiveWindow()
   If lhWnd = m_hWndParentForm Then
      ' is active
      getTheActiveWindow = True
   Else
      lhWnd = GetProp(lhWnd, TOOLWINDOWPARENTWINDOWHWND)
      If lhWnd = m_hWndParentForm Then
         ' is active
         getTheActiveWindow = True
      End If
   End If
End Function
Friend Sub pMenuClick(ByVal hWndToolbar As Long, ByVal iButton As Long)
Dim lR As Long
   
   Debug.Print "MENUCLICK", iButton
   If Not m_lPtrMenu = 0 Then
      PopupObject.CreateSubClass m_hWndParentForm
   End If
   
   If Not m_cMenu Is Nothing Then
      m_bMenuLoop = False
      m_bMenuShown = True
      m_cMenu.MenuAlignLeft = (m_eDropDownAlign = CTBDropDownAlignLeft)
      m_cMenu.CoolMenuAttach m_hWndParentForm, hWndToolbar, m_hMenu, m_lPtrMenu
      Debug.Print "Calling Track Popup:"
      lR = m_cMenu.TrackPopup(iButton)
      m_cMenu.CoolMenuDetach
      setDroppedButton 0, False
      m_bMenuShown = False
      If (m_cMenu.EscapeWasPressed) Then
         Debug.Print "ESCAPE WAS PRESSED:"
         m_bMenuLoop = True
         ButtonHot(iButton) = True
      End If
   End If
   
   If Not m_lPtrMenu = 0 Then
      If lR <> 0 Then
         ' Debug.Print "THAT WAS MENU ITEM: ", lR
         PopupObject.EmulateMenuClick lR
      End If
      PopupObject.DestroySubClass
   End If
   
End Sub
Friend Sub setDroppedButton(ByVal iButton As Long, ByVal bState As Boolean)
   m_tBInfo(iButton).bDropped = bState
End Sub

Private Property Get PopupObject() As Object
Dim oTemp As Object
   CopyMemory oTemp, m_lPtrMenu, 4
   Set PopupObject = oTemp
   CopyMemory oTemp, 0&, 4
End Property

Public Property Get AutosizeButtonPadding() As Long
Attribute AutosizeButtonPadding.VB_Description = "Gets/sets the number of pixels by which to pad out buttons with the CTBAutosize property set."
   ' NB Only applies to autosize buttons
   If m_hWndToolBar <> 0 Then
      AutosizeButtonPadding = (SendMessageLong(m_hWndToolBar, TB_GETPADDING, 0, 0) And &H7FFF&)
   End If
End Property
Public Property Let AutosizeButtonPadding(ByVal lPadding As Long)
Dim lxy As Long
   If m_hWndToolBar <> 0 Then
      lxy = (lPadding And &H7FFF&) Or (lPadding And &H7FFF& * &H10000)
      SendMessageLong m_hWndToolBar, TB_SETPADDING, 0, lxy
   End If
End Property

Public Sub GetComCtrlVersionInfo( _
      ByRef lMajor As Long, _
      ByRef lMinor As Long, _
      Optional ByRef lBuild As Long _
   )
Attribute GetComCtrlVersionInfo.VB_Description = "Returns the system's COMCTL32.DLL version."
   lMajor = m_lMajorVer
   lMinor = m_lMinorVer
   lBuild = m_lBuild
   End Sub
      

Public Property Get ButtonCount() As Long
Attribute ButtonCount.VB_Description = "Returns the number of buttons in a toolbar."
   If m_hWndToolBar <> 0 Then
      ButtonCount = plButtonCount(m_hWndToolBar)
   End If
End Property
Private Property Get plButtonCount(ByVal hWndToolbar As Long) As Long
   plButtonCount = SendMessageLong(hWndToolbar, TB_BUTTONCOUNT, 0, 0)
End Property

Public Property Get ButtonToolTip(ByVal vButton As Variant) As String
Attribute ButtonToolTip.VB_Description = "Gets/sets the tool tip shown for a button."
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB > -1) Then
        ButtonToolTip = m_tBInfo(iB).sTipText
    End If
End Property
Public Property Let ButtonToolTip(ByVal vButton As Variant, ByVal sToolTip As String)
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB > -1) Then
        m_tBInfo(iB).sTipText = sToolTip
    End If
End Property
Private Function pbGetIndexForID(ByVal iBtnId As Long) As Long
Dim iB As Long
    pbGetIndexForID = -1
    For iB = 0 To UBound(m_tBInfo)
        If (m_tBInfo(iB).wId = iBtnId) Then
            pbGetIndexForID = iB
            Exit For
        End If
    Next iB
End Function

Public Property Get ButtonImage(ByVal vButton As Variant) As Long
Attribute ButtonImage.VB_Description = "Gets/sets the zero based index of a button's image."
Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB <> -1) Then
      ButtonImage = m_tBInfo(iB).iImage
   End If
End Property
Public Property Let ButtonImage(ByVal vButton As Variant, ByVal iImage As Long)
Dim iB As Long

   ' If we are running pre 4.71 we must remove the button and add it again.
   ' 4.71+ we can use the TB_SETBUTTONINFO method to change it on the fly:
   If (m_lMajorVer > 4) Or ((m_lMajorVer = 4) And (m_lMinorVer > 70)) Then
      Dim tBI As TBBUTTONINFO
      Dim iID As Long
      
      iB = ButtonIndex(vButton)
      If (iB <> -1) Then
         iID = m_tBInfo(iB).wId
         tBI.cbSize = Len(tBI)
         tBI.dwMask = TBIF_IMAGE
         tBI.iImage = iImage
         If (SendMessage(m_hWndToolBar, TB_SETBUTTONINFO, iID, tBI) <> 0) Then
         End If
         m_tBInfo(iB).iImage = iImage
      End If
   Else
      iB = ButtonIndex(vButton)
      If (iB <> -1) Then
         ' Delete this button...
         'RemoveButton iB
         '
      End If
      
   End If
End Property

Public Property Get ButtonCaption(ByVal vButton As Variant) As String
Attribute ButtonCaption.VB_Description = "Gets/sets the caption of a button."
Dim iB As Long
    iB = ButtonIndex(vButton)
    If (iB <> -1) Then
        ButtonCaption = m_tBInfo(iB).sCaption
    End If
End Property
Public Property Let ButtonCaption(ByVal vButton As Variant, ByVal sCaption As String)
Dim iB As Integer
Dim bEnd As Boolean

   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      
   
      If ((m_lMajorVer > 4) Or ((m_lMajorVer = 4) And (m_lMinorVer > 70))) And sCaption <> "" Then
         Dim tBI As TBBUTTONINFO
         Dim sBuf As String
         Dim iID As Long
         
         If iB <> -1 Then
            ' Remove any existing accelerator associated with caption:
            plRemoveString m_tBInfo(iB).sCaption
         
            ' don't add too many strings...
            plAddStringIfRequired m_hWndToolBar, sCaption
            If m_tBInfo(iB).bShowText Then
               sBuf = sCaption
               sBuf = sBuf & String$(80 - Len(sBuf), 0)
            Else
               sBuf = String$(80, 0)
            End If
            sBuf = StrConv(sBuf, vbFromUnicode)
            
            iID = m_tBInfo(iB).wId
            tBI.cbSize = Len(tBI)
            tBI.pszText = StrPtr(sBuf)
            tBI.dwMask = TBIF_TEXT
            If (SendMessage(m_hWndToolBar, TB_SETBUTTONINFO, iID, tBI) <> 0) Then
               m_tBInfo(iB).sCaption = sCaption
            End If
            
         End If
      Else
      
         ' Hmmm.  YOu can't remove any of the captions that have
         ' been added to the toolbar control, so if we keep on
         ' adding the damn things...  Don't change button captions
         ' to too many different things!
         Dim tBInfo As ButtonInfoStore
         LSet tBInfo = m_tBInfo(iB)
         If iB = ButtonCount - 1 Then
            bEnd = True
         End If
         RemoveButton iB
         If bEnd Then
            AddButton tBInfo.sTipText, tBInfo.iImage, , tBInfo.iLarge, sCaption, tBInfo.eStyle, tBInfo.sKey
         Else
            AddButton tBInfo.sTipText, tBInfo.iImage, iB, tBInfo.iLarge, sCaption, tBInfo.eStyle, tBInfo.sKey
         End If
      End If
   End If

End Property
Public Property Get ButtonTextVisible(ByVal vButton As Variant) As Boolean
Attribute ButtonTextVisible.VB_Description = "Gets/sets whether the caption for a button is visible or not."
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      ButtonTextVisible = m_tBInfo(iB).bShowText
   End If
End Property
Public Property Let ButtonTextVisible(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iB As Integer
Dim tBI As ButtonInfoStore
Dim bEnd As Boolean
Dim bChecked As Boolean
Dim bEnabled As Boolean
Dim bVisible As Boolean, bSet As Boolean
Dim lStyle As Long, lR As Long

   lStyle = GetWindowLong(m_hWndToolBar, GWL_STYLE)
   If (lStyle And TBSTYLE_LIST) <> TBSTYLE_LIST Then
   
      lR = SendMessageLong(m_hWndToolBar, TB_GETTEXTROWS, 0, 0)
      If bState Then
         If lR < 1 Then
            SendMessageLong m_hWndToolBar, TB_SETMAXTEXTROWS, 1, 0
            bSet = True
         End If
      Else
         If lR > 0 Then
            SendMessageLong m_hWndToolBar, TB_SETMAXTEXTROWS, 0, 0
            bSet = True
         End If
      End If
      If bSet Then
         For iB = 0 To ButtonCount - 1
            m_tBInfo(iB).bShowText = bState
         Next iB
      End If
      
   Else
   
      iB = ButtonIndex(vButton)
      If iB > -1 Then
         ShowWindow m_hWndToolBar, SW_HIDE
         If Not (m_tBInfo(iB).bControl) Then
            If bState <> m_tBInfo(iB).bShowText Then
            
               ' Hide/show text for this button:
               bChecked = ButtonChecked(iB)
               bEnabled = ButtonEnabled(iB)
               bVisible = ButtonVisible(iB)
               
               LSet tBI = m_tBInfo(iB)
               bEnd = (iB = (ButtonCount - 1))
               
               RemoveButton iB
               
               If bEnd Then
                  If bState Then
                     iB = plAddButton(m_hWndToolBar, NewButtonID, tBI.sTipText, tBI.iImage, , tBI.iLarge, tBI.sCaption, tBI.eStyle, tBI.sKey)
                  Else
                     iB = plAddButton(m_hWndToolBar, NewButtonID, tBI.sTipText, tBI.iImage, , tBI.iLarge, , tBI.eStyle, tBI.sKey)
                  End If
               Else
                  If bState Then
                     iB = plAddButton(m_hWndToolBar, NewButtonID, tBI.sTipText, tBI.iImage, iB, tBI.iLarge, tBI.sCaption, tBI.eStyle, tBI.sKey)
                  Else
                     iB = plAddButton(m_hWndToolBar, NewButtonID, tBI.sTipText, tBI.iImage, iB, tBI.iLarge, , tBI.eStyle, tBI.sKey)
                  End If
               End If
               m_tBInfo(iB).sCaption = tBI.sCaption
               
               ButtonEnabled(iB) = bEnabled
               ButtonChecked(iB) = bChecked
               ButtonVisible(iB) = bVisible
               m_tBInfo(iB).bShowText = bState
               m_tBInfo(iB).hSubMenu = tBI.hSubMenu
                              
            End If
         End If
         ShowWindow m_hWndToolBar, SW_SHOW
      End If
   End If
End Property

Public Property Get ButtonIndex(ByVal vButton As Variant) As Integer
Attribute ButtonIndex.VB_Description = "Returns the zero based index of a button given its key or position."
Dim iB As Integer
Dim iIndex As Integer
    iIndex = -1
    If (IsNumeric(vButton)) Then
        iIndex = CInt(vButton)
    Else
        For iB = 0 To UBound(m_tBInfo)
            If (m_tBInfo(iB).sKey = vButton) Then
                iIndex = iB
                Exit For
            End If
        Next iB
    End If
    If (iIndex > -1) And (iIndex <= UBound(m_tBInfo)) Then
        ButtonIndex = iIndex
    Else
        ' error
        debugmsg m_sCtlName & ",Button index failed"
        ButtonIndex = -1
    End If
    
End Property
Public Property Get ButtonKey(ByVal iButton As Long) As String
Attribute ButtonKey.VB_Description = "Returns the key of a button given its position."
   If (iButton > -1) And (iButton < ButtonCount) Then
      ButtonKey = m_tBInfo(iButton).sKey
   End If
End Property

Public Property Get ButtonEnabled(ByVal vButton As Variant) As Boolean
Attribute ButtonEnabled.VB_Description = "Gets/sets whether a button is enabled."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        ButtonEnabled = pbGetState(iID, TBSTATE_ENABLED)
    End If
End Property
Public Property Let ButtonEnabled(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
Dim lEnable As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        pbSetState iID, TBSTATE_ENABLED, bState
    End If
End Property
Public Property Get ButtonVisible(ByVal vButton As Variant) As Boolean
Attribute ButtonVisible.VB_Description = "Gets/sets whether a button is visible in the toolbar."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        ButtonVisible = Not (pbGetState(iID, TBSTATE_HIDDEN))
    End If
End Property
Public Property Let ButtonVisible(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
Dim i As Long
Dim j As Long
Dim bPriorSeparator As Boolean
Dim bNextSeparator As Boolean
Dim bHiddenSeparator As Boolean
Dim iNextSeparator As Long
    
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        
        pbSetState iID, TBSTATE_HIDDEN, Not (bState)
        
        If (m_tBInfo(iButton).eStyle <> CTBSeparator) Then
            If Not (bState) Then
               ' if the prior visible button is a separator, and the next one is also,
               ' then we hide the next separator:
               bPriorSeparator = True
               For i = iButton - 1 To 0 Step -1
                  If (ButtonVisible(i)) Then
                     If (m_tBInfo(i).eStyle = CTBSeparator) Then
                        bPriorSeparator = True
                     Else
                        bPriorSeparator = False
                     End If
                     Exit For
                  End If
               Next i
               
               bNextSeparator = False
               For i = iButton + 1 To ButtonCount - 1
                  If (ButtonVisible(i)) Then
                     If (m_tBInfo(i).eStyle = CTBSeparator) Then
                        bNextSeparator = True
                        iNextSeparator = i
                     End If
                     Exit For
                  End If
               Next i
               
               If (bPriorSeparator And bNextSeparator) Then
                  pbSetState m_tBInfo(iNextSeparator).wId, TBSTATE_HIDDEN, True
               End If
               
            Else
               ' check for a hidden separator followed by a visible button:
               For i = iButton + 1 To ButtonCount - 1
                  If (ButtonVisible(i)) Then
                     Exit For
                  Else
                     If (m_tBInfo(i).eStyle = CTBSeparator) Then
                        bHiddenSeparator = True
                        iNextSeparator = i
                        Exit For
                     End If
                  End If
               Next i
               
               If (bHiddenSeparator) Then
                  ' check that the next visible button is not also a separator
                  For i = iNextSeparator + 1 To ButtonCount - 1
                     If (ButtonVisible(i)) Then
                        If (m_tBInfo(i).eStyle = CTBSeparator) Then
                           bHiddenSeparator = False
                        End If
                     End If
                     Exit For
                  Next i
                  If (bHiddenSeparator) Then
                     pbSetState m_tBInfo(iNextSeparator).wId, TBSTATE_HIDDEN, False
                  End If
               End If
               
            End If
        End If
        
        ResizeToolbar
    End If
    
End Property
Private Property Get plButtonVisible(ByVal hWndToolbar As Long, ByVal lBtnIndex As Long) As Boolean
Dim tBB As TBBUTTON
      
   SendMessage m_hWndToolBar, TB_GETBUTTON, lBtnIndex, tBB
   plButtonVisible = (SendMessageLong(hWndToolbar, TB_ISBUTTONHIDDEN, tBB.idCommand, 0) = 0)

End Property
Public Property Get ButtonWidth(ByVal vButton As Variant)
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wId, tR
      ButtonWidth = tR.right - tR.left
      moveChildWindow iButton
   End If
End Property
Public Property Let ButtonWidth(ByVal vButton As Variant, ByVal lWidth As Variant)
' the width parameter should be a long for pixels, but the original was
' compiled with the property Get as a variant... forgot to type the
' vartype - doh!
Dim iButton As Long
Dim tR As RECT
Dim tWR As RECT
Dim lhWnd As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      Dim tBB As TBBUTTONINFO
      tBB.cbSize = LenB(tBB)
      tBB.dwMask = TBIF_SIZE
      SendMessage m_hWndToolBar, TB_GETBUTTONINFO, m_tBInfo(iButton).wId, tBB
      If Not (tBB.cx = lWidth) Then
         tBB.cx = lWidth
         SendMessage m_hWndToolBar, TB_SETBUTTONINFO, m_tBInfo(iButton).wId, tBB
         If Not (m_tBInfo(iButton).hWndCapture = 0) Then
            moveChildWindow iButton
         End If
      End If
   End If
End Property
Public Property Get ButtonHeight(ByVal vButton As Variant) As Long
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wId, tR
      ButtonHeight = tR.bottom - tR.top
   End If
End Property
Public Property Get ButtonLeft(ByVal vButton As Variant) As Long
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wId, tR
      ButtonLeft = tR.left
   End If
End Property
Public Property Get ButtonTop(ByVal vButton As Variant) As Long
Dim iButton As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage m_hWndToolBar, TB_GETRECT, m_tBInfo(iButton).wId, tR
      ButtonTop = tR.top
   End If
End Property
Public Property Get ButtonHot(ByVal vButton As Variant) As Boolean
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      ButtonHot = (SendMessageLong(m_hWndToolBar, TB_GETHOTITEM, 0, 0) = iB)
   End If
End Property
Public Property Let ButtonHot(ByVal vButton As Variant, ByVal bHot As Boolean)
Attribute ButtonHot.VB_Description = "Gets/sets whether a button in a flat toolbar appears in the ""hot"" state (i.e. looks like the mouse is over it)"
Dim iB As Integer
   iB = ButtonIndex(vButton)
   If iB > -1 Then
      If ButtonHot(iB) Then
         If Not bHot Then
            SendMessageLong m_hWndToolBar, TB_SETHOTITEM, -1, 0
         End If
      Else
         If bHot Then
            SendMessageLong m_hWndToolBar, TB_SETHOTITEM, iB, 0
         End If
      End If
   End If
End Property
Public Property Get MaxButtonWidth() As Long
Attribute MaxButtonWidth.VB_Description = "Gets/sets the maximum allowable button width."
Dim i As Long
Dim lW As Long
Dim lMaxW As Long
   For i = 0 To ButtonCount - 1
      lW = ButtonWidth(i)
      If lW > lMaxW Then
         lMaxW = lW
      End If
   Next i
   MaxButtonWidth = lMaxW
End Property
Public Property Get MaxButtonHeight() As Long
Attribute MaxButtonHeight.VB_Description = "Gets/sets the maximum allowable button height."
Dim i As Long
Dim lH As Long
Dim lMaxH As Long
   For i = 0 To ButtonCount - 1
      lH = ButtonHeight(i)
      If lH > lMaxH Then
         lMaxH = lH
      End If
   Next i
   MaxButtonHeight = lMaxH
End Property
Public Property Get ButtonChecked(ByVal vButton As Variant) As Boolean
Attribute ButtonChecked.VB_Description = "Gets/sets whether a button is checked or not (if the button has the checked or check group style)"
   ButtonChecked = plButtonChecked(m_hWndToolBar, vButton)
End Property
Private Property Get plButtonChecked(ByVal hWndToolbar As Long, ByVal vButton As Variant) As Boolean
Dim iButton As Long
Dim iID As Long
Dim tBB As TBBUTTON
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage hWndToolbar, TB_GETBUTTON, iButton, tBB
      iID = tBB.idCommand 'm_tBInfo(iButton).wID
      plButtonChecked = pbGetState2(hWndToolbar, iID, TBSTATE_CHECKED)
   End If
End Property
Public Property Let ButtonChecked(ByVal vButton As Variant, ByVal bState As Boolean)
   plButtonChecked(m_hWndToolBar, vButton) = bState
End Property
Private Property Let plButtonChecked(ByVal hWndToolbar As Long, ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
Dim tBB As TBBUTTON
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      SendMessage hWndToolbar, TB_GETBUTTON, iButton, tBB
      iID = tBB.idCommand
      'Check the button
      SendMessageLong hWndToolbar, TB_CHECKBUTTON, iID, Abs(bState)
      If (ButtonPressed(iButton) <> bState) Then
         SendMessageLong hWndToolbar, TB_CHECKBUTTON, iID, Abs(bState)
      End If
   End If
End Property
Public Property Get ButtonPressed(ByVal vButton As Variant) As Boolean
Attribute ButtonPressed.VB_Description = "Gets/sets whether a button is pressed."
   ButtonPressed = plButtonPressed(m_hWndToolBar, vButton)
End Property
Private Property Get plButtonPressed(ByVal hWndToolbar As Long, ByVal vButton As Variant) As Boolean
Dim iButton As Long
Dim iID As Long
Dim tBB As TBBUTTON
   If (hWndToolbar = m_hWndToolBar) Then
      iButton = ButtonIndex(vButton)
   Else
      iButton = vButton
   End If
   If (iButton <> -1) Then
      SendMessage hWndToolbar, TB_GETBUTTON, iButton, tBB
      iID = tBB.idCommand
      plButtonPressed = pbGetState2(hWndToolbar, iID, TBSTATE_PRESSED)
   End If
End Property
Public Property Let ButtonPressed(ByVal vButton As Variant, ByVal bState As Boolean)
   plButtonPressed(m_hWndToolBar, vButton) = bState
End Property
Private Property Let plButtonPressed(ByVal hWndToolbar As Long, ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        pbSetState2 hWndToolbar, iID, TBSTATE_PRESSED, bState
    End If
End Property
Public Property Get ButtonStyle(ByVal vButton As Variant) As ECTBToolButtonSyle
Dim iButton As Long
Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      Dim tBI As TBBUTTONINFO
      iID = m_tBInfo(iButton).wId
      tBI.cbSize = LenB(tBI)
      tBI.dwMask = TBIF_STYLE
      If (SendMessage(m_hWndToolBar, TB_GETBUTTONINFO, iID, tBI) = iButton) Then
         ButtonStyle = tBI.fsStyle
      End If
   End If
End Property
Public Property Let ButtonStyle(ByVal vButton As Variant, ByVal eStyle As ECTBToolButtonSyle)
Dim iButton As Long
Dim iID As Long
Dim tR As RECT
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      Dim tBI As TBBUTTONINFO
      iID = m_tBInfo(iButton).wId
      tBI.cbSize = LenB(tBI)
      tBI.dwMask = TBIF_STYLE
      tBI.fsStyle = eStyle
      If m_tBInfo(iButton).bShowText = False And (GetWindowLong(m_hWndToolBar, GWL_STYLE) And TBSTYLE_LIST) = TBSTYLE_LIST Then
         tBI.dwMask = tBI.dwMask Or TBIF_SIZE
         SendMessage m_hWndToolBar, TB_GETITEMRECT, iButton, tR
         tBI.cx = tR.right - tR.left
      End If
      SendMessage m_hWndToolBar, TB_SETBUTTONINFO, iID, tBI
      m_tBInfo(iButton).eStyle = tBI.fsStyle
   End If
End Property
Public Property Get ButtonControl(ByVal vButton As Variant) As Long
Dim iButton As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      ButtonControl = m_tBInfo(iButton).hWndCapture
   End If
End Property
Public Property Get ButtonTextWrap(ByVal vButton As Variant) As Boolean
Attribute ButtonTextWrap.VB_Description = "Gets/sets whether button text will wrap onto a newline if it is too long."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        ButtonTextWrap = pbGetState(iID, TBSTATE_WRAP)
    End If
End Property
Public Property Let ButtonTextWrap(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        pbSetState iID, TBSTATE_WRAP, bState
    End If
End Property
Public Property Get ButtonTextEllipses(ByVal vButton As Variant) As Boolean
Attribute ButtonTextEllipses.VB_Description = "Gets/sets whether button text will be truncated if the button text is too long."
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        ButtonTextEllipses = pbGetState(iID, TBSTATE_ELLIPSES)
    End If
End Property
Public Property Let ButtonTextEllipses(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iButton As Long
Dim iID As Long
    iButton = ButtonIndex(vButton)
    If (iButton <> -1) Then
        iID = m_tBInfo(iButton).wId
        pbSetState iID, TBSTATE_ELLIPSES, bState
    End If
End Property
Private Function pbGetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates) As Boolean
Dim fState As Long
    fState = SendMessageLong(m_hWndToolBar, TB_GETSTATE, iIDBtn, 0)
    pbGetState = ((fState And fStateFlag) = fStateFlag)
End Function
Private Function pbGetState2(ByVal hWndToolbar As Long, ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates) As Boolean
Dim fState As Long
    fState = SendMessageLong(hWndToolbar, TB_GETSTATE, iIDBtn, 0)
    pbGetState2 = ((fState And fStateFlag) = fStateFlag)
End Function
Private Function pbSetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates, ByVal bState As Boolean)
Dim fState As Long
    fState = SendMessageLong(m_hWndToolBar, TB_GETSTATE, iIDBtn, 0)
    If (bState) Then
        fState = fState Or fStateFlag
    Else
        fState = fState And Not fStateFlag
    End If
    If (SendMessageLong(m_hWndToolBar, TB_SETSTATE, iIDBtn, fState) = 0) Then
        debugmsg m_sCtlName & ",Button state failed"
    Else
        pbSetState = True
    End If
End Function
Private Function pbSetState2(ByVal hWndToolbar As Long, ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates, ByVal bState As Boolean)
Dim fState As Long
    fState = SendMessageLong(hWndToolbar, TB_GETSTATE, iIDBtn, 0)
    If (bState) Then
        fState = fState Or fStateFlag
    Else
        fState = fState And Not fStateFlag
    End If
    If (SendMessageLong(hWndToolbar, TB_SETSTATE, iIDBtn, fState) = 0) Then
        debugmsg m_sCtlName & ",Button state failed"
    Else
        pbSetState2 = True
    End If
End Function
 
 
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the window handle of the control."
    hwnd = m_hWndToolBar
End Property

Public Property Get TitleBarModifier() As Boolean
   TitleBarModifier = g_bTitleBarModifier
End Property
Public Property Let TitleBarModifier(ByVal bState As Boolean)
   g_bTitleBarModifier = bState
   If bState Then
      'AttachTitleBarMod m_hWndParentForm
   Else
      'DetachTitleBarMod m_hWndParentForm
   End If
End Property

Public Sub DestroyToolBar()
Attribute DestroyToolBar.VB_Description = "Destroys the toolbar and all resources associated with it."
Dim i As Long
Dim iU As Long

'On Error Resume Next
'We need to clean up our windows
   debugmsg m_sCtlName & ",DestroyToolBar"
   ' Chevron:
   If Not m_cChevronWindow Is Nothing Then
      m_cChevronWindow.Destroy
      Set m_cChevronWindow = Nothing
   End If
   
   pSubClass False
   If (m_hWndToolBar <> 0) Then
      ' Remove from tooltip:
      RemoveFromToolTip m_hWndToolBar
            
      ' Clear me from keyboard hook:
      DetachKeyboardHook Me
      
      If Not (m_lPtrMenu = 0) Then
         RemoveProp m_hWndToolBar, "vbalTbar:OwnsMenu:" & m_lPtrMenu
         m_lPtrMenu = 0
      End If
      ' Can't use button count - the buttons can all be removed before
      ' we get here!
      iU = UBound(m_tBInfo)
      For i = 0 To iU
         If Not (m_tBInfo(i).hWndCapture = 0) Then
            debugmsg m_sCtlName & ",Resetting parent:" & m_tBInfo(i).hWndCapture
            If Not (IsWindow(m_tBInfo(i).hWndParentOrig) = 0) Then
               SetParent m_tBInfo(i).hWndCapture, m_tBInfo(i).hWndParentOrig
            End If
         End If
      Next i
      ShowWindow m_hWndToolBar, SW_HIDE
      SetParent m_hWndToolBar, 0
      DestroyWindow m_hWndToolBar
      RemoveProp m_hWndToolBar, "vbalTbar:ControlPtr"
      RemoveProp m_hWndToolBar, "vbalTBar:MDIClient"
      RemoveProp m_hWndToolBar, "vbalTBar:NotifyWindow"
      m_hWndToolBar = 0
   End If
   If Not (m_hWndParentForm = 0) Then
      RemoveProp m_hWndParentForm, "vbalTbar:MDIClient"
      m_hWndParentForm = 0
   End If
   Set m_cMenu = Nothing
   
   Err.Clear
   On Error GoTo 0
End Sub
Public Sub CreateFromMenu( _
      ByRef cMenu As Object _
   )
Attribute CreateFromMenu.VB_Description = "Sets up a toolbar based on a cPopupMenu object so the toolbar can act as the form's menu."
   CreateFromMenu2 cMenu, CTBMenuStyle
   m_bCreateFromMenu2 = False
End Sub
Public Sub CreateFromMenu2( _
      ByRef cMenu As Object, _
      Optional ByVal eStyle As ECTBToolbarFromMenuStyle, _
      Optional ByVal sMenuParentKey As String _
   )
Dim i As Long
Dim lIndexSearch As Long
Dim hSubMenu As Long
Dim sCaption As String
Dim iPos As Long
Dim bEnabled As Boolean
Dim bVisible As Boolean
Dim sKey As String
Dim iIcon As Long
Dim tMII As MENUITEMINFO
Dim lR As Long
Dim lID As Long
Dim iB As Long
Dim eBtnStyle As ECTBToolButtonSyle
Dim lThisGroupCount As Long
Dim lThisGroup() As Long
Dim iThisGroupCheckIndex As Long
Dim iThis As Long
Dim sHelpText As String
Dim lhWndLock As Long
   
   If Not (m_lPtrMenu = 0) Then
      RemoveProp m_hWndParentForm, "vbalTbar:OwnsMenu:" & m_lPtrMenu
      m_lPtrMenu = 0
   End If
   
   If m_hWndToolBar = 0 Then
      If eStyle = CTBMenuStyle Then
         If (DrawStyle = CTBDrawStandard) Then
            DrawStyle = CTBDrawNoVisualStyles
         End If
         CreateToolbar , True, True, True, 0
      Else
         CreateToolbar , True, True, True
      End If
   Else
      If IsWindowVisible(m_hWndToolBar) Then
         LockWindowUpdate m_hWndToolBar
         lhWndLock = m_hWndToolBar
      End If
      ' remove all buttons:
      For i = ButtonCount - 1 To 0 Step -1
         RemoveButton i
      Next i
   End If
   
   iThisGroupCheckIndex = -1
   
   ' Now add buttons according to menu:
   With cMenu
      
      If .count > 0 Then
         
         If sMenuParentKey <> "" Then
            lIndexSearch = .IndexForKey(sMenuParentKey)
            For i = 1 To .count
               If (.ItemParentIndex(i) = lIndexSearch) Then
                  m_hMenu = .hMenu(i)
                  Exit For
               End If
            Next i
         Else
            m_hMenu = .hMenu(1)
         End If
         m_eCreateFromMenuStyle = eStyle
         m_bCreateFromMenu2 = True
         
         For i = 1 To .count
            
            ' Is top level menu?
            If .hMenu(i) = m_hMenu Then
            
               ' Get info about menu item:
               iB = -1
               sCaption = .Caption(i)
               sKey = .ItemKey(i)
               sHelpText = .HelpText(i)
               lID = .IDForItem(i)
               bVisible = .Visible(i)
               ' Find if this menu has submenus:
               tMII.fMask = MIIM_SUBMENU Or MIIM_STATE
               tMII.cbSize = LenB(tMII)
               lR = GetMenuItemInfo(.hMenu(i), lID, False, tMII)
               hSubMenu = tMII.hSubMenu
               bEnabled = ((tMII.fState And &H1) = &H0)
                                                            
               If (sCaption = "-") Then
                  eBtnStyle = CTBSeparator
               Else
                  eBtnStyle = CTBAutoSize
                  If eStyle = CTBToolbarStyle Then
                     If Not (hSubMenu = 0) Then
                        eBtnStyle = eBtnStyle Or CTBDropDown
                     End If
                  End If
               End If
                                                            
               ' Add the button:
               If eStyle = CTBMenuStyle Then
                  iB = plAddButton(m_hWndToolBar, NewButtonID, , , , , sCaption, CTBAutoSize, sKey)
               Else
                  iIcon = .ItemIcon(i)
                  iB = plAddButton(m_hWndToolBar, NewButtonID, sHelpText, iIcon, , , sCaption, eBtnStyle, sKey)
                  If eBtnStyle = CTBSeparator Then
                     If iThisGroupCheckIndex > -1 Then
                        For iThis = 1 To lThisGroupCount
                           ButtonStyle(lThisGroup(iThis)) = CTBCheckGroup Or CTBAutoSize  'ButtonStyle(lThisGroup(iThis) Or CTBCheckGroup)
                        Next iThis
                        ButtonChecked(iThisGroupCheckIndex) = True
                     End If
                     lThisGroupCount = 0
                     iThisGroupCheckIndex = -1
                  Else
                     lThisGroupCount = lThisGroupCount + 1
                     ReDim Preserve lThisGroup(1 To lThisGroupCount) As Long
                     lThisGroup(lThisGroupCount) = iB
                     If .RadioCheck(i) Then
                        iThisGroupCheckIndex = iB
                     ElseIf .Checked(i) Then
                        ButtonChecked(iB) = True
                     End If
                  End If
               End If
               ButtonVisible(iB) = bVisible
               
               'Debug.Print "Added " & sCaption, iB, bEnabled
               
               If iB > -1 Then
                  m_tBInfo(iB).hSubMenu = hSubMenu
                  ButtonEnabled(iB) = bEnabled
                  If eStyle = CTBToolbarStyle Then
                     If (GetWindowLong(m_hWndToolBar, GWL_STYLE) And TBSTYLE_LIST) = TBSTYLE_LIST Then
                        ButtonTextVisible(iB) = False
                     End If
                  End If
               End If
            End If
            
         Next i

      End If
   End With
   
   If lhWndLock <> 0 Then
      LockWindowUpdate 0
   End If
   
   ' Store a reference to the item:
   m_lPtrMenu = ObjPtr(cMenu)
   SetProp m_hWndParentForm, "vbalTbar:OwnsMenu:" & m_lPtrMenu, ObjPtr(Me)
   
End Sub
Public Property Get DropDownAlign() As ECTBDropDownAlign
   '
   DropDownAlign = m_eDropDownAlign
   '
End Property
Public Property Let DropDownAlign(ByVal eAlign As ECTBDropDownAlign)
   m_eDropDownAlign = eAlign
End Property
Public Sub CreateToolbar( _
      Optional ButtonSize As Integer = 16, _
      Optional StyleList As Boolean, _
      Optional WithText As Boolean, _
      Optional Wrappable As Boolean, _
      Optional PicSize As Integer)
Attribute CreateToolbar.VB_Description = "Initialises a toolbar for use."
On Error Resume Next
Dim Button As TBBUTTON
Dim lParam As Long
Dim ListButtons As Boolean
Dim dwStyle As Long
Dim dwExStyle As Long
Dim lExStyle As Long
Dim lhWndClient As Long
Dim hWndParent As Long

   DestroyToolBar

   m_bWrappable = Wrappable
   m_bListStyle = StyleList

   m_bWithText = WithText

   hWndParent = getFormParenthWnd(UserControl.hwnd)

   dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN
   dwStyle = dwStyle Or CCS_NOPARENTALIGN Or CCS_NORESIZE Or CCS_NODIVIDER
   dwStyle = dwStyle Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT
   'dwStyle = dwStyle Or CCS_ADJUSTABLE
   If (StyleList) Then
      dwStyle = dwStyle Or TBSTYLE_LIST
   End If
   If (Wrappable) Then
      dwStyle = dwStyle Or TBSTYLE_WRAPABLE
   End If

   dwExStyle = WS_EX_TOOLWINDOW
   lExStyle = GetWindowLong(hWndParent, GWL_EXSTYLE)
   lExStyle = lExStyle And (WS_EX_RIGHT Or WS_EX_RTLREADING)
   dwExStyle = dwExStyle Or lExStyle

   m_hWndToolBar = CreateWindowEX(dwExStyle, "ToolbarWindow32", "", _
         dwStyle, _
         0, 0, 0, 0, UserControl.hwnd, 0&, App.hInstance, 0&)
         
   If Not (m_hWndToolBar = 0) Then
   
      DrawStyle = m_eVisualStyle
      m_cNCM.GetMetrics
    
      SendMessageLong m_hWndToolBar, TB_SETPARENT, UserControl.hwnd, 0
  
      m_lR = SendMessageLong(m_hWndToolBar, TB_BUTTONSTRUCTSIZE, LenB(Button), 0)
     
      AddBitmapIfRequired m_hWndToolBar
      m_lOrigButtonSize = ButtonSize
      If m_eImageSourceType <> -1 Then
         lParam = ButtonSize + (ButtonSize * &H10000)
      Else
         lParam = 0
      End If
      m_lR = SendMessageLong(m_hWndToolBar, TB_SETBITMAPSIZE, 0, lParam)

      SetProp m_hWndToolBar, "vbalTbar:ControlPtr", ObjPtr(Me)
      m_hWndParentForm = hWndParent
      lhWndClient = FindWindowEx(m_hWndParentForm, 0, "MDIClient", ByVal 0&)
      SetProp m_hWndToolBar, "vbalTbar:MDIClient", lhWndClient
      SetProp m_hWndToolBar, "vbalTbar:NotifyWindow", UserControl.hwnd
   
      pSubClass True, UserControl.hwnd
      AddToToolTip m_hWndToolBar
      
      ' Start checking for accelerator key presses here:
      AttachKeyboardHook Me

      Set m_cMenu = New cTbarMenu
      
   End If
   
End Sub
Public Property Get ListStyle() As Boolean
   ListStyle = pbIsStyle(TBSTYLE_LIST)
End Property
Public Property Let ListStyle(ByVal bState As Boolean)
   pbSetStyle TBSTYLE_LIST, bState
   m_bListStyle = bState
End Property
Public Property Get Wrappable() As Boolean
   Wrappable = pbIsStyle(TBSTYLE_WRAPABLE)
End Property
Public Property Let Wrappable(ByVal bState As Boolean)
   pbSetStyle TBSTYLE_WRAPABLE, bState
   m_bWrappable = bState
End Property
Private Function pbSetStyle(ByVal lStyleBit As Long, ByVal bState As Boolean) As Boolean
Dim lS As Long
Dim iB As Long
   If Not pbIsStyle(lStyleBit) = bState Then
      lS = GetWindowLong(m_hWndToolBar, GWL_STYLE)
      If bState Then
         lS = lS Or lStyleBit
      Else
         lS = lS And Not lStyleBit
      End If
      SetWindowLong m_hWndToolBar, GWL_STYLE, lS
      ShowWindow m_hWndToolBar, 0
      Dim i As Long
      For iB = 0 To ButtonCount - 1
         ButtonTextVisible(iB) = Not (ButtonTextVisible(iB))
         ButtonTextVisible(iB) = Not (ButtonTextVisible(iB))
      Next iB
      ShowWindow m_hWndToolBar, 1
      ResizeToolbar
   End If
End Function
Private Function pbIsStyle(ByVal lStyleBit As Long) As Boolean
Dim lS As Long
   If m_hWndToolBar <> 0 Then
      lS = GetWindowLong(m_hWndToolBar, GWL_STYLE)
      If (lS And lStyleBit) = lStyleBit Then
         pbIsStyle = True
      End If
   End If
End Function
Public Property Let ImageSource( _
        ByVal eType As ECTBImageSourceTypes _
    )
Attribute ImageSource.VB_Description = "Sets the type of image (file, picture, resource, image list or standard image list) to be used as the source of the button's images."
    m_eImageSourceType = eType
End Property
Public Property Let ImageResourceID(ByVal lResourceId As Long)
Attribute ImageResourceID.VB_Description = "Sets a resource id to be used as the source of the button's images."
    m_lResourceID = lResourceId
End Property
Public Property Let ImageResourcehInstance(ByVal hInstance As Long)
Attribute ImageResourcehInstance.VB_Description = "Sets the hInstance of the binary containing the resource specified in ImageResourceID."
   m_hInstance = hInstance
End Property
Public Property Let ImageFile(ByVal sFile As String)
Attribute ImageFile.VB_Description = "Sets a bitmap file to be used as the source of the buttons images."
    m_sFileName = sFile
End Property
Public Sub SetImageList( _
      ByVal vThis As Variant, _
      Optional ByVal eType As ECTBImageListTypes = CTBImageListNormal _
   )
Attribute SetImageList.VB_Description = "Sets the image list to be used for standard, hot or disabled button images."
Dim hIml As Long
    
    m_ptrVb6ImageList = 0

   ' Set the ImageList handle property either from a VB
   ' image list or directly:
   If VarType(vThis) = vbObject Then
       ' Assume VB ImageList control.  Note that unless
       ' some call has been made to an object within a
       ' VB ImageList the image list itself is not
       ' created.  Therefore hImageList returns error. So
       ' ensure that the ImageList has been initialised by
       ' drawing into nowhere:
      On Error Resume Next
      ' Get the image list initialised..
      vThis.ListImages(1).Draw 0, 0, 0, 1
      hIml = vThis.hImageList
      If (Err.Number <> 0) Then
         Err.Clear
         hIml = vThis.hIml
         If Err.Number <> 0 Then
             hIml = 0
         End If
      Else
         ' Check for VB6 image list:
         If (TypeName(vThis) = "ImageList") Then
             If (vThis.ListImages.count <> ImageList_GetImageCount(hIml)) Then
                 Dim o As Object
                 Set o = vThis
                 If (eType = CTBImageListNormal) Then
                     m_ptrVb6ImageList = ObjPtr(o)
                 End If
             End If
         End If
      End If
      On Error GoTo 0
       
   ElseIf VarType(vThis) = vbLong Then
       ' Assume ImageList handle:
       hIml = vThis
   Else
       Err.Raise vbObjectError + 1049, "cToolbar." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
   End If
    
   If Not (hIml = 0) Then
      If (m_ptrVb6ImageList <> 0) Then
         m_lIconHeight = vThis.ImageHeight
         m_lIconWidth = vThis.ImageWidth
         m_lTransColor = vThis.BackColor
      Else
         Dim rc As RECT
         ImageList_GetImageRect hIml, 0, rc
         m_lIconHeight = rc.bottom - rc.top
         m_lIconWidth = rc.right - rc.left
         m_lTransColor = -1
      End If
   End If
    
   ' If we have a valid image list, then associate it with the control:
   Select Case eType
   Case CTBImageListDisabled
      m_hImlDis = hIml
   Case CTBImageListHot
      m_hImlHot = hIml
   Case CTBImageListNormal
      m_hIml = hIml
   End Select
   
   If Not (m_hWndToolBar = 0) Then
      AddBitmapIfRequired m_hWndToolBar
   End If
      
End Sub
Public Property Let ImagePicture(ByVal picThis As StdPicture)
Attribute ImagePicture.VB_Description = "Sets a picture object to be used as the source of the button's images."
    Set m_pic = picThis
End Property
Public Property Set ImagePicture(ByVal picThis As StdPicture)
    Set m_pic = picThis
End Property
Public Property Let ImageStandardBitmapType(ByVal eType As ECTBStandardImageSourceTypes)
Attribute ImageStandardBitmapType.VB_Description = "Sets the standard image list bitmap to be used to generate the button images."
   m_eStandardType = eType
End Property


Private Sub AddBitmapIfRequired(ByVal lhWndToolbar As Long)
Dim tbab As TBADDBITMAP
    
   Set m_cMemDC = Nothing
   
   Select Case m_eImageSourceType
   Case CTBStandardImageSources
      SendMessageLong lhWndToolbar, TB_LOADIMAGES, m_eStandardType, HINST_COMMCTRL
   Case CTBPicture
      tbab.hInst = 0
      tbab.nID = hBmpFromPicture(m_pic)
      Set m_cMemDC = New cAlphaDibSection
      m_cMemDC.CreateFromPicture m_pic
      m_cMemDC.MakeTransparent
      ' Add the bitmap containing button images to the toolbar.
      m_lR = SendMessage(lhWndToolbar, TB_ADDBITMAP, 54, tbab)
   Case CTBLoadFromFile
      tbab.hInst = 0
      tbab.nID = LoadImage(0, m_sFileName, IMAGE_BITMAP, 0, 0, _
                   LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
      Set m_cMemDC = New cAlphaDibSection
      m_cMemDC.CreateFromHBitmap tbab.nID
      m_cMemDC.MakeTransparent
      m_lR = SendMessage(lhWndToolbar, TB_ADDBITMAP, 54, tbab)
   Case CTBResourceBitmap
      tbab.hInst = 0
      tbab.nID = LoadImageLong(m_hInstance, m_lResourceID, IMAGE_BITMAP, 0, 0, _
                    LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
      Set m_cMemDC = New cAlphaDibSection
      m_cMemDC.CreateFromHBitmap tbab.nID
      m_cMemDC.MakeTransparent
      m_lR = SendMessage(lhWndToolbar, TB_ADDBITMAP, 54, tbab)
   Case CTBExternalImageList
      ' Get the size of the image list:
      If m_hIml <> 0 Then
         Set m_cMemDC = New cAlphaDibSection
         m_cMemDC.Create m_lIconWidth, m_lIconHeight
         SendMessageLong lhWndToolbar, CTBImageListNormal, 0, m_hIml
      End If
      If m_hImlHot <> 0 Then
         SendMessageLong lhWndToolbar, CTBImageListHot, 0, m_hImlHot
      End If
      If m_hImlDis <> 0 Then
         SendMessageLong lhWndToolbar, CTBImageListDisabled, 0, m_hImlDis
      End If
   End Select
    
End Sub

Public Sub RemoveButton(ByVal vButton As Variant)
Attribute RemoveButton.VB_Description = "Removes a button from the toolbar."
Dim iB As Integer
Dim iCount As Long
Dim iNewCount As Long
Dim i As Long
Dim iT As Long
Dim sCaption As String
   
   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      iCount = ButtonCount
      
      If iCount <= 0 Then
         Debug.Assert iCount > 0
      Else
         If Not (m_tBInfo(iB).hWndCapture = 0) Then
            'SetParent m_tBInfo(iB).hWndCapture, m_tBInfo(iB).hWndParentOrig
         End If
      
         sCaption = m_tBInfo(iB).sCaption
         m_lR = SendMessageLong(m_hWndToolBar, TB_DELETEBUTTON, iB, 0)
         If m_lMajorVer < 4 Or (m_lMajorVer = 4 And m_lMinorVer < 71) Then
            iNewCount = ButtonCount
            If iNewCount = 0 Then
               Erase m_tBInfo
            Else
               For i = iB To iNewCount - 1
                  LSet m_tBInfo(i) = m_tBInfo(i + 1)
               Next i
               ReDim Preserve m_tBInfo(0 To iNewCount - 1) As ButtonInfoStore
            End If
            plRemoveString sCaption
         End If
      End If
   End If
   
End Sub

Public Sub AddControl( _
      ByVal lhWnd As Long, _
      Optional ByVal vButtonBefore As Variant, _
      Optional ByVal sKey As String = "" _
    )
Attribute AddControl.VB_Description = "Adds a control (such as a combo box) to the toolbar, optionally setting the control's key and which button it is added before."
Dim lButton As Long
   lButton = plAddButton(m_hWndToolBar, NewButtonID, , , vButtonBefore, , , CTBNormal, sKey)
   If lButton > -1 Then
      SetControlSub lhWnd, lButton
   End If
End Sub

Public Sub SetControl( _
      ByVal lhWnd As Long, _
      ByVal vButton As Variant _
   )
Attribute SetControl.VB_Description = "Places a control over the specified button.  Similar to AddControl, but modifies an existing button."
Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB <> -1) Then
      SetControlSub lhWnd, iB
   End If
End Sub
   
Private Sub SetControlSub(ByVal lhWnd As Long, ByVal lButton As Long)
Dim tR As RECT
Dim lhWndParent As Long
   ButtonEnabled(lButton) = False
   GetWindowRect lhWnd, tR
   ButtonWidth(lButton) = tR.right - tR.left
   If Not (lhWnd = 0) Then
      lhWndParent = GetParent(lhWnd)
      SetParent lhWnd, m_hWndToolBar
   End If
   With m_tBInfo(lButton)
      .bControl = True
      .hWndCapture = lhWnd
      .hWndParentOrig = lhWndParent
      .xWidth = tR.right - tR.left + 2
   End With
   If Not (lhWnd = 0) Then
      moveChildWindow lButton
   End If
End Sub

Public Property Get ControlStretch(ByVal vButton As Variant) As Boolean
Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB <> -1) Then
      ControlStretch = m_tBInfo(iB).bStretch
   End If
End Property
Public Property Let ControlStretch(ByVal vButton As Variant, ByVal bState As Boolean)
Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB <> -1) Then
      m_tBInfo(iB).bStretch = bState
   End If
End Property
Private Function plAddButton( _
      ByVal hWndToolbar As Long, _
      ByVal lIDCommand As Long, _
      Optional ByVal sTip As String = "", _
      Optional ByVal iImage As Integer = -1, _
      Optional ByVal vButtonBefore As Variant, _
      Optional ByVal xLarge As Integer = 0, _
      Optional ByVal sButtonText As String, _
      Optional ByVal eButtonStyle As ECTBToolButtonSyle, _
      Optional ByVal sKey As String = "" _
   ) As Long
Dim tB As TBBUTTON
Dim lParam As Long
Dim iB As Integer, i As Integer
Dim bInsert As Boolean
Dim iCount As Long
Dim idString As Long

   plAddButton = -1

   iCount = plButtonCount(hWndToolbar)
   If iCount = 0 Then
      ' Make sure we can have drop-down buttons:
      SendMessageLong hWndToolbar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS
   End If

   ' Are we adding or inserting?
   If Not (IsMissing(vButtonBefore)) Then
      iB = ButtonIndex(vButtonBefore)
      If (iB > -1) Then
         bInsert = True
      End If
   End If
     
   ' Do we need to add a new string for this button?
   idString = -1
   If Len(sButtonText) > 0 Then
      idString = plAddStringIfRequired(hWndToolbar, sButtonText)
   End If
 
   With tB
      .iBitmap = iImage
      .idCommand = lIDCommand
      .fsState = TBSTATE_ENABLED
      .fsStyle = eButtonStyle
      .dwData = 0
      .iString = idString
   End With
   
   If (bInsert) Then
      m_lR = SendMessage(hWndToolbar, TB_INSERTBUTTON, iB, tB)
      If (m_lR <> 0) Then
         If hWndToolbar = m_hWndToolBar Then
            ' We need to insert into the structure:
            ReDim Preserve m_tBInfo(0 To iCount) As ButtonInfoStore
            For i = iCount To iB + 1 Step -1
               LSet m_tBInfo(i) = m_tBInfo(i - 1)
            Next i
            With m_tBInfo(iB)
               .wId = tB.idCommand
               .iImage = iImage
               .sTipText = sTip
               .iLarge = xLarge
               .sKey = sKey
               .bShowText = m_bWithText
               .sCaption = sButtonText
               .eStyle = eButtonStyle
               .hWndCapture = 0
               .hWndParentOrig = 0
               .bControl = False
               .bStretch = False
               .hSubMenu = 0
            End With
            plAddButton = iB
         End If
      End If
   Else
      m_lR = SendMessage(hWndToolbar, TB_ADDBUTTONS, 1, tB)
      If (m_lR <> 0) Then
         ' Add this button to the list:
         If hWndToolbar = m_hWndToolBar Then
            ReDim Preserve m_tBInfo(0 To iCount) As ButtonInfoStore
            With m_tBInfo(iCount)
               .wId = tB.idCommand
               .iImage = iImage
               .sTipText = sTip
               .iLarge = xLarge
               .sKey = sKey
               .bShowText = m_bWithText
               .sCaption = sButtonText
               .eStyle = eButtonStyle
               .hWndCapture = 0
               .hWndParentOrig = 0
               .bControl = False
               .bStretch = False
               .hSubMenu = 0
            End With
            plAddButton = iCount
         End If
      End If
   End If
   
   ' Size window:
   pResizeToolbar hWndToolbar
    
End Function
Public Sub AddButton( _
      Optional ByVal sTip As String = "", _
      Optional ByVal iImage As Integer = -1, _
      Optional ByVal vButtonBefore As Variant, _
      Optional ByVal xLarge As Integer = 0, _
      Optional ByVal sButtonText As String, _
      Optional ByVal eButtonStyle As ECTBToolButtonSyle, _
      Optional ByVal sKey As String = "" _
   )
Attribute AddButton.VB_Description = "Adds a button to the toolbar, optionally setting the buttons text, tool tip, image and style at the same time."
   plAddButton m_hWndToolBar, NewButtonID, sTip, iImage, vButtonBefore, xLarge, sButtonText, eButtonStyle, sKey
End Sub
Private Function plAddStringIfRequired(ByVal hWndToolbar As Long, ByVal sString As String) As Long
Dim id As Long
Dim i As Long
Dim b() As Byte
Dim sAccel As String

   ' Signal default:
   id = -1
   
   If hWndToolbar = m_hWndToolBar Then
      ' Check if we already have the string - if we do, then use that
      For i = 1 To m_lStringIDCount
         If (m_sString(i) = sString) Then
            id = m_lStringID(i)
            Exit For
         End If
      Next i
   End If
   
   ' If string not found, then add one:
   If (id = -1) Then
      b = StrConv(sString, vbFromUnicode)
      i = UBound(b) + 2
      ReDim Preserve b(0 To i) As Byte
      b(i - 1) = 0
      b(i) = 0
      
      id = SendMessage(hWndToolbar, TB_ADDSTRING, 0, b(0))
      
      If m_hWndToolBar = hWndToolbar Then
         m_lStringIDCount = m_lStringIDCount + 1
         ReDim Preserve m_sString(1 To m_lStringIDCount) As String
         ReDim Preserve m_lStringID(1 To m_lStringIDCount) As Long
         m_sString(m_lStringIDCount) = sString
         m_lStringID(m_lStringIDCount) = id
      End If
      
   End If
   
   ' Return the Id:
   plAddStringIfRequired = id
   
End Function
Private Function psGetAccelerator(ByVal sString As String) As String
Dim iPos As Long
   iPos = InStr(sString, "&")
   If iPos <> 0 And iPos <> InStr(sString, "&&") Then
      If iPos < Len(sString) Then
         psGetAccelerator = Chr$(CharToKeyCode(UCase$(Mid$(sString, iPos + 1, 1))))
      End If
   End If
End Function
Private Function plRemoveString(ByVal sCaption As String)
   ' unfortunately you cannot remove a string
   ' from the toolbar itself (because, as MSJ puts it,
   ' ".. the toolbar is braindead ..")
   
End Function
Public Sub ResizeToolbar()
Attribute ResizeToolbar.VB_Description = "Resizes the toolbar."
   pResizeToolbar m_hWndToolBar
End Sub
Private Sub pResizeToolbar(ByVal hWndToolbar As Long)
Dim tR As RECT, tPR As RECT, tCR As RECT
Dim tP As POINTAPI
Dim lCount As Long
Dim i As Long
Dim Button As TBBUTTON
Dim lW As Long, lH As Long
Dim bInRebar As Boolean
Dim lhWnd As Long
   
   ' Get number of buttons:
   lCount = SendMessageLong(hWndToolbar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the total length:
      lW = plToolbarWidth(hWndToolbar)
      lH = plToolbarHeight(hWndToolbar)
      
      ' Get rectangle for toolbar.  Unfortunately the rebar doesn't
      ' seem to like ClientToScreen and gives the wrong answer!  So
      ' do it manually:
      GetWindowRect hWndToolbar, tR
      GetWindowRect GetParent(hWndToolbar), tPR
      GetClientRect GetParent(hWndToolbar), tCR
      
      'Debug.Print tR.Top, tPR.Top, tCR.Top
      tP.x = tR.left - tPR.left - 2
      tP.y = tR.top - tPR.top - 2
      
      ' Make window correct size:
      If (m_bWrappable) Then
         SetWindowPos hWndToolbar, 0, tP.x, tP.y, lW, lH, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
      Else
         SetWindowPos hWndToolbar, 0, tP.x, tP.y, lW, lH, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
      End If
      
      If hWndToolbar = m_hWndToolBar Then
         For i = 0 To lCount - 1
            If Not (m_tBInfo(i).hWndCapture = 0) Then
               moveChildWindow i
            End If
         Next i
         correctRebarIfExists
      End If
                 
    End If
End Sub
Private Sub correctRebarIfExists()
Dim lhWnd As Long
Dim sBuf As String
Dim iPos As Long
   If IsWindowVisible(m_hWndToolBar) Then
      lhWnd = GetParent(m_hWndToolBar)
      sBuf = String$(256, 0)
      GetClassName lhWnd, sBuf, 255
      iPos = InStr(sBuf, vbNullChar)
      If iPos > 1 Then sBuf = left$(sBuf, iPos - 1)
      'Debug.Print lhWnd, sBuf
      
      If sBuf = REBARCLASSNAME Then
         SendMessageLong lhWnd, WM_SIZE, 0, 0
         Exit Sub
      Else
         SendMessageLong lhWnd, WM_SIZE, 0, 0
      End If
      
      lhWnd = GetParent(lhWnd)
      sBuf = String$(256, 0)
      GetClassName lhWnd, sBuf, 255
      iPos = InStr(sBuf, vbNullChar)
      If iPos > 1 Then sBuf = left$(sBuf, iPos - 1)
      'Debug.Print lhWnd, sBuf
            
      'If sBuf = REBARCLASSNAME Then
         SendMessageLong lhWnd, WM_SIZE, 0, 0
         Exit Sub
      'End If
      
   End If
End Sub
Public Property Get ToolbarWidth() As Long
Attribute ToolbarWidth.VB_Description = "Gets the width of the toolbar."
   ToolbarWidth = plToolbarWidth(m_hWndToolBar)
End Property
Private Property Get plToolbarWidth(ByVal hWndToolbar As Long) As Long
Dim lSize As Long
Dim lCount As Long
Dim lWidth As Long
Dim i As Long
Dim rc As RECT

   ' Get number of buttons:
   lCount = SendMessageLong(hWndToolbar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the total length:
      For i = 0 To lCount - 1
         If (plButtonVisible(hWndToolbar, i)) Then
            If (m_tBInfo(i).bControl) Then
               ButtonWidth(i) = m_tBInfo(i).xWidth
               moveChildWindow i
            Else
               SendMessage hWndToolbar, TB_GETITEMRECT, i, rc
               lSize = lSize + rc.right - rc.left
            End If
         End If
      Next i
      plToolbarWidth = lSize + 2
   End If
   
End Property
Public Property Get ToolbarHeight() As Long
Attribute ToolbarHeight.VB_Description = "Gets the height of the toolbar."
   ToolbarHeight = plToolbarHeight(m_hWndToolBar)
End Property
Private Property Get plToolbarHeight(ByVal hWndToolbar As Long) As Long
Dim lSize As Long
Dim lCount As Long
Dim i As Long
Dim rc As RECT
   ' Get number of buttons:
   lCount = SendMessageLong(hWndToolbar, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get the height:
      i = 0
      Do While plButtonVisible(hWndToolbar, i) = False
         i = i + 1
         If i >= lCount Then
            Exit Do
         End If
      Loop
      SendMessage hWndToolbar, TB_GETITEMRECT, i, rc
      plToolbarHeight = rc.bottom
   End If
End Property

Public Sub ButtonSize(xWidth As Integer, xHeight As Integer)
Attribute ButtonSize.VB_Description = "Gets the rectangle of a button."
   m_iButtonWidth = xWidth
   m_iButtonHeight = xHeight
   SendMessageLong m_hWndToolBar, TB_AUTOSIZE, 0, 0
   ResizeToolbar
End Sub
Public Sub GetDropDownPosition( _
        ByVal id As Integer, _
        ByRef x As Long, _
        ByRef y As Long _
    )
Attribute GetDropDownPosition.VB_Description = "Returns the position to show a drop-down menu for a button in response to the DropDownPress event."
Dim rc As RECT
Dim tP As POINTAPI
Dim i As Long
Dim lMappedID As Long
    
   If Not m_hWndChevronToolbar = 0 Then
      ' need to modify ID so it is relative to the chevron toolbar,
      ' rather than the
      For i = 1 To m_iChevronIDMapCount
         If id = m_iChevronIDMap(i) Then
            lMappedID = i - 1
            Exit For
         End If
      Next i
      SendMessage m_hWndChevronToolbar, TB_GETITEMRECT, lMappedID, rc
      tP.x = rc.left
      tP.y = rc.bottom
      MapWindowPoints m_hWndChevronToolbar, m_hWndParentForm, tP, 1
   Else
      SendMessage m_hWndToolBar, TB_GETITEMRECT, id, rc
      tP.x = rc.left
      tP.y = rc.bottom
      MapWindowPoints m_hWndToolBar, m_hWndParentForm, tP, 1
   End If
   x = tP.x * Screen.TwipsPerPixelX
   y = tP.y * Screen.TwipsPerPixelY
    
End Sub

Private Sub pInitialise()
Dim tIccex As CommonControlsEx

   If Not (UserControl.Ambient.UserMode) Then
     ' We are in design mode:
     lblInfo.Caption = "Toolbar Control: " & UserControl.Extender.Name
   Else
      UserControl.Extender.Visible = False
      lblInfo.Visible = False
      UserControl.Extender.left = -UserControl.width * 2
      ' We are in run
      With tIccex
          .dwSize = LenB(tIccex)
          .dwICC = ICC_BAR_CLASSES
      End With
      'We need to make this call to make sure the common controls are loaded
      InitCommonControlsEx tIccex
      m_hWndToolBar = 0
   End If
   
End Sub
Private Sub pSubClass(ByVal bState As Boolean, Optional ByVal lhWnd As Long = 0)
Static s_lhWndSave As Long

    If (m_bInSubClass <> bState) Then
        If (bState) Then
            'Debug.Print "Subclassing:Start"
            Debug.Assert (lhWnd <> 0)
            If (s_lhWndSave <> 0) Then
                pSubClass False
            End If
            s_lhWndSave = lhWnd
            pAttMsg lhWnd, WM_COMMAND
            pAttMsg lhWnd, WM_MOUSEMOVE
            pAttMsg lhWnd, WM_LBUTTONDOWN
            pAttMsg lhWnd, WM_LBUTTONUP
            pAttMsg lhWnd, WM_RBUTTONDOWN
            pAttMsg lhWnd, WM_RBUTTONUP
            pAttMsg lhWnd, WM_MBUTTONDOWN
            pAttMsg lhWnd, WM_MBUTTONUP
            pAttMsg lhWnd, WM_NOTIFY
            pAttMsg m_hWndToolBar, WM_SIZE
            pAttMsg m_hWndToolBar, WM_WINDOWPOSCHANGING
            pAttMsg m_hWndToolBar, WM_WINDOWPOSCHANGED
            pAttMsg m_hWndToolBar, WM_SHOWWINDOW
            pAttMsg m_hWndToolBar, WM_DESTROY
            pAttMsg lhWnd, WM_PARENTNOTIFY
            pAttMsg lhWnd, WM_DESTROY
            s_lhWndSave = lhWnd
            m_bInSubClass = True
        Else
            'Debug.Print "Subclassing:End"
            pDelMsg s_lhWndSave, WM_COMMAND
            pDelMsg s_lhWndSave, WM_MOUSEMOVE
            pDelMsg s_lhWndSave, WM_LBUTTONDOWN
            pDelMsg s_lhWndSave, WM_LBUTTONUP
            pDelMsg s_lhWndSave, WM_RBUTTONDOWN
            pDelMsg s_lhWndSave, WM_RBUTTONUP
            pDelMsg s_lhWndSave, WM_MBUTTONDOWN
            pDelMsg s_lhWndSave, WM_MBUTTONUP
            pDelMsg s_lhWndSave, WM_NOTIFY
            pDelMsg m_hWndToolBar, WM_SIZE
            pDelMsg m_hWndToolBar, WM_WINDOWPOSCHANGING
            pDelMsg m_hWndToolBar, WM_WINDOWPOSCHANGED
            pDelMsg m_hWndToolBar, WM_SHOWWINDOW
            pDelMsg m_hWndToolBar, WM_DESTROY
            pDelMsg s_lhWndSave, WM_PARENTNOTIFY
            pDelMsg s_lhWndSave, WM_DESTROY
            s_lhWndSave = 0
            m_bInSubClass = False
        End If
    End If
End Sub
Private Sub pTerminate()
    ' Clear toolbar window:
   DestroyToolBar
   ' Background picture -> nothing if any:
   Set m_pic = Nothing
End Sub
Private Sub pAttMsg(ByVal lhWnd As Long, ByVal lMsg As Long)
    AttachMessage Me, lhWnd, lMsg
End Sub
Private Sub pDelMsg(ByVal lhWnd As Long, ByVal lMsg As Long)
    DetachMessage Me, lhWnd, lMsg
End Sub

Public Function RaiseButtonClick(ByVal iIDButton As Long)
Attribute RaiseButtonClick.VB_Description = "Causes a button click to occur."
   ' Required as part of the WM_COMMAND handler:
   SendMessageLong m_hWndParentForm, WM_CANCELMODE, 0, 0
   RaiseEvent ButtonClick(iIDButton)
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If (CurrentMessage = WM_NOTIFY) Then
      ISubclass_MsgResponse = emrConsume
   Else
      ISubclass_MsgResponse = emrPreprocess
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim msgStruct As Msg
Dim hdr As NMHDR
Dim ttt As ToolTipText
Dim pt32 As POINTAPI
Dim ptx As Long
Dim pty As Long
Dim hWndOver As Long
Dim b() As Byte
Dim iB As Long
Dim ib2 As Long
Dim iBRaise As Long
Dim lButton As Long
Dim lPtr As Long
Dim iOld As Long, iNew As Long
Dim eReason As ECTBHotItemChangeReasonConstants
Dim bS As Boolean
Dim bCanInsert As Boolean
Dim bCanDelete As Boolean
Dim tR As RECT, tBR As RECT, tWR As RECT
Dim lAW As Long
Dim iStretchCount As Long
Dim bStretch As Boolean
Dim bControl As Boolean
Dim bSubMenu As Boolean
Dim wId As Long
Dim iNewCount As Long
Dim fwEvent As Long
Dim lIDChild As Long
Dim hWndCHild As Long
Dim lhWnd As Long
Dim tWP As WINDOWPOS
Dim lFlag As Long
Dim lStyle As Long
Dim i As Long
  
On Error Resume Next

   Select Case iMsg
   Case WM_PARENTNOTIFY
      
      fwEvent = (wParam And &HFFFF&)
      lIDChild = (wParam And &H7FFF0000)
      hWndCHild = lParam
      If fwEvent = WM_DESTROY Then
         debugmsg m_sCtlName & ",Parent Notify:Destroy"
         For lButton = ButtonCount - 1 To 0 Step -1
            If m_tBInfo(lButton).hWndCapture = hWndCHild Then
               RemoveButton lButton
            End If
         Next lButton
      End If
      
   Case WM_DESTROY, WM_CLOSE, WM_SYSCOMMAND
      If iMsg = WM_SYSCOMMAND Then
         'Debug.Print wParam, SC_CLOSE
         If wParam <> SC_CLOSE Then
            Exit Function
         End If
      End If
      debugmsg m_sCtlName & ",cToolbar:WM_DESTROY"
      'pSubClass False
      pTerminate
   
   Case WM_SHOWWINDOW
      m_bMenuLoop = False
      If wParam = 0 Then
         'Debug.Print "Hiding"
         m_bVisible = False
         lFlag = SW_HIDE
      Else
         'Debug.Print "Showing"
         m_bVisible = True
         lFlag = SW_SHOW
      End If
      ' hiding window
      For lButton = 0 To ButtonCount - 1
         If m_tBInfo(lButton).hWndCapture <> 0 Then
            ShowWindow m_tBInfo(lButton).hWndCapture, lFlag
            'lStyle = GetWindowLong(m_tBInfo(lButton).hWndCapture, GWL_STYLE)
            'If (wParam = 0) Then
            '   lStyle = lStyle And Not WS_VISIBLE
            'Else
            '   lStyle = lStyle Or WS_VISIBLE
            'End If
            'SetWindowLong m_tBInfo(lButton).hWndCapture, GWL_STYLE, lStyle
         End If
      Next lButton
      
   Case WM_COMMAND
      If (lParam = m_hWndToolBar) Or (lParam = m_hWndChevronToolbar) Then

         ' This is the index of the button in the toolbar, which can be different if the
         ' toolbar is a chevron:
         iB = SendMessageLong(lParam, TB_COMMANDTOINDEX, wParam, 0)
         ' And this is the actual index of the button in the proper toolbar:
         iBRaise = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, wParam, 0)
         
         If iB > -1 Then
            bSubMenu = Not (m_tBInfo(iB).hSubMenu = 0)
            If bSubMenu Then
               If (m_tBInfo(iB).eStyle And CTBDropDown) = CTBDropDown Then
                  ' sub menu is only accessible via drop down
                  bSubMenu = False
               End If
            End If
         
            If bSubMenu Then
               bS = ButtonPressed(iB)
               ButtonPressed(iB) = True
               ' First tell the client we're about to show the menu
               RaiseButtonClick iBRaise
               ' Now show the menu:
               pMenuClick lParam, iB
               ButtonPressed(iB) = False
               ISubclass_WindowProc = 0
               SendMessageLong m_hWndParentForm, WM_EXITMENULOOP, 0, 0
               SendMessageLong m_hWndToolBar, TB_SETHOTITEM, -1, 0
            Else
               'Debug.Print "Items", m_tBInfo(iBRaise).sKey, m_tBInfo(iBRaise).eStyle And &H2
               pbSetState2 lParam, wParam, TBSTATE_PRESSED, True
               If Not (m_hWndToolBar = lParam) Then
                  pbSetState2 m_hWndToolBar, wParam, TBSTATE_PRESSED, True
                  If ((m_tBInfo(iBRaise).eStyle And CTBCheck) = CTBCheck) Then
                     bS = (pbGetState2(lParam, wParam, TBSTATE_CHECKED))
                     'Debug.Print "Chevron Window Checked: "; bS
                     ButtonChecked(iBRaise) = bS
                     'Debug.Print "Toolbar Checked: "; ButtonChecked(iBRaise)
                  End If
               End If
               RaiseButtonClick iBRaise
               pbSetState2 lParam, wParam, TBSTATE_PRESSED, False
               If Not (lParam = m_hWndToolBar) Then
                  pbSetState2 m_hWndToolBar, wParam, TBSTATE_PRESSED, False
               End If
               If lParam = m_hWndChevronToolbar Then
                  SendMessageLong m_hWndParentForm, WM_EXITMENULOOP, 0, 0
               End If
               ISubclass_WindowProc = 0
            End If
            
            If (lParam = m_hWndToolBar) Then
               If m_hMenu <> 0 Then
                  If m_bCreateFromMenu2 Then ' don't break existing apps
                     If ((m_tBInfo(iBRaise).hSubMenu = 0) Or (m_tBInfo(iBRaise).eStyle And CTBDropDown) = CTBDropDown) Then
                        Dim cMenu As Object
                        Dim cT As Object
                        Dim iID As Long
                        CopyMemory cT, m_lPtrMenu, 4
                        Set cMenu = cT
                        CopyMemory cT, 0&, 4
                        iID = cMenu.IDForItem(cMenu.IndexForKey(ButtonKey(iB)))
                        cMenu.EmulateMenuClick iID
                     End If
                  End If
               End If
            End If
            
         End If
      End If
   
   Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP
      With msgStruct
         .lParam = lParam
         .wParam = wParam
         .message = iMsg
         .hwnd = hwnd
      End With
      
      'Pass the structure
      SendMessage hwndToolTip, TTM_RELAYEVENT, 0, msgStruct
   
   Case WM_SIZE, WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED
      ' time to adjust any captured controls to match:
      'GetWindowRect m_hWndToolBar, tR
      m_bMenuLoop = False
      If iMsg = WM_SIZE Then
         lAW = lParam And &HFFFF& 'tR.right - tR.left + 1 'tWP.cx
      Else
         CopyMemory tWP, ByVal lParam, Len(tWP)
         lAW = tWP.cx
      End If
      For lButton = 0 To ButtonCount - 1
         If ButtonVisible(iB) Then
            If m_tBInfo(lButton).bControl Then
               bControl = True
               bStretch = bStretch Or m_tBInfo(lButton).bStretch
               If m_tBInfo(lButton).bStretch Then
                  iStretchCount = iStretchCount + 1
               Else
                  SendMessage m_hWndToolBar, TB_GETITEMRECT, lButton, tR
                  lAW = lAW - (tR.right - tR.left)
               End If
            Else
               SendMessage m_hWndToolBar, TB_GETITEMRECT, lButton, tR
               lAW = lAW - (tR.right - tR.left)
            End If
         End If
      Next lButton
      
      If bControl Then
         If bStretch Then
            lAW = (lAW \ iStretchCount) - 1
            'Debug.Print "WidthChange:", lAW
            For lButton = 0 To ButtonCount - 1
               If ButtonVisible(iB) Then
                  If m_tBInfo(lButton).bControl Then
                     'Debug.Print lAW, m_tBInfo(lButton).xWidth
                     If (m_tBInfo(lButton).bStretch) Then
                        If lAW >= m_tBInfo(lButton).xWidth Then
                           If ButtonWidth(lButton) <> lAW Then
                              ButtonWidth(lButton) = lAW
                           End If
                        Else
                           If ButtonWidth(lButton) <> m_tBInfo(lButton).xWidth Then
                              ButtonWidth(lButton) = m_tBInfo(lButton).xWidth
                           End If
                        End If
                     Else
                        If ButtonWidth(lButton) <> m_tBInfo(lButton).xWidth Then
                           ButtonWidth(lButton) = m_tBInfo(lButton).xWidth
                        Else
                           SendMessage m_hWndToolBar, TB_GETITEMRECT, lButton, tR
                           If Not (m_tBInfo(lButton).hWndCapture = 0) Then
                              GetWindowRect m_tBInfo(lButton).hWndCapture, tWR
                              LSet tBR = tR
                              MapWindowPoints m_hWndToolBar, HWND_DESKTOP, tBR, 2
                              If tWR.left <> tBR.left Or tWR.top <> tBR.top Or tWR.right <> tBR.right Or tWR.bottom <> tBR.bottom Then
                                 moveChildWindow lButton
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            Next lButton
            
         Else
            For lButton = 0 To ButtonCount - 1
               If ButtonVisible(iB) Then
                  If m_tBInfo(lButton).bControl Then
                     SendMessage m_hWndToolBar, TB_GETITEMRECT, lButton, tR
                     If Not (m_tBInfo(lButton).hWndCapture = 0) Then
                        GetWindowRect m_tBInfo(lButton).hWndCapture, tWR
                        LSet tBR = tR
                        MapWindowPoints m_hWndToolBar, HWND_DESKTOP, tBR, 2
                        'If tWR.left <> tBR.left Or tWR.top <> tBR.top Or tWR.right <> tBR.right Or tWR.bottom <> tBR.bottom Then
                           moveChildWindow lButton
                        'End If
                     End If
                  End If
               End If
            Next lButton
         End If
      End If
   
   Case WM_NOTIFY
      CopyMemory hdr, ByVal lParam, Len(hdr)
         
      Select Case hdr.code
      Case VBALCHEVRONMENUCONST
         If (hdr.hwndFrom = m_hWndToolBar) Then
            Dim iIDType As Long, iBtn As Long
            
            iID = hdr.idfrom
            iIDType = iID And &H7FFF0000
            Select Case iIDType
            Case 0
               ' button visible
               iBtn = iID And &HFFFF&
               ButtonVisible(iBtn) = Not (ButtonVisible(iBtn))
               ISubclass_WindowProc = findFirstNonVisibleButton()
               
            Case 1
               ' customise
               RaiseEvent CustomiseBegin
               
            Case 2
               '  reset
               RaiseEvent CustomiseResetPressed
            
            Case 3
               ' ?
               
            End Select
            '
         End If
                  
      Case TTN_NEEDTEXT
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            Dim idNum As Integer
            idNum = hdr.idfrom
            On Error Resume Next
            
            iB = pbGetIndexForID(idNum)
            If (iB > -1) Then
               msToolTipBuffer = StrConv(ButtonToolTip(iB), vbFromUnicode)
               If Err.Number = 0 Then
                  If (Len(msToolTipBuffer) > 0) Then
                     msToolTipBuffer = msToolTipBuffer & vbNullChar
                     ' Debug.Print "Show tool tip", ButtonToolTip(iB)
                     CopyMemory ttt, ByVal lParam, Len(ttt)
                     ttt.lpszText = StrPtr(msToolTipBuffer)
                     CopyMemory ByVal lParam, ttt, Len(ttt)
                  End If
               Else
                  Err.Clear
               End If
            End If
         
      Case TBN_DROPDOWN
         
         If (hdr.hwndFrom = m_hWndToolBar) Or (hdr.hwndFrom = m_hWndChevronToolbar) Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            Dim nmTb As NMTOOLBAR_SHORT
            CopyMemory nmTb, ByVal lParam, Len(nmTb)
            iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTb.iItem, 0)
            bSubMenu = Not (m_tBInfo(iB).hSubMenu = 0)
            
            Debug.Print "Setting Dropped for button", iB
            For i = 0 To ButtonCount - 1
               m_tBInfo(i).bDropped = (i = iB)
            Next i
            If bSubMenu Then
               bS = ButtonPressed(iB)
               ButtonPressed(iB) = True
               ' Now show the menu:
               pMenuClick hdr.hwndFrom, iB
               ButtonPressed(iB) = False
               ISubclass_WindowProc = 0
               SendMessageLong m_hWndParentForm, WM_CANCELMODE, 0, 0
               SendMessageLong m_hWndToolBar, TB_SETHOTITEM, -1, 0
            Else
               RaiseEvent DropDownPress(iB)
               If hdr.hwndFrom = m_hWndChevronToolbar Then
                  SendMessageLong m_hWndParentForm, WM_CANCELMODE, 0, 0
               End If
            End If
            m_tBInfo(iB).bDropped = False
            
         End If
         
      Case TBN_HOTITEMCHANGE
         If (hdr.hwndFrom = m_hWndToolBar) Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            If m_lMajorVer > 4 Or (m_lMajorVer = 4 And m_lMinorVer >= 70) Then
               Dim nmTBHI As NMTBHOTITEM
               CopyMemory nmTBHI, ByVal lParam, Len(nmTBHI)
               eReason = nmTBHI.dwFlags
               iOld = -1: iNew = -1
               If (eReason And HICF_ENTERING) <> HICF_ENTERING Then
                  iOld = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTBHI.idOld, 0)
               End If
               If (eReason And HICF_LEAVING) <> HICF_LEAVING Then
                  iNew = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTBHI.idNew, 0)
               End If
               RaiseEvent HotItemChange(iNew, iOld, eReason)
            End If
            ISubclass_WindowProc = 0
         End If
         
      Case TBN_BEGINADJUST
         ' begin adjust:
         If (hdr.hwndFrom = m_hWndToolBar) Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            RaiseEvent CustomiseBegin
         End If
         
      Case TBN_QUERYINSERT
         ' toolbar is asking whether a button can be inserted to the left of the
         ' button specified in the NMTOOLBAR structure:
         If (hdr.hwndFrom = m_hWndToolBar) Then
            g_lCustomiseResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            CopyMemory nmTb, ByVal lParam, Len(nmTb)
            iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTb.iItem, 0)
            bCanInsert = True
            RaiseEvent CustomiseCanInsertBefore(iB, bCanInsert)
            If bCanInsert Then
               g_lCustomiseResponse = 1
               ISubclass_WindowProc = 1
            Else
               g_lCustomiseResponse = 0
               ISubclass_WindowProc = 0
            End If
         End If
      
      Case TBN_QUERYDELETE
         ' toolbar is asking if button can be deleted:
         If (hdr.hwndFrom = m_hWndToolBar) Then
            g_lCustomiseResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            
            CopyMemory nmTb, ByVal lParam, Len(nmTb)
            iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, nmTb.iItem, 0)
            bCanDelete = True
            RaiseEvent CustomiseCanDelete(iB, bCanDelete)
            If bCanDelete Then
               g_lCustomiseResponse = 1
            Else
               g_lCustomiseResponse = 0
            End If
            ISubclass_WindowProc = g_lCustomiseResponse
         End If
                  
      Case TBN_GETBUTTONINFO
         If (hdr.hwndFrom = m_hWndToolBar) Then
            g_lCustomiseResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            Dim nmTBF As NMTOOLBAR
            CopyMemory nmTBF, ByVal lParam, Len(nmTBF)
            'Debug.Print "TBN_GETBUTTONINFO", nmTBF.iItem, nmTBF.cchText, nmTBF.lpszString
            ReDim b(0 To nmTBF.cchText) As Byte
            CopyMemory b(0), ByVal nmTBF.lpszString, nmTBF.cchText
            'Debug.Print StrConv(b, vbUnicode)
            
            g_lCustomiseResponse = 1
         End If
         
      Case TBN_CUSTHELP
         If (hdr.hwndFrom = m_hWndToolBar) Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            RaiseEvent CustomiseHelpPressed
         End If
         
      Case TBN_RESET
         If (hdr.hwndFrom = m_hWndToolBar) Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            RaiseEvent CustomiseResetPressed
         End If
         
      Case TBN_DELETINGBUTTON
         If (hdr.hwndFrom = m_hWndToolBar) Then
            CopyMemory nmTb, ByVal lParam, Len(nmTb)
            wId = nmTb.iItem
            iB = SendMessageLong(m_hWndToolBar, TB_COMMANDTOINDEX, wId, 0)
            If iB > -1 Then
               If Not (m_tBInfo(iB).hWndCapture = 0) Then
                  'SetParent m_tBInfo(iB).hWndCapture, m_tBInfo(iB).hWndParentOrig
               End If
               iNewCount = ButtonCount
               If iNewCount = 0 Then
                  Erase m_tBInfo
               Else
                  For lButton = iB To iNewCount - 1
                     LSet m_tBInfo(lButton) = m_tBInfo(lButton + 1)
                  Next lButton
                  ReDim Preserve m_tBInfo(0 To iNewCount - 1) As ButtonInfoStore
               End If
            End If
         End If
         
      Case NM_CUSTOMDRAW
         If (hdr.hwndFrom = m_hWndToolBar) Or (hdr.hwndFrom = m_hWndChevronToolbar) Then
            ISubclass_WindowProc = CustomDraw(hdr.hwndFrom, lParam)
         End If
      End Select
      
   End Select
    
End Function

Private Function CustomDraw(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim tTBCD As NMTBCUSTOMDRAW
   CopyMemory tTBCD, ByVal lParam, Len(tTBCD)
   'Debug.Print UserControl.Extender.Name + ":CustomDraw", m_eVisualStyle, tTBCD.nmcd.dwDrawStage
   Select Case tTBCD.nmcd.dwDrawStage
   Case CDDS_PREPAINT
      'Debug.Print UserControl.Extender.Name + ":CustomDraw", m_eVisualStyle, tTBCD.nmcd.dwDrawStage, "Returning CDRF_NOTIFYITEMDRAW"
      CustomDraw = CDRF_NOTIFYITEMDRAW
   Case CDDS_ITEMPREPAINT
      CustomDraw = CDRF_DODEFAULT
      If (m_eVisualStyle = CTBDrawOfficeXPStyle) Then
         CustomDraw = OfficeXpCustomDraw(hwnd, tTBCD)
      End If
   End Select
End Function

Private Function OfficeXpCustomDraw(ByVal hwnd As Long, tTBCD As NMTBCUSTOMDRAW) As Long
Dim bHot As Boolean
Dim bSelected  As Boolean
Dim bDisabled As Boolean
Dim sText As String
Dim sDrawText As String
Dim bShowText As Boolean
Dim lIndex As Long
Dim wId As Long
Dim rc As RECT
Dim rcNoArrow As RECT
Dim rcText As RECT
Dim lHDC As Long
Dim hBr As Long
Dim hPen As Long
Dim hPenOld As Long
Dim tJunk As POINTAPI
Dim iPos As Long
Dim hFont As Long
Dim hFontOld As Long
Dim lX As Long
Dim lY As Long
Dim lSrcX As Long
Dim lFmt As Long
Dim lImgWidth As Long
Dim lImgHeight As Long
Dim bOver As Boolean
Dim bChecked As Boolean
Dim bListStyle As Boolean

   OfficeXpCustomDraw = CDRF_SKIPDEFAULT
   
   bListStyle = IIf(hwnd = m_hWndChevronToolbar, True, m_bListStyle)

   wId = tTBCD.nmcd.dwItemSpec
   lIndex = pbGetIndexForID(wId)
   lHDC = tTBCD.nmcd.hdc
   LSet rc = tTBCD.nmcd.rc

   bHot = ((tTBCD.nmcd.uItemState And CDIS_HOT) = CDIS_HOT)
   bSelected = ((tTBCD.nmcd.uItemState And CDIS_SELECTED) = CDIS_SELECTED)
   bDisabled = ((tTBCD.nmcd.uItemState And CDIS_DISABLED) = CDIS_DISABLED)
   bShowText = m_tBInfo(lIndex).bShowText
   bChecked = ButtonChecked(lIndex)
   sText = m_tBInfo(lIndex).sCaption
   If ((m_tBInfo(lIndex).eStyle And CTBDropDown) = CTBDropDown) Or _
      ((m_tBInfo(lIndex).eStyle And CTBDropDownArrow) = CTBDropDownArrow) Then
      If (m_bMenuShown) Then
         bHot = m_tBInfo(lIndex).bDropped
      End If
   End If
   
      
   If (bChecked) Then
      If (bHot) Then
         hBr = CreateSolidBrush(VSNetPressedColor)
      Else
         hBr = CreateSolidBrush(VSNetCheckedColor)
      End If
      FillRect lHDC, rc, hBr
      DeleteObject hBr
      
      hPen = CreatePen(PS_SOLID, 1, VSNetBorderColor)
      hPenOld = SelectObject(lHDC, hPen)
      MoveToEx lHDC, rc.left, rc.bottom - 1, tJunk
      LineTo lHDC, rc.left, rc.top
      LineTo lHDC, rc.right - 1, rc.top
      LineTo lHDC, rc.right - 1, rc.bottom - 1
      LineTo lHDC, rc.left, rc.bottom - 1
      SelectObject lHDC, hPenOld
      DeleteObject hPen
      
   ElseIf (bSelected) Then

      If (m_tBInfo(lIndex).eStyle And CTBDropDown) = CTBDropDown Then
         If (m_tBInfo(lIndex).bDropped) Then
            hBr = CreateSolidBrush(VSNetControlColor)
            FillRect lHDC, rc, hBr
            DeleteObject hBr
            DrawEdge lHDC, rc, BDR_RAISED, BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_FLAT
         Else
            hBr = CreateSolidBrush(VSNetSelectionColor)
            FillRect lHDC, rc, hBr
            DeleteObject hBr
   
            hBr = CreateSolidBrush(VSNetPressedColor)
            LSet rcNoArrow = rc
            rcNoArrow.right = rcNoArrow.right - DROPDOWN_ARROW_WIDTH + 1
            FillRect lHDC, rcNoArrow, hBr
            DeleteObject hBr
            
            hPen = CreatePen(PS_SOLID, 1, VSNetBorderColor)
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, rc.left, rc.bottom - 1, tJunk
            LineTo lHDC, rc.left, rc.top
            LineTo lHDC, rc.right - 1, rc.top
            LineTo lHDC, rc.right - 1, rc.bottom - 1
            LineTo lHDC, rc.left, rc.bottom - 1
            MoveToEx lHDC, rc.right - DROPDOWN_ARROW_WIDTH, rc.top, tJunk
            LineTo lHDC, rc.right - DROPDOWN_ARROW_WIDTH, rc.bottom
            SelectObject lHDC, hPenOld
            DeleteObject hPen
         End If
         
      ElseIf (m_tBInfo(lIndex).eStyle And CTBDropDownArrow) = CTBDropDownArrow Then
         hBr = CreateSolidBrush(VSNetControlColor)
         FillRect lHDC, rc, hBr
         DeleteObject hBr
         DrawEdge lHDC, rc, BDR_RAISED, BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_FLAT
         
      Else
         If (m_eCreateFromMenuStyle = CTBMenuStyle) Then
            hBr = CreateSolidBrush(VSNetControlColor)
            FillRect lHDC, rc, hBr
            DeleteObject hBr
            DrawEdge lHDC, rc, BDR_RAISED, BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_FLAT
         Else
            hBr = CreateSolidBrush(VSNetPressedColor)
            FillRect lHDC, rc, hBr
            DeleteObject hBr
            
            hPen = CreatePen(PS_SOLID, 1, VSNetBorderColor)
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, rc.left, rc.bottom - 1, tJunk
            LineTo lHDC, rc.left, rc.top
            LineTo lHDC, rc.right - 1, rc.top
            LineTo lHDC, rc.right - 1, rc.bottom - 1
            LineTo lHDC, rc.left, rc.bottom - 1
            SelectObject lHDC, hPenOld
            DeleteObject hPen
         End If
      End If
            
   ElseIf (((m_tBInfo(lIndex).eStyle And CTBDropDown) = CTBDropDown Or (m_tBInfo(lIndex).eStyle And CTBDropDownArrow) = CTBDropDownArrow) And (m_tBInfo(lIndex).bDropped)) Then
      hBr = CreateSolidBrush(VSNetControlColor)
      FillRect lHDC, rc, hBr
      DeleteObject hBr
      DrawEdge lHDC, rc, BDR_RAISED, BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_FLAT
            
   ElseIf (bHot) Then
      hBr = CreateSolidBrush(VSNetSelectionColor)
      FillRect lHDC, rc, hBr
      DeleteObject hBr

      hPen = CreatePen(PS_SOLID, 1, VSNetBorderColor)
      hPenOld = SelectObject(lHDC, hPen)
      MoveToEx lHDC, rc.left, rc.bottom - 1, tJunk
      LineTo lHDC, rc.left, rc.top
      LineTo lHDC, rc.right - 1, rc.top
      LineTo lHDC, rc.right - 1, rc.bottom - 1
      LineTo lHDC, rc.left, rc.bottom - 1
      
      If ((m_tBInfo(lIndex).eStyle And CTBDropDown) = CTBDropDown) Then
         MoveToEx lHDC, rc.right - DROPDOWN_ARROW_WIDTH, rc.top, tJunk
         LineTo lHDC, rc.right - DROPDOWN_ARROW_WIDTH, rc.bottom - 1
      End If
      
      SelectObject lHDC, hPenOld
      DeleteObject hPen

   End If
   
   LSet rcText = rc
   InflateRect rcText, -1, -1
   
   ' Draw drop-down arrow if required
   If (m_tBInfo(lIndex).eStyle And CTBDropDown) = CTBDropDown Or (m_tBInfo(lIndex).eStyle And CTBDropDownArrow) = CTBDropDownArrow Then
      If (bDisabled) Then
         hPen = CreatePen(PS_SOLID, 1, BlendColor(vbMenuText, vbWindowBackground, 120))
      Else
         hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbMenuText))
      End If
      hPenOld = SelectObject(lHDC, hPen)
      lX = rc.right - 9
      lY = rc.top + (rc.bottom - rc.top) / 2
      MoveToEx lHDC, lX, lY, tJunk
      LineTo lHDC, lX + 5, lY
      MoveToEx lHDC, lX + 1, lY + 1, tJunk
      LineTo lHDC, lX + 4, lY + 1
      MoveToEx lHDC, lX + 2, lY, tJunk
      LineTo lHDC, lX + 2, lY + 3
      SelectObject lHDC, hPenOld
      DeleteObject hPen
      rcText.right = rcText.right - DROPDOWN_ARROW_WIDTH + 2
   End If
   
   ' Now draw the icon:
   If (m_tBInfo(lIndex).iImage > -1) Then
      If Not (m_cMemDC Is Nothing) Then
         ' If using an image list:
         If Not (m_hIml = 0) Then
            lImgWidth = m_lIconWidth
            lImgHeight = m_lIconHeight
            If Not (m_ptrVb6ImageList = 0) Then
               m_cMemDC.Clear 0, 0, 0, 0
               ' Draw from the VB image list onto the mem dc
               On Error Resume Next
               Dim o As Object
               Set o = ObjectFromPtr(m_ptrVb6ImageList)
               o.ListImages(m_tBInfo(lIndex).iImage + 1).Draw m_cMemDC.hdc, 0, 0, 0
               m_cMemDC.MakeTransparent m_lTransColor
            Else
               m_cMemDC.Clear 1, 3, 5, 0
               ' Draw the Image List item onto the mem dc:
               ImageList_Draw m_hIml, m_tBInfo(lIndex).iImage, m_cMemDC.hdc, 0, 0, ILD_TRANSPARENT Or ILD_PRESERVEALPHA
               m_cMemDC.MakeTransparent RGB(1, 3, 5)
            End If
            lSrcX = 0
         Else
            lImgWidth = m_cMemDC.height
            lImgHeight = m_cMemDC.height
            lSrcX = m_tBInfo(lIndex).iImage * lImgWidth
         End If
      
         If (bListStyle) Then
            lX = rcText.left + 2
            lY = rcText.top + (rcText.bottom - rcText.top - lImgWidth) / 2
            rcText.left = lX + lImgWidth + 4
         Else
            lX = rcText.left + (rcText.right - rcText.left - lImgHeight) / 2 + 1
            lY = rcText.top + 2
            rcText.top = lY + lImgHeight
         End If
         
         If (bHot And Not (bChecked) And Not (bSelected)) Then
            bOver = Not (bDisabled)
         End If
         
         ' Find the bitmap for this item:
         Dim cA As New cAlphaDibSection
         If (bOver Or bDisabled) Then
            Set cA = m_cMemDC.DisabledVersion( _
               lSrcX, 0, _
               lImgWidth, lImgHeight)
            If (bOver) Then
               lX = lX + 1
               lY = lY + 1
            End If
            cA.AlphaPaintPicture lHDC, _
               lX, lY, lImgWidth, lImgHeight
            If (bOver) Then
               lX = lX - 2
               lY = lY - 2
            End If
         End If
         If Not (bDisabled) Then
            If (bOver) Or (bChecked) Or (bSelected) Then
               m_cMemDC.AlphaPaintPicture lHDC, _
                  lX, _
                  lY, _
                  lImgWidth, lImgHeight, _
                  lSrcX, 0
            Else
               m_cMemDC.AlphaPaintPicture lHDC, _
                  lX, _
                  lY, _
                  lImgWidth, lImgHeight, _
                  lSrcX, 0, _
                  160
            End If
         End If
      End If
   End If
   
   
   ' Draw Text
   If (bShowText) Then
      If (bListStyle) Then
         lFmt = DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
      Else
         lFmt = DT_CENTER Or DT_SINGLELINE
         rcText.bottom = rcText.bottom - 2
      End If
      sDrawText = sText
      If Not ((m_bMenuShown) Or (m_bMenuLoop)) Or m_bAltPressed Then
         iPos = InStr(sDrawText, "&")
         If (iPos = 1) Then
            sDrawText = Mid(sDrawText, iPos + 1)
         ElseIf (iPos > 1) Then
            sDrawText = left(sDrawText, iPos - 1) & Mid(sDrawText, iPos + 1)
         End If
      End If
      SetBkMode lHDC, TRANSPARENT
      If (bDisabled) Then
         SetTextColor lHDC, BlendColor(vbMenuText, vbWindowBackground, 120)
      Else
         SetTextColor lHDC, TranslateColor(vbMenuText)
      End If
      hFont = m_cNCM.FontHandle(MenuFOnt)
      hFontOld = SelectObject(lHDC, hFont)
      DrawText lHDC, sDrawText, -1, rcText, lFmt
      SelectObject lHDC, hFontOld
   End If

End Function

Private Sub moveChildWindow(ByVal lButton As Long)
Dim lhWnd As Long
Dim tR As RECT
Dim tWR As RECT
Dim iB As Long

   iB = findFirstNonVisibleButton()
   If iB < 0 Then iB = ButtonCount()

   If lButton >= iB Then
      ShowWindow m_tBInfo(lButton).hWndCapture, SW_HIDE
   End If

   SendMessage m_hWndToolBar, TB_GETITEMRECT, lButton, tR
   lhWnd = GetParent(m_tBInfo(lButton).hWndCapture)
   MapWindowPoints m_hWndToolBar, lhWnd, tR, 2
   GetWindowRect m_tBInfo(lButton).hWndCapture, tWR
   If tWR.left <> tR.left Or tWR.right <> tR.right Or tWR.top <> tR.top Or tWR.bottom <> tR.bottom Then
      SetWindowPos m_tBInfo(lButton).hWndCapture, 0, _
         tR.left, tR.top + ((tR.bottom - tR.top) - (tWR.bottom - tWR.top)) / 2, _
         tR.right - tR.left, tR.bottom - tR.top, _
         SWP_FRAMECHANGED Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
   End If
   
   If lButton < iB Then
      If IsWindowVisible(m_tBInfo(lButton).hWndCapture) = 0 Then
         ShowWindow m_tBInfo(lButton).hWndCapture, SW_SHOW
      End If
   End If
      
End Sub
Private Function findFirstNonVisibleButton() As Long
Dim lhWnd As Long
Dim sBuf As String
Dim tWR As RECT
Dim iNotVisibleIndex As Long
Dim i As Long
Dim tR As RECT
   
   iNotVisibleIndex = -1
   lhWnd = GetParent(m_hWndToolBar)
   sBuf = String$(255, 0)
   GetClassName lhWnd, sBuf, 255
   
   'Debug.Print sBuf
   If LCase$(left$(sBuf, 7)) = "thunder" Then ' VB Control or Form
      GetClientRect lhWnd, tWR
'   ElseIf (left$(sBuf, Len(REBARCLASSNAME))) = REBARCLASSNAME Then
'      LSet tWR = m_tRebarBand
'      OffsetRect tWR, -tWR.left, -tWR.top
   Else
      GetClientRect m_hWndToolBar, tWR
   End If
   If Not (m_bVisible) Then
      findFirstNonVisibleButton = 0
      Exit Function
   End If
     
   For i = 0 To ButtonCount - 1
      SendMessage m_hWndToolBar, TB_GETITEMRECT, i, tR
      If tR.right > tWR.right Then
         If Not (m_tBInfo(i).eStyle = CTBSeparator) Then
            iNotVisibleIndex = i
            Exit For
         End If
      ElseIf tR.bottom > tWR.bottom Then
         If Not (m_tBInfo(i).eStyle = CTBSeparator) Then
            iNotVisibleIndex = i
            Exit For
         End If
      End If
   Next i
   findFirstNonVisibleButton = iNotVisibleIndex

End Function

Private Sub UserControl_Initialize()
   debugmsg "cToolbar:Initialize"
   If Not (ComCtlVersion(m_lMajorVer, m_lMinorVer, m_lBuild)) Then
      m_lMajorVer = 4
      m_lMinorVer = 0
      m_lBuild = 0
   End If
   m_eImageSourceType = -1
   m_sChevronAdditionalButton(CTBChevronAdditionalAddorRemove) = "&Add or Remove Buttons..."
   m_sChevronAdditionalButton(CTBChevronAdditionalReset) = "&Reset Toolbar..."
   m_sChevronAdditionalButton(CTBChevronAdditionalCustomise) = "&Customise..."
   m_bVisible = True
   m_eCreateFromMenuStyle = -1
End Sub

Private Sub UserControl_InitProperties()
    ' Initialise the control
    pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Read properties:
    
    On Error Resume Next
    m_sCtlName = UserControl.Extender.Name
    Err.Clear
    On Error GoTo 0
    
    ' Initialise the control
    DrawStyle = PropBag.ReadProperty("DrawStyle", CTBDrawStandard)
    pInitialise
    
End Sub

Private Sub UserControl_Terminate()
    debugmsg m_sCtlName & ",cToolbar:Enter Terminate"
    pTerminate
    debugmsg m_sCtlName & ",cToolbar:Exit Terminate"
    'MsgBox "cToolbar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Write properties:
    PropBag.WriteProperty "DrawStyle", DrawStyle, CTBDrawStandard
End Sub
