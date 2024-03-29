VERSION 5.00
Begin VB.UserControl cReBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   4905
   ToolboxBitmap   =   "cReBar.ctx":0000
   Begin VB.Label lblRebar 
      Caption         =   "'Rebar Control'"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "cReBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =========================================================================
' vbAccelerator Rebar control v3.0
' Copyright © 1998-2000 Steve McMahon (steve@vbaccelerator.com)
'
' This is a complete rebar implementation.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================
Private Type NMREBAR
    hdr As NMHDR
    dwMask As Long
    uBand As Long
    fStyle As Long
    wId As Long
    lParam As Long
End Type
Private Type NMRBAUTOSIZE
    hdr As NMHDR
    fChanged As Long
    rcTarget As RECT
    rcActual As RECT
End Type
Private Type NMREBARCHILDSIZE
    hdr As NMHDR
    uBand As Long
    wId As Long
    rcChild As RECT
    rcBand As RECT
End Type
Private Type NMREBARCHEVRON
   hdr As NMHDR
   uBand As Long
   wId As Long
   lParam As Long
   rcChevron As RECT
End Type
Private Type REBARINFO
    cbSize As Integer
    fMask As Integer
    hIml As Long
End Type
Private Type REBARBANDINFO
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Long
    hWndCHild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long
    wId As Long
End Type
Private Type REBARBANDINFO_471
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Integer 'Image
    hWndCHild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wId As Long
    cyChild As Long
    cyMaxChild As Long
    cyIntegral As Long
    cxIdeal As Long
    lParam As Long
    cxHeader As Long
End Type
Private Type REBARBANDINFO_NOTEXT_471
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As Long
    cch As Long
    iImage As Integer 'Image
    hWndCHild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wId As Long
    cyChild As Long
    cyMaxChild As Long
    cyIntegral As Long
    cxIdeal As Long
    lParam As Long
    cxHeader As Long
End Type

'Rebar Styles
Private Const RBS_TOOLTIPS = &H100&
Private Const RBS_VARHEIGHT = &H200&
Private Const RBS_BANDBORDERS = &H400&
Private Const RBS_FIXEDORDER = &H800&
Private Const RBS_AUTOSIZE = &H2000&
Private Const RBS_VERTICALGRIPPER = &H4000& '  // this always has the vertical gripper (default for horizontal mode)
Private Const RBS_DBLCLKTOGGLE = &H8000&

Private Const RBBS_BREAK = &H1               ' break to new line
Private Const RBBS_FIXEDSIZE = &H2           ' band can't be sized
Private Const RBBS_CHILDEDGE = &H4           ' edge around top & bottom of child window
Private Const RBBS_NOVERT = &H10             ' don't show when vertical
Private Const RBBS_FIXEDBMP = &H20           ' bitmap doesn't move during band resize
Private Const RBBS_VARIABLEHEIGHT = &H40
Private Const RBBS_GRIPPERALWAYS = &H80      ' always show the gripper
Private Const RBBS_NOGRIPPER = &H100 '// never show the gripper
Private Const RBBS_CHEVRON = &H200& ' // If you set cxIdeal, version 5.00 only...

Private Const RBS_EX_OFFICE9 = &H1&     '// new gripper, chevron, focus handling

Private Const RBBIM_COLORS = &H2
Private Const RBBIM_TEXT = &H4
Private Const RBBIM_IMAGE = &H8
Private Const RBBIM_CHILDSIZE = &H20
Private Const RBBIM_SIZE = &H40
Private Const RBBIM_BACKGROUND = &H80
Private Const RBBIM_ID = &H100
' 4.72 +
Private Const RBBIM_IDEALSIZE = &H200
Private Const RBBIM_LPARAM = &H400
Private Const RBBIM_HEADERSIZE = &H800

Private Const RB_INSERTBANDA = (WM_USER + 1)
Private Const RB_DELETEBAND = (WM_USER + 2)
Private Const RB_GETBARINFO = (WM_USER + 3)
Private Const RB_SETBARINFO = (WM_USER + 4)
Private Const RB_SETBANDINFOA = (WM_USER + 6)
Private Const RB_SETPARENT = (WM_USER + 7)
Private Const RB_HITTEST = (WM_USER + 8)
Private Const RB_GETRECT = (WM_USER + 9)
Private Const RB_INSERTBANDW = (WM_USER + 10)
Private Const RB_SETBANDINFOW = (WM_USER + 11)
Private Const RB_GETROWCOUNT = (WM_USER + 13)
Private Const RB_GETROWHEIGHT = (WM_USER + 14)

Private Const RB_IDTOINDEX = (WM_USER + 16)    '// wParam == id
Private Const RB_GETTOOLTIPS = (WM_USER + 17)
Private Const RB_SETTOOLTIPS = (WM_USER + 18)
Private Const RB_SETBKCOLOR = (WM_USER + 19)
Private Const RB_GETBKCOLOR = (WM_USER + 20)
Private Const RB_SETTEXTCOLOR = (WM_USER + 21)
Private Const RB_GETTEXTCOLOR = (WM_USER + 22)
Private Const RB_SIZETORECT = (WM_USER + 23)   '// resize the rebar/break bands and such to this rect (lparam)

Private Const RB_BEGINDRAG = (WM_USER + 24)
Private Const RB_ENDDRAG = (WM_USER + 25)
Private Const RB_DRAGMOVE = (WM_USER + 26)
Private Const RB_GETBARHEIGHT = (WM_USER + 27)

Private Const RB_GETBANDINFOA = (WM_USER + 29)

Private Const RB_MINIMIZEBAND = (WM_USER + 30)
Private Const RB_MAXIMIZEBAND = (WM_USER + 31)

Private Const RB_SHOWBAND = (WM_USER + 35)         '// show/hide band
Private Const RB_SETPALETTE = (WM_USER + 37)
Private Const RB_GETPALETTE = (WM_USER + 38)
Private Const RB_MOVEBAND = (WM_USER + 39)         ' // move band

Private Const RB_SETBANDFOCUS = (WM_USER + 40) '// (UINT) wParam == band index      lParam == TRUE/FALSE
                                        '// returns TRUE if gave band focus, else FALSE
Private Const RB_GETBANDFOCUS = (WM_USER + 41) '// returns index of band with focus (-1 if none)
Private Const RB_CYCLEFOCUS = (WM_USER + 42)    '// (UINT) wParam == band index      (BOOL) lParam == back/forward
                                                '// returns index of band that got focus (-1 if none)
Private Const RB_SETEXTENDEDSTYLE = (WM_USER + 43)


Private Const RBHT_NOWHERE = &H1
Private Const RBHT_CAPTION = &H2
Private Const RBHT_CLIENT = &H3
Private Const RBHT_GRABBER = &H4
Private Const RBHT_CHEVRON = &H8

Private Const RB_INSERTBAND = RB_INSERTBANDA
Private Const RB_SETBANDINFO = RB_SETBANDINFOA
Private Const RB_GETBANDINFO471 = RB_GETBANDINFOA

Private Const RBN_FIRST = H_MAX - 831                  '// rebar
Private Const RBN_LAST = H_MAX - 859
Private Const RBN_HEIGHTCHANGE = (RBN_FIRST - 0)
Private Const RBN_GETOBJECT = (RBN_FIRST - 1)
Private Const RBN_LAYOUTCHANGED = (RBN_FIRST - 2)
Private Const RBN_AUTOSIZE = (RBN_FIRST - 3)
Private Const RBN_BEGINDRAG = (RBN_FIRST - 4)
Private Const RBN_ENDDRAG = (RBN_FIRST - 5)
Private Const RBN_DELETINGBAND = (RBN_FIRST - 6)       '// Uses NMREBAR
Private Const RBN_DELETEDBAND = (RBN_FIRST - 7)        '// Uses NMREBAR
Private Const RBN_CHILDSIZE = (RBN_FIRST - 8)
Private Const RBN_SETFOCUS = (RBN_FIRST - 9)            '// Uses NMREBAR
Private Const RBN_CHEVRONPUSHED = (RBN_FIRST - 10)

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

' ==============================================================================
' INTERFACE
' ==============================================================================
' Enumerations:
Public Enum ERBPositionConstants
   erbPositionTop
   erbPositionLeft
   erbPositionRight
   erbPositionBottom
End Enum
Public Enum ECRBImageSourceTypes
    CRBResourceBitmap
    CRBLoadFromFile
    CRBPicture
End Enum

' Internal Implementation:
Private m_hWnd As Long ' Rebar
Private m_hWndCtlParent As Long ' Rebar window parent
Private m_hWndMsgParent As Long ' Where messages are sent
Private m_bSubClassing As Boolean
Private m_bInTerminate As Boolean
Private m_bKillChildren As Boolean
Private m_lMajor As Long, m_lMinor As Long

' Position:
Private m_ePosition As ERBPositionConstants

' Background imaage:
Private m_sPicture As String
Private m_lResourceID As Long
Private m_hInstance As Long
Private m_pic As StdPicture
Private m_hBmp As Long
Private m_eImageSourceType As ECRBImageSourceTypes

' Band original location information:
Private Type tRebarWndStore
   hwndItem As Long
   hWndItemParent As Long
   tR As RECT
End Type
Private m_tWndStore() As tRebarWndStore
Private m_iWndItemCount As Integer

' Band keys:
Private Type tRebarDataStore
   wId As Long
   vData As Variant
   bFixedSize As Boolean
   sBandText As String
End Type
Private m_tDataStore() As tRebarDataStore
Private m_lIDCount As Long

Private m_bVisible As Boolean

Private m_sCtlName As String

Implements ISubclass

' Events:
Public Event HeightChanged(lNewHeight As Long)
Attribute HeightChanged.VB_Description = "Raised whenever the height of the rebar changes, for example when the user moves the bands around. "
Public Event BeginBandDrag(ByVal wId As Long, ByRef bCancel As Boolean)
Attribute BeginBandDrag.VB_Description = "Raised when the user is about to start dragging a band."
Public Event EndBandDrag(ByVal wId As Long)
Attribute EndBandDrag.VB_Description = "Raised when the user has completed dragging a band within the rebar."
Public Event BandChildResize(ByVal wId As Long, ByVal lBandLeft As Long, ByVal lBandTop As Long, ByVal lBandRight As Long, ByVal lBandBottom As Long, ByRef lChildLeft As Long, ByRef lChildTop As Long, ByRef lChildRight As Long, ByRef lChildBottom As Long)
Attribute BandChildResize.VB_Description = "Raised whenever a child is resized because of a change in size of a band."
Public Event LayoutChanged()
Attribute LayoutChanged.VB_Description = "Raised whenever the layout of the rebar bands changes, due to either the rebar being resized or the user dragging the bands."
Public Event ChevronPushed(ByVal wId As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Attribute ChevronPushed.VB_Description = "Raised when a band chevron is pressed."

Public Function SaveLayout() As String
Dim i As Long
Dim lBandId As Long
Dim lBandPos As Long
Dim lLeft As Long
Dim lTop As Long
Dim lRight As Long
Dim lBottom As Long
Dim doc As DOMDocument
Dim nodWork As IXMLDOMElement
Dim nodBand As IXMLDOMElement
Dim attr As IXMLDOMAttribute
Dim lIndexInt As Long
Dim j As Long
   
   Set doc = New DOMDocument
   Set nodWork = doc.createElement("Bands")
   Set attr = doc.createAttribute("count")
   attr.Value = BandCount
   nodWork.Attributes.setNamedItem attr
   Set attr = doc.createAttribute("rowCount")
   attr.Value = RowCount
   nodWork.Attributes.setNamedItem attr
   
   For i = 0 To BandCount - 1
      
      Set nodBand = doc.createElement("Band")
      lBandId = BandIDForIndex(i)
      Set attr = doc.createAttribute("id")
      attr.Value = lBandId
      nodBand.Attributes.setNamedItem attr
      
      Set attr = doc.createAttribute("data")
      attr.Value = BandData(lBandId)
      nodBand.Attributes.setNamedItem attr
      
      GetBandRectangle i, lLeft, lTop, lRight, lBottom
      
      Set attr = doc.createAttribute("left")
      attr.Value = lLeft
      nodBand.Attributes.setNamedItem attr
      Set attr = doc.createAttribute("top")
      attr.Value = lTop
      nodBand.Attributes.setNamedItem attr
      Set attr = doc.createAttribute("right")
      attr.Value = lRight
      nodBand.Attributes.setNamedItem attr
      Set attr = doc.createAttribute("bottom")
      attr.Value = lBottom
      nodBand.Attributes.setNamedItem attr

      Set attr = doc.createAttribute("bandChildMinWidth")
      attr.Value = BandChildMinWidth(i)
      nodBand.Attributes.setNamedItem attr
      
      Set attr = doc.createAttribute("bandChildIdealWidth")
      attr.Value = BandChildIdealWidth(i)
      nodBand.Attributes.setNamedItem attr
      
      Set attr = doc.createAttribute("chevron")
      attr.Value = IIf(BandChevron(i), "y", "n")
      nodBand.Attributes.setNamedItem attr
      
      lIndexInt = 0
      For j = 1 To m_lIDCount
         If (m_tDataStore(j).wId = lBandId) Then
            lIndexInt = j
            Exit For
         End If
      Next j
      
      If (lIndexInt > 0) Then
         Set attr = doc.createAttribute("text")
         attr.Value = m_tDataStore(lIndexInt).sBandText
         nodBand.Attributes.setNamedItem attr
         
         Set attr = doc.createAttribute("fixedBand")
         attr.Value = IIf(m_tDataStore(lIndexInt).bFixedSize, "y", "n")
         nodBand.Attributes.setNamedItem attr
         
      End If
      
      nodWork.appendChild nodBand
      
   Next i
   
   doc.appendChild nodWork
   
   SaveLayout = doc.xml
End Function
Public Sub RestoreLayout(ByVal sXml As String, ByRef sData() As String, ByRef lhWnd() As Long)
Dim dom As DOMDocument
Dim nodTop As IXMLDOMElement
Dim nodBand As IXMLDOMElement
Dim attr As IXMLDOMAttribute
Dim lLastTop As Long
Dim lLeft As Long
Dim lTop As Long
Dim lRight As Long
Dim lBottom As Long
Dim lMinWidth As Long
Dim lIdealWidth As Long
Dim sBandText As String
Dim sThisData As String
Dim lItemIndex As Long
Dim bNewLine As Boolean
Dim bChevron As Boolean
Dim lAddIdealWidth As Long
Dim i As Long
Dim bFixed As Boolean
Dim lIdealWidthA() As Long
Dim lMinWidthA() As Long

   Set dom = New DOMDocument
   dom.LoadXml sXml
   Set nodTop = dom.SelectSingleNode("Bands")
   If Not nodTop Is Nothing Then
      
      ReDim lIdealWidthA(0 To UBound(sData) - LBound(sData) + 1) As Long
      ReDim lMinWidthA(0 To UBound(sData) - LBound(sData) + 1) As Long
      
      For Each nodBand In nodTop.selectNodes("Band")
         
         lLeft = 0
         lTop = 0
         lRight = 0
         lBottom = 0
         lMinWidth = 0
         lIdealWidth = 0
         sThisData = ""
         sBandText = ""
         bFixed = False
         bChevron = False
         
         For Each attr In nodBand.Attributes
            Select Case attr.Name
            Case "data"
               sThisData = attr.Value
            Case "left"
               lLeft = CLng(attr.Value)
            Case "top"
               lTop = CLng(attr.Value)
            Case "right"
               lRight = CLng(attr.Value)
            Case "bottom"
               lBottom = CLng(attr.Value)
            Case "bandChildMinWidth"
               lMinWidth = CLng(attr.Value)
            Case "bandChildIdealWidth"
               lIdealWidth = CLng(attr.Value)
            Case "text"
               sBandText = attr.Value
            Case "fixedBand"
               bFixed = IIf(attr.Value = "y", True, False)
            Case "chevron"
               bChevron = IIf(attr.Value = "y", True, False)
            End Select
         Next
         
         If (lTop <> lLastTop) Then
            bNewLine = True
         Else
            bNewLine = False
         End If
         lLastTop = lTop
         
         lItemIndex = -1
         For i = LBound(sData) To UBound(sData)
            If (sData(i) = sThisData) Then
               lItemIndex = i
               Exit For
            End If
         Next i
         
         If (lItemIndex > -1) Then
            lAddIdealWidth = lRight - lLeft - 11
            AddBandByHwnd lhWnd(lItemIndex), sBandText, bNewLine, bFixed, sThisData
            lIdealWidthA(BandCount - 1) = lIdealWidth
            lMinWidthA(BandCount - 1) = lMinWidth
            
            BandChevron(BandCount - 1) = bChevron
            BandChildIdealWidth(BandCount - 1) = lAddIdealWidth
            BandChildMinWidth(BandCount - 1) = lAddIdealWidth * 0.6
            BandMaximise BandCount - 1
            BandChildMinWidth(i) = lMinWidthA(i)
            BandChildIdealWidth(i) = lIdealWidthA(i)
         End If
      Next
            
   End If
End Sub

Public Property Get RowCount() As Long
   RowCount = SendMessage(m_hWnd, RB_GETROWCOUNT, 0&, ByVal 0&)
End Property
Public Property Get RowHeight(ByVal lBand As Long) As Long
   If (lBand >= 0) And (lBand <= BandCount) Then
      RowHeight = SendMessage(m_hWnd, RB_GETROWHEIGHT, lBand, ByVal 0&)
   Else
      ' IncorrectBand
   End If
End Property

Public Sub Autosize()
Attribute Autosize.VB_Description = "Attempts to automatically move the Rebar bands so they best fit the specified rectangle (in pixels relative to the rebar's container).   Not available for COMCTL32.DLL version below 4.71."
Dim lWidth As Long
Dim lHeight As Long
Dim rc As RECT, rcP As RECT
   If (m_ePosition = erbPositionBottom) Or (m_ePosition = erbPositionTop) Then
      GetWindowRect m_hWndCtlParent, rcP
      lWidth = rcP.right - rcP.left
      lHeight = RebarHeight
   Else
      GetWindowRect m_hWndCtlParent, rcP
      lHeight = rcP.bottom - rcP.top
      lWidth = RebarWidth
   End If
   rc.right = lWidth
   rc.bottom = lHeight
   SendMessage m_hWnd, RB_SIZETORECT, 0, rc
End Sub

Public Property Get Position() As ERBPositionConstants
Attribute Position.VB_Description = "Gets/sets the orientation of the rebar on its container."
Attribute Position.VB_MemberFlags = "400"
   Position = m_ePosition
End Property
Public Property Let Position(ByVal ePosition As ERBPositionConstants)
Dim dwStyle As Long
Dim dwNewStyle As Long
Dim hWndP As Long
Dim rc As RECT
   If (m_ePosition <> ePosition) Then
      m_ePosition = ePosition
      
      If (m_hWnd <> 0) Then
         SetProp m_hWnd, "vbal:cRebarPosition", m_ePosition
         
         ' Move...
         dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
         dwNewStyle = dwStyle
         dwNewStyle = dwNewStyle And Not (CCS_LEFT Or CCS_TOP Or CCS_RIGHT Or CCS_BOTTOM)
         Select Case m_ePosition
         Case erbPositionTop
            dwNewStyle = dwNewStyle Or CCS_TOP
         Case erbPositionRight
            dwNewStyle = dwNewStyle Or CCS_RIGHT
         Case erbPositionLeft
            dwNewStyle = dwNewStyle Or CCS_LEFT
         Case erbPositionBottom
            dwNewStyle = dwNewStyle Or CCS_BOTTOM
         End Select
         If dwNewStyle <> dwStyle Then
            SetWindowLong m_hWnd, GWL_STYLE, dwNewStyle
         End If
         
         RebarSize
         RaiseEvent HeightChanged(RebarHeight)
         RebarSize
         
      End If
      
   End If
End Property

Private Sub pCreateSubClass()
   If Not (m_bSubClassing) Then
      If m_hWnd <> 0 Then
         m_hWndMsgParent = getFormParenthWnd(UserControl.hwnd)
         If (m_hWndMsgParent > 0) Then
            ' Debug.Print "Subclassing window: " & m_hWndMsgParent
            AttachMessage Me, m_hWndMsgParent, WM_NOTIFY
            AttachMessage Me, m_hWnd, WM_DESTROY
            AttachMessage Me, m_hWndMsgParent, WM_DESTROY
            m_bSubClassing = True
         End If
         SendMessageLong m_hWnd, RB_SETPARENT, m_hWndMsgParent, 0
      End If
   End If
End Sub

Private Sub pDestroySubClass()
   If (m_bSubClassing) Then
      DetachMessage Me, m_hWndMsgParent, WM_NOTIFY
      DetachMessage Me, m_hWnd, WM_DESTROY
      DetachMessage Me, m_hWndMsgParent, WM_DESTROY
      m_hWndMsgParent = 0
      m_bSubClassing = False
   End If
End Sub

' Interface properties
Private Property Get ISubclass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_DESTROY
      ISubclass_MsgResponse = emrPreprocess
   Case Else
      ISubclass_MsgResponse = emrPreprocess
   End Select
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
   '
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Dim lHeight As Long
Dim tNMH As NMHDR
Dim tNMR As NMREBAR
Dim tNMRBA As NMRBAUTOSIZE
Dim tNMRCS As NMREBARCHILDSIZE
Dim tNMRC As NMREBARCHEVRON
Dim tNMMouse As NMMOUSE
Dim tR As RECT
Dim bCancel As Boolean
Dim rcChild As RECT
Dim i As Long
Dim lhWnd As Long
Dim wId As Long
   
   ' Don't try to raise events when the control is terminating -
   ' you will crash!
   'If Not (m_bInTerminate) And Not (m_hWnd = 0 Or m_hWndMsgParent = 0) Then
   
      If iMsg = WM_NOTIFY Then
         CopyMemory tNMH, ByVal lParam, Len(tNMH)
         If tNMH.hwndFrom = m_hWnd Then
         
            Select Case tNMH.code
            Case NM_NCHITTEST
               ' NC hittest.  Apparently we can return alternative HT_ values
               ' here but I cannot get it to do anything
               CopyMemory tNMMouse, ByVal lParam, Len(tNMMouse)
               ' ...
               
            Case RBN_HEIGHTCHANGE
               ' Height change notification:
               RebarSize
               lHeight = RebarHeight
               RaiseEvent HeightChanged(lHeight)
            
            Case RBN_AUTOSIZE
               ' Autosize notification, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  CopyMemory tNMRBA, ByVal lParam, Len(tNMRBA)
                  ' This event isn't of any use because the CCS_NORESIZE style
                  ' is set.  I do not recommend turning CCS_NORESIZE off as it
                  ' is very easy to get infinite loops during resize code without
                  ' it...
               End If
               
            Case RBN_BEGINDRAG, RBN_ENDDRAG
               ' Band dragging notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' user began dragging a band:
                  CopyMemory tNMR, ByVal lParam, Len(tNMR)
                  If tNMR.uBand > -1 Then
                     If tNMH.code = RBN_BEGINDRAG Then
                        bCancel = False
                        RaiseEvent BeginBandDrag(tNMR.wId, bCancel)
                        If bCancel Then
                           ISubclass_WindowProc = 1
                        Else
                           ISubclass_WindowProc = 0
                        End If
                     Else
                        RaiseEvent EndBandDrag(tNMR.wId)
                     End If
                  Else
                     ' no band affected.
                     RaiseEvent EndBandDrag(-1)
                  End If
               End If
            
            Case RBN_CHILDSIZE
               ' Child size change notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' user began dragging a band:
                  CopyMemory tNMRCS, ByVal lParam, Len(tNMRCS)
                  LSet rcChild = tNMRCS.rcChild
                  RaiseEvent BandChildResize(tNMRCS.wId, tNMRCS.rcBand.left, tNMRCS.rcBand.top, tNMRCS.rcBand.right, tNMRCS.rcBand.bottom, rcChild.left, rcChild.top, rcChild.right, rcChild.bottom)
                  If rcChild.left <> tNMRCS.rcChild.left Or rcChild.top <> tNMRCS.rcChild.top Or rcChild.right <> tNMRCS.rcChild.right Or rcChild.bottom <> tNMRCS.rcChild.bottom Then
                     LSet tNMRCS.rcChild = rcChild
                     CopyMemory ByVal lParam, tNMRCS, Len(tNMRCS)
                  End If
                  'Debug.Print tNMRCS.rcBand.left, tNMRCS.rcBand.top, tNMRCS.rcBand.right, tNMRCS.rcBand.bottom
                  ISubclass_WindowProc = 1
               End If
            
            Case RBN_DELETEDBAND, RBN_DELETINGBAND
               ' band deletion notifications, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  ' A band has just been deleted:
                  CopyMemory tNMR, ByVal lParam, Len(tNMR)
                  If tNMH.code = RBN_DELETEDBAND Then
                     pRemoveID tNMR.wId
                  Else
                     lhWnd = plGetHwndOfBandChild(m_hWnd, tNMR.uBand, wId)
                     If lhWnd <> 0 Then
                        pResetParent lhWnd
                     End If
                  End If
               End If
                     
            Case RBN_LAYOUTCHANGED
               ' layout changed notification, 4.71+
               If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
                  RaiseEvent LayoutChanged
               End If
            
            Case RBN_CHEVRONPUSHED
               'Debug.Print "Chevron Pushed"
               SendMessageLong getFormParenthWnd(m_hWnd), WM_CANCELMODE, 0, 0
               If m_lMajor >= 5 Then
                  CopyMemory tNMRC, ByVal lParam, Len(tNMRC)
                  LSet tR = tNMRC.rcChevron
                  MapWindowPoints m_hWnd, HWND_DESKTOP, tR, 2
                  tR.left = tR.left * Screen.TwipsPerPixelX
                  tR.top = tR.top * Screen.TwipsPerPixelY
                  tR.right = tR.right * Screen.TwipsPerPixelX
                  tR.bottom = tR.bottom * Screen.TwipsPerPixelY
                  RaiseEvent ChevronPushed(tNMRC.wId, tR.left, tR.top, tR.right, tR.bottom)
               End If
            
            'Case Else
            '   Debug.Print tNMH.code
               
            End Select
         
         Else
            Select Case tNMH.code
            Case TBN_QUERYINSERT
               ISubclass_WindowProc = g_lCustomiseResponse
            Case TBN_QUERYDELETE
               ISubclass_WindowProc = g_lCustomiseResponse
            End Select
         End If
         
      ElseIf iMsg = WM_DESTROY Then
         debugmsg m_sCtlName & ":WM_DESTROY," & Hex$(hwnd)
         DestroyRebar
         
      End If
   
   'End If

End Function

Public Property Get BandVisible(ByVal lBand As Long) As Boolean
Attribute BandVisible.VB_Description = "Gets/sets whether a rebar band is visible or not.  Not available for COMCTL32.DLL version below 4.71."
Dim lStyle As Long
    If (lBand >= 0) And (lBand < BandCount) Then
        If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
            BandVisible = ((lStyle And RBBS_HIDDEN) <> RBBS_HIDDEN)
        End If
    Else
        BandVisible = False
    End If
   
End Property
Public Property Let BandVisible(ByVal lBand As Long, ByVal bState As Boolean)
Dim lS As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      lS = Abs(bState)
      SendMessageLong m_hWnd, RB_SHOWBAND, lBand, lS
   End If
End Property
Public Property Get BandChildEdge(ByVal lBand As Long) As Boolean
Attribute BandChildEdge.VB_Description = "Gets/sets whether a band draws a narrow  internal border around the child control."
Dim lStyle As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandChildEdge = ((lStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE)
          End If
      Else
          BandChildEdge = False
      End If
   Else
      'Unsupported
   End If
   
End Property
Public Property Let BandChildEdge(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle And Not RBBS_CHILDEDGE
               Else
                  lStyle = lStyle Or RBBS_CHILDEDGE
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE
                  .fStyle = lStyle
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Get BandGripper(ByVal lBand As Long) As Boolean
Attribute BandGripper.VB_Description = "Gets/sets whether a rebar band has a gripper or not.  (COMCTL32.DLL v5 or higher only)"
Dim lStyle As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandGripper = ((lStyle And RBBS_NOGRIPPER) <> RBBS_NOGRIPPER)
          End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Let BandGripper(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_NOGRIPPER) <> RBBS_NOGRIPPER)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle Or RBBS_NOGRIPPER
               Else
                  lStyle = lStyle And Not RBBS_NOGRIPPER
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE
                  .fStyle = lStyle
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Get BandChevron(ByVal lBand As Long) As Boolean
Attribute BandChevron.VB_Description = "Gets/sets whether a band will show  a chevron if it is sized too small for the contents to fit. (COMCTL32.DLL v5 or higher only)"
Dim lStyle As Long
   If m_lMajor >= 5 Then
      If (lBand >= 0) And (lBand < BandCount) Then
          If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE, fStyle:=lStyle)) Then
              BandChevron = ((lStyle And RBBS_CHEVRON) = RBBS_CHEVRON)
          End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property
Public Property Let BandChevron(ByVal lBand As Long, ByVal bState As Boolean)
Dim lStyle As Long
Dim lCX As Long
Dim bCurrent As Boolean
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471

   If m_lMajor >= 5 Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_STYLE Or RBBIM_CHILDSIZE, cxMinChild:=lCX, fStyle:=lStyle) Then
            bCurrent = ((lStyle And RBBS_CHEVRON) = RBBS_CHEVRON)
            If bState <> bCurrent Then
               If bCurrent Then
                  lStyle = lStyle And Not RBBS_CHEVRON
               Else
                  lStyle = lStyle Or RBBS_CHEVRON
               End If
               With tRbbi471
                  .cbSize = LenB(tRbbi471)
                  .fMask = RBBIM_STYLE Or RBBIM_IDEALSIZE
                  .fStyle = lStyle
                  .cxIdeal = lCX
               End With
               SendMessage m_hWnd, RB_SETBANDINFO, lBand, tRbbi471
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      'Unsupported
   End If
End Property

Public Property Get BandChildMinHeight(ByVal lBand As Long) As Long
Attribute BandChildMinHeight.VB_Description = "Gets/sets the minimum height of a rebar band."
Dim cy As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cyMinChild:=cy)) Then
         BandChildMinHeight = cy
      End If
   Else
      BandChildMinHeight = -1
      ' IncorrectBand
   End If
End Property
Public Property Let BandChildMinHeight(ByVal lBand As Long, lHeight As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      Dim tRbbi As REBARBANDINFO_NOTEXT
      Dim lR As Long
      tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
      tRbbi.cbSize = Len(tRbbi)
      lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         If (tRbbi.hWndCHild <> 0) Then
            tRbbi.fMask = RBBIM_CHILDSIZE
            tRbbi.cyMinChild = lHeight
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
         End If
      End If
   Else
      ' IncorrectBand
   End If
End Property
Public Property Get BandChildMaxHeight(ByVal lBand As Long) As Long
Attribute BandChildMaxHeight.VB_Description = "Gets/sets the maximum height a band can size to (COMCTL32.DLL v5 or higher only)"
Dim cy As Long
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cyMaxChild:=cy)) Then
            BandChildMaxHeight = cy
         End If
      Else
         BandChildMaxHeight = -1
         ' IncorrectBand
      End If
   Else
      ' Unsupported
   End If
End Property
Public Property Let BandChildMaxHeight(ByVal lBand As Long, lHeight As Long)
   If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
      If (lBand >= 0) And (lBand < BandCount) Then
         Dim tRbbi As REBARBANDINFO_NOTEXT_471
         Dim lR As Long
         tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
         tRbbi.cbSize = Len(tRbbi)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
         If (lR <> 0) Then
            If (tRbbi.hWndCHild <> 0) Then
               tRbbi.fMask = RBBIM_CHILDSIZE
               tRbbi.cyMaxChild = lHeight
               lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            End If
         End If
      Else
         ' IncorrectBand
      End If
   Else
      ' Unsupported
   End If
End Property
Public Property Get BandChildMinWidth(ByVal lBand As Long) As Long
Attribute BandChildMinWidth.VB_Description = "Gets/sets the minimum width of rebar band."
Dim cx As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cxMinChild:=cx)) Then
         BandChildMinWidth = cx
      End If
   Else
      BandChildMinWidth = -1
      ' IncorrectBand
   End If

End Property
Public Property Let BandChildMinWidth(ByVal lBand As Long, lWidth As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      Dim tRbbi As REBARBANDINFO_NOTEXT
      Dim lR As Long
      Dim tR As RECT
      
      tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
      tRbbi.cbSize = Len(tRbbi)
      lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         If (tRbbi.hWndCHild <> 0) Then
            tRbbi.fMask = RBBIM_CHILDSIZE
            tRbbi.cxMinChild = lWidth
            lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
         End If
      End If
   Else
      ' IncorrectBand
   End If
End Property
Public Property Get BandChildIdealWidth(ByVal lBand As Long) As Long
Dim cx As Long
   If (lBand >= 0) And (lBand < BandCount) Then
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         Dim tRbbi As REBARBANDINFO_NOTEXT_471
         Dim lR As Long
         Dim tR As RECT
      
         tRbbi.fMask = RBBIM_IDEALSIZE Or RBBIM_CHILD
         tRbbi.cbSize = Len(tRbbi)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
         If (lR <> 0) Then
            BandChildIdealWidth = tRbbi.cxIdeal
         End If
      Else
         ' unsupported
      End If
   Else
      BandChildIdealWidth = -1
      ' IncorrectBand
   End If

End Property
Public Property Let BandChildIdealWidth(ByVal lBand As Long, lWidth As Long)
Static s_bLock As Boolean
Dim j As Long
   
   If (lBand >= 0) And (lBand < BandCount) Then
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         Dim tRbbi As REBARBANDINFO_NOTEXT_471
         Dim lR As Long
         Dim tR As RECT
      
         tRbbi.fMask = RBBIM_IDEALSIZE Or RBBIM_CHILD
         tRbbi.cbSize = Len(tRbbi)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
         If (lR <> 0) Then
            If (tRbbi.hWndCHild <> 0) Then
               tRbbi.fMask = RBBIM_IDEALSIZE
               tRbbi.cxIdeal = lWidth
               lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
               SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
            End If
         End If
      Else
         ' unsupported
      End If
   Else
      ' IncorrectBand
   End If
End Property

Public Sub BandChildResized(ByVal lBand As Long, ByVal lWidth As Long, ByVal lHeight As Long)
   If (lBand >= 0) And (lBand < BandCount) Then
      
      Dim tRBandNT As REBARBANDINFO_NOTEXT
      Dim tRBandNT471 As REBARBANDINFO_NOTEXT_471
      Dim lR As Long
      Dim tR As RECT
      
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         tRBandNT471.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD Or RBBIM_STYLE
         tRBandNT471.cbSize = LenB(tRBandNT471)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRBandNT471)
      Else
         tRBandNT.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD Or RBBIM_STYLE
         tRBandNT.cbSize = Len(tRBandNT)
         lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRBandNT)
      End If
      
      If (lR <> 0) Then
         If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
            If (tRBandNT471.hWndCHild <> 0) Then
               tRBandNT471.cyMinChild = lHeight
               tRBandNT471.cx = lWidth
               If (tRBandNT471.fStyle And RBBS_CHEVRON) = RBBS_CHEVRON Then
                  tRBandNT471.cxMinChild = 24
                  tRBandNT471.fMask = tRBandNT471.fMask Or RBBIM_IDEALSIZE
                  tRBandNT471.cxIdeal = lWidth
               Else
                  tRBandNT471.cxMinChild = lWidth
               End If
               lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRBandNT471)
               GetWindowRect m_hWnd, tR
               MapWindowPoints 0, GetParent(m_hWnd), tR, 2
               MoveWindow m_hWnd, tR.left, tR.top, tR.right - tR.left + 2, tR.bottom - tR.top + 1, 1
               MoveWindow m_hWnd, tR.left, tR.top, tR.right - tR.left + 1, tR.bottom - tR.top + 1, 1
               SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
            End If
         Else
            If (tRBandNT.hWndCHild <> 0) Then
               tRBandNT.fMask = RBBIM_CHILDSIZE
               tRBandNT.cxMinChild = lWidth
               tRBandNT.cyMinChild = lHeight
               tRBandNT.cx = lWidth
               lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRBandNT)
               SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
            End If
         End If
      End If
   End If
End Sub

Public Sub BandMove(ByVal lBand As Long, ByVal lIndexTo As Long)
Attribute BandMove.VB_Description = "Moves a band from one position to another.  All bands in lower positions are moved up.   Not available for COMCTL32.DLL version below 4.71."
    If (lBand >= 0) And (lBand < BandCount) Then
      If (lIndexTo >= 0) And (lIndexTo < BandCount) Then
         SendMessageLong m_hWnd, RB_MOVEBAND, lBand, lIndexTo
      Else
         ' Incorrectband
      End If
   Else
      ' Incorrectband
   End If
End Sub
Public Sub BandMinimise(ByVal lBand As Long)
Attribute BandMinimise.VB_Description = "Minimises a rebar band in the current layout."
    If (lBand >= 0) And (lBand < BandCount) Then
        SendMessageLong m_hWnd, RB_MINIMIZEBAND, lBand, 0
    Else
      ' IncorrectBand
    End If
End Sub
Public Sub BandMaximise(ByVal lBand As Long)
Attribute BandMaximise.VB_Description = "Maximises a rebar band in the current layout."
    If (lBand >= 0) And (lBand < BandCount) Then
        SendMessageLong m_hWnd, RB_MAXIMIZEBAND, lBand, 1 'BandChildIdealWidth(lBand)
    Else
      ' IncorrectBand
    End If
End Sub
Public Sub GetBandRectangle( _
      ByVal lBand As Long, _
      Optional ByRef lLeft As Long, _
      Optional ByRef lTop As Long, _
      Optional ByRef lRight As Long, _
      Optional ByRef lBottom As Long _
   )
Attribute GetBandRectangle.VB_Description = "Returns the internal bounding rectangle for a rebar band. Not available for COMCTL32.DLL version below 4.71."
Dim tR As RECT
   If (lBand >= 0) And (lBand <= BandCount) Then
      SendMessage m_hWnd, RB_GETRECT, lBand, tR
      lLeft = tR.left
      lTop = tR.top
      lRight = tR.right
      lBottom = tR.bottom
   Else
      ' IncorrectBand
   End If
End Sub
Property Get BandCount() As Long
Attribute BandCount.VB_Description = "Returns the number of bands in the rebar."
    BandCount = SendMessage(m_hWnd, RB_GETBANDCOUNT, 0&, ByVal 0&)
End Property

Private Function pbGetBandInfo( _
        ByVal lhWnd As Long, _
        ByVal lBand As Long, _
        Optional ByRef fMask As Long, _
        Optional ByRef fStyle As Long, _
        Optional ByRef clrFore As Long, _
        Optional ByRef clrBack As Long, _
        Optional ByRef cch As Long, _
        Optional ByRef iImage As Integer, _
        Optional ByRef hWndCHild As Long, _
        Optional ByRef cxMinChild As Long, _
        Optional ByRef cyMinChild As Long, _
        Optional ByRef cx As Long, _
        Optional ByRef hbmpBack As Long, _
        Optional ByRef wId As Long, _
        Optional ByRef cyIntegral As Long, _
        Optional ByRef cyChild As Long, _
        Optional ByRef cyMaxChild As Long, _
        Optional ByRef lParam As Long, _
        Optional ByRef cxHeader As Long _
    ) As Boolean
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim tRbbi471 As REBARBANDINFO_NOTEXT_471
Dim lR As Long

   If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
      ' Use old version
      tRbbi.cbSize = LenB(tRbbi)
      tRbbi.fMask = fMask
      lR = SendMessage(lhWnd, RB_GETBANDINFO, lBand, tRbbi)
      If (lR <> 0) Then
         With tRbbi
            fMask = .fMask
            fStyle = .fStyle
            clrFore = .clrFore
            clrBack = .clrBack
            cch = .cch
            iImage = .iImage
            hWndCHild = .hWndCHild
            cxMinChild = .cxMinChild
            cyMinChild = .cyMinChild
            cx = .cx
            hbmpBack = .hbmBack
            wId = .wId
         End With
         pbGetBandInfo = True
      End If
   Else
      tRbbi471.cbSize = LenB(tRbbi471)
      tRbbi471.fMask = fMask
      lR = SendMessage(lhWnd, RB_GETBANDINFO471, lBand, tRbbi471)
      If (lR <> 0) Then
         With tRbbi471
            fMask = .fMask
            fStyle = .fStyle
            clrFore = .clrFore
            clrBack = .clrBack
            cch = .cch
            iImage = .iImage
            hWndCHild = .hWndCHild
            cxMinChild = .cxMinChild
            cyMinChild = .cyMinChild
            cx = .cx
            hbmpBack = .hbmBack
            cyIntegral = .cyIntegral
            cyChild = .cyChild
            cyMaxChild = .cyMaxChild
            cyMinChild = .cyMinChild
            cxHeader = .cxHeader
            lParam = .lParam
            wId = .wId
         End With
         pbGetBandInfo = True
       End If
   End If
End Function
Public Property Get HasBitmap() As Boolean
Attribute HasBitmap.VB_Description = "Returns whether a background bitmap is loaded into the rebar or not."
   HasBitmap = (BackgroundBitmapHandle <> 0)
End Property

Public Property Let ImageSource( _
        ByVal eType As ECRBImageSourceTypes _
    )
Attribute ImageSource.VB_Description = "Specifies which type of bitmap source (file, picture or resource) should be used as the source of the rebar's background bitmap."
    m_eImageSourceType = eType
End Property
Public Property Let ImageResourceID(ByVal lResourceId As Long)
Attribute ImageResourceID.VB_Description = "Sets a resource id to be used  to be used as the source of the rebar's background bitmap."
   ClearPicture
   m_lResourceID = lResourceId
End Property
Public Property Let ImageResourcehInstance(ByVal hInstance As Long)
Attribute ImageResourcehInstance.VB_Description = "Specifies the hInstance from which to load the resource set by the ImageResourceID property."
   m_hInstance = hInstance
End Property
Public Property Let ImageFile(ByVal sFile As String)
Attribute ImageFile.VB_Description = "Sets a bitmap file to be used as the source of the rebar's background bitmap."
   ClearPicture
   m_sPicture = sFile
End Property
Public Property Let ImagePicture(ByVal picThis As StdPicture)
Attribute ImagePicture.VB_Description = "ets a picture object to be used as the source of the rebar's background bitmap."
   ClearPicture
   Set m_pic = picThis
End Property
Public Property Set ImagePicture(ByVal picThis As StdPicture)
   ClearPicture
   Set m_pic = picThis
End Property
Public Property Get BackgroundBitmap() As String
Attribute BackgroundBitmap.VB_Description = "Gets/sets the background bitmap file.  Has no effect unless it is called before the rebar is created.  Note: you can't recreate a rebar at run-time if you have COMCTL32.DLL version lower than 4.71."
   BackgroundBitmap = m_sPicture
End Property
Public Property Let BackgroundBitmap(ByVal sFile As String)
   ImageSource = CRBLoadFromFile
   ImageFile = sFile
End Property
Private Property Get BackgroundBitmapHandle() As Long

   ' Set up the picture if we don't already have one:
   If (m_hBmp = 0) Then
      Select Case m_eImageSourceType
      Case CRBPicture
         If Not (m_pic Is Nothing) Then
            m_hBmp = hBmpFromPicture(m_pic)
         End If
      Case CTBLoadFromFile
         If (m_sPicture <> "") Then
            m_hBmp = LoadImage(0, m_sPicture, IMAGE_BITMAP, 0, 0, _
                     LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
         End If
      Case CTBResourceBitmap
         m_hBmp = LoadImageLong(m_hInstance, m_lResourceID, IMAGE_BITMAP, 0, 0, _
                     LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
      End Select
   End If

   BackgroundBitmapHandle = m_hBmp
   
End Property

Public Function AddBandByHwnd( _
        ByVal hwnd As Long, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByVal vData As Variant _
    ) As Long
Attribute AddBandByHwnd.VB_Description = "Adds a band to the rebar and sets the band to contain the window with the specified hWnd."
Dim hBmp As Long
Dim lX As Long
Dim lBand As Long
Dim hWndP As Long
Dim wId As Long
    
   If (m_hWnd = 0) Then
      debugmsg m_sCtlName & ",Call To AddBandByHWnd before rebar created."
   End If
   
   If (m_hWnd <> 0) Then
      hBmp = BackgroundBitmapHandle()
      
      hWndP = GetParent(hwnd)
      If (hWndP <> 0) Then
         pAddWnds hwnd, hWndP
      End If
      wId = plAddId(vData, bFixedSize, sBandText)
      If (Not (pbRBAddBandByhWnd(m_hWnd, wId, hwnd, sBandText, hBmp, bBreakLine, bFixedSize, lBand))) Then
         debugmsg m_sCtlName & ",Failed to add Band"
         pRemoveID wId
      Else
         AddBandByHwnd = wId
         If Not (m_bSubClassing) Then
             ' Start subclassing:
             'Debug.Print "Start subclassing"
             pCreateSubClass
         End If
         RebarSize
      End If
   End If
End Function
Private Function pbRBAddBandByhWnd( _
        ByVal hWndRebar As Long, _
        ByVal wId As Long, _
        Optional ByVal hWndCHild As Long = 0, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal hBmp As Long = 0, _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByRef ltRBand As Long _
    ) As Boolean

If hWndRebar = 0 Then
    MsgBox "No hWndRebar!"
    Exit Function
End If

Dim sClassName As String
Dim hWndReal As Long
Dim tRBand As REBARBANDINFO
Dim tRBand471 As REBARBANDINFO_471
Dim tRBandNT As REBARBANDINFO_NOTEXT
Dim tRBandNT471 As REBARBANDINFO_NOTEXT_471
Dim bNoText As Boolean
Dim rct As RECT
Dim fMask As Long
Dim fStyle As Long
Dim dwStyle As Long
Dim bListStyle As Boolean

   hWndReal = hWndCHild
   
   If Not (hWndCHild = 0) Then
      'Check to see if it's a toolbar (so we can
      'make if flat)
      fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE
      sClassName = Space$(255)
      GetClassName hWndCHild, sClassName, 255
      'see if it's a real Windows toolbar
      If InStr(UCase$(sClassName), "TOOLBARWINDOW32") Then
         dwStyle = GetWindowLong(hWndCHild, GWL_STYLE)
         dwStyle = dwStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT
         SetWindowLong hWndCHild, GWL_STYLE, dwStyle
      End If
      'Could be a VB Toolbar -- make it flat anyway.
      If InStr(UCase$(sClassName), "TOOLBARWNDCLASS") Then
         hWndReal = GetWindow(hWndCHild, GW_CHILD)
         dwStyle = GetWindowLong(hWndReal, GWL_STYLE)
         dwStyle = dwStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT
         SetWindowLong hWndReal, GWL_STYLE, dwStyle
      End If
   End If
   
   GetWindowRect hWndReal, rct
   
   If hBmp <> 0 Then
       fMask = fMask Or RBBIM_BACKGROUND
   End If
   fMask = fMask Or RBBIM_STYLE Or RBBIM_ID Or RBBIM_COLORS Or RBBIM_SIZE
   If sBandText <> "" Then
      fMask = fMask Or RBBIM_TEXT
      tRBand.lpText = sBandText
      tRBand.cch = Len(sBandText)
   Else
      bNoText = True
   End If
   
   fStyle = RBBS_FIXEDBMP ' or RBBS_CHILDEDGE
   If bBreakLine = True Then
      fStyle = fStyle Or RBBS_BREAK
   End If
   If bFixedSize = True Then
      fStyle = fStyle Or RBBS_FIXEDSIZE
   Else
      fStyle = fStyle And Not RBBS_FIXEDSIZE
   End If
   
   If (bNoText) Then
      With tRBandNT
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndCHild = hWndReal
            If m_ePosition = erbPositionLeft Or m_ePosition = erbPositionRight Then
               .cxMinChild = rct.bottom - rct.top
               .cyMinChild = rct.right - rct.left
            Else
               .cxMinChild = rct.right - rct.left
               .cyMinChild = rct.bottom - rct.top
            End If
         End If
         'Set the rest OK
         .wId = wId
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBmp
         'The length of the type
         .cbSize = LenB(tRBandNT)
      End With
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         CopyMemory tRBandNT471, tRBandNT, LenB(tRBandNT)
         tRBandNT471.cbSize = LenB(tRBandNT471)
         tRBandNT471.fMask = tRBandNT471.fMask Or RBBIM_IDEALSIZE
         tRBandNT471.cxIdeal = tRBandNT471.cxMinChild
         tRBandNT471.fStyle = tRBandNT471.fStyle Or RBBS_CHEVRON
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBandNT471) <> 0)
      Else
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBandNT) <> 0)
      End If
   Else
      With tRBand
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndCHild = hWndReal
            If m_ePosition = erbPositionLeft Or m_ePosition = erbPositionRight Then
               .cxMinChild = rct.bottom - rct.top
               .cyMinChild = rct.right - rct.left
            Else
               .cxMinChild = rct.right - rct.left
               .cyMinChild = rct.bottom - rct.top
            End If
         End If
         'Set the rest OK
         .wId = wId
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBmp
         'The length of the type
         .cbSize = LenB(tRBand)
      End With
      If m_lMajor > 4 Or (m_lMajor = 4 And m_lMinor >= 71) Then
         CopyMemory tRBand471, tRBand, LenB(tRBandNT)
         tRBand471.cbSize = LenB(tRBand471)
         tRBand471.fStyle = tRBand471.fStyle Or RBBS_CHEVRON
         tRBand471.fMask = tRBand471.fMask Or RBBIM_IDEALSIZE
         tRBand471.cxIdeal = tRBand471.cxMinChild
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBand471) <> 0)
      Else
         pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBand) <> 0)
      End If
   End If
   
   ltRBand = BandCount

End Function

Private Sub pRemoveID( _
        ByVal wId As Long _
    )
Dim lItem As Long
Dim lTarget As Long
    
   For lItem = 1 To m_lIDCount
      If (m_tDataStore(lItem).wId = wId) Then
      Else
         lTarget = lTarget + 1
         If (lTarget <> lItem) Then
            LSet m_tDataStore(lTarget) = m_tDataStore(lItem)
         End If
      End If
   Next lItem
   If lTarget = 0 Then
      debugmsg m_sCtlName & ",Removed all IDs and data"
      m_lIDCount = 0
      Erase m_tDataStore
   Else
      If (lTarget <> m_lIDCount) Then
         debugmsg m_sCtlName & ",Reduced ID Count to : " & lTarget
         m_lIDCount = lTarget
         ReDim Preserve m_tDataStore(1 To m_lIDCount) As tRebarDataStore
      End If
   End If
    
End Sub
Public Property Get BandIndexForId( _
        ByVal wId As Long _
    ) As Long
Attribute BandIndexForId.VB_Description = "Returns the internal index of a band given the band's id."
Dim lItem As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lIndex As Long
Dim lR As Long

   If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
      lIndex = -1
      tRbbi.cbSize = Len(tRbbi)
      tRbbi.fMask = RBBIM_ID
      For lItem = 0 To BandCount - 1
          lR = SendMessage(m_hWnd, RB_GETBANDINFO, lItem, tRbbi)
          If (lR <> 0) Then
              If (wId = tRbbi.wId) Then
                  lIndex = lItem
                  Exit For
              End If
          End If
      Next lItem
      BandIndexForId = lIndex
   Else
      BandIndexForId = SendMessageLong(m_hWnd, RB_IDTOINDEX, wId, 0)
   End If
End Property
Public Property Get BandIDForIndex( _
      ByVal lIndex As Long _
   ) As Long
Attribute BandIDForIndex.VB_Description = "Gets the ID of band given its 0-based index in the rebar."
Dim lR As Long
Dim tRbbi As REBARBANDINFO_NOTEXT

   tRbbi.cbSize = Len(tRbbi)
   tRbbi.fMask = RBBIM_ID
   lR = SendMessage(m_hWnd, RB_GETBANDINFO, lIndex, tRbbi)
   BandIDForIndex = tRbbi.wId
   
End Property
Public Property Get BandData( _
      ByVal wId As Long _
   ) As Variant
Attribute BandData.VB_Description = "Gets/sets a variant value associated with a band in the rebar."
Dim lItem As Long
   For lItem = 1 To m_lIDCount
      If m_tDataStore(lItem).wId = wId Then
         BandData = m_tDataStore(lItem).vData
         Exit For
      End If
   Next lItem
End Property

Public Property Get BandIndexForData( _
        ByVal vData As Variant _
    ) As Long
Attribute BandIndexForData.VB_Description = "Returns the index of a band given the band's key."
Dim lItem As Long
Dim lAt As Long
Dim vitem As Variant
On Error Resume Next
    lAt = -1
    For lItem = 1 To m_lIDCount
      If IsMissing(m_tDataStore(lItem).vData) Then
         vitem = ""
      ElseIf IsObject(m_tDataStore(lItem).vData) Then
         If (vData Is m_tDataStore(lItem).vData) Then
            lAt = lItem
            Exit For
         End If
      Else
         If vData = m_tDataStore(lItem).vData Then
            lAt = lItem
            Exit For
         End If
      End If
      
    Next lItem
    If (lAt > 0) Then
        lAt = BandIndexForId(m_tDataStore(lAt).wId)
    End If
    BandIndexForData = lAt
End Property
Private Function plAddId( _
        ByVal vData As Variant, _
        ByVal bFixedSize As Boolean, _
        ByVal sBandText As String _
    ) As Long
   m_lIDCount = m_lIDCount + 1
   ReDim Preserve m_tDataStore(1 To m_lIDCount) As tRebarDataStore
   With m_tDataStore(m_lIDCount)
      .wId = m_lIDCount
      .vData = vData
      .bFixedSize = bFixedSize
      .sBandText = sBandText
   End With
   plAddId = m_lIDCount
End Function
Private Sub pAddWnds( _
        ByVal hwndItem As Long, _
        ByVal hWndParent As Long _
    )
   m_iWndItemCount = m_iWndItemCount + 1
   ReDim Preserve m_tWndStore(1 To m_iWndItemCount) As tRebarWndStore
   With m_tWndStore(m_iWndItemCount)
      .hwndItem = hwndItem
      .hWndItemParent = hWndParent
      GetWindowRect hwndItem, .tR
   End With
End Sub
Private Sub pResetParent( _
        ByVal hwndItem As Long _
    )
Dim iItem As Long
Dim iTarget As Long
Dim bSuccess As Boolean
    
   For iItem = 1 To m_iWndItemCount
      If (m_tWndStore(iItem).hwndItem = hwndItem) Then
         ' Set the parent back to the original:
         SetParent m_tWndStore(iItem).hwndItem, m_tWndStore(iItem).hWndItemParent
         ' send a message to destroy the object:
         If m_bKillChildren Then
            ShowWindow m_tWndStore(iItem).hwndItem, SW_HIDE
            SendMessageLong m_tWndStore(iItem).hwndItem, WM_DESTROY, 0, 0
         End If
         ' Reset the size to original:
         SetWindowPos m_tWndStore(iItem).hwndItem, 0, m_tWndStore(iItem).tR.left, m_tWndStore(iItem).tR.top, m_tWndStore(iItem).tR.right - m_tWndStore(iItem).tR.left, m_tWndStore(iItem).tR.bottom - m_tWndStore(iItem).tR.top, SWP_NOREDRAW Or SWP_NOZORDER Or SWP_NOOWNERZORDER
         'MoveWindow m_tWndStore(iItem).hWndItem, m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.Top, m_tWndStore(iItem).tR.Right - m_tWndStore(iItem).tR.Left, m_tWndStore(iItem).tR.Bottom - m_tWndStore(iItem).tR.Top, 1
         bSuccess = True
      Else
         iTarget = iTarget + 1
         If iTarget <> iItem Then
            LSet m_tWndStore(iTarget) = m_tWndStore(iItem)
         End If
      End If
   Next iItem
   
   If (iTarget = 0) Then
      debugmsg m_sCtlName & ",Successfully reset all parents"
      m_iWndItemCount = 0
      Erase m_tWndStore
   Else
      If iTarget <> m_iWndItemCount Then
         debugmsg m_sCtlName & ",Decrease wnd count to " & iTarget
         m_iWndItemCount = iTarget
         ReDim Preserve m_tWndStore(1 To m_iWndItemCount) As tRebarWndStore
      End If
   End If
   
   
   If Not bSuccess Then
      debugmsg m_sCtlName & ",Failed to reset parent.."
      ' At least ensure it won't stop the rebar terminating:
      ShowWindow hwndItem, SW_HIDE
      SetParent hwndItem, 0
   End If
End Sub
Public Sub RebarSize()
Attribute RebarSize.VB_Description = "Sizes the rebar to the parent object."
Dim lLeft As Long, lTop As Long
Dim cx As Long, cy As Long
Dim rc As RECT, rcb As RECT, rcI As RECT, rcP As RECT
   
   If (m_hWnd <> 0) Then
      GetWindowRect m_hWnd, rcb
      OffsetRect rcb, -rcb.left, -rcb.top
      GetClientRect m_hWndCtlParent, rcP
      If (m_ePosition = erbPositionBottom) Or (m_ePosition = erbPositionTop) Then
         cx = rcP.right - rcP.left
         cy = RebarHeight
         If m_ePosition = erbPositionBottom Then
            lTop = rcP.bottom - rc.top - cy
         End If
         AdjustForOtherRebars m_hWnd, lLeft, lTop, cx, cy
         SetWindowPos m_hWnd, 0, lLeft, lTop, cx, cy, SWP_NOZORDER Or SWP_NOACTIVATE
      Else
         cy = rcP.bottom - rcP.top
         cx = RebarHeight
         If m_ePosition = erbPositionRight Then
            lLeft = rcP.right - rcP.left - cx
         End If
         AdjustForOtherRebars m_hWnd, lLeft, lTop, cx, cy
         SetWindowPos m_hWnd, 0, lLeft, lTop, cx, cy, SWP_NOZORDER Or SWP_NOACTIVATE
      End If
      GetWindowRect m_hWnd, rc
      OffsetRect rc, -rc.left, -rc.top
      UnionRect rcI, rc, rcb
      InvalidateRect m_hWnd, rcI, True
      UpdateWindow m_hWnd
   End If
   
End Sub
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the window handle of the control.  Use RebarhWnd to get the handle of the Rebar itself."
    hwnd = UserControl.hwnd
End Property
Public Property Get RebarHwnd() As Long
Attribute RebarHwnd.VB_Description = "Returns the windows handle of the Rebar window."
    RebarHwnd = m_hWnd
End Property
Public Property Get RebarHeight() As Long
Attribute RebarHeight.VB_Description = "Gets the current height of the rebar."
Dim tc As RECT
    'If (m_hWnd <> 0) Then
    '  GetWindowRect m_hWnd, tc
    '  RebarHeight = (tc.Bottom - tc.Top)
    'End If
    ' Get the height that would be good for the rebar:
   If m_bVisible Then
      RebarHeight = SendMessageLong(m_hWnd, RB_GETBARHEIGHT, 0, 0) + 4
   Else
      RebarHeight = 0
   End If
End Property
Public Property Get RebarWidth() As Long
Dim tc As RECT
   If (m_hWnd <> 0) Then
      If m_bVisible Then
         GetWindowRect m_hWnd, tc
         RebarWidth = (tc.right - tc.left)
      Else
         RebarWidth = 0
      End If
   End If
End Property
Private Function pbLoadCommCtls() As Boolean
Dim ctEx As CommonControlsEx

    ctEx.dwSize = Len(ctEx)
    ctEx.dwICC = ICC_COOL_CLASSES Or _
        ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
    
    pbLoadCommCtls = (InitCommonControlsEx(ctEx) <> 0)

End Function

Public Function CreateRebar(ByVal hWndParent As Long) As Boolean
Attribute CreateRebar.VB_Description = "Initialises a rebar for use and allows you to specify the host window for the rebar.  For a standard form, this should be the form.  For an MDI form, this should be a PictureBox aligned to the top of the MDI form."
   If (UserControl.Ambient.UserMode) Then
      DestroyRebar
      ' Set up the rebar:
      If (pbCreateRebar(hWndParent)) Then
         SetProp m_hWnd, "vbal:cRebarPosition", m_ePosition
         m_hWndCtlParent = hWndParent
         AddRebar m_hWnd, m_hWndCtlParent
      End If
   End If
End Function
Public Function AddResizeObject(ByVal hWndParent As Long, ByVal hwnd As Long, ByVal ePosition As ERBPositionConstants)
Attribute AddResizeObject.VB_Description = "Adds a control to the list of objects to be considered when resizing a rebar on screen.  Other rebars are automatically taken into account."
   AddRebar hwnd, hWndParent
   SetProp hwnd, "vbal:cRebarPosition", ePosition
End Function
Private Function pbCreateRebar(ByVal hWndParent As Long) As Boolean
Dim lWidth As Long
Dim lHeight As Long
Dim bVertical As Boolean
Dim hwndCoolBar As Long
Dim lResult As Long
Dim dwStyle As Long
Dim dwExStyle As Long
Dim lExStyle As Long
Dim rc As RECT

    If (UserControl.Ambient.UserMode) Then
    
      ' Try to load the Common Controls support for the
      ' rebar control:
      If (pbLoadCommCtls()) Then
         'Debug.Print "Loaded Coolbar support"
         ' If we have done this, then build a rebar:
         GetWindowRect hWndParent, rc
         lWidth = rc.right - rc.left
         lHeight = rc.bottom - rc.top

         ComCtlVersion m_lMajor, m_lMinor
         dwStyle = WS_CHILD Or WS_BORDER Or _
             WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
             WS_VISIBLE
         Select Case m_ePosition
         Case erbPositionTop
            dwStyle = dwStyle Or CCS_TOP
         Case erbPositionRight
            dwStyle = dwStyle Or CCS_RIGHT Or CCS_VERT
         Case erbPositionLeft
            dwStyle = dwStyle Or CCS_LEFT Or CCS_VERT
         Case erbPositionBottom
            dwStyle = dwStyle Or CCS_BOTTOM
         End Select
         dwStyle = dwStyle Or CCS_NORESIZE
         dwStyle = dwStyle Or CCS_NODIVIDER
         
         dwStyle = dwStyle Or RBS_DBLCLKTOGGLE
         dwStyle = dwStyle Or RBS_VARHEIGHT Or RBS_BANDBORDERS
         dwStyle = dwStyle Or RBS_AUTOSIZE
   
         dwExStyle = WS_EX_TOOLWINDOW
         lExStyle = GetWindowLong(hWndParent, GWL_EXSTYLE)
         lExStyle = lExStyle And (WS_EX_RIGHT Or WS_EX_RTLREADING)
         dwExStyle = dwExStyle Or lExStyle
   
         m_hWnd = CreateWindowEX(dwExStyle, _
                              REBARCLASSNAME, "", _
                              dwStyle, 0, 0, lWidth, lHeight, _
                              hWndParent, ICC_COOL_CLASSES, App.hInstance, ByVal 0&)
         If (m_hWnd <> 0) Then
            ' Debug.Print "Created Rebar Window"
            AddToToolTip m_hWnd
            If m_lMajor >= 5 Then
               SendMessageLong m_hWnd, RB_SETEXTENDEDSTYLE, 0, RBS_EX_OFFICE9
            End If
            pbCreateRebar = True
         End If
      End If
    End If
    
End Function
Public Sub DestroyRebar()
Attribute DestroyRebar.VB_Description = "Removes all bands from a rebar and clears all resources associated with it."
   pDestroyRebar True
End Sub
Public Sub DestroyRebarDontDestroyChildren()
Attribute DestroyRebarDontDestroyChildren.VB_Description = "Removes all bands from a rebar and clears all resources associated with it without posting a destroy window message to any children."
   pDestroyRebar False
End Sub
Private Sub pDestroyRebar(ByVal bKillChildren As Boolean)
   If (m_hWnd <> 0) Then
      m_bKillChildren = bKillChildren
      
      debugmsg m_sCtlName & ",pDestroyRebar"
      RemoveRebar m_hWnd
      
      RemoveFromToolTip m_hWnd
      RemoveAllRebarBands
      
      DeleteObject m_hBmp
      m_hBmp = 0
      
      pDestroySubClass
      RemoveProp m_hWnd, "vbal:cRebarPosition"
      
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      m_hWnd = 0
      m_hWndCtlParent = 0
      
      m_bKillChildren = True
   End If
End Sub

Public Sub RemoveAllRebarBands()
Attribute RemoveAllRebarBands.VB_Description = "Removes all bands from the rebar.  To prevent controls not terminating when a form unloads because they are contained by a different parent, call this method."
Dim lBands As Long
Dim lBand As Long
    If (m_hWnd <> 0) Then
        lBands = BandCount
        For lBand = 0 To lBands - 1
            RemoveBand 0
        Next lBand
        pDestroySubClass
    End If
End Sub
Public Sub RemoveBand( _
        ByVal lBand As Long _
    )
Attribute RemoveBand.VB_Description = "Removes a specified band from the rebar control."
Dim lhWnd As Long
Dim wId As Long

    If (m_hWnd <> 0) Then
        ' If a valid band:
        If (lBand >= 0) And (lBand < BandCount) Then
            If m_lMajor < 4 Or (m_lMajor = 4 And m_lMinor < 71) Then
               ' Remove the child from this band:
               lhWnd = plGetHwndOfBandChild(m_hWnd, lBand, wId)
               If (lhWnd <> 0) Then
                   pResetParent lhWnd
                   
               End If
               ' Remove the band:
               SendMessageLong m_hWnd, RB_DELETEBAND, lBand, 0&
               ' Remove the id for this band:
               pRemoveID wId
               ' No bands left? Stop subclassing:
               If (BandCount = 0) Then
                  debugmsg m_sCtlName & ",All bands destroyed"
                  pDestroySubClass
               End If
            Else
               SendMessageLong m_hWnd, RB_DELETEBAND, lBand, 0&
               If BandCount = 0 Then
                  debugmsg m_sCtlName & ",All bands destroyed"
                  pDestroySubClass
               End If
            End If
        End If
    End If
End Sub
Private Function plGetHwndOfBandChild( _
        ByVal lhWnd As Long, _
        ByVal lBand As Long, _
        ByRef wId As Long _
    ) As Long
Dim lParam As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lR As Long

    tRbbi.cbSize = Len(tRbbi)
    tRbbi.fMask = RBBIM_CHILD Or RBBIM_ID
    lR = SendMessage(lhWnd, RB_GETBANDINFO, lBand, tRbbi)
    If (lR <> 0) Then
        plGetHwndOfBandChild = tRbbi.hWndCHild
        wId = tRbbi.wId
    End If
End Function

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Gets/sets whether the entire rebar will be visible or not."
   Visible = m_bVisible
End Property
Public Property Let Visible(ByVal bState As Boolean)
   m_bVisible = bState
   If m_hWnd <> 0 Then
      If Not bState Then
         ShowWindow m_hWnd, SW_HIDE
         RaiseEvent HeightChanged(0)
      Else
         ShowWindow m_hWnd, SW_SHOW
         RaiseEvent HeightChanged(RebarHeight)
      End If
   End If
   PropertyChanged "Visible"
End Property

Private Sub ClearPicture()
   If (m_hBmp <> 0) Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   m_sPicture = ""
   m_lResourceID = 0
   Set m_pic = Nothing
End Sub

Private Sub UserControl_Initialize()
    debugmsg "cRebar:Initialise"
    m_lMajor = 4
    m_lMinor = 0
    m_bVisible = True
    m_bKillChildren = True
End Sub

Private Sub UserControl_InitProperties()
   ' If init properties we must be in design mode.
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read in properties here:
    ' ...
    On Error Resume Next
    m_sCtlName = UserControl.Extender.Name
    Err.Clear
    On Error GoTo 0
    
End Sub

Private Sub UserControl_Resize()
   If (UserControl.Ambient.UserMode) Then
      UserControl.width = 0
      UserControl.height = 0
   End If
End Sub

Private Sub UserControl_Terminate()
    m_bInTerminate = True
    DestroyRebar
    ClearPicture
    debugmsg m_sCtlName & ",cRebar:Terminate"
    'MsgBox "cRebar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Write properties here:
    ' ...
    
End Sub


