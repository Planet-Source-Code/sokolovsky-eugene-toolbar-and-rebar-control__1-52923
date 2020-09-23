VERSION 5.00
Object = "{1680BB0E-7877-4B44-9176-F05BDD9F114A}#1.0#0"; "vbalTbar6.ocx"
Begin VB.MDIForm mfrmTest 
   BackColor       =   &H8000000C&
   Caption         =   "vbAccelerator Toolbar/Rebar MDI Demonstration"
   ClientHeight    =   5655
   ClientLeft      =   1110
   ClientTop       =   1425
   ClientWidth     =   8235
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8235
      TabIndex        =   3
      Top             =   5340
      Width           =   8235
   End
   Begin vbalTBar6.cReBar cReBar1 
      Left            =   480
      Top             =   1800
      _ExtentX        =   7223
      _ExtentY        =   979
   End
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin vbalTBar6.cToolbarHost tbhMenu 
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         BorderStyle     =   0
      End
      Begin vbalTBar6.cToolbar tbrMenu 
         Height          =   1125
         Left            =   4020
         Top             =   300
         Width           =   3000
         _ExtentX        =   1508
         _ExtentY        =   661
      End
      Begin vbalTBar6.cToolbar cToolbar1 
         Height          =   1125
         Left            =   180
         Top             =   300
         Width           =   3000
         _ExtentX        =   6694
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboTest 
         Height          =   315
         Left            =   6060
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Menu mnuDrop 
      Caption         =   "mnuDrop"
      Visible         =   0   'False
      Begin VB.Menu mnuTest1 
         Caption         =   "Test1"
      End
      Begin VB.Menu mnuTest2 
         Caption         =   "Test2"
      End
   End
End
Attribute VB_Name = "mfrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Const WM_MDINEXT = &H224
Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long

Private Sub pBackgroundBitmap(ByVal bState As Boolean)
   ' To change the background bitmap, we remove all bands
   ' and add them in again.
   ' In order to prevent flickering whilst the rebar builds,
   ' use LockWindowUpdate.  See Tips on vbAccelerator for
   ' more info.
   LockWindowUpdate Me.hwnd
   With cReBar1
      .ImageSource = CRBLoadFromFile
      If (bState) Then
         .DestroyRebarDontDestroyChildren
         .ImageFile = App.Path & "\iebar2.bmp"
         .CreateRebar picHolder.hwnd
         ' Add the bands:
         .AddBandByHwnd tbhMenu.hwnd, , , , "MenuBar"
         .BandChildMinWidth(0) = 64
         .AddBandByHwnd cToolbar1.hwnd, , , , "Toolbar1"
         .BandChildMinWidth(1) = 64
         .AddBandByHwnd cboTest.hwnd, "Style", True, , "Stylebar"
         tbhMenu.BackgroundBitmap = LoadPicture(App.Path & "\iebar2.bmp")
      Else
         .DestroyRebarDontDestroyChildren
         .ImageFile = ""
         .CreateRebar picHolder.hwnd
         ' Add the bands:
         .AddBandByHwnd tbhMenu.hwnd, , , , "MenuBar"
         .BandChildMinWidth(0) = 64
         .AddBandByHwnd cToolbar1.hwnd, , , , "Toolbar1"
         .BandChildMinWidth(1) = 64
         .AddBandByHwnd cboTest.hwnd, "Style", True, , "Stylebar"
         tbhMenu.ClearPicture
      End If
   End With
   LockWindowUpdate 0

End Sub
Private Sub pFileMenu(ByVal lIndex As Long, ByVal sKey As String)
Dim lItemIndex As Long
   lItemIndex = CLng(Mid$(sKey, 9, 1))
   Select Case lItemIndex
   Case 0
      PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
   End Select
End Sub
Private Sub cReBar1_ChevronPushed(ByVal wID As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Dim v As Variant
   v = cReBar1.BandData(wID)
   If Not IsMissing(v) Then
      Debug.Print lRight, lTop
      Select Case v
      Case "MenuBar"
         tbrMenu.ChevronPress lRight \ Screen.TwipsPerPixelX + 1, lTop \ Screen.TwipsPerPixelY
      Case "Toolbar1"
         cToolbar1.ChevronPress lRight \ Screen.TwipsPerPixelX + 1, lTop \ Screen.TwipsPerPixelY
      End Select
   End If
End Sub
Private Sub cReBar1_HeightChanged(lNewHeight As Long)
   If picHolder.Align = 1 Or picHolder.Align = 2 Then
      picHolder.Height = lNewHeight * Screen.TwipsPerPixelY
   Else
      picHolder.Width = lNewHeight * Screen.TwipsPerPixelY
   End If
End Sub
Private Sub cToolbar1_DropDownPress(ByVal lButton As Long)
Dim x As Long, y As Long
   cToolbar1.GetDropDownPosition lButton, x, y
   'y = y - picHolder.Height - 2 * Screen.TwipsPerPixelY
   Me.PopupMenu mnuDrop, , x, y
End Sub
Private Sub MDIForm_Load()
   ' NB: If you place a picture box control on your MDI,
   ' the Rebar will be hosted in this.  The effect is much smoother that
   ' if the rebar is hosted directly in the MDI.
   With cToolbar1
      .ImageSource = CTBLoadFromFile
      .ImageFile = App.Path & "\small256.bmp"
      .CreateToolbar 16, , , True
      .AddButton "New", 0, , , "New", CTBDropDown
      .AddButton "Open", 1, , , "Open", CTBNormal
      .AddButton "Save", 2, , , "Save"
      .AddButton "", -1, , , , CTBSeparator
      .AddButton "Cut", 3, , , "Cut"
      .AddButton "Copy", 4, , , "Copy"
      .AddButton "Paste", 5, , , "Paste"
      .AddButton "", -1, , , , CTBSeparator
      .AddButton "CheckBox", 6, , , "Check", CTBCheck
      .AddButton "", -1, , , , CTBSeparator
      .AddButton "Print", 7, , , "Print", CTBDropDown
      .AddButton "", -1, , , , CTBSeparator
      .AddButton "Help", 8, , , "Help", CTBCheckGroup
      .AddButton "Whats This", 9, , , "Desktop", CTBCheckGroup
   End With
   Dim i As Long
   For i = 1 To cToolbar1.ButtonCount
     cboTest.AddItem cToolbar1.ButtonCaption(i - 1)
   Next i
   cboTest.ListIndex = 0
   With cReBar1
      ' Create the rebar:
      .ImageSource = CRBLoadFromFile
      .CreateRebar picHolder.hwnd
      ' Add the bands:
      .AddBandByHwnd tbhMenu.hwnd, , , , "MenuBar"
      .BandChildMinWidth(0) = 64
      .AddBandByHwnd cToolbar1.hwnd, , , , "Toolbar1"
      .BandChildMinWidth(1) = 64
      .AddBandByHwnd cboTest.hwnd, "Style", True, , "Stylebar"
   End With
   ' NB: Don't show child forms until you have initialised the rebar!
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cReBar1.RemoveAllRebarBands
End Sub
Private Sub MDIForm_Resize()
   '
End Sub
Private Sub picHolder_Resize()
   cReBar1.RebarSize
   If picHolder.Align = 1 Or picHolder.Align = 2 Then
      picHolder.Height = cReBar1.RebarHeight * Screen.TwipsPerPixelY
   Else
      picHolder.Width = cReBar1.RebarHeight * Screen.TwipsPerPixelY
   End If
End Sub

Private Sub WindowMenuClick(ByVal lItemData As Long)
Dim f As Object
Dim lhWnd As Long
Dim lhWndMDIClient
   Select Case lItemData
   Case Is >= 0
      Forms(lItemData).SetFocus
   Case -8001
      Me.Arrange vbTileHorizontal
   Case -8002
      Me.Arrange vbCascade
   Case -8003
      Me.ActiveForm.WindowState = vbMaximized
   Case -8004
      For Each f In Forms
         If Not f Is Me Then
            If f.Visible And f.MDIChild Then
               f.WindowState = vbMinimized
            End If
         End If
      Next f
   Case -8005
      lhWndMDIClient = FindWindowEx(Me.hwnd, 0, "MDIClient", ByVal 0&)
      lhWnd = Me.ActiveForm.hwnd
      PostMessage lhWndMDIClient, WM_MDINEXT, lhWnd, 0
   Case -8006
      lhWndMDIClient = FindWindowEx(Me.hwnd, 0, "MDIClient", ByVal 0&)
      lhWnd = Me.ActiveForm.hwnd
      PostMessage lhWndMDIClient, WM_MDINEXT, lhWnd, 1
   End Select
End Sub


Private Sub tbrMenu_ButtonClick(ByVal lButton As Long)
   Select Case tbrMenu.ButtonKey(lButton)
   Case "mnuWindowTOP"
   End Select
End Sub
Private Sub tbrMenu_DropDownPress(ByVal lButton As Long)
    PopupMenu mnuDrop
End Sub
