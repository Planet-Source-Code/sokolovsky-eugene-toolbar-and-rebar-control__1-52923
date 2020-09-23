VERSION 5.00
Object = "{1680BB0E-7877-4B44-9176-F05BDD9F114A}#1.0#0"; "vbalTbar6.ocx"
Begin VB.Form frmTEST 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
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
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1410
      ScaleWidth      =   10140
      TabIndex        =   1
      Top             =   0
      Width           =   10140
      Begin VB.ComboBox cboTest 
         Height          =   315
         Left            =   6060
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   300
         Width           =   1995
      End
      Begin vbalTBar6.cToolbar cToolbar1 
         Height          =   540
         Left            =   90
         Top             =   135
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   953
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10140
      TabIndex        =   0
      Top             =   6750
      Width           =   10140
   End
   Begin vbalTBar6.cReBar cReBar1 
      Left            =   480
      Top             =   1800
      _ExtentX        =   7223
      _ExtentY        =   979
   End
   Begin VB.Menu mnuDrop 
      Caption         =   "mnuDrop"
      Visible         =   0   'False
      Begin VB.Menu mnuDrop1 
         Caption         =   "mnuDrop1"
      End
      Begin VB.Menu mnuDrop2 
         Caption         =   "mnuDrop2"
      End
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cToolbar1_ButtonClick(ByVal lButton As Long)
    MsgBox ("Button Pressed " & lButton)
End Sub

Private Sub cToolbar1_DropDownPress(ByVal lButton As Long)
Dim x As Long, y As Long
    cToolbar1.GetDropDownPosition lButton, x, y
    PopupMenu mnuDrop, , x, y
End Sub

Private Sub Form_Load()
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
      .AddBandByHwnd cToolbar1.hwnd, , , , "Toolbar1"
      .BandChildMinWidth(1) = 64
      .AddBandByHwnd cboTest.hwnd, "Style", True, , "Stylebar"
   End With
End Sub

Private Sub mnuDrop1_Click()
    MsgBox ("Drop1 Clicked!")
End Sub
Private Sub mnuDrop2_Click()
    MsgBox ("Drop2 Clicked!")
End Sub
