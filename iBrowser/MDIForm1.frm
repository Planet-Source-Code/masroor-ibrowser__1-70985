VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "iBrowser"
   ClientHeight    =   8925
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10590
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":441E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5072
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":569F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":62EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6921
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6FBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   7320
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   1050
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdAddTab 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14760
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   1085
         TabWidthStyle   =   2
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   706
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Browser 1"
               Object.Tag             =   "0"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Description     =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Description     =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Description     =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Reload the page"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Description     =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Description     =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            Key             =   "Favorites"
            Description     =   "Favorites"
            Object.ToolTipText     =   "Shows your favorites"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mail"
            Key             =   "Mail"
            Description     =   "Mail"
            Object.ToolTipText     =   "Open's Outlook Express"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Media"
            Key             =   "Media"
            Description     =   "Media"
            Object.ToolTipText     =   "Media"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print's the file"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewTab 
         Caption         =   "New Tab"
         Shortcut        =   ^T
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuInternetOptions 
         Caption         =   "Internet Options"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBO 
         Caption         =   "Browser Options"
      End
      Begin VB.Menu mnuWindowsUpdate 
         Caption         =   "Windows Update"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupManager 
         Caption         =   "Popup Manager"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpWin(1 To 100) As Form1
Dim BrowserID As Integer


Private Sub cmdAddTab_Click()
mnuNewTab_Click
End Sub

Private Sub cmdbollywood_Click()
MDIForm1.ActiveForm.w.Navigate "www.apunkabollywood.net/browser"
End Sub

Private Sub cmdFunMaza_Click()
MDIForm1.ActiveForm.w.Navigate "www.funmaza.com"
End Sub

Private Sub cmdNumair4u_Click()
MDIForm1.ActiveForm.w.Navigate "www.numair4u.tk"
End Sub

Private Sub cmdRajaAli_Click()
MDIForm1.ActiveForm.w.Navigate "www.rajaali.tk"
End Sub

Private Sub cmdSongsPK_Click()
MDIForm1.ActiveForm.w.Navigate "www.songs.pk"
End Sub

Private Sub MDIForm_Load()
Call InitializeWinInfo
Set ExpWin(1) = New Form1
WinInfo(1).winid = 1
WinInfo(1).tabID = 1
ExpWin(1).Tag = 1
BrowserID = 1
TabStrip1.Tabs(1).Tag = 1
ExpWin(1).w.Navigate frmOptions.txtHome.Text
ExpWin(1).w.Navigate frmOptions.txtHome.Text
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuBO_Click()
frmOptions.Show
End Sub

Private Sub mnuCopy_Click()
Me.ActiveForm.w.SetFocus
    On Error Resume Next
    Me.ActiveForm.w.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuCut_Click()
Me.ActiveForm.w.SetFocus
    On Error Resume Next
    Me.ActiveForm.w.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFB_Click()
If Toolbar3.Visible = True Then
Toolbar3.Visible = False
mnuFB.Checked = False
Else
Toolbar3.Visible = True
mnuFB.Checked = True
End If
End Sub

Private Sub mnuFind_Click()
Form1.w.SetFocus
 SendKeys "^f"
End Sub

Private Sub mnuInternetOptions_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub mnuNewTab_Click()
TabStrip1.Tabs.Add
TabStrip1.Tabs(TabStrip1.Tabs.Count).Caption = "<untitled>"
Dim winid As Integer
winid = FreeID()
TabStrip1.Tabs(TabStrip1.Tabs.Count).Tag = winid
On Error Resume Next
  Set ExpWin(winid) = New Form1
  With WinInfo(winid)
    .winid = winid
    .tabID = winid
  End With
ExpWin(winid).Tag = winid
  ExpWin(winid).Show
  TabStrip1.Tabs.Item(winid).Selected = True
    ExpWin(winid).w.Navigate App.Path & "\tab.htm"
    ExpWin(winid).w.Navigate App.Path & "\tab.htm"
End Sub

Private Sub mnuOpen_Click()
Me.ActiveForm.w.SetFocus
    On Error Resume Next
    CDialog1.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    CDialog1.ShowOpen
    If CDialog1.FileName = "" Then
        Exit Sub
        Else
    Me.ActiveForm.w.Navigate (CDialog1.FileName)
    End If
End Sub

Private Sub mnuPageSetup_Click()
 Me.ActiveForm.w.SetFocus
    On Error Resume Next
   Me.ActiveForm.w.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuPaste_Click()
Me.ActiveForm.w.SetFocus
On Error Resume Next
Me.ActiveForm.w.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuPopupManager_Click()
frmPopupManager.Show
End Sub

Private Sub mnuPrint_Click()
 Me.ActiveForm.w.SetFocus
    On Error Resume Next
    Me.ActiveForm.w.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuPrintPreview_Click()
    Me.ActiveForm.w.SetFocus
    On Error Resume Next
    Me.ActiveForm.w.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuProperties_Click()
Me.ActiveForm.w.SetFocus
    On Error Resume Next
    Me.ActiveForm.w.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuSelectAll_Click()
Me.ActiveForm.w.SetFocus
    On Error Resume Next
 Me.ActiveForm.w.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuToolbar_Click()
If Toolbar1.Visible = True Then
Toolbar1.Visible = False
mnuToolbar.Checked = False
Else
Toolbar1.Visible = True
mnuToolbar.Checked = True
End If
End Sub

Private Sub mnuWindowsUpdate_Click()
MDIForm1.ActiveForm.w.Navigate "http://windowsupdate.microsoft.com/"
End Sub



Private Sub TabStrip1_Click()
ExpWin(TabStrip1.SelectedItem.Tag).ZOrder (0)
BrowserID = TabStrip1.SelectedItem.Tag
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim starting
On Error Resume Next

Select Case Button.Key

Case "Back"
MDIForm1.ActiveForm.w.GoBack
Case "Forward"
MDIForm1.ActiveForm.w.GoForward
Case "Stop"
MDIForm1.ActiveForm.w.Stop
Case "Refresh"
MDIForm1.ActiveForm.w.Refresh
Case "Home"
MDIForm1.ActiveForm.w.Navigate frmOptions.txtHome.Text
Case "Search"
frmSearch.Show
Case "Favorites"
frmFavorites.Show
Case "Mail"
Shell "C:\Program Files\Outlook Express\MSIMN.EXE"
Case "Media"
frmMedia.Show
Case "Print"
 Me.ActiveForm.w.SetFocus
On Error Resume Next
Me.ActiveForm.w.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Select
End Sub


