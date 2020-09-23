VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "New Window"
   ClientHeight    =   6720
   ClientLeft      =   510
   ClientTop       =   825
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   315
      Left            =   10080
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
   Begin MSComctlLib.ProgressBar p 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6465
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "7/9/2008"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:53 PM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   3
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   10186
      ViewMode        =   6
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusText As String
Private Type PopUpData
    URL As String * 256
    Action As String * 9
End Type



Private Sub cmdGo_Click()
If Combo1.Text = "" Then
        'do nothing
    Else
        
              w.Navigate (Combo1.Text)
    End If
    strURL = Combo1.Text
    If Left(LCase(strURL), 7) = "http://" Or Left(LCase(strURL), 6) = "ftp://" Then
        Combo1.Text = strURL
    Else
        If Left(strURL, 7) <> "http://" Then
            Combo1.Text = "http://" & strURL
        Else
            If Left(strURL, 6) = "ftp://" Then
                Combo1.Text = "ftp://" & strURL
            End If
        End If
    End If
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cmdGo_Click
    End If
End Sub


Private Sub Form_Activate()
    MDIForm1.Caption = "iBrowser"
    'MDIForm1.Combo1.Text = MDIForm1.ActiveForm.w.LocationURL
    winid = CInt(MDIForm1.ActiveForm.Tag)
    tabcount = MDIForm1.TabStrip1.Tabs.Count
    For i = 1 To tabcount
        If MDIForm1.TabStrip1.Tabs(i).Tag = CStr(winid) Then
            MDIForm1.TabStrip1.Tabs(i).Caption = MDIForm1.ActiveForm.w.LocationName
            Exit For
        End If
    Next
End Sub

Private Sub Form_Click()
'MDIForm1.Combo1.Text = MDIForm1.ActiveForm.w.LocationURL
End Sub

Private Sub Form_GotFocus()
'MDIForm1.Caption = Me.Caption & " - CyberBrowser"
'MDIForm1.cboURL.Text = MDIForm1.ActiveForm.w.LocationURL
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Height = 6300
Form1.Width = 11700
Me.Caption = w.LocationName
End Sub




Private Sub Form_Paint()
'Combo1.Width = Form1.Width - Label1.Width - cmdGo.Width - 300

End Sub

Private Sub Form_Resize()
w.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - 800
Combo1.Width = w.Width - Label1.Width - 600
cmdGo.Left = Combo1.Left + Combo1.Width + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim winid As Integer
'Set a = MDIForm1.ActiveForm
winid = CStr(MDIForm1.ActiveForm.Tag)
'MsgBox winid
WinInfo(winid).tabID = 101
WinInfo(winid).winid = 101
'winid = CStr(MDIForm1.ActiveForm.Tag)
tabcount = MDIForm1.TabStrip1.Tabs.Count
'Dim w1 As Integer
'w1 = 1
For w1 = 1 To tabcount Step 1
    If MDIForm1.TabStrip1.Tabs.Item(w1).Tag = CStr(winid) Then Exit For
Next
If w1 = tabcount + 1 Then w1 = tabcount
If w1 > 0 Then MDIForm1.TabStrip1.Tabs.Remove (w1)
Unload Me
End Sub

Private Sub w_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Combo1.AddItem URL
If Me.Caption <> "" Then
    'MDIForm1.TabStrip1.Tabs(CInt(Form.Tag)).Caption = Me.Caption
Else
    'MDIForm1.TabStrip1.Tabs(CInt(Form.Tag)).Caption = "<untitled>"
End If
End Sub

Private Sub w_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Combo1.Text = w.LocationURL
'Me.Caption = w.LocationName
End Sub





Private Sub w_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    winid = CInt(MDIForm1.ActiveForm.Tag)
    tabcount = MDIForm1.TabStrip1.Tabs.Count
    For i = 1 To tabcount
        If MDIForm1.TabStrip1.Tabs(i).Tag = CStr(winid) Then
            MDIForm1.TabStrip1.Tabs(i).Caption = MDIForm1.ActiveForm.w.LocationName
            Exit For
        End If
    Next
End Sub

Private Sub w_NewWindow2(ppDisp As Object, Cancel As Boolean)
Dim fFile As Long, pObj As PopUpData, Found As Boolean
fFile = FreeFile
Open App.Path & "\PopUp.TXT" For Random As fFile Len = Len(pObj)
Do While Not EOF(fFile)
    DoEvents
    Get #fFile, , pObj
    If InStr(w.LocationURL, "?") > 0 Then
        If Trim(pObj.URL) = Left(LCase(Trim(w.LocationURL)), InStr(w.LocationURL, "?") - 1) Then Found = True
    Else
        If Trim(pObj.URL) = LCase(Trim(w.LocationURL)) Then Found = True
    End If
    If Found Then Exit Do
Loop
Close fFile

If Found Then
    If pObj.Action = "Block all" Then Cancel = True
Else
    frmPopupBlocker.lblSrc1 = "Source: " & w.LocationURL
    frmPopupBlocker.lblSrc2 = "Source: " & w.LocationURL
    frmPopupBlocker.Show vbModal
    
    Select Case frmPopupBlocker.Result
    Case "Allow"
    
    Case "Allow all"
        If InStr(w.LocationURL, "?") > 0 Then
            pObj.URL = Left(LCase(Left(w.LocationURL, 256)), InStr(w.LocationURL, "?") - 1)
        Else
            pObj.URL = LCase(Left(w.LocationURL, 256))
        End If
        pObj.Action = "Allow all"
        
        fFile = FreeFile
        Open App.Path & "\PopUp.TXT" For Random As fFile Len = Len(pObj)
        Put #fFile, LOF(fFile) / Len(pObj) + 1, pObj
        Close #fFile
    Case "Block"
        Cancel = True
    Case "Block all"
        Cancel = True
        
        If InStr(w.LocationURL, "?") > 0 Then
            pObj.URL = Left(LCase(Left(w.LocationURL, 256)), InStr(w.LocationURL, "?") - 1)
        Else
            pObj.URL = LCase(Left(w.LocationURL, 256))
        End If
        pObj.Action = "Block all"
        
        fFile = FreeFile
        Open App.Path & "\PopUp.TXT" For Random As fFile Len = Len(pObj)
        Put #fFile, LOF(fFile) / Len(pObj) + 1, pObj
        Close #fFile
    End Select
End If
End Sub

Private Sub w_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Form1.p.Max = ProgressMax
Form1.p.Value = Progress
Form1.p.Refresh
End Sub

Private Sub w_StatusTextChange(ByVal Text As String)
On Error Resume Next
Form1.StatusBar.Panels(5).Text = Text & " " & p.Value & "  %"
End Sub

Private Sub w_TitleChange(ByVal Text As String)
    MDIForm1.ActiveForm.Caption = Text
End Sub

