VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.ComboBox cboSearchBox 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Google"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtSearchBox 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Search Engine"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If cboSearchBox.Text = "Select a Search Engine" Then
MsgBox "Please select a search engine to search", vbInformation, "Select Engine"
End If
If txtSearchBox.Text = "" Then
MsgBox "Please enter at least 1 word to search for", vbInformation, "Enter Word"
End If

On Error Resume Next
Select Case cboSearchBox.Text
    
Case "MSN"
    MDIForm1.ActiveForm.w.Navigate ("http://search.msn.com/results.asp?RS=CHECKED&FORM=MSNH&v=1&q=" & txtSearchBox.Text)
Case "Excite"
    MDIForm1.ActiveForm.w.Navigate ("http://search.excite.com/search.gw?search=" & txtSearchBox.Text)
Case "Google"
    MDIForm1.ActiveForm.w.Navigate ("http://www.google.com/search?q=" & txtSearchBox.Text & "&meta=lr%3D%26hl%3Den&btnG=Google+Search")
Case "Yahoo"
    MDIForm1.ActiveForm.w.Navigate ("http://ink.yahoo.com/bin/query?p=" & txtSearchBox.Text & "&z=2&hc=0&hs=0")
Case "Altavista"
    MDIForm1.ActiveForm.w.Navigate ("http://www.altavista.com/cgi-bin/query?pg=q&kl=XX&stype=stext&q=" & txtSearchBox.Text)
Case "Lycos"
    MDIForm1.ActiveForm.w.Navigate ("http://www.lycos.com/srch/?lpv=1&loc=searchhp&query=" & txtSearchBox.Text)
Case "About.COM"
    MDIForm1.ActiveForm.w.Navigate ("http://search.about.com/fullsearch.htm?terms=" & "&PM=59_0100_S&Action.x=9&Action.y=7 ")

End Select
Unload Me
End Sub

Private Sub Form_Load()
cboSearchBox.AddItem "MSN"
cboSearchBox.AddItem "Excite"
cboSearchBox.AddItem "Yahoo!"
cboSearchBox.AddItem "Lycos"
cboSearchBox.AddItem "About.com"
End Sub
