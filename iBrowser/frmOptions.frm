VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "Current Page"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdBlank 
      Caption         =   "Blank"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
      Begin VB.CheckBox Check5 
         Caption         =   "Play sounds in webpage"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show Pictures"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Use iPopup"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Use Media Player"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Value           =   2  'Grayed
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use Java"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.TextBox txtHome 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "www.numair4u.net.tc"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Homepage:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlank_Click()
txtHome.Text = "about:blank"
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCurrent_Click()
txtHome.Text = MDIForm1.ActiveForm.Combo1.Text
End Sub

Private Sub cmdOK_Click()
Dim sfile As String
Dim i As Long
sfile = App.Path + "\" + "homepage.ib"
On Error GoTo error
Open sfile For Output As #1
     Write #1, txtHome.Text
error:
    Close #1
    Unload Me
End Sub

Private Sub Form_Load()
Dim sfile, a As String
    sfile = App.Path + "\" + "homepage.ib"
    On Error GoTo serr
    Open sfile For Input As #1           'this is a self loading INI file ID system
    Do Until EOF(1)                      'Specific for the Favs list this is working...
    Input #1, a$
    txtHome.Text = a$
    Loop
    Close #1
serr:
    Close #1
End Sub

