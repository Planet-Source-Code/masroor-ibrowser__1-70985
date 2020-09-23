VERSION 5.00
Begin VB.Form frmPopupManager 
   Caption         =   "Popup Manager"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "frmPopupManager.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6495
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ListBox lstAllow 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ListBox lstBlock 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      ItemData        =   "frmPopupManager.frx":382A
      Left            =   4080
      List            =   "frmPopupManager.frx":382C
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdBlockAll 
      BackColor       =   &H008080FF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdBlock 
      BackColor       =   &H00C0C0FF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton cmdAllow 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdAllowAll 
      BackColor       =   &H0080FF80&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblSrc2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   5025
      TabIndex        =   14
      Top             =   1575
      Width           =   1395
   End
   Begin VB.Label lblSrc2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allow list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "frmPopupManager.frx":382E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPopupManager.frx":40F8
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   7575
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   7455
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   7455
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Popup Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   855
      TabIndex        =   11
      Top             =   15
      Width           =   3750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Popup Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D6AB8B&
      Height          =   585
      Left            =   840
      TabIndex        =   10
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label lblSrc1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allow list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   0
      Left            =   1065
      TabIndex        =   9
      Top             =   1560
      Width           =   1350
   End
   Begin VB.Label lblSrc1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   1
      Left            =   5010
      TabIndex        =   8
      Top             =   1560
      Width           =   1395
   End
End
Attribute VB_Name = "frmPopupManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type PopUpData
    URL As String * 256
    Action As String * 9
End Type
Dim DeleteFrom As String

Private Sub cmdAllow_Click()
If lstBlock.ListCount = 0 Then Exit Sub

lstAllow.AddItem lstBlock.List(lstBlock.ListIndex)
lstBlock.RemoveItem lstBlock.ListIndex

If lstBlock.ListCount > 0 Then lstBlock.ListIndex = 0
End Sub

Private Sub cmdAllowAll_Click()
If lstBlock.ListCount = 0 Then Exit Sub

Dim a As Long
For a = 0 To lstBlock.ListCount
    lstAllow.AddItem lstBlock.List(a)
Next
lstBlock.Clear
End Sub

Private Sub cmdBlock_Click()
If lstAllow.ListCount = 0 Then Exit Sub

lstBlock.AddItem lstAllow.List(lstAllow.ListIndex)
lstAllow.RemoveItem lstAllow.ListIndex

If lstAllow.ListCount > 0 Then lstAllow.ListIndex = 0
End Sub

Private Sub cmdBlockAll_Click()
If lstAllow.ListCount = 0 Then Exit Sub

Dim a As Long
For a = 0 To lstAllow.ListCount
    lstBlock.AddItem lstAllow.List(a)
Next
lstAllow.Clear
End Sub

Private Sub cmdDelete_Click()
If DeleteFrom = "lstAllow" Then
    If lstAllow.ListCount = 0 Then Exit Sub
    lstAllow.RemoveItem lstAllow.ListIndex
    If lstAllow.ListCount > 0 Then lstAllow.ListIndex = 0
Else
    If lstBlock.ListCount = 0 Then Exit Sub
    lstBlock.RemoveItem lstBlock.ListIndex
    If lstBlock.ListCount > 0 Then lstBlock.ListIndex = 0
End If
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim fFile As Long, pObj As PopUpData, a As Long

fFile = FreeFile
Open App.Path & "\PopUp.TXT" For Random As fFile Len = Len(pObj)
While Not EOF(fFile)
    DoEvents
    Get #fFile, , pObj
    If pObj.Action = "Allow all" Then
        lstAllow.AddItem Trim(pObj.URL)
    End If
    If pObj.Action = "Block all" Then
        lstBlock.AddItem Trim(pObj.URL)
    End If
Wend
Close fFile

If lstAllow.ListCount > 0 Then lstAllow.ListIndex = 0
If lstBlock.ListCount > 0 Then lstBlock.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim fFile As Long, pObj As PopUpData, a As Long

If Dir(App.Path & "\PopUp.TXT") > "" Then Kill App.Path & "\PopUp.TXT"
fFile = FreeFile
Open App.Path & "\PopUp.TXT" For Random As fFile Len = Len(pObj)

If lstAllow.ListCount > 0 Then
    For a = 0 To lstAllow.ListCount - 1
        pObj.URL = lstAllow.List(a)
        pObj.Action = "Allow all"
        Put #fFile, , pObj
    Next
End If

If lstBlock.ListCount > 0 Then
    For a = 0 To lstBlock.ListCount - 1
        pObj.URL = lstBlock.List(a)
        pObj.Action = "Block all"
        Put #fFile, , pObj
    Next
End If
Close fFile
End Sub

Private Sub lstAllow_Click()
DeleteFrom = "lstAllow"
End Sub

Private Sub lstBlock_Click()
DeleteFrom = "lstBlock"
End Sub

