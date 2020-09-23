VERSION 5.00
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favorites"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   735
   End
   Begin VB.ListBox List1 
      Columns         =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
frmAddURL.Show
End Sub

Private Sub cmdClearAll_Click()
If List1.ListCount = 0 Then
MsgBox "Empty List can't be deleted. ERROR 210", vbCritical
Else
List1.Clear
End If
End Sub

Private Sub Form_Load()
 Dim sfile, a As String
    sfile = App.Path + "\" + "favorites.ib"
    On Error GoTo serr
    Open sfile For Input As #1           'this is a self loading INI file ID system
    Do Until EOF(1)                      'Specific for the Favs list this is working...
    Input #1, a$
    List1.AddItem a$
    Loop
    Close #1
serr:
    Close #1
If List1.Text = "default" Then
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sfile As String
Dim i As Long
List1.Visible = False
sfile = App.Path + "\" + "favorites.ib"
On Error GoTo error
Open sfile For Output As #1
For i = 0 To 100
    List1.ListIndex = i
     Write #1, List1.Text
Next i
error:
    Close #1

End Sub

Private Sub List1_DblClick()
MDIForm1.ActiveForm.w.Navigate List1.Text
End Sub
