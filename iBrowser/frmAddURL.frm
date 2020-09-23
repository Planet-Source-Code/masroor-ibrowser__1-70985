VERSION 5.00
Begin VB.Form frmAddURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add URL"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddURL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdGetURL 
      Caption         =   "Get URL"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmAddURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetURL_Click()
txtURL.Text = MDIForm1.ActiveForm.Combo1.Text
End Sub

Private Sub cmdOK_Click()
If txtURL.Text = "" Then
MsgBox "Please type a URL that you want to add in Favorites List.EROOR 1080", vbCritical
Else
frmFavorites.List1.AddItem txtURL.Text
Unload Me
End If
End Sub
