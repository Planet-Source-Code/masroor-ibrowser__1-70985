VERSION 5.00
Begin VB.Form frmPopupBlocker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iPopup Alert"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPopupBlocker.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBlockAll 
      Caption         =   "&Block All"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdBlock 
      Caption         =   "&Block"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdAllowAll 
      Caption         =   "&Allow All"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "&Allow"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPopupBlocker.frx":382A
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label lblSrc1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source: HTTP://"
      Height          =   795
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSrc2 
      BackStyle       =   0  'Transparent
      Caption         =   "Source: HTTP://"
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Popup Found..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPopupBlocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Result As String

Private Sub cmdAllow_Click()
Result = "Allow"
Unload Me
End Sub

Private Sub cmdAllowAll_Click()
Result = "Allow all"
Unload Me
End Sub

Private Sub cmdBlock_Click()
Result = "Block"
Unload Me
End Sub

Private Sub cmdBlockAll_Click()
Result = "Block all"
Unload Me
End Sub

Private Sub Form_Load()
Me.Refresh

End Sub

