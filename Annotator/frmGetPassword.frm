VERSION 5.00
Begin VB.Form frmGetPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lotus Notes"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1515
   End
   Begin VB.TextBox txtPW 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please, enter you password (case sensitive):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetPassword.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmGetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bPassed As Boolean


Private Sub cmdCancel_Click()
    sNOTESPASSWORD = "_CANCEL"
    bPassed = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    sNOTESPASSWORD = Trim(txtPW.Text)
    bPassed = True
    Unload Me
End Sub

Private Sub Form_Load()
    bPassed = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bPassed Then sNOTESPASSWORD = "_CANCEL"
End Sub

Private Sub txtPW_GotFocus()
    cmdOK.Default = True
End Sub
