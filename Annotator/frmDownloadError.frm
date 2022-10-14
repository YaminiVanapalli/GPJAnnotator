VERSION 5.00
Begin VB.Form frmDownloadError 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download Error..."
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      ScaleHeight     =   375
      ScaleWidth      =   5775
      TabIndex        =   7
      Top             =   3960
      Width           =   5775
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Never ask me again for this application"
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set this option! "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Full Access"
      Height          =   315
      Left            =   750
      TabIndex        =   6
      Top             =   2490
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5820
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact the GPJ Helpdesk and they will help you resolve the issue."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   4980
      TabIndex        =   4
      Top             =   4980
      Width           =   2280
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If you continue to have problems..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   4980
      TabIndex        =   3
      Top             =   4560
      Width           =   2205
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDownloadError.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You were not able to download the selected files.  Typically this is due to incorrect ICA Client File Security settings."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set this option! "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   0
      Top             =   2460
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   120
      Picture         =   "frmDownloadError.frx":0102
      Top             =   120
      Width           =   4275
   End
End
Attribute VB_Name = "frmDownloadError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
