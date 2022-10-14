VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3675
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   120
      Width           =   6195
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblCurrStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   765
         Left            =   210
         TabIndex        =   6
         Top             =   1500
         Width           =   5700
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   60
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   -300
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   4680
         Picture         =   "frmSplash.frx":030A
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5040
         TabIndex        =   4
         Top             =   60
         Width           =   945
      End
      Begin VB.Label lblApp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   4815
         TabIndex        =   3
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5220
         TabIndex        =   2
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GPJ Annotator"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   810
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   2640
         Width           =   4545
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblCompany = "Property of:  " & App.CompanyName
    lblApp = "Application Name:  " & App.ProductName
    lblVersion.Caption = "Version:  " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblCurrStatus = frmAnnotator.lblUpdateNote.Caption
End Sub
