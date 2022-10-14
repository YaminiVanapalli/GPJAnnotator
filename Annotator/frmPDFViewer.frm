VERSION 5.00
Object = "{23319180-2253-11D7-BD2E-08004608C318}#3.0#0"; "XpdfViewerCtrl.ocx"
Begin VB.Form frmCubing 
   Caption         =   "GPJ Annotator PDF Viewer"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPDFViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin XpdfViewerCtl.XpdfViewer xpdf1 
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   900
      Width           =   3795
      showScrollbars  =   -1  'True
      showBorder      =   -1  'True
      showPasswordDialog=   -1  'True
   End
   Begin VB.PictureBox picSelect 
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   120
      Width           =   1995
      Begin VB.Label lblSelect 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close PDF Viewer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         MouseIcon       =   "frmPDFViewer.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Graphics Viewer is loading..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   270
      Left            =   2220
      TabIndex        =   0
      Top             =   232
      Width           =   3150
   End
End
Attribute VB_Name = "frmCubing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Screen.MousePointer = 11
    With xpdf1
        .Top = 675
        .Left = 120
'''        .Width = Me.ScaleWidth - 240
'''        .Height = Me.ScaleHeight - 795
    End With
    xpdf1.loadFile (sPDFFile)
    
    lblWelcome = frmAnnotator.lblWelcome
    Me.WindowState = frmAnnotator.WindowState
    If Me.WindowState = 0 Then
        Me.Width = frmAnnotator.Width
        Me.Height = frmAnnotator.Height
        Me.Top = frmAnnotator.Top
        Me.Left = frmAnnotator.Left
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSelect.FontBold = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        xpdf1.Width = Me.ScaleWidth - (xpdf1.Left * 2)
        xpdf1.Height = Me.ScaleHeight - xpdf1.Top - xpdf1.Left
    End If
End Sub

Private Sub lblSelect_Click()
    Unload Me
End Sub

Private Sub lblSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSelect.FontBold = True
End Sub

Private Sub picSelect_Click()
    Unload Me
End Sub

Private Sub picSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSelect.FontBold = True
End Sub
