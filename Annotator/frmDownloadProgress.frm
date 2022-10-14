VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDownloadProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download Progress..."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7050
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
   ScaleHeight     =   1530
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation ani1 
      Height          =   735
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   257
      FullHeight      =   49
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size..."
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copying..."
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   795
   End
End
Attribute VB_Name = "frmDownloadProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lSize As Long, lProgress As Long

Dim pSRC As String, pDES As String

Public Property Get PassSRCFILE() As String
    PassSRCFILE = pSRC
End Property
Public Property Let PassSRCFILE(ByVal vNewValue As String)
    pSRC = vNewValue
End Property

Public Property Get PassDESFILE() As String
    PassDESFILE = pDES
End Property
Public Property Let PassDESFILE(ByVal vNewValue As String)
    pDES = vNewValue
End Property



Private Sub Form_Activate()
    Dim fs As New Scripting.FileSystemObject
    
    Me.Refresh
'    Timer1.Enabled = True
    
    fs.CopyFile pSRC, pDES, True
    
    ani1.AutoPlay = False
    
    Me.Caption = "File Download Successful..."
    lblFile.Caption = "File Copied to " & pDES
    lblSize.Caption = "Download Complete..."
    
    Set fs = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    lSize = FileLen(pSRC)
    lblSize.Caption = Format(lSize / 1000, "#,##0") & " KB"
    lblFile.Caption = "Copying to " & UCase(pDES)
    ani1.Open App.Path & "\FILECOPY.AVI"
    
End Sub

Private Sub Form_Resize()
    If lblFile.Width > ani1.Width Then
        Me.Width = (Me.Width - Me.ScaleWidth) + lblFile.Width + (lblFile.Left * 2)
'        ani1.Left = (Me.ScaleWidth - ani1.Width) / 2
'        lblFile.Left = (Me.ScaleWidth - lblFile.Width) / 2
'        lblSize.Left = (Me.ScaleWidth - lblSize.Width) / 2
    End If
    
End Sub
