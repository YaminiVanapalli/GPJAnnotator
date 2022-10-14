VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmViewer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
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
   ScaleHeight     =   4035
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4080
      ScaleHeight     =   285
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   3690
      Width           =   1335
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close Viewer"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         MouseIcon       =   "frmViewer.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   1275
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
      ExtentX         =   2355
      ExtentY         =   2249
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   1720
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pFile As String
Dim pLeft As Long, pTop As Long

Public Property Get PassFile() As String
    PassFile = pFile
End Property
Public Property Let PassFile(ByVal vNewValue As String)
    pFile = vNewValue
End Property

Public Property Get PassLeft() As Long
    PassLeft = pLeft
End Property
Public Property Let PassLeft(ByVal vNewValue As Long)
    pLeft = vNewValue
End Property

Public Property Get PassTop() As Long
    PassTop = pTop
End Property
Public Property Let PassTop(ByVal vNewValue As Long)
    pTop = vNewValue
End Property





Private Sub Form_Load()
    Dim iDot As Integer
    Dim sFormat As String
    
    Me.Top = pTop
    Me.Left = pLeft
    
    For iDot = Len(pFile) To 0 Step -1
        If Mid(pFile, iDot, 1) = "." Then
            sFormat = UCase(Mid(pFile, iDot + 1))
            Exit For
        End If
    Next iDot
    
    Select Case sFormat
        Case "AVI", "MPG"
            wmp1.Top = 0
            wmp1.Left = 0
            wmp1.Width = Me.Width
            wmp1.Height = Me.Height
            wmp1.URL = pFile
            wmp1.Visible = True
        Case Else
            web1.Top = 0
            web1.Left = 0
            web1.Width = Me.Width
            web1.Height = Me.Height
            web1.Navigate2 pFile
            web1.Visible = True
    End Select
    
    
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub
