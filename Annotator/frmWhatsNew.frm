VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWhatsNew 
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10320
   ControlBox      =   0   'False
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
   ScaleHeight     =   6885
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShow 
      Caption         =   "Stop showing this screen at start up"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   3075
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   10821
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close ""What's New..."""
      Height          =   435
      Left            =   8040
      TabIndex        =   0
      Top             =   6360
      Width           =   2175
   End
End
Attribute VB_Name = "frmWhatsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bShow As Boolean


Public Property Get PassCheck() As Boolean
    PassCheck = bShow
End Property
Public Property Let PassCheck(ByVal vNewValue As Boolean)
    bShow = vNewValue
End Property



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Not bShow Then chkShow.Visible = False
    web1.Navigate App.Path & "\WhatsNew.htm"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strInsert As String
    Dim rstL As ADODB.Recordset
    Dim lSessionID As Long
    
    If chkShow.value = 1 Then
        lSessionID = GetAnoSeq
        
        strInsert = "INSERT INTO " & ANOSession & " " & _
                    "(SESSIONID, USER_SEQ_ID, APPID, SESSIONDESC, " & _
                    "UPDUSER, UPDDTTM, UPDCNT) VALUES " & _
                    "(" & lSessionID & ", " & UserID & ", 1002, 'AnnoWhatsNew', " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
    End If
        
End Sub
