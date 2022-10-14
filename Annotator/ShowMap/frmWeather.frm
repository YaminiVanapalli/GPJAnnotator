VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWeather 
   Caption         =   "GPJ Weather Line"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
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
   ScaleHeight     =   8040
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   2475
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      ExtentX         =   4895
      ExtentY         =   4366
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tState As String

Public Property Get PassSTATE() As String
    PassSTATE = tState
End Property
Public Property Let PassSTATE(ByVal vNewValue As String)
    tState = vNewValue
End Property



Private Sub Form_Load()
    Dim sURL As String
    
    web1.Top = 240
    web1.Left = 240
    sURL = "http://iwin.nws.noaa.gov/iwin/" & tState & "/" & tState & ".html"
    web1.Navigate sURL
End Sub

Private Sub Form_Resize()
    web1.Width = Me.ScaleWidth - 480
    web1.Height = Me.ScaleHeight - 480
End Sub
