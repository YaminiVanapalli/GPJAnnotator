VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{02BF25D2-8C17-4B23-BC80-D3488ABDDC6B}#2.0#0"; "QTPlugin.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin QTActiveXPluginCtl.QTActiveXPlugin qtm1 
      Height          =   675
      Left            =   780
      TabIndex        =   2
      Top             =   3540
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   7200
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   9600
      ExtentX         =   16933
      ExtentY         =   12700
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
      Location        =   ""
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   3375
      Left            =   2940
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   4395
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
      _cx             =   7752
      _cy             =   5953
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuTest 
         Caption         =   "AVI Test"
         Index           =   0
      End
      Begin VB.Menu mnuTest 
         Caption         =   "MPG Test"
         Index           =   1
      End
      Begin VB.Menu mnuTest 
         Caption         =   "MOV Test"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPath As String

Private Sub Form_load()
    sPath = "C:\Data\VB Projects\GPJAnnotator\2003-11-06 Annotator (Priority 2 mods)\"
    
    web1.Top = 120: web1.Left = 120
    wmp1.Top = 120: wmp1.Left = 120
    qtm1.Top = 120: qtm1.Left = 120
End Sub

Private Sub Form_Resize()
    web1.Width = Me.ScaleWidth - (web1.Left * 2)
    web1.Height = Me.ScaleHeight - web1.Top - web1.Left
    
    wmp1.Width = web1.Width
    wmp1.Height = web1.Height
    
    qtm1.Width = web1.Width
    qtm1.Height = web1.Height
End Sub

Private Sub mnuMenu_Click()
'''    wmp1.Close
'''    web1.Navigate2 "blank.htm"
'''    web1.Visible = False
'''    wmp1.Visible = False
End Sub

Private Sub mnuTest_Click(Index As Integer)
    wmp1.Close
    web1.Navigate2 "blank.htm"
    web1.Visible = False
    wmp1.Visible = False
    Select Case Index
        Case 0
            wmp1.Visible = True
'''            web1.Navigate2 (sPath & "GPJ Events Teaser.avi")
            wmp1.URL = sPath & "GPJ Events Teaser.avi"
        Case 1
            wmp1.Visible = True
            wmp1.URL = sPath & "Kurzfassung.mpg"
'''            web1.Navigate2 (sPath & "Kurzfassung.mpg")
        Case 2
'''            qtm1.Visible = True
'''            qtm1.SetURL (sPath & "GPJ-EADS.mov")
'''
''''            qtm1.GetRectangle
            Me.Height = 8550
            Me.Width = 9960
            web1.Visible = True
            web1.Navigate2 (sPath & "GPJ-EADS.mov")
            
    End Select
    
'''    Select Case Index
'''        Case 0: wmp1.URL = sPath & "GPJ Events Teaser.avi"
'''        Case 1: wmp1.URL = sPath & "Kurzfassung.mpg"
'''        Case 2: wmp1.URL = sPath & "GPJ-EADS.mov"
'''    End Select
End Sub
