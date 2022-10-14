VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPrintOut 
   Caption         =   "Report View"
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
   Icon            =   "frmPrintOut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email Report..."
      Height          =   435
      Left            =   9240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1600
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   5235
      ExtentX         =   9234
      ExtentY         =   7435
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
End
Attribute VB_Name = "frmPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile As String, sHDR As String

Public Property Get PassPATH() As String
    PassPATH = sFile
End Property
Public Property Let PassPATH(ByVal vNewValue As String)
    sFile = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = sHDR
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    sHDR = vNewValue
End Property



'''Private Sub cmdEMail_Click()
'''    frmEmailReport.PassFile = sFile
'''    frmEmailReport.PassSub = sHDR
'''    frmEmailReport.PassUser = LogName
'''    frmEmailReport.Show 1, Me
'''End Sub

Private Sub Form_Load()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim dDate As Date
    Dim iDay As Integer, i As Integer
    
    web1.Left = 120
    web1.Top = 120
    
    
    web1.Navigate (sFile)
    
    Me.Caption = sHDR
    
'''    strSelect = "SELECT TIMEDT FROM TIMETRACK"
'''    Set rst = AConn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        dDate = rst.Fields("TIMEDT")
'''    End If
'''    rst.Close: Set rst = Nothing
'''    iDay = CInt(Format(dDate, "w"))
'''    dDate = DateAdd("d", CDbl(7 - iDay), dDate)
'''    cboWE.AddItem Format(dDate, "mmm dd, yyyy")
'''    i = 0
'''    Do While dDate < Now
'''        dDate = DateAdd("ww", 1, dDate)
'''        cboWE.AddItem Format(dDate, "mmm dd, yyyy")
'''        i = i + 1
'''    Loop
    
End Sub

'''''SELECT TASKMASTER.PROJID, TASKMASTER.CLIENT, TASKMASTER.TASKNAME,
'''''TIMETRACK.TASKID, FORMAT(TIMETRACK.TIMEDT, "W") AS DAYOFWEEK, TIMETRACK.HOURS
'''''From TIMETRACK, TASKMASTER
'''''Where TIMETRACK.TIMEDT >= #8/12/2001#
'''''AND TIMETRACK.TIMEDT <= #8/18/2001#
'''''AND TIMETRACK.TASKID = TASKMASTER.TASKID
'''''ORDER BY TASKMASTER.CLIENT, TIMETRACK.TIMEDT;
Private Sub Form_Resize()
    web1.Width = Me.ScaleWidth - 240
    web1.Height = Me.ScaleHeight - 240
    cmdEmail.Left = Me.Width - 2220
End Sub
