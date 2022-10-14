VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownloader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloading..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5550
   FillColor       =   &H000040C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClose 
      Caption         =   "Close Window upon Completion"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2760
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.ListBox lstFile 
      Height          =   270
      Left            =   4500
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstTarget 
      Height          =   270
      Left            =   3600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstSource 
      Height          =   270
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar pbr1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1620
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.Animation ani1 
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1614
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   345
      FullHeight      =   61
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   6
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   780
      TabIndex        =   5
      Top             =   2100
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   2100
      Width           =   315
   End
   Begin VB.Label lblDLCnt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading 1 of 12"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1320
      Width           =   5235
   End
End
Attribute VB_Name = "frmDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Conn As New ADODB.Connection
Public Shortname As String, Logname As String
Dim UserID As Long
Dim bClose As Boolean


Private Sub chkClose_Click()
    bClose = CBool(chkClose.Value)
End Sub

Private Sub Form_Activate()
    Dim fs '' New Scripting.FileSystemObject
    Dim i As Integer
    Dim sSRC As String, sTARG As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Me.Refresh
'    Timer1.Enabled = True
    For i = 0 To lstSource.ListCount - 1
        sSRC = lstSource.List(i)
        sTARG = lstTarget.List(i)
        lblDLCnt.Caption = "Downloading " & i + 1 & " of " & lstSource.ListCount
        lblDLCnt.Refresh
        lblFile.Caption = lstFile.List(i)
        lblFile.Refresh
        lblSize.Caption = Format(FileLen(sSRC), "#,###,##0") & " bytes"
        lblSize.Refresh
        
        pbr1.Value = i
        pbr1.Refresh
        
        Me.Refresh
        
        fs.CopyFile sSRC, sTARG, True
        
    Next i
    
    pbr1.Value = pbr1.Max
    ani1.AutoPlay = False
    lblDLCnt.Caption = "Download Complete"
    lblFile.Caption = ""
    lblSize.Caption = ""
    
    Me.Caption = "File Download Successful..."
    
    Set fs = Nothing
    
    Call ClearDownload
    
    If bClose Then Unload Me
    
End Sub

Private Sub Form_Load()
'''    Dim Conn ''As Object
    Dim s$, cnt&, dl&
    Dim strSelect As String
    Dim rst '''As ADODB.Recordset
    
    bClose = True
    
    cnt& = 199
    s$ = String$(200, 0)
    dl& = GetUserName(s$, cnt)
    Shortname = Left$(s$, cnt)
    If Asc(Right(Shortname, 1)) = 0 Then Shortname = UCase(Left(Shortname, Len(Shortname) - 1))
    
    On Error Resume Next
    
'    MsgBox "About to Create Conn object"
'    Set Conn = CreateObject("ADODB.Connection")
'    MsgBox "Conn created, open next"
'''    Conn.Open ("DSN=JDE;UID=ANNOTATOR;PWD=ANNOTATOR")
    Conn.Open ("DSN=JDETEST;UID=ANNOTATOR_APP_USER;PWD=q2eNqsgHxcKqre3")
'    If Err = 0 Then
'        MsgBox "Conn open"
'    Else
'        MsgBox "Conn open error (" & Err.Number & " - " & Err.Description & ")"
'    End If
    
    strSelect = "SELECT (TRIM(U.NAME_FIRST) || ' ' || TRIM(U.NAME_LAST))LOGNAME, U.USER_SEQ_ID " & _
                "FROM IGL_USER U " & _
                "WHERE U.NAME_LOGON = '" & Shortname & "'"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Logname = Trim(rst.Fields("LOGNAME"))
        UserID = rst.Fields("USER_SEQ_ID")
    Else
        rst.Close: Set rst = Nothing
        MsgBox "Unable to download at this time", vbCritical, "Sorry..."
        Unload Me
        Exit Sub
    End If
    rst.Close: Set rst = Nothing
                
    
    Call GetDownload(UserID)
    pbr1.Max = lstSource.ListCount
    pbr1.Value = 0
    
    ani1.Open App.Path & "\FILECOPY.AVI"
End Sub

Public Sub GetDownload(pUID As Long)
    Dim strSelect As String, strUpdate As String
    Dim rst As ADODB.Recordset
    On Error Resume Next
    lstSource.Clear
    lstTarget.Clear
    lstFile.Clear
    
    strSelect = "SELECT DLID, SOURCE_PATH, DL_PATH, FILE_NAME " & _
                "FROM ANO_DOWNLOAD " & _
                "WHERE USER_SEQ_ID = " & pUID & " " & _
                "AND DLSTATUS = 5"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstSource.AddItem Trim(rst.Fields("SOURCE_PATH"))
        lstSource.ItemData(lstSource.NewIndex) = rst.Fields("DLID")
        lstTarget.AddItem Trim(rst.Fields("DL_PATH")) & Trim(rst.Fields("FILE_NAME"))
        lstTarget.ItemData(lstTarget.NewIndex) = rst.Fields("DLID")
        lstFile.AddItem Trim(rst.Fields("FILE_NAME"))
        
        strUpdate = "UPDATE ANO_DOWNLOAD SET " & _
                    "DLSTATUS = 10, " & _
                    "UPDDTTM = SYSDATE, " & _
                    "UPDCNT = UPDCNT + 1 " & _
                    "WHERE DLID = " & rst.Fields("DLID")
        Conn.Execute (strUpdate)
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub

Public Sub ClearDownload()
    Dim i As Integer
    Dim strDelete As String
    Dim lDLID As Long
    
    For i = lstSource.ListCount - 1 To 0 Step -1
        lDLID = lstSource.ItemData(i)
        strDelete = "DELETE FROM ANO_DOWNLOAD " & _
                    "WHERE DLID = " & lDLID
        Conn.Execute (strDelete)
    Next i
End Sub
