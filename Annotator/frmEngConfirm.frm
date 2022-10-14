VERSION 5.00
Begin VB.Form frmEngConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEngConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Show Year Approval"
      Height          =   1395
      Left            =   150
      TabIndex        =   2
      Top             =   2880
      Width           =   2235
      Begin VB.ComboBox cboSHYR 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   660
         Width           =   2040
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve && Close"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3540
      Width           =   2055
   End
   Begin VB.Label lblMess 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   5
      Top             =   180
      Width           =   4695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEngConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tSHNM As String
Dim tSHCD As Long

Public Property Get PassSHNM() As String
    PassSHNM = tSHNM
End Property
Public Property Let PassSHNM(ByVal vNewValue As String)
    tSHNM = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = tSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    tSHCD = vNewValue
End Property



Private Sub cboSHYR_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If cboSHYR.Text <> "" Then
        strSelect = "SELECT IGL_JDEDATE_TOCHAR(SHY56BEGDT, 'MON DD, YYYY')BEG_DATE " & _
                    "FROM " & F5601 & " " & _
                    "WHERE SHY56SHYR = " & cboSHYR.Text & " " & _
                    "AND SHY56SHCD = " & tSHCD
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            cmdApprove.Enabled = True
            lblYear = "Show begin date for " & cboSHYR.Text & " show is " & _
                        Trim(rst.Fields("BEG_DATE"))
        Else
            cmdApprove.Enabled = False
            lblYear = "No show record was found for this year"
        End If
        rst.Close: Set rst = Nothing
    Else
        cmdApprove.Enabled = False
    End If
End Sub

Private Sub cmdApprove_Click()
    Dim strSelect As String, strUpdate As String, strInsert As String
    Dim rst As ADODB.Recordset, rstN As ADODB.Recordset
    Dim lConfID As Long
    
    Screen.MousePointer = 11
    strSelect = "SELECT ENGCONFID FROM " & SRAEngCodeConf & " " & _
                "WHERE SHYR = " & cboSHYR.Text & " " & _
                "AND AN8_SHCD = " & tSHCD
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        lConfID = rst.Fields("ENGCONFID")
        rst.Close: Set rst = Nothing
        '///// UPDATE \\\\\
        strUpdate = "UPDATE " & SRAEngCodeConf & " " & _
                    "SET CONFIRMSTATUS = 1, " & _
                    "CONFIRMUSER = '" & ShortName & "', " & _
                    "CONFIRMDTTM = SYSDATE, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE ENGCONFID = " & lConfID
        Conn.Execute (strUpdate)
    Else
        rst.Close: Set rst = Nothing
        '///// INSERT \\\\\
        Set rstN = Conn.Execute("SELECT " & SRASeq & ".NEXTVAL FROM DUAL")
        lConfID = rstN.Fields("NEXTVAL")
        rstN.Close: Set rstN = Nothing
        strInsert = "INSERT INTO " & SRAEngCodeConf & " " & _
                    "(ENGCONFID, SHYR, AN8_SHCD, " & _
                    "CONFIRMSTATUS, CONFIRMUSER, CONFIRMDTTM, " & _
                    "UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lConfID & ", " & cboSHYR.Text & ", " & tSHCD & ", " & _
                    "1, '" & ShortName & "', SYSDATE, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
    End If
    
    frmShow.lblConfUser = "Confirmed for " & cboSHYR.Text & " Show on " & _
                format(Now, "DD-MMM-YYYYY") & " by " & LogName
    
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strMess As String, strSelect As String
    Dim rst As ADODB.Recordset
    Dim iYear As Integer, i As Integer
    
    Me.Caption = tSHNM
    strMess = "By selecting 'APPROVE' below, you are stating that you have reviewed " & _
                "the Code Requirements attached to " & DblAmp(tSHNM) & ", and confirmed " & _
                "them to be correct for the show year selected below." & _
                vbNewLine & vbNewLine & _
                "If a new Code Requirement is added, following your confirmation, " & _
                "you will be alerted of the change via an automated email, " & _
                "and your confirmation will be inactivated.  NOTE: If you are the one " & _
                "adding a new Code Requirement, no email will be sent.  " & _
                "Instead, an alert will remind you of the inactivated confirmation."
    lblMess = strMess
    
    strMess = "Show Year is the calendar year of the Show's begin date."
    lblYear = strMess
    
    iYear = CInt(format(Now, "YYYY"))
    cboSHYR.Clear
    strSelect = "SELECT SHY56SHYR FROM " & F5601 & " " & _
                "WHERE SHY56SHCD = " & tSHCD & " " & _
                "AND SHY56SHYR >= " & iYear - 1 & " " & _
                "ORDER BY SHY56SHYR"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboSHYR.AddItem rst.Fields("SHY56SHYR")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

