VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Begin VB.Form frmStartUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPJ Annotator"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VOLOVIEWXLibCtl.AvViewX vol1 
      Height          =   435
      Left            =   12060
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   555
      _cx             =   979
      _cx             =   979
      src             =   ""
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      UserMode        =   "Pan"
      HighlightLinks  =   0   'False
      LayersOff       =   ""
      LayersOn        =   ""
      SrcTemp         =   ""
      SupportPath     =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
      FontPath        =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
      NamedView       =   ""
      BackgroundColor =   "DefaultColors"
      GeometryColor   =   "DefaultColors"
      PrintBackgroundColor=   "16777215"
      PrintGeometryColor=   "0"
      ShadingMode     =   "Gouraud"
      ProjectionMode  =   "Parallel"
      EnableUIMode    =   "DefaultUI"
      Layout          =   ""
   End
   Begin VB.PictureBox picHdr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      Picture         =   "frmStartUp.frx":08CA
      ScaleHeight     =   1440
      ScaleWidth      =   11370
      TabIndex        =   14
      Top             =   0
      Width           =   11370
   End
   Begin VB.PictureBox picFly 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   3015
      TabIndex        =   11
      Top             =   6780
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lblOptions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   20.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001CAF6F&
         Height          =   405
         Index           =   1
         Left            =   60
         MouseIcon       =   "frmStartUp.frx":7DB4
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "Click to Close Options Menu"
         Top             =   240
         Width           =   405
      End
      Begin VB.Image imgOptions 
         Height          =   495
         Left            =   420
         MouseIcon       =   "frmStartUp.frx":80BE
         MousePointer    =   99  'Custom
         Picture         =   "frmStartUp.frx":83C8
         Stretch         =   -1  'True
         ToolTipText     =   "Options..."
         Top             =   120
         Width           =   495
      End
      Begin VB.Image imgEmailTeam 
         Height          =   615
         Left            =   960
         MouseIcon       =   "frmStartUp.frx":880A
         MousePointer    =   99  'Custom
         Picture         =   "frmStartUp.frx":8B14
         Stretch         =   -1  'True
         ToolTipText     =   "Email Team Maintenance"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgSecurity 
         Height          =   615
         Left            =   2340
         MouseIcon       =   "frmStartUp.frx":8E1E
         MousePointer    =   99  'Custom
         Picture         =   "frmStartUp.frx":9128
         Stretch         =   -1  'True
         ToolTipText     =   "Annotator Security Interface"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgUserReports 
         Height          =   615
         Left            =   1620
         MouseIcon       =   "frmStartUp.frx":9432
         MousePointer    =   99  'Custom
         Picture         =   "frmStartUp.frx":973C
         Stretch         =   -1  'True
         ToolTipText     =   "Annotator Usage Reports"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facilities"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   6
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":9A46
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2400
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George P. Johnson Annotator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      MouseIcon       =   "frmStartUp.frx":9D50
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   4665
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   10560
      MouseIcon       =   "frmStartUp.frx":A05A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label lblWelcome 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   6780
      Width           =   1980
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   20.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001CAF6F&
      Height          =   405
      Index           =   0
      Left            =   60
      MouseIcon       =   "frmStartUp.frx":A364
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Click to access Options"
      Top             =   7020
      Width           =   405
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10620
      MouseIcon       =   "frmStartUp.frx":A66E
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2160
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Image Library"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   5
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":A978
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label lblDriver 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11160
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6180
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Showplans & Show Abstracts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":AC82
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3000
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Graphic Importer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   4
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":AF8C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6000
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Graphics Handler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   2
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":B296
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engineering Drawings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":B5A0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5400
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Space Plan Viewing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   1
      Left            =   480
      MouseIcon       =   "frmStartUp.frx":B8AA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3600
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label lblback2 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   1440
      Width           =   11370
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuPassword 
         Caption         =   "Reset Password..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConnSpeed 
         Caption         =   "Connection Speed"
         Visible         =   0   'False
         Begin VB.Menu mnuSpeed 
            Caption         =   "Low Speed (Dial-Up, 28.8k, 56k...)"
            Index           =   0
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "High Speed (ISDN, T1, Cable Modem...)"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuPrinterDrivers 
         Caption         =   "Printer Drivers..."
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWhatsNew 
         Caption         =   "What's New..."
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/////---------- FORM MODIFICATIONS LOG ---------\\\\\
''06-SEP-2001   SAW     CHANGES MADE TO HELP RESOLVE CITRIX PRINTING ISSUES
''
''
''
'\\\\\----------------------------------------------/////

Dim sPath As String, sDriver As String
'''Dim lColor As Long
Dim iLoginTrys As Integer, iLabel As Integer
Dim bUserCheck As Boolean

'''Private Sub cmdCancel_Click()
'''    fraNewPassword.Visible = False
'''    cmdGo.Default = True
'''End Sub

'''Private Sub cmdOK_Click()
'''    Dim strSelect As String, strUpdate As String
'''    Dim rst As ADODB.Recordset
'''    If txtNewPassword(0).Text = txtNewPassword(1).Text _
'''                And Len(txtNewPassword(0).Text) >= 8 _
'''                And Len(txtNewPassword(0).Text) <= 16 Then
'''        strSelect = "SELECT U.NAME_LOGON, R.USER_SEQ_ID " & _
'''                    "FROM " & IGLUserAR & " R, " & IGLUser & " U " & _
'''                    "WHERE U.NAME_LOGON = '" & UCase(txtLogName.Text) & "' " & _
'''                    "AND U.USER_SEQ_ID = R.USER_SEQ_ID " & _
'''                    "AND R.APP_ID = 1002 " & _
'''                    "AND UPPER(R.PCODE) = '" & UCase(txtExistingPassword.Text) & "'"
'''        Set rst = Conn.Execute(strSelect)
'''        If Not rst.EOF Then
'''            On Error Resume Next
'''            strUpdate = "UPDATE " & IGLUserAR & " " & _
'''                        "SET PCODE = '" & txtNewPassword(0).Text & "', " & _
'''                        "UPDUSER = '" & rst.Fields("NAME_LOGON") & "', " & _
'''                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
'''                        "WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
'''                        "AND APP_ID = 1002"
'''            Conn.Execute (strUpdate)
'''            rst.Close: Set rst = Nothing
'''            If Err Then
'''                MsgBox "Error:  " & Err.Description, vbExclamation, "Could not change password..."
'''            Else
'''                fraNewPassword.Visible = False
'''                cmdOK.Enabled = False
'''                cmdOK.Default = False
'''                cmdGo.Default = True
'''                txtPassword.Text = txtNewPassword(0).Text
'''                txtExistingPassword.Text = ""
'''                txtNewPassword(0).Text = ""
'''                txtNewPassword(1).Text = ""
'''                With txtPassword
'''                    .SelStart = 0
'''                    .SelLength = Len(txtPassword.Text)
'''                    .SetFocus
'''                End With
'''            End If
'''        Else
'''            rst.Close: Set rst = Nothing
'''            MsgBox "The password entered as your 'Existing Password' is not correct.", _
'''                        vbExclamation, "Incorrect Password..."
'''        End If
'''
'''    End If
'''End Sub

'''Private Sub cmdResetPassword_Click()
'''    Dim strSelect As String
'''    Dim rst As ADODB.Recordset
''''''    If Len(txtLogName.Text) = 8 Then
'''        strSelect = "SELECT R.PCODE " & _
'''                    "FROM " & IGLUserAR & " R, " & IGLUser & " U " & _
'''                    "WHERE U.NAME_LOGON = '" & UCase(txtLogName.Text) & "' " & _
'''                    "AND U.USER_SEQ_ID = R.USER_SEQ_ID " & _
'''                    "AND R.APP_ID = 1002"
'''        Set rst = Conn.Execute(strSelect)
'''        If Not rst.EOF Then
'''            rst.Close
'''            Set rst = Nothing
'''            fraNewPassword.Visible = True
'''            If txtPassword.Text <> "" Then
'''                txtExistingPassword.Text = txtPassword
'''                txtNewPassword(0).SetFocus
'''            Else
'''                txtExistingPassword.SetFocus
'''            End If
'''            txtNewPassword(0).Text = ""
'''            txtNewPassword(1).Text = ""
'''            cmdGo.Default = False
'''        Else
'''            rst.Close
'''            Set rst = Nothing
'''            MsgBox "Login name is not a registered User.", vbExclamation, "Invalid User..."
'''        End If
''''''    End If
'''End Sub



'''Private Sub cboViewUsage_Click()
'''    Dim strSelect As String
'''    Dim sSource As String, sName As String, sMess As String, _
'''                sShow As String, sNewShow As String, sType As String
'''    Dim rst As ADODB.Recordset
'''    Dim iOff As Integer, iCase As Integer, iLen As Integer, iQty As Integer
'''
'''
'''    iQty = 0
'''    If cboViewUsage.Text = "" Then Exit Sub
'''    If cboViewUsage.ItemData(cboViewUsage.ListIndex) = 0 Then Exit Sub
'''
'''    Me.MousePointer = 11
'''
'''    iCase = cboViewUsage.ItemData(cboViewUsage.ListIndex)
'''    Select Case iCase
'''        Case 12, 13, 14, 15
'''            iOff = iCase - 12
'''            strSelect = "SELECT ADDUSER, COUNT(GID) AS GCOUNT " & _
'''                        "From GFX_MASTER " & _
'''                        "WHERE TO_CHAR(ADDDTTM, 'DD-MON-YY') = '" & _
'''                        UCase(format(DateAdd("d", (iCase - 12) * -1, Date), "DD-MMM-YY")) & "' " & _
'''                        "GROUP BY ADDUSER"
'''            sMess = ""
'''            Set rst = Conn.Execute(strSelect)
'''            Do While Not rst.EOF
'''                iQty = iQty + rst.Fields("GCOUNT")
'''                sMess = sMess & vbTab & rst.Fields("GCOUNT") & vbTab & Trim(rst.Fields("ADDUSER")) & vbNewLine
'''                rst.MoveNext
'''            Loop
'''            rst.Close: Set rst = Nothing
'''            If sMess = "" Then
'''                sMess = "THERE WERE NO GRAPHICS POSTED"
'''            Else
'''                sMess = "THE FOLLOWING QUANTITIES OF GRAPHICS WERE POSTED:" & vbNewLine & sMess
'''            End If
'''
'''        Case 16, 17
'''            If iCase = 16 Then iLen = -7 Else iLen = -30
'''            strSelect = "SELECT ADDUSER, COUNT(GID) AS GCOUNT " & _
'''                        "From GFX_MASTER " & _
'''                        "WHERE ADDDTTM BETWEEN " & _
'''                        "TO_DATE('" & UCase(format(DateAdd("d", iLen, Date), "DD-MMM-YY")) & "', 'DD-MON-YY') " & _
'''                        "AND TO_DATE('" & UCase(format(Date, "DD-MMM-YY")) & "', 'DD-MON-YY') " & _
'''                        "GROUP BY ADDUSER"
'''            sMess = ""
'''            Set rst = Conn.Execute(strSelect)
'''            Do While Not rst.EOF
'''                iQty = iQty + rst.Fields("GCOUNT")
'''                sMess = sMess & vbTab & rst.Fields("GCOUNT") & vbTab & Trim(rst.Fields("ADDUSER")) & vbNewLine
'''                rst.MoveNext
'''            Loop
'''            rst.Close: Set rst = Nothing
'''            If sMess = "" Then
'''                sMess = "THERE WERE NO GRAPHICS POSTED"
'''            Else
'''                sMess = "THE FOLLOWING QUANTITIES OF GRAPHICS WERE POSTED:" & vbNewLine & sMess
'''            End If
'''
'''    End Select
'''
'''    Me.MousePointer = 0
'''
'''    With frmUsage
'''        .PassMess = sMess
'''        Select Case iCase
'''            Case 16: .PassTitle = "POSTING LOG: Quantity of Graphics posted during past 7 Days"
'''            Case 17: .PassTitle = "POSTING LOG: Quantity of Graphics posted during past 30 Days"
'''            Case Else
'''                .PassTitle = "USAGE LOG: " & format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
'''        End Select
'''        .PassQty = iQty
'''        .Show 1
'''    End With
'''
''''''    MsgBox sMess, vbInformation, "LOG: " & format(Now, "MMMM D, YYYY")
'''
'''End Sub



Private Sub Form_Load()
    Dim sIniFile As String, sName As String, gsUser As String, sAddr As String, sDirPath As String
    Dim lRet As Long
    Dim istr1 As Integer, istr2 As Integer, i As Integer
    Dim Aspect As Double
    Dim IDFile As String, sCheck As String
    Dim NeedToAsk As Boolean
    Dim dWhen As Date
    Dim RetVal As Variant
'''    Dim ConnStr As String
    Dim s$, cnt&, dl&
    Dim sUser As String
    Dim strSelect As String, strDelete As String
    Dim rst As ADODB.Recordset
    Dim lErr As Long
    Dim sEnviron As String
    
'''    lErr = LockWindowUpdate(Me.hwnd)
    bDebug = False '' True
    bHideRN = False
    
    iIconSize = 1
    lIconX = 1600
    lIconY = 1200

    If bDebug Then MsgBox "Form_Load begin"

    sSupDocPath = "\\DETMSFS01\GPJAnnotator\Graphics\SupDocs\"
    strHTMLPath = "\\DETMSFS01\GPJAnnotator\Support\HTML\"
    
    cnt& = 199
    s$ = String$(200, 0)
    dl& = GetUserName(s$, cnt)
    sUser = Left$(s$, cnt)
    If Asc(Right(sUser, 1)) = 0 Then sUser = Left(sUser, Len(sUser) - 1)
    
    sOUser = sUser
    
    If bDebug Then MsgBox "sUser = " & sUser
    
    sUser = "GBUTEYN"
'    sUser = "JRomero"
'    sUser = "LHanna"
'    sUser = "DLing"
'    sUser = "Metaxp3"
'    sUser = "MLeblanc"
'    sUser = "RFunk"
'    sUser = "JBarlage"
'    sUser = "CAlberts"
'    sUser = "BRocha"
'    sUser = "MKaiser"
'    sUser = "DWickwar"
'    sUser = "LKoch"
'    sUser = "BPalm"
'    sUser = "DFord"
'    sUser = "ALee"
'    sUser = "CTXTSTAH"
'    sUser = "DBacon"
'    sUser = "MMears"
'    sUser = "AHeuser"
'    sUser = "NJackson"
'    sUser = "TDEV3"
'    sUser = "LBargen"
'    sUser = "SBOOTH"
'    sUser = "TDEV3"
'    sUser = "SEATON"
'    sUser = "NOCONNEL"
'    sUser = "SGINGELL"
'    sUser = "LMILLER"
'    sUser = "RADELMAN"
'    sUser = "SCREAL"
'    sUser = "MMETCALF"
'    sUser = "KEGAN"
'    sUser = "KBUCKLAN"
    
    
    iCurrSHYR = CInt(Format(Now, "YYYY"))
'''    If CInt(format(Now, "m")) < 5 Then
'''        iCurrSHYR = CInt(format(Now, "YYYY"))
'''    Else
'''        iCurrSHYR = CInt(format(Now, "YYYY")) + 1
'''    End If
    
'''    fraNewPassword.Left = fraLogin.Left
'''    fraNewPassword.Top = fraLogin.Top
'''    iLoginTrys = 0
    AppWindowState = Me.WindowState
    bAnnoOpen = False: bGfxOpen = False: bConstOpen = False
    bDo_Printer_Check = True
    
    
'    Me.BackColor = RGB(199, 211, 215)
    lColor = RGB(182, 182, 124)
'''    lColor2 = RGB(102, 153, 0)
    lGeo_Back = RGB(56, 88, 14) '' RGB(42, 66, 11) '' RGB(56, 88, 14) '' RGB(30, 30, 21)
    lGeo_Fore = RGB(72, 117, 17) '' RGB(84, 132, 21) '' RGB(153, 153, 51) ''RGB(84, 132, 21) '' RGB(111, 175, 28) '' RGB(100, 100, 68)
    lGeo_Bright = RGB(111, 175, 28)
    lGeo_Dark = RGB(42, 66, 11)
'''    lblExit.ForeColor = lGeo_Back '' lColor
'''    lblHelp.ForeColor = lGeo_Back '' lColor
    
'''    picFly.BackColor = lGeo_Back '' RGB(111, 175, 28)
    
'''    txtLogName.BackColor = lColor
'''    txtPassword.BackColor = lColor
'''    For i = 2 To lbl1.Count - 1
'''        lbl1(i).ForeColor = lColor
'''    Next i
'''    cmdGo.BackColor = lColor
'''    cmdResetPassword.BackColor = lColor
'''    txtExistingPassword.BackColor = lColor
'''    For i = 0 To 1
'''        txtNewPassword(i).BackColor = lColor
'''    Next i
'''    cmdCancel.BackColor = lColor
'''    cmdOK.BackColor = lColor
'    Shape1.Visible = True
    
'''    shpSpeed.BorderColor = lColor
'''    lblSpeed.ForeColor = lColor
'''    optSpeed(0).ForeColor = lColor
'''    optSpeed(1).ForeColor = lColor
'''
    For i = 0 To 6 '''5 '''4
        Label1(i).ForeColor = lGeo_Bright '' lGeo_Fore '' lGeo_Back ''lColor
    Next i
    
'''''    MsgBox "About to create ADO Object"
'''    Set Conn = New ADODB.Connection
'''''    MsgBox "ADO Object created"
    
    If bDebug Then MsgBox "About to Connect to Oracle"
    
'''''    ConnStr = "DSN=GPJProd;UID=System;PWD=ou812"
'''''    ConnStr = "DSN=JDE;UID=ANNOTATOR;PWD=ANNOTATOR"
    ConnStr = "DSN=JDETEST;UID=ANNOTATOR_APP_USER;PWD=q2eNqsgHxcKqre3"
    
'''''    MsgBox "About to try connecting (" & ConnStr & ")"
    Call OpenConn(ConnStr)
    
    If bDebug Then MsgBox "Conn open"
    
'''    Conn.Open (ConnStr)
'''    ConnOpen = True
    
'''''    MsgBox "Connection is open"
    
    '///// SET THESE FOR GRAPHIC USERS \\\\\
    iView = 1
    iRes = 2
    iRows = 4
    
''''''    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    On Error Resume Next
    
'''    Dim X As Printer
'''    i = 0
'''    For Each X In Printers
'''        i = i + 1
'''        MsgBox i & ".   " & X.DeviceName & "  [" & X.DriverName & "]"
'''    Next


'''    sPrinter = Printer.DeviceName
''''''    MsgBox sPrinter
'''    If Err Then sPrinter = ""
''''''    '\\\\\-------------------------------------------------------- /////
    
    bUserCheck = SetUser(sUser)
    
    If bUserCheck = False Then
        Unload Me
        Exit Sub
    End If
    
    
    iAppConn = 0
    If Dir(App.Path & "\Citrix.gpj", vbNormal) = "" Then
        ''SET MAIL FILE ID = SHORTNAME''
'        sNOTESID = Shortname
'        sNOTESPASSWORD = ""
'        Call SetLotusVars(Shortname, Me)
'
'        If strMailSrvr = "" Or strMailFile = "" Then
'            MsgBox "Will not be able to run the Annotator without Lotus Notes", vbCritical, "Sorry..."
'            Unload Me
'            Exit Sub
'        End If
        
        If Dir(App.Path & "\Thin-Client.gpj", vbNormal) <> "" Then
            iAppConn = 2
            bCitrix = False
        Else
        End If
    Else
        iAppConn = 1
        bCitrix = True
        ''SET MAIL FILE ID = "GANNOTAT"''
        sNOTESID = "GANNOTAT"
        sNOTESPASSWORD = "XPD8Notes"
        strMailSrvr = "Global_Links/IBM/GPJNotes"
        strMailFile = "mail\gannotat.nsf"
        
    End If
    
    
    
    Select Case iAppConn
        Case 0, 1 ''APP RUNNING LOCAL, OR ON CITRIX''
            i = 1
            Do
               sEnviron = Environ(i)
               If Left(sEnviron, 5) = "TEMP=" Then
                  sTempPath = Mid(Environ(i), 6)
                  Exit Do
               Else
                  i = i + 1
               End If
            Loop Until sEnviron = ""
            If sEnviron = "" Then
                If iAppConn = 0 Then
                    sTempPath = "C:\Temp"
                    sGIPath = "C"
                Else
                    sTempPath = "C:\Temp" ''"U:\Temp"
                    sGIPath = "C" ''"U"
                End If
                On Error Resume Next: Err = 0
                If Dir(sTempPath, vbDirectory) = "" Then MkDir (sTempPath)
                If Err Then
                    MsgBox "The GPJ Annotator cannot run as it is configured.  " & _
                                "The app must be able to create reports in a Temp folder, " & _
                                "but no Temp folder was found, nor was one able to be created.", _
                                vbCritical, "Sorry..."
                    Unload Me
                    Exit Sub
                End If
            Else
                Select Case iAppConn
                    Case 0, 2: sGIPath = "C"
                    Case 1: sGIPath = "C" ''"U"
                End Select
            End If
        
        Case 2: ''APP IS RUNNING ON THIN-CLIENT''
            sTempPath = "n:\Temp"
            sGIPath = "c"
            On Error Resume Next: Err = 0
            If Dir(sTempPath, vbDirectory) = "" Then MkDir sTempPath
            If Err Then
                MsgBox "The GPJ Annotator cannot run as it is configured.  " & _
                            "The app must be able to create reports in a Temp folder, " & _
                            "but no Temp folder was found, nor was one able to be created.", _
                            vbCritical, "Sorry..."
                Unload Me
                Exit Sub
            End If
            
    End Select
    sTempPath = sTempPath & "\"
    
'    MsgBox "Graphic Importer Path:  " & sGIPath & _
'                vbNewLine & _
'                "Temp File Path:  " & sTempPath
                
'    MsgBox "sGIPath = " & sGIPath & vbNewLine & _
'                "sTempPath = " & sTempPath & vbNewLine & _
'                "sNOTESID = " & sNOTESID & vbNewLine & _
'                "sNOTESPASSWORD = " & sNOTESPASSWORD
    
    
    '///// CREATE LINK STRING VARIABLE FOR ANNOTATOR ACCESS \\\\\
    '///// THIS IS USED IN ALL AUTOMATED EMAILS THAT POINT TO A POSTING \\\\\
'''    sLink = "For your convenience, you can use this link to launch the Annotator:  " & _
'''                "http://63.79.167.165:9980/GPJgeneric_icaclient.asp?TargetURL='GPJAnnotator.ica'"
    '' CHANGE 050102 LINK TO NEW DNS FOR CTXO2 ''
'''    sLink = "For your convenience, you can use this link to launch the Annotator:  " & _
'''                "http://appserver.gpjco.com:9980/GPJgeneric_icaclient.asp?TargetURL='GPJAnnotator.ica'"
    sLink = "For your convenience, you can use this link to launch the Annotator:  " & _
                "https://gpjapps.gpjco.com"
    sGLLink = "For your convenience, you can use this link to launch Global Links:  " & _
                "http://globallinks.gpjco.com"
    
    ''///// CHECK WHETHER TO SHOW WHAT'S NEW \\\\\''
    strSelect = "SELECT VALUE FROM ANNOTATOR.ANO_SESSION " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND APPID = 1002 " & _
                "AND SESSIONDESC = 'AnnoWhatsNew'"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close
        frmWhatsNew.PassCheck = True
        frmWhatsNew.Show 1
    Else
        rst.Close
    End If
    
    ''///// THIS IS A TEST... \\\\\''
    strSelect = "SELECT VALUE, SESSIONDESC FROM ANNOTATOR.ANO_SESSION " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND APPID = 1002 " & _
                "AND SESSIONDESC IN ('LinkToFloorplan', 'LinkToGraphic', 'LinkToDIL') " & _
                "ORDER BY SESSIONID DESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        bPassIn = True
        sPassInValue = Trim(rst.Fields("VALUE"))
        Select Case UCase(Trim(rst.Fields("SESSIONDESC")))
            Case "LINKTOFLOORPLAN": iPassIn = 1
            Case "LINKTOGRAPHIC": iPassIn = 2
            Case "LINKTODIL": iPassIn = 5
        End Select
    Else
        bPassIn = False
        iPassIn = 0
    End If
    rst.Close: Set rst = Nothing
    If bPassIn Then
        ''DELETE RECORD''
        strDelete = "DELETE FROM ANNOTATOR.ANO_SESSION " & _
                    "WHERE USER_SEQ_ID = " & UserID & " " & _
                    "AND APPID = 1002 " & _
                    "AND SESSIONDESC IN ('LinkToFloorplan', 'LinkToGraphic', 'LinkToDIL')"
        Conn.Execute (strDelete)
        
        Call Label1_Click(iPassIn)
    End If
    
    
'''    iIconSize = 1
'''    lIconX = 1600
'''    lIconY = 1200
    
''''''''''    bWatching = True ''SHOULD THIS ONLY BE SET IF PASSING IN?''
    
'''    If bPerm(61) Then
'''        With cboViewUsage
'''            .AddItem "<Gfx Posting Report>"
'''            .ItemData(.NewIndex) = 0
'''            .AddItem "Posted Gfx - Today"
'''            .ItemData(.NewIndex) = 12
'''            .AddItem "Posted Gfx - Yesterday"
'''            .ItemData(.NewIndex) = 13
'''            .AddItem "Posted Gfx - 2 Days Ago"
'''            .ItemData(.NewIndex) = 14
'''            .AddItem "Posted Gfx - 3 Days Ago"
'''            .ItemData(.NewIndex) = 15
'''            .AddItem "Posted Gfx - For Week"
'''            .ItemData(.NewIndex) = 16
'''            .AddItem "Posted Gfx - For Month"
'''            .ItemData(.NewIndex) = 17
'''
'''            .ListIndex = 0
'''        End With
'''    End If
    
    
'''    lErr = LockWindowUpdate(0)
Exit Sub
NoNotes:
'''    lErr = LockWindowUpdate(0)
    Unload Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    picHdr.Width = Me.ScaleWidth
    shpHDR.Width = Me.ScaleWidth
    lblExit.Left = Me.ScaleWidth - 300 - lblExit.Width
    lblHelp.Left = lblExit.Left
    
    lblWelcome.Left = Me.ScaleWidth - 300 - lblWelcome.Width
    lblWelcome.Top = Me.ScaleHeight - 300 - lblWelcome.Height
    
End Sub

'''Private Sub Form_Resize()
'''    imgClose.Left = Me.ScaleWidth - imgClose.Width
'''    lblHelp.Left = imgClose.Left + (imgClose.Width / 2) - (lblHelp.Width / 2)
'''    lblExit.Left = imgClose.Left + (imgClose.Width / 2) - (lblExit.Width / 2)
'''End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strUpdate As String
    
    If UCase(Shortname) = UCase(sOUser) Then
        strUpdate = "UPDATE " & ANOLockLog & " " & _
                    "SET LOCKCLOSEDTTM = SYSDATE, " & _
                    "LOCKSTATUS = 0, " & _
                    "UPDDTTM = SYSDATE, " & _
                    "UPDCNT = UPDCNT + 1 " & _
                    "WHERE LOCKID = " & lOpenID
        Conn.Execute (strUpdate)
    End If
    
    If ConnOpen Then
        Conn.Close
        Set Conn = Nothing
        ConnOpen = False
    End If
End Sub

Private Sub Image3_Click()

End Sub

'''Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblHelp.ForeColor = lGeo_Back '' vbWhite
'''    lblExit.ForeColor = lGeo_Back '' vbWhite
'''End Sub

Private Sub imgEmailTeam_Click()
    picFly.Visible = False
    lblWelcome.Visible = False
'''    lblWelcome = "George P. Johnson Annotator Menu"
    frmEmailTeam.Show 1
End Sub

Private Sub imgOptions_Click()
    picFly.Visible = False
    Me.PopupMenu mnuOptions
'''    frmPassword.Show 1
End Sub

Private Sub imgUserReports_Click()
    picFly.Visible = False
    frmUsageReports.Show 1, Me
End Sub

Private Sub lblHelp_Click()
    lblHelp.ForeColor = vbBlack ''vbWhite '' lColor
    frmHelp.Show 1
End Sub

Private Sub imgSecurity_Click()
    picFly.Visible = False
    lblWelcome.Visible = False
'''    lblWelcome = "George P. Johnson Annotator Menu"
    frmSecurity.Show 1
End Sub

''''''Private Sub lblBack2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''''    Dim i As Integer
''''''
''''''    If iLabel <> -1 Then
''''''        Label1(iLabel).ForeColor = lColor
''''''        iLabel = -1
''''''    End If
''''''End Sub

Private Sub Label1_Click(Index As Integer)
    Dim RetVal
    Dim sTFile As String
    
'''    Screen.MousePointer = 11
    lblWelcome.Visible = False
    Select Case Index
        Case 0
            frmShow.Show 1
'''            ShellExecute 0&, vbNullString, "http://igl/menu.asp", vbNullString, _
'''                        vbNullString, 1
        Case 1
'''            lblMess = "...Populating Client List with available floorplans"
'''            lblMess.Refresh
            frmAnnotator.Show 1
'''            lblMess = ""
        Case 2
'            If bPerm(24) Or bPerm(25) Then frmGraphics.Show Else frmGraphics.Show 1
            frmGraphics.Show 1
'''            frmGraphics.Show
        Case 3: frmConst.Show 1
        Case 4
'''            sTFile = "M:\Temp\Exporter.gpj"
            sTFile = sGIPath & ":\Program Files\GraphicExporter\Exporter.gpj"
'''            MsgBox "sTFile = " & sTFile
'''            sTFile = "C:\Temp\Exporter.gpj"
            Open sTFile For Output As #1
            Write #1, Shortname, LogAddress
            Close #1
'''            RetVal = Shell("C:\Program Files\GraphicExporter\GraphicExporter.exe", 1)
'''            RetVal = Shell("M:\Program Files\GraphicExporter\GraphicExporter.exe", 1)
'''            MsgBox "sGIFile = " & sGIPath & ":\Program Files\GraphicExporter\GraphicExporter.exe"
            RetVal = Shell(sGIPath & ":\Program Files\GraphicExporter\GraphicExporter.exe", 1)
'''''            RetVal = Shell("D:\Data\VB Projects\GPJAnnotator\GraphicExporter\GraphicExporter.exe", 1)
        Case 5: frmDIL.Show 1
        Case 6
            MsgBox "The Facilities interface is brand new.  At this time, " & _
                        "our Facilities have not yet been imported.  The sample files " & _
                        "available for viewing are development samples and do not reflect " & _
                        "actual Facilities." & vbCr & vbCr & _
                        "Actual Facilities - along with hyperlinked on-site photos - will begin populating " & _
                        "soon.  In the mean time, feel free to browse the samples to " & _
                        "get a feeling how the interface will work.", vbInformation, "TAKE NOTE..."
            frmFacil.Show 1
    End Select
'''    Screen.MousePointer = 0
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If iLabel <> Index And iLabel <> -1 Then Label1(iLabel).ForeColor = lGeo_Bright '' lGeo_Fore ''lGeo_Back ''vbWhite '' lColor
    iLabel = Index
    Label1(Index).ForeColor = vbBlack '' vbWhite ''lGeo_Bright  ''vbWhite ''vbYellow ''vbBlack
'''''''    If Shape1.Visible = False Then Shape1.Visible = True
End Sub

'''''''Private Sub lblback2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''''    Shape1.Visible = False
'''''''End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

'''Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblExit.ForeColor = vbWhite
'''End Sub

'''Private Sub lblExitBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblExit.ForeColor = lGeo_Back '' lColor
'''    lblHelp.ForeColor = lGeo_Back '' lColor
'''End Sub

'''Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblHelp.ForeColor = vbWhite ''vbYellow
'''
'''End Sub

Private Sub lblOptions_Click(Index As Integer)
    Select Case Index
        Case 0: picFly.Visible = True
        Case 1: picFly.Visible = False
    End Select
End Sub

''Private Sub mnuPassword_Click()
''    frmPassword.Show 1
''End Sub

Private Sub mnuPrinterDrivers_Click()
    frmPrinterDrivers.Show 1
End Sub

Private Sub mnuSpeed_Click(Index As Integer)
    iView = Index
    mnuSpeed(Index).Checked = True
    mnuSpeed(Abs(Index - 1)).Checked = False
    Debug.Print "iView = " & iView
End Sub

'''Private Sub txtNewPassword_Change(Index As Integer)
'''    If txtNewPassword(0).Text = txtNewPassword(1).Text And txtNewPassword(0).Text <> "" Then
'''        cmdOK.Enabled = True
'''        cmdOK.Default = True
'''    Else
'''        cmdOK.Enabled = False
'''    End If
'''End Sub

Public Sub SetPermissionBools(sTotal As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iCnt As Integer, i As Integer ''', iInt As Integer
    Dim subT As String
    
    subT = sTotal
    
    strSelect = "SELECT COUNT(*) AS PERMCNT FROM ANNOTATOR.ANO_PERM"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        iCnt = CInt(rst.Fields("PERMCNT")) - 1
        ReDim bPerm(iCnt)
    Else
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
    rst.Close
    Set rst = Nothing
    
    For i = iCnt To 0 Step -1
        strSelect = "SELECT TO_CHAR(MOD(" & sTotal & ", POWER(2, " & i & ")), '" & String(38, "9") & "') FROM DUAL"
        Set rst = Conn.Execute(strSelect)
        subT = Trim(rst.Fields(0))
        If subT = sTotal Then
            bPerm(i) = False
        Else
            bPerm(i) = True
            sTotal = subT
        End If
'        Debug.Print "bPerm(" & i & ") = " & bPerm(i)
    Next i
    
'''    bPerm(39) = False ''TEMP FOR TESTING''
    
'''    For i = iCnt To 0 Step -1
'''        iInt = Int(subT / (2 ^ i))
'''        bPerm(i) = CBool(iInt * -1)
'''        Debug.Print "bPerm(" & i & ") = " & bPerm(i)
'''        subT = subT - (iInt * (2 ^ i))
'''    Next i
    
End Sub

Public Sub GetClientPrivilege(UID As Long)
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    
    '///// DETERMINE IF USER HAS ACCESS TO ALL CLIENTS \\\\\
    strSelect = "SELECT CUNO_GROUP_ID " & _
                "FROM " & IGLUserCR & " " & _
                "WHERE USER_SEQ_ID = " & UID
    Set rst = Conn.Execute(strSelect)
    bClientAll_Enabled = False
    If Not rst.EOF Then
        If CInt(rst.Fields("CUNO_GROUP_ID")) = -1 Then bClientAll_Enabled = True
    End If
    rst.Close
    Set rst = Nothing
    
    Debug.Print "bClientAll_Enabled = " & bClientAll_Enabled
    If bClientAll_Enabled Then Exit Sub
    
    '///// IF NOT ALL, GET CLIENT LIST \\\\\
    strSelect = "SELECT AN8_CUNO FROM " & IGLUserCR & " " & _
                "WHERE USER_SEQ_ID = " & UID & " " & _
                "AND CUNO_GROUP_ID = 0 " & _
                "UNION " & _
                "SELECT GR.AN8_CUNO " & _
                "FROM " & IGLUserCR & " CR, " & IGLCGR & " GR " & _
                "WHERE CR.USER_SEQ_ID = " & UID & " " & _
                "AND CR.CUNO_GROUP_ID = GR.CUNO_GROUP_ID " & _
                "ORDER BY AN8_CUNO"
    Set rst = Conn.Execute(strSelect)
    strCunoList = ""
    If Not rst.EOF Then
        strCunoList = CStr(rst.Fields("AN8_CUNO"))
        rst.MoveNext
        Do While Not rst.EOF
            strCunoList = strCunoList & ", " & CStr(rst.Fields("AN8_CUNO"))
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
        
    Debug.Print strCunoList
End Sub

Public Function GetGFXReviewer(lID As Long) As Boolean
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
'''    strSelect = "SELECT TEAM_ID " & _
'''                "FROM " & ANOETeamUR & " " & _
'''                "WHERE USER_SEQ_ID = " & lID & " " & _
'''                "AND RECIPIENT_FLAG1 = 1"
    strSelect = "SELECT TEAM_ID " & _
                "FROM " & ANOETeamUR & " " & _
                "WHERE USER_SEQ_ID = " & lID & " " & _
                "AND (RECIPIENT_FLAG1 = 1 " & _
                "OR EXTCLIENTAPPROVER_FLAG = 1)"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        GetGFXReviewer = True
    Else
        rst.Close
        ''CHECK IF GRAPHIC APPROVER FOR SPECIFIC GRAPHIC''
        If bGPJ Then
            strSelect = "SELECT DISTINCT AN8_CUNO " & _
                        "FROM " & GFXMas & " " & _
                        "WHERE GID > 0 " & _
                        "AND GSTATUS IN (10, 20, 27) " & _
                        "AND GAPPROVER_ID = " & lID
        Else
            strSelect = "SELECT DISTINCT AN8_CUNO " & _
                        "FROM " & GFXMas & " " & _
                        "WHERE GID > 0 " & _
                        "AND GSTATUS IN (20, 27) " & _
                        "AND GAPPROVER_ID = " & lID
        End If
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            GetGFXReviewer = True
        Else
            GetGFXReviewer = False
        End If
    End If
    rst.Close: Set rst = Nothing
    
    If bPerm(71) Then GetGFXReviewer = True
End Function



Public Function SetUser(sUser As String) As Boolean
    Dim i As Integer, iVis As Integer
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lFirst As Long
    
    
    bICAUser = False
    strSelect = "SELECT R.PCODE, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, U.USER_SEQ_ID, " & _
                "TO_CHAR(UT.USERTYPEVALUE, '" & String(38, "9") & "') AS PERM, U.EMPLOYER, UT.USERTYPE " & _
                "FROM " & IGLUserAR & " R, " & IGLUser & " U, " & ANOUserType & " UT " & _
                "WHERE U.NAME_LOGON = '" & UCase(Trim(sUser)) & "' " & _
                "AND U.USER_STATUS > 0 " & _
                "AND U.USER_SEQ_ID = R.USER_SEQ_ID " & _
                "AND R.APP_ID = 1002 " & _
                "AND R.PERMISSION_STATUS > 0 " & _
                "AND R.USER_PERMISSION_ID = UT.USERTYPEID"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Shortname = UCase(sUser)
        LogName = StrConv(Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")), vbProperCase)
        LogFirstName = StrConv(Trim(rst.Fields("NAME_FIRST")), vbProperCase)
        LogAddress = Trim(rst.Fields("EMAIL_ADDRESS"))
        UserID = rst.Fields("USER_SEQ_ID")
        UserType = Trim(rst.Fields("USERTYPE"))
        If UCase(Left(rst.Fields("EMPLOYER"), 3)) = "GPJ" Then
            bGPJ = True
        Else
            bGPJ = False
            If UCase(Left(rst.Fields("EMPLOYER"), 9)) = "COMPUWARE" Then bICAUser = True
        End If
        
        
        
        '///// TEMP FIX FOR CITRIX TEST \\\\\
'''        NotesServer = "detsrv1/det"
'''        If UCase(NotesServer) = "detsrv1/det" Then bGPJNotes = True Else bGPJNotes = False
'''        bBos = False

        ''SET LINK DISCLAIMER''
        sLink_Disclaimer = "NOTE: Any direct access hyperlinks contained in this email " & _
                    "have been specifically created for access by the original email recipient only.  " & _
                    "Please, do not forward this document.  If you have been " & _
                    "forwarded this document, the link will not function properly for you, " & _
                    "unless you are able to log in as the original email recipient."
                
        Debug.Print "Permission:  " & Trim(rst.Fields("PERM"))
        Call SetPermissionBools(Trim(rst.Fields("PERM")))
        
        rst.Close: Set rst = Nothing
        
        ''CHECK FOR CUNO PREFERENCE''
        strSelect = "SELECT UP.AN8_CUNO AS CUNO, AB.ABALPH AS FBCN " & _
                    "FROM " & ANOUPref & " UP, " & F0101 & " AB " & _
                    "WHERE UP.USER_SEQ_ID > 0 " & _
                    "AND UP.USER_SEQ_ID = " & UserID & " " & _
                    "AND UP.APP_ID = 1002 " & _
                    "AND UP.AN8_CUNO = AB.ABAN8"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            defCUNO = rst.Fields("CUNO")
            defFBCN = Trim(rst.Fields("FBCN"))
        Else
            defCUNO = 0
            defFBCN = ""
        End If
        rst.Close: Set rst = Nothing
        
        '///// TEMPORARY PRINTER DISABLE SOLUTION 11-SEP-2001 \\\\\
'''''        bENABLE_PRINTERS = False
'        bENABLE_PRINTERS = True
'''                If ShortName = "SWESTERH" Then bENABLE_PRINTERS = True
        '\\\\\ ---------------------------------------------------- /////
        
                
        iLabel = 0
'''        Shape1.Visible = False
'''        lblBack.Enabled = True
'''        lblback2.Enabled = True
'        Shape1.Visible = True
        
        lblWelcome = "Welcome " & LogFirstName & "..."
'        lblHelp.Visible = True
        
        If bPerm(0) Then imgEmailTeam.Visible = True Else imgEmailTeam.Visible = False
        If bPerm(41) Then imgSecurity.Visible = True Else imgSecurity.Visible = False
        lFirst = 3000 ''2400
        iVis = 0
        For i = 0 To 6
            Select Case i
                Case 0
                    If bPerm(48) Then
                        Label1(0).Top = lFirst + (iVis * 600)
                        Label1(0).Visible = True
                        iVis = iVis + 1
                    End If
                Case 1
                    If bPerm(1) Then
                        Label1(1).Top = lFirst + (iVis * 600)
                        Label1(1).Visible = True
                        iVis = iVis + 1
                    End If
                Case 2
''                    If bPerm(23) Then
''                        Label1(2).Top = lFirst + (iVis * 600)
''                        Label1(2).Visible = True
''                        iVis = iVis + 1
''                    End If
                Case 3
                    If bPerm(32) Then
                        Label1(3).Top = lFirst + (iVis * 600)
                        Label1(3).Visible = True
                        iVis = iVis + 1
                    End If
                Case 4
''                    If bPerm(31) Then
''                        Label1(4).Top = lFirst + (iVis * 600)
''                        Label1(4).Visible = True
''                        iVis = iVis + 1
''                    End If
'''                Case 5: If bPerm(57) Then Label1(5).Visible = True
                Case 5
''                    Label1(5).Visible = False
'''                    If bGPJ Then
'''                        Label1(5).Top = lFirst + (iVis * 600)
'''                        Label1(5).Visible = True
'''                        iVis = iVis + 1
'''                    End If
                Case 6
''                    Label1(6).Top = 2400
''                    Label1(6).Visible = True
            End Select
        Next i
        
'''        If bPerm(61) Then cboViewUsage.Visible = True Else cboViewUsage.Visible = False
        If bPerm(62) Then imgUserReports.Visible = True Else imgUserReports.Visible = False
        
        '///// GET CLIENT PRIVILEGE \\\\\
        Call GetClientPrivilege(UserID)
        
        '///// GET GFX REVIEWER STATUS \\\\\
        bGFXReviewer = GetGFXReviewer(UserID)
        
'''        rst.Close: Set rst = Nothing
        
        ''CLEANUP ON ANO_LOCKLOG''
        If UCase(Shortname) = "SWESTERH" Then
'            Call CleanUpAnnoLog
        End If
        
        If UCase(Shortname) = UCase(sOUser) Then
            '///// WRITE ANNO_OPEN TO ANO_LOCKLOG \\\\\'
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lOpenID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            strInsert = "INSERT INTO " & ANOLockLog & " " & _
                        "(LOCKID, LOCKREFID, LOCKREFSOURCE, USER_SEQ_ID, " & _
                        "LOCKOPENDTTM, LOCKSTATUS, ADDUSER, ADDDTTM, " & _
                        "UPDUSER, UPDDTTM, UPDCNT) VALUES " & _
                        "(" & lOpenID & ", 1002, 'ANNO_OPEN', " & UserID & ", " & _
                        "SYSDATE, 1, '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
        End If
        
        SetUser = True
    Else
        rst.Close: Set rst = Nothing
        MsgBox UCase(sUser) & " is not recognized as an approved GPJ Annotator user." & _
                    vbCr & vbCr & "Verify you have entered the Login Name correctly.", vbExclamation, _
                    "Unrecognized Login Name..."
        
        SetUser = False
    End If
    
End Function

Private Sub mnuWhatsNew_Click()
    frmWhatsNew.PassCheck = False
    frmWhatsNew.Show 1
End Sub

