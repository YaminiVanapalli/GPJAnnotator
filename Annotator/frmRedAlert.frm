VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRedAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Annotation Notification..."
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRedAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTeamEmail 
      Height          =   255
      Index           =   1
      Left            =   7140
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstShortname 
      Height          =   255
      Index           =   1
      Left            =   7140
      TabIndex        =   9
      Top             =   2580
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstShortname 
      Height          =   255
      Index           =   0
      Left            =   7140
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtComment 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2340
      Width           =   3195
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   3900
      Width           =   1095
   End
   Begin VB.ListBox lstTeamEmail 
      Height          =   255
      Index           =   0
      Left            =   7140
      TabIndex        =   0
      Top             =   2100
      Visible         =   0   'False
      Width           =   135
   End
   Begin TabDlg.SSTab sstEmail 
      Height          =   4215
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3175
      TabCaption(0)   =   "Team Members"
      TabPicture(0)   =   "frmRedAlert.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTeam"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "GPJ Personnel"
      TabPicture(1)   =   "frmRedAlert.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstGPJ"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstGPJ 
         Height          =   3660
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   420
         Width           =   3435
      End
      Begin VB.ListBox lstTeam 
         Height          =   3660
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   3435
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2100
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblHdr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblHdr"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3195
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRedAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''iType As Integer, sHDR As String, tmpBCC As String, tmpSHCD As Long
Dim iType As Integer, pSHYR As Integer
Dim lBCC As Long, lSHCD As Long
Dim sHDR As String
Dim bSent As Boolean
Dim pGID As Long


Public Property Get PassBCC() As Long
    PassBCC = lBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    lBCC = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = sHDR
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    sHDR = vNewValue
End Property

Public Property Get PassType() As Integer
    PassType = iType
End Property
Public Property Let PassType(ByVal vNewValue As Integer)
    iType = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = lSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    lSHCD = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = pSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    pSHYR = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property


Private Sub cmdOK_Click()
    Dim MessBody As String, MessHdr As String
    Dim i As Integer, iAdd As Integer, iVal As Integer
    Dim tYear As String, tClient As String, tShow As String, sWho As String, _
                sEMail As String, sList As String, sInternet As String
    Dim iStr As Integer, iEnd As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sColumn As String
    Dim sPreLink As String, sPostLink As String
    
     '///// EXECUTE E-MAIL \\\\\
'''''    Dim myNotes As New Domino.NotesSession
'''''    Dim myDB As New Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    
    Dim Address() As String
    Dim Shortname() As String
    
    Dim MailMan As New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
    
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
    
    Dim Email As New ChilkatEmail2
    
    
    
'    sPreLink = "http://gpjapps02.gpjco.com/LinksToAnno.asp?name_logon="
'    Select Case iType
'        Case 0: sPostLink = "&shyr=" & pSHYR & "&shcd=" & lSHCD & "&cuno=" & lBCC
'        Case 1: sPostLink = "&gid=" & pGID
'        Case 2: sPostLink = "" ''COMING SOON''
'    End Select
    
    iAdd = -1: sList = ""
    For i = 0 To lstTeam.ListCount - 1
        If lstTeam.Selected(i) = True Then
            iAdd = iAdd + 1
            ReDim Preserve Address(iAdd)
            Address(iAdd) = lstTeamEmail(0).List(i)
            
            Email.AddTo Address(iAdd), Address(iAdd)
            
            ReDim Preserve Shortname(iAdd)
            Shortname(iAdd) = lstShortname(0).List(i)
            
            sList = sList & vbTab & lstTeam.List(i) & vbNewLine
            
        End If

    Next i
    If sstEmail.TabVisible(1) Then
        For i = 0 To lstGPJ.ListCount - 1
            If lstGPJ.Selected(i) = True Then
                iAdd = iAdd + 1
                ReDim Preserve Address(iAdd)
                Address(iAdd) = lstTeamEmail(1).List(i)
                
                Email.AddTo Address(iAdd), Address(iAdd)
                
                ReDim Preserve Shortname(iAdd)
                Shortname(iAdd) = lstShortname(1).List(i)
                
                sList = sList & vbTab & lstGPJ.List(i) & vbNewLine
                
            End If
        Next i
    End If
    
    
    If iAdd = -1 Then GoTo GetOut
    
    MessHdr = sHDR & " Redlines"
    MessBody = "Redline Annotations have been drawn and saved by " & _
                LogName & " for " & sHDR & ".  "
    If iType = 2 Then MessBody = MessBody & "Drawing file:  " & frmConst.sOpenFile
    MessBody = MessBody & vbNewLine & vbNewLine & _
                "The following Team members are being alerted through this email:" & _
                vbNewLine & vbNewLine & sList & vbNewLine
    If txtComment.Text <> "" Then
        MessBody = MessBody & vbNewLine & "Comment from Redliner (" & LogName & "):" & vbNewLine & _
                    String(75, "=") & vbNewLine & vbNewLine & _
                    txtComment.Text & vbNewLine & vbNewLine & String(75, "=") & vbNewLine
    End If
        
    Email.subject = MessHdr
    Email.Body = MessBody
    Email.FromAddress = LogAddress
    Email.fromName = LogName
    
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
    
    
'    If Not bCitrix Then
'        ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
'        Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
'        Dim myDB As Object '' LOTUS.NotesDatabase
'
'
'        On Error Resume Next
'        Set myNotes = GetObject(, "Notes.NotesSession")
'
'        If Err Then
'            Err.Clear
'            Set myNotes = CreateObject("Notes.NotesSession")
'            If Err Then
'                MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
'                GoTo GetOut
'            End If
'        End If
'        On Error GoTo 0
'        Set myDB = myNotes.GetDatabase("", "")
'        myDB.OPENMAIL
'
'    Else
'        ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
'        Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
'        Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
'
'
'        myDom.Initialize (sGAnnoPW)
'        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'
'    End If
'
'    On Error Resume Next
'    For i = 0 To iAdd
'        If Not bCitrix Then
'            Set myDoc = myDB.CreateDocument
'        Else
'            Set myDoc = myDomDB.CreateDocument
'            Call myDoc.ReplaceItemValue("Principal", LogName)
'            Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'        End If
'
'
''''''        Set myDoc = myDB.CreateDocument
''''''        If sNOTESID = "GANNOTAT" Then Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
''''''        If sNOTESID = "GANNOTAT" Then Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'
'        Set myRichText = myDoc.CreateRichTextItem("Body")
'        Select Case iType
'            Case 0, 1
'                myRichText.AppendText MessBody ''& vbNewLine & _
''                            "Click for direct access to file:  " & _
''                            sPreLink & LCase(Shortname(i)) & sPostLink & _
''                            vbNewLine & vbNewLine & sLink_Disclaimer
'            Case 2
'                myRichText.AppendText MessBody & vbNewLine
'        End Select
'        myDoc.AppendItemValue "SENDTO", Address(i)
'''        myDoc.SaveMessageOnSend = True
'
'        Call myDoc.Send(False, Address(i))
'
'        Set myReply = Nothing
'        Set myRichText = Nothing
'        Set myItem = Nothing
'        Set myDoc = Nothing
'    Next i
'    If Err Then
'        MsgBox "An error has occurred while attempting to send out the automated " & _
'                    "notification, possibly resulting in no one being alerted of your Annotation.  " & _
'                    "Please, contact the GPJ Help Desk if this persists." & vbCr & vbCr & _
'                    "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                    vbExclamation, "Unable to send out Notification..."
'        Err = 0
'
'    Else
'
'        If bCitrix Then
'            If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'            If Not myDom Is Nothing Then Set myDom = Nothing
'        Else
'            If Not myDB Is Nothing Then Set myDB = Nothing
'            If Not myNotes Is Nothing Then Set myNotes = Nothing
'        End If
'
'    End If
    
GetOut:
    bSent = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSelect As String, sEmployer As String
    Dim rst As ADODB.Recordset
    Dim sColumn As String
    
    bSent = False
    
    lblHdr.Caption = "A notification will be sent out to those mandatory recipients " & _
                "checked on the list at right.  If you wish to alert any other Team " & _
                "members of your Annotations, check their name as well." & vbNewLine & _
                vbNewLine & _
                "If you would like to include a comment with the automated notification, " & _
                "enter it in the textbox below."
    
    '///// FIRST, GET TEAM \\\\\
    '///// SEE IF CLIENT-SHOW TEAM EXISTS \\\\\
    Select Case iType
        Case 0: sColumn = "RECIPIENT_FLAG0"
        Case 1: sColumn = "RECIPIENT_FLAG1"
        Case 2: sColumn = "RECIPIENT_FLAG2"
    End Select
    
    strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, U.EMPLOYER, " & _
                "R.RECIPIENT_FLAG0, R.RECIPIENT_FLAG1, R.RECIPIENT_FLAG2 " & _
                "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                "WHERE T.AN8_CUNO = " & lBCC & " " & _
                "AND T.AN8_SHCD = " & lSHCD & " " & _
                "AND T.MCU IS NULL " & _
                "AND T.TEAM_ID = R.TEAM_ID " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, U.EMPLOYER, " & _
                    "R.RECIPIENT_FLAG0, R.RECIPIENT_FLAG1, R.RECIPIENT_FLAG2 " & _
                    "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                    "WHERE T.AN8_CUNO = " & lBCC & " " & _
                    "AND T.AN8_SHCD IS NULL " & _
                    "AND T.MCU IS NULL " & _
                    "AND T.TEAM_ID = R.TEAM_ID " & _
                    "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                    "AND U.USER_STATUS > 0 " & _
                    "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Set rst = Conn.Execute(strSelect)
    End If
    lstTeam.Clear: lstTeamEmail(0).Clear: lstShortname(0).Clear
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) <> "GPJ" Then
            sEmployer = " (" & Trim(rst.Fields("EMPLOYER")) & ")"
        Else
            sEmployer = ""
        End If
        lstTeam.AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & _
                    UCase(Trim(rst.Fields("NAME_LAST"))) & sEmployer
        lstTeam.ItemData(lstTeam.NewIndex) = rst.Fields(sColumn)
        lstTeam.Selected(lstTeam.NewIndex) = CBool(rst.Fields(sColumn) * -1)
        lstTeamEmail(0).AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstShortname(0).AddItem LCase(Trim(rst.Fields("NAME_LOGON")))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    Call UnselectSelf(LogName, Me.lstTeam)
    
    '///// GET GPJ PERSONNEL \\\\\
    If bGPJ Then
        sstEmail.TabVisible(1) = True
        Call GetGPJEmail
    Else
        sstEmail.TabVisible(1) = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bSent Then
'''        Cancel = 1
        cmdOK_Click
    End If
End Sub

Private Sub lstTeam_ItemCheck(Item As Integer)
    Dim i As Integer
    Debug.Print lstTeam.Selected(Item)
    If lstTeam.ItemData(Item) = 1 And lstTeam.Selected(Item) = False Then
        lstTeam.Selected(Item) = True
        MsgBox "The person selected is a Mandatory Recipient.", _
                    vbExclamation, "Member cannot be Unchecked..."
    End If
End Sub

Public Sub GetGPJEmail()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    lstGPJ.Clear: lstTeamEmail(1).Clear: lstShortname(1).Clear
    '///// POP GPJ EMPLOYEE LIST \\\\\
    strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LOGON, " & _
                "U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
                "FROM " & IGLUser & " U, " & IGLUserAR & " R " & _
                "WHERE U.USER_STATUS = 1 " & _
                "AND SUBSTR(U.EMPLOYER, 1, 3) = 'GPJ' " & _
                "AND U.EMAIL_ADDRESS IS NOT NULL " & _
                "AND U.USER_SEQ_ID = R.USER_SEQ_ID " & _
                "AND R.APP_ID = 1002 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstGPJ.AddItem Trim(rst.Fields("NAME_FIRST")) & " " & _
                    Trim(rst.Fields("NAME_LAST"))
        lstTeamEmail(1).AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstShortname(1).AddItem LCase(Trim(rst.Fields("NAME_LOGON")))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub


