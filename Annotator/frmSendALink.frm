VERSION 5.00
Begin VB.Form frmSendALink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send-A-Link..."
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendALink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstUserNames 
      Height          =   2085
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5955
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   315
   End
   Begin VB.Frame fraEmail 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   540
         Width           =   5955
      End
      Begin VB.TextBox txtMess 
         Height          =   1635
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1200
         Width           =   6675
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   5655
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send Email"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4740
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   690
      End
   End
   Begin VB.ListBox lstShortname 
      Height          =   1815
      ItemData        =   "frmSendALink.frx":08CA
      Left            =   780
      List            =   "frmSendALink.frx":08D1
      TabIndex        =   2
      Top             =   4200
      Width           =   555
   End
   Begin VB.ListBox lstUsers 
      Height          =   1815
      ItemData        =   "frmSendALink.frx":08E3
      Left            =   6240
      List            =   "frmSendALink.frx":08E5
      TabIndex        =   1
      Top             =   4260
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox lstEmail 
      Height          =   1815
      ItemData        =   "frmSendALink.frx":08E7
      Left            =   120
      List            =   "frmSendALink.frx":08E9
      TabIndex        =   0
      Top             =   4200
      Width           =   555
   End
End
Attribute VB_Name = "frmSendALink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sEMail() As String
Dim sUsers() As String
Dim sShortnames() As String

Dim pFrom As String, pSub As String
Dim pBCC As Long, pGID As Long, pSHCD As Long
Dim pSHYR As Integer

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property

Public Property Get PassSub() As String
    PassSub = pSub
End Property
Public Property Let PassSub(ByVal vNewValue As String)
    pSub = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = pSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    pSHYR = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = pSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    pSHCD = vNewValue
End Property




Public Sub GetUsers(tFrom As String, lBCC As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    lstUserNames.Clear
    lstUsers.Clear
    lstEmail.Clear
    Select Case tFrom
        Case "DIL"
            strSelect = "SELECT U.USER_SEQ_ID, U.EMAIL_ADDRESS, U.EMPLOYER, " & _
                        "(TRIM(U.NAME_LAST)||', '||TRIM(U.NAME_FIRST)) AS LISTNAME, " & _
                        "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS USERNAME " & _
                        "FROM IGLPROD.IGL_USER U, IGLPROD.IGL_USER_APP_R A " & _
                        "Where U.USER_SEQ_ID > 0 " & _
                        "AND U.USER_STATUS > 0 " & _
                        "AND UPPER(SUBSTR(U.EMPLOYER, 1, 3)) = 'GPJ' " & _
                        "AND U.USER_SEQ_ID = A.USER_SEQ_ID " & _
                        "AND A.APP_ID = 1002 " & _
                        "AND A.PERMISSION_STATUS > 0 " & _
                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Case "GH", "FP"
            strSelect = "SELECT USER_SEQ_ID, EMAIL_ADDRESS, EMPLOYER, " & _
                        "(TRIM(NAME_LAST)||', '||TRIM(NAME_FIRST)) AS LISTNAME, " & _
                        "(TRIM(NAME_FIRST)||' '||TRIM(NAME_LAST)) AS USERNAME " & _
                        "From IGLPROD.IGL_USER WHERE USER_SEQ_ID IN " & _
                        "(SELECT U.USER_SEQ_ID " & _
                        "FROM IGLPROD.IGL_USER U, IGLPROD.IGL_USER_APP_R A Where U.USER_SEQ_ID > 0 " & _
                        "AND U.USER_STATUS > 0 " & _
                        "AND UPPER(SUBSTR(U.EMPLOYER, 1, 3)) = 'GPJ' " & _
                        "AND U.USER_SEQ_ID = A.USER_SEQ_ID " & _
                        "AND A.APP_ID = 1002 " & _
                        "AND A.PERMISSION_STATUS > 0 "
            strSelect = strSelect & _
                        "Union " & _
                        "(SELECT IU.USER_SEQ_ID " & _
                        "FROM IGLPROD.IGL_USER_CUNO_R IUCR, IGLPROD.IGL_USER IU " & _
                        "Where IUCR.AN8_CUNO = " & lBCC & " " & _
                        "AND IUCR.USER_SEQ_ID = IU.USER_SEQ_ID " & _
                        "AND IU.USER_SEQ_ID > 0 " & _
                        "AND IU.USER_STATUS > 0 " & _
                        "AND IU.EMAIL_ADDRESS IS NOT NULL "
            strSelect = strSelect & _
                        "Union " & _
                        "SELECT IU.USER_SEQ_ID " & _
                        "FROM IGLPROD.IGL_CUNO_GROUP_R ICGR, IGLPROD.IGL_USER_CUNO_R IUCR, IGLPROD.IGL_USER IU " & _
                        "Where ICGR.AN8_CUNO = " & lBCC & " " & _
                        "AND ICGR.CUNO_GROUP_ID = IUCR.CUNO_GROUP_ID " & _
                        "AND IUCR.USER_SEQ_ID = IU.USER_SEQ_ID " & _
                        "AND IU.USER_SEQ_ID > 0 AND IU.USER_STATUS > 0 " & _
                        "AND IU.EMAIL_ADDRESS IS NOT NULL)) " & _
                        "ORDER BY LISTNAME"
    End Select
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) = "GPJ" Then
            lstUserNames.AddItem Trim(rst.Fields("LISTNAME"))
        Else
            lstUserNames.AddItem Trim(rst.Fields("LISTNAME")) & _
                        " [" & Trim(rst.Fields("EMPLOYER")) & "]"
        End If
        lstUserNames.ItemData(lstUserNames.NewIndex) = rst.Fields("USER_SEQ_ID")
        lstUsers.AddItem StrConv(Trim(rst.Fields("USERNAME")), vbProperCase)
        lstUsers.ItemData(lstUsers.NewIndex) = rst.Fields("USER_SEQ_ID")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Private Sub cmdUsers_Click()
    If lstUserNames.Visible = False Then
        fraEmail.Enabled = False
        lstUserNames.Visible = True
    Else
        lstUserNames.Visible = False
        fraEmail.Enabled = True
    End If
End Sub

Private Sub lstUsernames_Click()
    Dim i As Integer
    Dim sMess As String
    
    For i = 0 To lstUserNames.ListCount - 1
        If lstUserNames.Selected(i) Then
            If sMess = "" Then
                sMess = lstUsers.List(i)
            Else
                sMess = sMess & ", " & lstUsers.List(i)
            End If
        End If
    Next i
    txtEmail.Text = sMess
    lstUserNames.Visible = False
    fraEmail.Enabled = True
    
'''    If Trim(txtEmail.Text) = "" Then
'''        txtEmail.Text = lstUsers.List(lstUserNames.ListIndex)
'''    Else
'''        txtEmail.Text = txtEmail.Text & ", " & lstUsers.List(lstUserNames.ListIndex)
'''    End If
'''    txtEmail.SelStart = Len(txtEmail.Text)
End Sub

Private Sub cmdSend_Click()
    Dim bSuccess As Boolean
    
    If Trim(txtEmail.Text) = "" Then Exit Sub
    
    Call ParseNames(Trim(txtEmail.Text))
    
    ''COMMENTED OUT.  IF SendTheLink NEEDED, REENABLE FUNCTION''
    ''bSuccess = SendTheLink
    
    If bSuccess Then
        Unload Me
    Else
        MsgBox "The link was not successfully sent to all listed Users", _
                    vbExclamation, "Sorry..."
    End If
End Sub

Private Sub Form_Load()
    Call GetUsers(pFrom, pBCC)
    txtSubject.Text = pSub
End Sub

Private Sub txtEmail_Change()
    If Trim(txtEmail.Text) = "" Then cmdSend.Enabled = False Else cmdSend.Enabled = True
End Sub

Public Sub ParseNames(sNames As String)
    Dim iArray As Integer, iCom As Integer, i As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bSuccess As Boolean
    
    
    iArray = 0
    iCom = 1
    Do While iCom > 0
        iCom = InStr(iCom, sNames, ",")
        ReDim Preserve sUsers(iArray)
        If iCom > 0 Then
            sUsers(iArray) = Mid(sNames, 1, iCom - 1)
            sNames = Trim(Mid(sNames, iCom + 1))
            iArray = iArray + 1
            iCom = 1
        Else
            sUsers(iArray) = Mid(sNames, 1)
            iCom = 0
        End If
    Loop
    
    lstShortname.Clear
    lstEmail.Clear
    For i = LBound(sUsers) To UBound(sUsers)
        strSelect = "SELECT USER_SEQ_ID, NAME_LOGON, EMAIL_ADDRESS " & _
                    "From IGLPROD.IGL_USER " & _
                    "Where USER_SEQ_ID > 0 " & _
                    "AND (TRIM(NAME_FIRST)||' '||TRIM(NAME_LAST)) = '" & UCase(sUsers(i)) & "'"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            lstEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
            lstEmail.ItemData(lstEmail.NewIndex) = rst.Fields("USER_SEQ_ID")
            lstShortname.AddItem Trim(rst.Fields("NAME_LOGON"))
            lstShortname.ItemData(lstShortname.NewIndex) = rst.Fields("USER_SEQ_ID")
        Else
            MsgBox sUsers(i) & " is not recognized as an Annotator User", _
                        vbExclamation, "Skipping..."
        End If
        rst.Close
    Next i
    
End Sub

''Public Function SendTheLink() As Boolean
''    Dim i As Integer
''    Dim iTo As Integer
''    Dim sLink As String, sMess As String, sNote As String
'''''    Dim myNotes As New Domino.NotesSession
'''''    Dim myDB As New Domino.NotesDatabase
''    Dim myItem  As Object ''' NOTESITEM
''    Dim myReply  As Object ''' NOTESITEM
''    Dim myDoc As Object ''' NOTESDOCUMENT
''    Dim myRichText As Object ''' NOTESRICHTEXTITEM
''    Dim bSuccess As Boolean
''
''    bSuccess = True
''
''    If lstShortname.ListCount > 0 Then
''
''        Dim MailMan As New ChilkatMailMan2 '' New ChilkatMailMan2
''        MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
''
''        MailMan.SmtpHost = "mail.gpjco.com"
''
''        Dim Email As New ChilkatEmail2
''
''
'''        If Not bCitrix Then
'''            ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
'''            Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
'''            Dim myDB As Object '' LOTUS.NotesDatabase
'''
'''
'''            On Error Resume Next
'''            Set myNotes = GetObject(, "Notes.NotesSession")
'''
'''            If Err Then
'''                Err.Clear
'''                Set myNotes = CreateObject("Notes.NotesSession")
'''                If Err Then
'''                    MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
'''                    GoTo GetOut
'''                End If
'''            End If
'''            On Error GoTo 0
'''            Set myDB = myNotes.GetDatabase("", "")
'''            myDB.OPENMAIL
'''        Else
'''            ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
'''            Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
'''            Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
'''
'''            myDom.Initialize (sGAnnoPW)
'''            Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'''        End If
''
''
''        For i = 0 To lstShortname.ListCount - 1
''            Select Case pFrom
''                Case "DIL"
''                    sLink = "http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
''                                "?name_logon=" & LCase(lstShortname.List(i)) & _
''                                "&dil=" & pGID
''                Case "GH"
''                    sLink = "http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
''                                "?name_logon=" & LCase(lstShortname.List(i)) & _
''                                "&gid=" & pGID
''                Case "FP"
''                    '''LinksToAnno.asp?name_logon=dling&shyr=2002&shcd=24643&cuno=1242'''
''                    sLink = "http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
''                                "?name_logon=" & LCase(lstShortname.List(i)) & _
''                                "&shyr=" & pSHYR & "&shcd=" & pSHCD & "&cuno=" & pBCC
''            End Select
''
''            If txtMess.Text <> "" Then
''                sMess = txtMess.Text & vbNewLine & vbNewLine & sLink
''            Else
''                sMess = sLink
''            End If
''
'''            If Not bCitrix Then
'''                Set myDoc = myDB.CreateDocument
'''            Else
'''                Set myDoc = myDomDB.CreateDocument
'''                Call myDoc.ReplaceItemValue("Principal", LogName)
'''                Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'''            End If
''
''
'''            Set myItem = myDoc.AppendItemValue("Subject", txtSubject.Text)
'''
'''            Set myRichText = myDoc.CreateRichTextItem("Body")
'''            myRichText.AppendText sMess & vbNewLine & vbNewLine ''& sLink_Disclaimer
'''            myDoc.AppendItemValue "SENDTO", lstEmail.List(i)
''
''
''            For iTo = 0 To lstEmail.ListCount - 1
''                Email.AddTo lstEmail.List(iTo), lstEmail.List(iTo) '' Address(i), Address(i) ''frmPartCreator.lstEmail.List(i), frmPartCreator.lstEmail.List(i)
''            Next iTo
''
''            Email.FromAddress = LogAddress
''            Email.fromName = LogName
''
''            Email.subject = txtSubject.Text
''            Email.Body = sMess
''
''            Dim Success As Integer
''            Success = MailMan.SendEmail(Email)
''            If (Success = 0) Then
''                MsgBox MailMan.LastErrorText
''            End If
''
''
''
''
'''            On Error Resume Next
'''            Call myDoc.Send(False, lstEmail.List(i))
'''            If Err Then
'''                MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'''                            vbExclamation, "Error Encountered with " & lstEmail.List(i)
'''                bSuccess = False
'''                Err = 0
'''            End If
''
'''            Set myReply = Nothing
'''            Set myRichText = Nothing
'''            Set myItem = Nothing
'''            Set myDoc = Nothing
''        Next i
''    End If
''
''GetOut:
'''    If bCitrix Then
'''        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'''        If Not myDom Is Nothing Then Set myDom = Nothing
'''    Else
'''        If Not myDB Is Nothing Then Set myDB = Nothing
'''        If Not myNotes Is Nothing Then Set myNotes = Nothing
'''    End If
''
''    SendTheLink = bSuccess
''End Function
