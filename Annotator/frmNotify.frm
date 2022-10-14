VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPJ Annotator - Notification Screen"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8100
      Width           =   1635
   End
   Begin VB.Frame fraNotification 
      Caption         =   "Notification"
      Height          =   8295
      Left            =   5580
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.ListBox lstClientEmail 
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstTeamClients 
         Height          =   2310
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CheckBox chkReceive 
         Alignment       =   1  'Right Justify
         Caption         =   "Check to receive a copy of the Email Notification"
         Height          =   375
         Left            =   1740
         TabIndex        =   17
         Top             =   3180
         Width           =   2055
      End
      Begin VB.ListBox lstTeamEmail 
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   2820
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send Notification"
         Height          =   555
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7560
         Width           =   1755
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Email Notification..."
         Height          =   555
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7560
         Width           =   1755
      End
      Begin VB.TextBox txtComment 
         Height          =   1155
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   6300
         Width           =   3615
      End
      Begin VB.ListBox lstTeam 
         Height          =   2310
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox cboTeams 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clients Users:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   29
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Comments:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   6060
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Email Notification Team:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graphic Selection"
      Height          =   7755
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstDates 
         Height          =   255
         Left            =   3840
         TabIndex        =   31
         Top             =   7260
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   375
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7200
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   2205
         Left            =   3300
         ScaleHeight     =   2145
         ScaleWidth      =   1635
         TabIndex        =   18
         Top             =   480
         Width           =   1695
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H80000005&
            Caption         =   "CANCELED"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   33
            Tag             =   "5"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H80000005&
            Caption         =   "APPROVED"
            Height          =   255
            Index           =   25
            Left            =   60
            TabIndex        =   26
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H80000005&
            Caption         =   "RELEASED"
            Height          =   255
            Index           =   15
            Left            =   60
            TabIndex        =   25
            Tag             =   "15"
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H80000005&
            Caption         =   "DRAFT"
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   24
            Tag             =   "5"
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.ComboBox cboUsers 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   60
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkUser 
            BackColor       =   &H80000005&
            Height          =   255
            Left            =   1740
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   32
            Top             =   300
            Width           =   1515
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter by Status:"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   23
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter by User:"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1020
         End
      End
      Begin VB.ListBox lstCUNOs 
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   7260
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Selections"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7200
         Width           =   1455
      End
      Begin VB.ListBox lstFiles 
         Height          =   4110
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   3000
         Width           =   4815
      End
      Begin VB.ListBox lstClients 
         Height          =   2205
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Filters:"
         Height          =   195
         Index           =   3
         Left            =   3300
         TabIndex        =   19
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graphic Files:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Graphic Files by Client:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bTeamMember As Boolean, bMandRecip As Boolean
Dim sSelect As String, sWhere As String, sStatus As String, sAddUser As String, sOrderBy As String
Dim lCUNO As Long

Private Sub cboTeams_Click()
    Dim strSelect As String, sEmployer As String
    Dim rst As ADODB.Recordset
    
    If cboTeams.Text = "" Then Exit Sub
        
    '/// 0=Floorplan,1=Graphics,2=Const Dwg \\\
    
    '///// FIRST, GET TEAM \\\\\
    '///// SEE IF CLIENT-SHOW TEAM EXISTS \\\\\
    bTeamMember = False: bMandRecip = False
    strSelect = "SELECT U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, " & _
                "U.EMPLOYER, R.RECIPIENT_FLAG1 " & _
                "FROM " & ANOETeamUR & " R, " & IGLUser & " U " & _
                "WHERE R.TEAM_ID = " & cboTeams.ItemData(cboTeams.ListIndex) & " " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    lstTeam.Clear: lstTeamEmail.Clear
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) <> "GPJ" Then
            sEmployer = " (" & Trim(rst.Fields("EMPLOYER")) & ")"
        Else
            sEmployer = ""
        End If
        If StrConv(Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")), vbProperCase) = LogName Then
            bTeamMember = True
            If rst.Fields("RECIPIENT_FLAG1") = 1 Then bMandRecip = True
        End If
        lstTeam.AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & _
                    UCase(Trim(rst.Fields("NAME_LAST"))) & sEmployer
        lstTeam.ItemData(lstTeam.NewIndex) = rst.Fields("RECIPIENT_FLAG1")
        lstTeam.Selected(lstTeam.NewIndex) = CBool(rst.Fields("RECIPIENT_FLAG1") * -1)
        lstTeamEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    
    strSelect = "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, " & _
                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME " & _
                "FROM IGL_USER_CUNO_R CR, IGL_CUNO_GROUP_R GR, " & _
                "IGL_USER US, IGL_USER_APP_R AR " & _
                "Where GR.AN8_CUNO = " & lCUNO & " " & _
                "AND GR.CUNO_GROUP_ID = CR.CUNO_GROUP_ID " & _
                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
                "AND AR.APP_ID = 1002 " & _
                "ORDER BY US.NAME_LAST"
    Set rst = Conn.Execute(strSelect)
    lstTeamClients.Clear: lstClientEmail.Clear
    Do While Not rst.EOF
        lstTeamClients.AddItem UCase(Trim(rst.Fields("FULLNAME")))
        lstClientEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    
    If bMandRecip Then
        chkReceive.value = 1
        chkReceive.value = 2
        chkReceive.Enabled = False
    ElseIf bTeamMember Then
        chkReceive.value = 1
        chkReceive.Enabled = True
    Else
        chkReceive.value = 0
        chkReceive.Enabled = True
    End If
    chkReceive.Tag = LogAddress
    
End Sub

'''Private Sub cboUsers_Click()
'''    If cboUsers.Text <> "" And chkUser.Value = 1 Then
'''        sAddUser = "AND ADDUSER = '" & cboUsers.Text & "' "
'''    Else
'''        sAddUser = ""
'''    End If
'''    PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
'''End Sub

Private Sub chkStatus_Click(Index As Integer)
    sStatus = ResetAnd
    PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
End Sub

'''Private Sub chkUser_Click()
'''    If chkUser.Value = 1 Then
'''        cboUsers.Enabled = True
'''        If cboUsers.Text <> "" Then
'''            sAddUser = "AND ADDUSER = '" & cboUsers.Text & "' "
'''        Else
'''            sAddUser = ""
'''        End If
'''    Else
'''        cboUsers.Enabled = False
'''        sAddUser = ""
'''    End If
'''
'''    PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
'''End Sub

'''Private Sub cmdApplyFilter_Click()
'''    PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
'''End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = False
    Next i
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    With frmUsage
        .PassMess = CreateMessage
        .PassTitle = "Email Notification Preview..."
        .Show 1
    End With
'''    MsgBox CreateMessage, vbInformation, "Email Notification Preview..."
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    
    For i = lstFiles.ListCount - 1 To 0 Step -1
        lstFiles.Selected(i) = True
    Next i
End Sub

Private Sub cmdSend_Click()
    Dim sMess As String, sList1 As String, sList2 As String, sList3 As String, _
                sHDR As String, strUpdate As String, strDelete As String, strSelect As String
    Dim rst As ADODB.Recordset
    Dim myNotes As New Domino.NotesSession
    Dim myDB As New Domino.NotesDatabase
    Dim myItem  As Object ''' NOTESITEM
    Dim myReply  As Object ''' NOTESITEM
    Dim myDoc As Object ''' NOTESDOCUMENT
    Dim myRichText As Object ''' NOTESRICHTEXTITEM
    Dim Address() As String
    Dim i As Integer, iAdd As Integer
    
    Screen.MousePointer = 11
    
    sMess = CreateMessage
    
    If sMess <> "" Then
        '///// CHECK FOR TEAM TO NOTIFY \\\\\'
        iAdd = 0
        For i = 0 To lstTeam.ListCount - 1
            If lstTeam.Selected(i) = True Then
                ReDim Preserve Address(iAdd + 1)
                Address(iAdd) = lstTeamEmail.List(i)
'''                sList = sList & vbTab & lstTeam.List(i) & vbNewLine
                iAdd = iAdd + 1
            End If
        Next i
        If iAdd = 0 Then
            MsgBox "No one has been selected to be notified.", vbExclamation, "Canceling..."
            Exit Sub
        End If
        
        ''///// FIRST, DELETE CANCELED FILES \\\\\''
        sList1 = "": sList2 = ""
        For i = 0 To lstFiles.ListCount - 1
            If lstFiles.Selected(i) And InStr(1, lstFiles.List(i), "[CANCELED]") <> 0 Then
                If lstDates.ItemData(i) = 4 Then
                    '///// SEE IF IT IS SCHEDULED \\\\\'
                    strSelect = "SELECT GS.SHOW_ID " & _
                                    "FROM " & GFXShow & " GS, " & F5601 & " SM " & _
                                    "WHERE GS.GID = " & CLng(lstFiles.ItemData(i)) & " " & _
                                    "AND GS.SHYR = SM.SHY56SHYR " & _
                                    "AND GS.AN8_SHCD = SM.SHY56SHCD " & _
                                    "AND SM.SHY56BEGDT < " & IGLToJDEDate(Now)
                    Set rst = Conn.Execute(strSelect)
                    If rst.EOF Then
                        If sList1 = "" Then sList1 = CStr(lstFiles.ItemData(i)) _
                                    Else sList1 = sList1 & ", " & CStr(lstFiles.ItemData(i))
                    Else
                        If sList2 = "" Then sList2 = CStr(lstFiles.ItemData(i)) _
                                    Else sList2 = sList2 & ", " & CStr(lstFiles.ItemData(i))
                    End If
                    rst.Close
                Else
                    If sList1 = "" Then sList1 = CStr(lstFiles.ItemData(i)) _
                                Else sList1 = sList1 & ", " & CStr(lstFiles.ItemData(i))
                End If
            End If
        Next i
        Set rst = Nothing
        '///// DO TRUE DELETES \\\\\'
        If sList1 <> "" Then
            strSelect = "SELECT GPATH FROM " & GFXMas & " " & _
                        "WHERE GID IN (" & sList1 & ")"
            Set rst = Conn.Execute(strSelect)
'            On Error Resume Next
            Do While Not rst.EOF
                Kill Trim(rst.Fields("GPATH"))
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            Err.Clear
            On Error GoTo 0
            
            strDelete = "DELETE FROM " & GFXElt & " " & _
                        "WHERE GID IN (" & sList1 & ")"
            Conn.Execute (strDelete)
            
            strDelete = "DELETE FROM " & GFXShow & " " & _
                        "WHERE GID IN (" & sList1 & ")"
            Conn.Execute (strDelete)
            
            strDelete = "DELETE FROM " & GFXMas & " " & _
                        "WHERE GID IN (" & sList1 & ")"
            Conn.Execute (strDelete)
            
        End If
        '///// DO STATUS RESETS \\\\\'
        If sList2 <> "" Then
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = 0, " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, " & _
                        "UPDCNT = UPDCNT + 1 " & _
                        "WHERE GID IN (" & sList2 & ")"
            Conn.Execute (strUpdate)
        End If
        
        ''///// NOW, RESET STATUSES OF ALL OTHERS \\\\\''
        sList3 = ""
        For i = 0 To lstFiles.ListCount - 1
            If lstFiles.Selected(i) And InStr(1, lstFiles.List(i), "[CANCELED]") = 0 Then
                If sList3 = "" Then sList3 = CStr(lstFiles.ItemData(i)) _
                            Else sList3 = sList3 & ", " & CStr(lstFiles.ItemData(i))
            End If
        Next i
        If sList3 <> "" Then
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = GSTATUS + 5, " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, " & _
                        "UPDCNT = UPDCNT + 1 " & _
                        "WHERE GID IN (" & sList3 & ")"
            Conn.Execute (strUpdate)
        End If
        
        If sList1 = "" And sList2 = "" And sList3 = "" Then GoTo GetOut
        
        ''///// NOW, SEND OUT ALERT \\\\\''
        sHDR = lstClients.Text & " -- GPJ Annotator Graphics Posting Alert"
'''        sMess = "cc:" & sList & vbNewLine & vbNewLine & sMess
        
'        myNotes.Initialize
        On Error Resume Next
        If sNOTESID = "GANNOTAT" Then
            myNotes.Initialize (sNOTESPASSWORD)
        Else
            If sNOTESPASSWORD = "" Then
                ''GET PASSWORD''
TryPWAgain:
                frmGetPassword.Show 1, Me
                Select Case sNOTESPASSWORD
                    Case "_CANCEL"
                        sNOTESPASSWORD = ""
                        MsgBox "No email will be sent", vbExclamation, "User Canceled..."
                        Set myNotes = Nothing
                        Set myDB = Nothing
                    Case Else
                        Err.Clear
                        myNotes.Initialize (sNOTESPASSWORD)
                        If Err Then
                            Err.Clear
                            GoTo TryPWAgain
                        End If
                End Select
            Else
                myNotes.Initialize (sNOTESPASSWORD)
            End If
        End If
    
        '/// ACTIVATE FOR CITRIX \\\
        Set myDB = myNotes.GetDatabase(strMailSrvr, strMailFile)
    '''    Set myDB = myNotes.GETDATABASE(strMailSrvr, strMailFile)
        Set myDoc = myDB.CreateDocument
        If sNOTESID = "GANNOTAT" Then Call myDoc.ReplaceItemValue("Principal", LogName)
        Set myItem = myDoc.AppendItemValue("Subject", sHDR)
        If sNOTESID = "GANNOTAT" Then Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
        Set myRichText = myDoc.CreateRichTextItem("Body")
        myRichText.AppendText sMess ''' & vbNewLine & vbNewLine & _
                    vbNewLine & vbNewLine & sLink
'''        Set myReply = myDoc.APPENDITEMVALUE("ReplyTo", sUser)
        myDoc.AppendItemValue "SENDTO", Address(i)
'''        myDoc.SaveMessageOnSend = True
        
        On Error Resume Next
        For i = 0 To iAdd - 1
            Call myDoc.Send(False, Address(i))
        Next i
        If Err Then
            MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
                        vbExclamation, "Error Encountered"
            Err = 0
            GoTo GetOut
        End If
        
        Set myReply = Nothing
        Set myRichText = Nothing
        Set myItem = Nothing
        Set myDoc = Nothing
        Set myDB = Nothing
        Set myNotes = Nothing
        
        ''///// REFRESH LIST \\\\\''
        PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
    Else
        MsgBox "No Files have been selected.", vbExclamation, "No Email has been sent..."
    End If
GetOut:
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim lCUNO As Long
    Dim sGStatus(2 To 25) As String
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(2) = "  [CANCELED]"
    sGStatus(3) = "  [CANCELED]"
    sGStatus(4) = "  [CANCELED]"
    sGStatus(5) = "  [DRAFT]"
    sGStatus(15) = "  [RELEASED]"
    sGStatus(25) = "  [APPROVED]"
    
'''    '///// POP ALL NON-NOTIFIED FILES (WITH CLIENTS) \\\\\'
'''    lCUNO = 0
    sSelect = "SELECT GID, GDESC, GSTATUS, GTYPE, AN8_CUNO AS CUNO, ADDDTTM " & _
                "FROM " & GFXMas & " "
   sWhere = "WHERE AN8_CUNO = " & lCUNO & " "
   sStatus = "AND GSTATUS IN (2, 3, 4, 5, 15, 25) "
   sAddUser = "AND UPDUSER = '" & LogName & "' "
   sOrderBy = "ORDER BY GSTATUS, UPPER(GDESC)"
   
'''    Set rst = Conn.Execute(strSelect)
'''    Do While Not rst.EOF
'''        If rst.Fields("CUNO") <> lCUNO Then
'''            lstClients.AddItem UCase(Trim(rst.Fields("CLIENT")))
'''            lstClients.ItemData(lstClients.NewIndex) = rst.Fields("CUNO")
'''            lCUNO = rst.Fields("CUNO")
'''        End If
'''        lstFiles.AddItem Trim(rst.Fields("GDESC")) & sGStatus(rst.Fields("GSTATUS"))
'''        lstFiles.ItemData(lstFiles.NewIndex) = rst.Fields("GID")
'''        lstCUNOs.AddItem rst.Fields("CUNO")
'''        rst.MoveNext
'''    Loop
'''    rst.Close
    
    '///// POP ALL CLIENTS WITH NON-NOTIFIED FILES \\\\\'
    lCUNO = 0
    strSelect = "SELECT DISTINCT AB.ABALPH AS CLIENT, AB.ABAN8 AS CUNO " & _
                "FROM " & GFXMas & " GM, " & F0101 & " AB " & _
                "WHERE GM.GSTATUS IN (2, 3, 4, 5, 15, 25) " & _
                "AND GM.UPDUSER = '" & LogName & "' " & _
                "AND GM.AN8_CUNO = AB.ABAN8 " & _
                "ORDER BY AB.ABALPH, AB.ABAN8"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
'''        If rst.Fields("CUNO") <> lCUNO Then
            lstClients.AddItem UCase(Trim(rst.Fields("CLIENT")))
            lstClients.ItemData(lstClients.NewIndex) = rst.Fields("CUNO")
'''            lCUNO = rst.Fields("CUNO")
'''        End If
'''        lstFiles.AddItem Trim(rst.Fields("GDESC")) & sGStatus(rst.Fields("GSTATUS"))
'''        lstFiles.ItemData(lstFiles.NewIndex) = rst.Fields("GID")
'''        lstCUNOs.AddItem rst.Fields("CUNO")
        rst.MoveNext
    Loop
    rst.Close
    
    '///// POP EMAIL TEAMS \\\\\'
    strSelect = "SELECT AB.ABALPH AS CLIENT, ET.TEAM_ID " & _
                "FROM " & ANOETeam & " ET, " & F0101 & " AB " & _
                "WHERE ET.AN8_CUNO = AB.ABAN8 " & _
                "ORDER BY UPPER(AB.ABALPH)"
    Set rst = Conn.Execute(strSelect)
    cboTeams.AddItem "<No Team Setup>"
    Do While Not rst.EOF
        cboTeams.AddItem Trim(rst.Fields("CLIENT"))
        cboTeams.ItemData(cboTeams.NewIndex) = rst.Fields("TEAM_ID")
        rst.MoveNext
    Loop
    rst.Close
    
    lblUser = LogName
    
'''    '///// POP GRAPHIC POSTERS \\\\\\
'''    strSelect = "SELECT DISTINCT ADDUSER " & _
'''                "FROM " & GFXMas & " " & _
'''                "WHERE GSTATUS IN (5, 15, 25)"
'''    Set rst = Conn.Execute(strSelect)
'''    Do While Not rst.EOF
'''        cboUsers.AddItem Trim(rst.Fields("ADDUSER"))
'''        rst.MoveNext
'''    Loop
'''    rst.Close: Set rst = Nothing
'''
'''    On Error Resume Next
'''    cboUsers.Text = sUser
End Sub

Private Sub lstClients_Click()
    Dim i As Integer
    
    lCUNO = lstClients.ItemData(lstClients.ListIndex)
    sWhere = "WHERE AN8_CUNO = " & lCUNO & " "
    
    PopFiles (sSelect & sWhere & sStatus & sAddUser & sOrderBy)
    
'''    For i = 0 To lstCUNOs.ListCount - 1
'''        If CLng(lstCUNOs.List(i)) = lCUNO Then
'''            lstFiles.Selected(i) = True
'''        Else
'''            lstFiles.Selected(i) = False
'''        End If
'''    Next i
    
    On Error Resume Next
    cboTeams.Text = lstClients.Text
    If Err Then
        cboTeams.Text = "<No Team Setup>"
        MsgBox "NOTE:  Prior to the selected files being available on the Annotator, " & _
                    "an Email Notification Team needs to be setup for " & lstClients.Text & ", " & _
                    "and a notification of their posting must be sent out." & vbNewLine & vbNewLine & _
                    "Please, contact the A/E to arrange for a Team to be setup.", _
                    vbExclamation, "No Email Notification Team..."
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

Public Function CreateMessage() As String
    Dim sList As String, strSelect As String, sMess As String, _
                sDefault As String, sComment As String, sCC As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iStatus As Integer
    Dim sGType(1 To 4) As String
    Dim sGStatus(2 To 25) As String
    
    '///// FILE TYPE VARIABLES \\\\\
    sGType(1) = "Photo File"
    sGType(2) = "Graphic File"
    sGType(3) = "Graphic Layout"
    sGType(4) = "Presentation File"
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(2) = "CANCELED"
    sGStatus(3) = "CANCELED"
    sGStatus(4) = "CANCELED"
    sGStatus(5) = "INTERNAL DRAFT"
    sGStatus(15) = "CLIENT DRAFT"
    sGStatus(25) = "APPROVED"
    
    sDefault = "Those Graphics posted with an 'INTERNAL DRAFT' status, need to be reviewed, " & _
                "and their status advanced to 'CLIENT DRAFT' or 'APPROVED', before " & _
                "they will be available for viewing by non-GPJ personnel."
    sList = ""
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) = True Then
            If sList = "" Then sList = CStr(lstFiles.ItemData(i)) _
                        Else sList = sList & ", " & CStr(lstFiles.ItemData(i))
        End If
    Next i
    
    If sList <> "" Then
        sCC = "cc:"
        For i = 0 To lstTeam.ListCount - 1
            If lstTeam.Selected(i) = True Then
                sCC = sCC & vbTab & lstTeam.List(i) & vbNewLine
            End If
        Next i
        If chkReceive.value = 1 Then sCC = sCC & vbTab & UCase(LogName) & vbNewLine
        sCC = sCC & vbNewLine & "From:" & vbTab & UCase(LogName) & vbNewLine
        
        sMess = "": iStatus = -1
        strSelect = "SELECT AB.ABALPH AS CLIENT, GM.GDESC, GM.GTYPE, " & _
                    "GM.GSTATUS, GM.ADDDTTM " & _
                    "FROM " & GFXMas & " GM, " & F0101 & " AB " & _
                    "WHERE GM.GID IN (" & sList & ") " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "ORDER BY AB.ABALPH, GM.GSTATUS, GM.GTYPE, GM.GDESC"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sMess = sMess & UCase(Trim(rst.Fields("CLIENT"))) & " GRAPHIC FILES:" & vbNewLine
            Do While Not rst.EOF
                If iStatus <> rst.Fields("GSTATUS") Then
                    iStatus = rst.Fields("GSTATUS")
                    Select Case iStatus
                        Case 2, 3, 4
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (The following Files have been Deleted)" & vbNewLine
                        Case 5
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Available for Viewing by internal GPJ Only)" & vbNewLine
                        Case 15
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Released for Client Viewing)" & vbNewLine
                        Case 25
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Reviewed and Approved for Production)" & vbNewLine
                    End Select
                End If
                Select Case iStatus
                    Case 2, 3, 4
                        sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                    UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                    vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & vbNewLine
                    Case Else
                        sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                    UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                    vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                    vbNewLine '''REMOVE THIS LINE'''& _
                                    vbtab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
'                                    "?name_logon=" & LCASE(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
'                                    vbNewLine
                End Select

                rst.MoveNext
            Loop
        End If
        rst.Close: Set rst = Nothing
        
        If txtComment.Text <> "" Then
            sComment = vbNewLine & vbNewLine & "POSTERS COMMENTS (" & LogName & "):" & _
                        vbNewLine & Trim(txtComment.Text)
        Else
            sComment = ""
        End If
        
        CreateMessage = sCC & vbNewLine & vbNewLine & _
                    "The following Graphic Files have been posted to the GPJ Annotator, " & _
                    "and are ready for your review.  " & sDefault & sComment & _
                    vbNewLine & vbNewLine & sLink & _
                    vbNewLine & vbNewLine & vbNewLine & sMess
    Else
        CreateMessage = ""
    End If
End Function

Public Function ResetAnd() As String
    Dim sIN As String
    sIN = ""
    If chkStatus(3).value = 1 Then sIN = "2, 3, 4"
    If chkStatus(5).value = 1 Then
        If sIN = "" Then sIN = "5" Else sIN = sIN & ", 5"
    End If
    If chkStatus(15).value = 1 Then
        If sIN = "" Then sIN = "15" Else sIN = sIN & ", 15"
    End If
    If chkStatus(25).value = 1 Then
        If sIN = "" Then sIN = "25" Else sIN = sIN & ", 25"
    End If
    If sIN = "" Then
        chkStatus(5).value = 1
        ResetAnd = "AND GSTATUS IN (5) "
    Else
        ResetAnd = "AND GSTATUS IN (" & sIN & ") "
    End If
End Function

Public Sub PopFiles(strSelect As String)
    Dim rst As ADODB.Recordset
    Dim sGStatus(2 To 25) As String
    Dim i As Integer
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(2) = "  [CANCELED]"
    sGStatus(3) = "  [CANCELED]"
    sGStatus(4) = "  [CANCELED]"
    sGStatus(5) = "  [DRAFT]"
    sGStatus(15) = "  [RELEASED]"
    sGStatus(25) = "  [APPROVED]"
    
    lstFiles.Clear: lstCUNOs.Clear
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstFiles.AddItem Trim(rst.Fields("GDESC")) & sGStatus(rst.Fields("GSTATUS"))
        lstFiles.ItemData(lstFiles.NewIndex) = rst.Fields("GID")
        lstCUNOs.AddItem rst.Fields("CUNO")
        lstDates.AddItem UCase(Format(rst.Fields("ADDDTTM"), "MMMM YYYY"))
        lstDates.ItemData(lstDates.NewIndex) = rst.Fields("GSTATUS")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing

    For i = lstFiles.ListCount - 1 To 0 Step -1
        lstFiles.Selected(i) = True
    Next i
End Sub

