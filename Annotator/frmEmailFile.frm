VERSION 5.00
Begin VB.Form frmEmailFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email File..."
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFiles 
      Caption         =   "Attachments"
      Height          =   3975
      Left            =   5940
      TabIndex        =   23
      Top             =   2400
      Width           =   3615
      Begin VB.ListBox lstPlans 
         Height          =   2985
         ItemData        =   "frmEmailFile.frx":08CA
         Left            =   180
         List            =   "frmEmailFile.frx":08CC
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment Size: "
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Files to Attach:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   4635
      Left            =   180
      TabIndex        =   17
      Top             =   2400
      Width           =   5595
      Begin VB.TextBox txtDefMessage 
         Height          =   1635
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   2520
         Width           =   5235
      End
      Begin VB.TextBox txtCustMessage 
         Enabled         =   0   'False
         Height          =   1635
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   540
         Width           =   5235
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "Include hyperlink to the Annotator"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   4200
         Width           =   2895
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "Customized Message"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "Include Default Message"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   2220
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Header Information"
      Height          =   2175
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox chkMessage 
         Caption         =   "Send Myself a Copy of the Email"
         Height          =   315
         Index           =   2
         Left            =   4560
         TabIndex        =   16
         Top             =   1320
         Width           =   2715
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   10
         Top             =   300
         Width           =   3255
      End
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   900
         TabIndex        =   9
         Top             =   1680
         Width           =   8295
      End
      Begin VB.ListBox lstTo 
         Height          =   840
         ItemData        =   "frmEmailFile.frx":08CE
         Left            =   900
         List            =   "frmEmailFile.frx":08D0
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   720
         Width           =   3555
      End
      Begin VB.ListBox lstEmail 
         Height          =   1425
         Left            =   9360
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Rem Sels"
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboTeam 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   300
         Width           =   3555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEmailFile.frx":08D2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5265
         TabIndex        =   15
         Top             =   180
         UseMnemonic     =   0   'False
         Width           =   3825
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type To:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   12
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgEmailEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Left            =   4560
         MouseIcon       =   "frmEmailFile.frx":09AC
         MousePointer    =   99  'Custom
         Picture         =   "frmEmailFile.frx":0CB6
         ToolTipText     =   "Click to Edit History List"
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.ListBox lstDelete 
      Height          =   1425
      Left            =   0
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   4980
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview Email..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   5940
      TabIndex        =   2
      Top             =   6540
      Width           =   1575
   End
   Begin VB.ListBox lstPaths 
      Height          =   1425
      Left            =   5640
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7620
      TabIndex        =   0
      Top             =   6540
      Width           =   1935
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Recipient"
      End
   End
   Begin VB.Menu mnuPreview 
      Caption         =   "mnuPreview"
      Visible         =   0   'False
      Begin VB.Menu mnuPDFPreview 
         Caption         =   "Preview PDF"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel01 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmEmailFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tSHYR As Integer
Dim tBCC As String
Dim tSHCD As Long, tDWGID As Long, pFCCD As Long, pGID As Long
Dim tFBCN As String
Dim tSHNM As String
Dim tFileType As String
Dim tSHDT As String
Dim FileAttach() As String
Dim tFrom As String
Dim TPos As Integer
Dim sCopyPath As String
Dim bLoading As Boolean
Dim tHDR As String, tTTT As String
Dim tTAB As Integer
Dim iPDFRow As Integer


Public Property Get PassHDR() As String
    PassHDR = tHDR
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    tHDR = vNewValue
End Property

Public Property Get PassTAB() As Integer
    PassTAB = tTAB
End Property
Public Property Let PassTAB(ByVal vNewValue As Integer)
    tTAB = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = tFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    tFrom = vNewValue
End Property

Public Property Get PassBCC() As String
    PassBCC = tBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As String)
    tBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = tFBCN
End Property
Public Property Let PassFBCN(ByVal vNewValue As String)
    tFBCN = vNewValue
End Property

Public Property Get PassSHNM() As String
    PassSHNM = tSHNM
End Property
Public Property Let PassSHNM(ByVal vNewValue As String)
    tSHNM = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = tSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    tSHYR = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = tSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    tSHCD = vNewValue
End Property

Public Property Get PassDWGID() As Long
    PassDWGID = tDWGID
End Property
Public Property Let PassDWGID(ByVal vNewValue As Long)
    tDWGID = vNewValue
End Property

Public Property Get PassFILETYPE() As String
    PassFILETYPE = tFileType
End Property
Public Property Let PassFILETYPE(ByVal vNewValue As String)
    tFileType = vNewValue
End Property

Public Property Get PassSHDT() As String
    PassSHDT = tSHDT
End Property
Public Property Let PassSHDT(ByVal vNewValue As String)
    tSHDT = vNewValue
End Property

Public Property Get PassTTT() As String
    PassTTT = tTTT
End Property
Public Property Let PassTTT(ByVal vNewValue As String)
    tTTT = vNewValue
End Property

Public Property Get PassFCCD() As Long
    PassFCCD = pFCCD
End Property
Public Property Let PassFCCD(ByVal vNewValue As Long)
    pFCCD = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property


Public Sub SendEmail_Chilkat(iState As Integer)
    ''iState 0 = Preview; iState 1 = Send''

    Dim strSelect As String
    Dim rst As ADODB.Recordset
    'Dim i As Integer
    Dim myDoc As Object ''' As NOTESDOCUMENT
    Dim myItem As Object ''' As NOTESITEM
    Dim myRichText As Object ''' As NOTESRICHTEXTITEM
    Dim myReply As Object
    Dim Address() As String '/// RESET TO (1 TO 2) \\\'TEMP_CHANGE
    
    Dim MessBody As String, MessHdr As String, sAddress As String, sChk As String, _
                sDisclaimer As String, sWhom As String
    Dim i As Integer, iAddress As Integer, iComma1 As Integer, iComma2 As Integer, iL As Integer


    Dim bFound As Boolean
    
    
    
    sDisclaimer = "This message, including any attachments, " & _
                "is intended solely to be used by the individual " & _
                "or entity to which it is addressed.  It may contain " & _
                "information which is privileged, confidential, " & _
                "and / or otherwise exempt by law from disclosure.  " & _
                "If the reader of this message is not the intended recipient, " & _
                "or an employee or agent responsible for delivering " & _
                "this message to its intended recipient, you are herewith " & _
                "notified that any dissemination, distribution or copying " & _
                "of this communication is strictly prohibited.  " & _
                "If you believe you have received this communication in error, " & _
                "please notify the sender immediately."
    ''PARSE THROUGH ADDRESS''
    ReDim Address(lstTo.ListCount - 1)
    For i = 0 To UBound(Address)
        Address(i) = lstTo.List(i)
    Next i
    
    If chkMessage(2).Value = 1 Then
        For i = 0 To UBound(Address)
            If UCase(Address(i)) = UCase(LogAddress) Then GoTo FoundIt
        Next i
        ReDim Preserve Address(i)
        Address(i) = LogAddress
FoundIt:
    End If
    
    MessHdr = txtSubject.Text
    MessBody = ""
    If chkMessage(1).Value = 1 Then MessBody = txtCustMessage.Text & vbNewLine
    If chkMessage(0).Value = 1 Then
        If MessBody <> "" Then MessBody = MessBody & vbNewLine
        MessBody = MessBody & txtDefMessage.Text & vbNewLine
    End If
    MessBody = MessBody & vbNewLine
    
    If chkMessage(3).Value = 1 Then
        MessBody = MessBody & vbNewLine & sLink & vbNewLine
    End If
    MessBody = MessBody & vbNewLine & vbNewLine & _
                LogName & vbNewLine & _
                "mailto:" & LogAddress & vbNewLine & _
                vbNewLine & vbNewLine & sDisclaimer
    
    
    
    If iState = 0 Then
        sWhom = ""
        For i = 0 To UBound(Address)
            If Address(i) = "" Then i = i + 1
            If i > UBound(Address) Then Exit For
            If sWhom = "" Then
                sWhom = "TO:   " & Address(i)
            Else
                sWhom = sWhom & ", " & Address(i)
            End If
        Next i
        If sWhom <> "" Then sWhom = sWhom & vbNewLine & vbNewLine
        
        MessBody = sWhom & "SUBJECT:   " & MessHdr & vbNewLine & vbNewLine & vbNewLine & _
                    MessBody
        With frmUsage
            .PassMess = MessBody
            .PassTitle = "Email Preview..."
            .Show 1
        End With
        
        Exit Sub
        
    Else
        Dim MailMan As New ChilkatMailMan2 '' New ChilkatMailMan2
        MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
        
        MailMan.SmtpSsl = 1
        MailMan.SmtpPort = 465
        MailMan.SmtpUsername = "smtp@project.com"
        MailMan.SmtpPassword = "Tosa5550"
        MailMan.SmtpHost = "smtp.gmail.com"
        
        Dim Email As New ChilkatEmail2
        Email.subject = MessHdr
        
    
        Email.Body = MessBody
        
        For i = 0 To UBound(Address) ''lstTo.ListCount - 1
            Email.AddTo Address(i), Address(i)
        Next i
        
        'Email.AddTo "Steve.Westerholm@gpj.com", "Steve.Westerholm@gpj.com"
        
        Email.FromAddress = LogAddress
        Email.fromName = LogName
        
        Call GetAttachments
        For i = 1 To UBound(FileAttach)
            If FileAttach(i) <> "" Then
                Email.AddFileAttachment (FileAttach(i))
            End If
        Next i
                
        Dim Success As Integer
        Success = MailMan.SendEmail(Email)
        If (Success = 0) Then
            MsgBox MailMan.LastErrorText, vbExclamation, "Error encountered..."
            
        Else
            ''CHECK FOR NEW HISTORY''
            For i = 0 To UBound(Address)
                sChk = Address(i)
        '        bFound = False
                For iL = 0 To lstEmail.ListCount - 1
                    If UCase(Address(i)) = UCase(lstEmail.List(iL)) Then
                        GoTo AlreadyOnList
                    End If
                Next iL
                Call AddToANO_Email_Address(UserID, CLng(tBCC), Address(i))
                lstEmail.AddItem Address(i)
AlreadyOnList:
            Next i
            
            MsgBox "Email Sent", vbInformation, "Confirmation..."
            
        End If
    
    End If
    

    
    
End Sub

'Public Sub SendEmail(iState As Integer)
'    ''iState 0 = Preview; iState 1 = Send''
'    Dim MessBody As String, MessHdr As String, sAddress As String, sChk As String, _
'                sDisclaimer As String, sWhom As String
'    Dim i As Integer, iAddress As Integer, iComma1 As Integer, iComma2 As Integer, iL As Integer
''''''    Dim myNotes As New Domino.NotesSession
''''''    Dim myDB As New Domino.NotesDatabase
'    Dim myDoc As Object '' Domino.NotesDocument '' Object '  NOTESDOCUMENT
'    Dim myItem As Object '  NOTESITEM
'    Dim myRichText As Object '  NOTESRICHTEXTITEM
'    Dim Address() As String
'    Dim myReply As Object
'    Dim bFound As Boolean
'
'
'
'    sDisclaimer = "This message, including any attachments, " & _
'                "is intended solely to be used by the individual " & _
'                "or entity to which it is addressed.  It may contain " & _
'                "information which is privileged, confidential, " & _
'                "and / or otherwise exempt by law from disclosure.  " & _
'                "If the reader of this message is not the intended recipient, " & _
'                "or an employee or agent responsible for delivering " & _
'                "this message to its intended recipient, you are herewith " & _
'                "notified that any dissemination, distribution or copying " & _
'                "of this communication is strictly prohibited.  " & _
'                "If you believe you have received this communication in error, " & _
'                "please notify the sender immediately."
'    ''PARSE THROUGH ADDRESS''
'    ReDim Address(lstTo.ListCount - 1)
'    For i = 0 To UBound(Address)
'        Address(i) = lstTo.List(i)
'    Next i
'
''''    sAddress = txtEmailAddress.Text
''''    iComma1 = InStr(1, sAddress, ",")
''''    If iComma1 > 0 Then
''''        iAddress = 1
''''        ReDim Preserve Address(iAddress)
''''        Address(iAddress) = Left(sAddress, iComma1 - 1)
''''        Do While iComma1 > 0
''''            iComma2 = InStr(iComma1 + 1, sAddress, ",")
''''            If iComma2 <> 0 Then
''''                iAddress = iAddress + 1
''''                ReDim Preserve Address(iAddress)
''''                Address(iAddress) = Trim(Mid(sAddress, iComma1 + 1, iComma2 - iComma1 - 1))
''''                iComma1 = iComma2
''''            Else
''''                sChk = Trim(Mid(sAddress, iComma1 + 1))
''''                If Len(sChk) > 0 Then
''''                    iAddress = iAddress + 1
''''                    ReDim Preserve Address(iAddress)
''''                    Address(iAddress) = Trim(Mid(sAddress, iComma1 + 1))
''''                End If
''''                Exit Do
''''            End If
''''        Loop
''''    Else
''''        ReDim Preserve Address(1)
''''        Address(0) = txtEmailAddress.Text
''''    End If
'
'
'    If chkMessage(2).Value = 1 Then
'        For i = 0 To UBound(Address)
'            If UCase(Address(i)) = UCase(LogAddress) Then GoTo FoundIt
'        Next i
'        ReDim Preserve Address(i)
'        Address(i) = LogAddress
'FoundIt:
'    End If
'
'
''''    Address = txtEmailAddress.Text ''FOR NOW, PARSE REQUIRED''
'
'    MessHdr = txtSubject.Text
'    MessBody = ""
'    If chkMessage(1).Value = 1 Then MessBody = txtCustMessage.Text & vbNewLine
'    If chkMessage(0).Value = 1 Then
'        If MessBody <> "" Then MessBody = MessBody & vbNewLine
'        MessBody = MessBody & txtDefMessage.Text & vbNewLine
'    End If
'    MessBody = MessBody & vbNewLine
'
'    If chkMessage(3).Value = 1 Then
'        MessBody = MessBody & vbNewLine & sLink & vbNewLine
'    End If
'
'    If iState = 0 Then
'        sWhom = ""
'        For i = 0 To UBound(Address)
'            If Address(i) = "" Then i = i + 1
'            If i > UBound(Address) Then Exit For
'            If sWhom = "" Then
'                sWhom = "TO:   " & Address(i)
'            Else
'                sWhom = sWhom & ", " & Address(i)
'            End If
'        Next i
'        If sWhom <> "" Then sWhom = sWhom & vbNewLine & vbNewLine
'
'        MessBody = sWhom & "SUBJECT:   " & MessHdr & vbNewLine & vbNewLine & vbNewLine & _
'                    MessBody & vbNewLine & vbNewLine & LogName & vbNewLine & _
'                    "mailto:" & LogAddress & vbNewLine & vbNewLine & "<ATTACHMENTS>"
'        MessBody = MessBody & vbNewLine & vbNewLine & sDisclaimer
'        With frmUsage
'            .PassMess = MessBody
'            .PassTitle = "Email Preview..."
'            .Show 1
'        End With
'
'        Exit Sub
'
'    ElseIf iState = 1 Then
'        Call GetAttachments
'
'        If Not bCitrix Then
'            ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
'            Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
'            Dim myDB As Object '' LOTUS.NotesDatabase
'
'
'            On Error Resume Next
'            Set myNotes = GetObject(, "Notes.NotesSession")
'
'            If Err Then
'                Err.Clear
'                Set myNotes = CreateObject("Notes.NotesSession")
'                If Err Then
'                    MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
'                    GoTo GetOut
'                End If
'            End If
'            On Error GoTo 0
'            Set myDB = myNotes.GetDatabase("", "")
'            myDB.OPENMAIL
'            Set myDoc = myDB.CreateDocument
'
'        Else
'            ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
'            Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
'            Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
'
'
'            myDom.Initialize (sGAnnoPW)
'            Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'            Set myDoc = myDomDB.CreateDocument
'
'            Call myDoc.ReplaceItemValue("Principal", LogName)
'            Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'        End If
'
''''''        Set myDoc = myDB.CreateDocument
''''''        If sNOTESID = "GANNOTAT" Then Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
'        Set myRichText = myDoc.CreateRichTextItem("Body")
'        With myRichText
'            .AppendText MessBody
'            .AddNewLine 2
'            .AppendText LogName
'            .AddNewLine 1
'            .AppendText "mailto:" & LogAddress
'            .AddNewLine 2
'            For i = 1 To UBound(FileAttach)
'                If FileAttach(i) <> "" Then
'                    .EmbedObject 1454, "", FileAttach(i)
'                    .AddNewLine 1
'                End If
'            Next i
'            .AddNewLine 2
'            .AppendText sDisclaimer
'        End With
'
''''''        If sNOTESID = "GANNOTAT" Then Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'        myDoc.AppendItemValue "SENDTO", Address
''''        myDoc.SaveMessageOnSend = True
'
'        On Error Resume Next
'        Call myDoc.Send(False, Address)
''        Call myDoc.SEND(True, Address)
'        If Err Then
'            MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                        vbExclamation, "Error Encountered"
'            Err = 0
'            GoTo GetOut
'        Else
'            MsgBox "Email Sent", vbInformation, "Confirmation..."
'        End If
'GetOut:
'
'    End If
'
'    Set myReply = Nothing
'    Set myRichText = Nothing
'    Set myItem = Nothing
'    Set myDoc = Nothing
'
'    If bCitrix Then
'        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'        If Not myDom Is Nothing Then Set myDom = Nothing
'    Else
'        If Not myDB Is Nothing Then Set myDB = Nothing
'        If Not myNotes Is Nothing Then Set myNotes = Nothing
'    End If
'
'    ''CHECK FOR NEW HISTORY''
'    For i = 0 To UBound(Address)
'        sChk = Address(i)
''        bFound = False
'        For iL = 0 To lstEmail.ListCount - 1
'            If UCase(Address(i)) = UCase(lstEmail.List(iL)) Then
'                GoTo AlreadyOnList
'            End If
'        Next iL
'        Call AddToANO_Email_Address(UserID, CLng(tBCC), Address(i))
'        lstEmail.AddItem Address(i)
'AlreadyOnList:
'    Next i
'
'End Sub

Private Sub cboTeam_Click()
    Dim i As Integer
    
    If cboTeam.Text = "" Or Left(cboTeam.Text, 3) = "---" Then Exit Sub
    
    For i = 0 To lstTo.ListCount - 1
        If UCase(lstTo.List(i)) = UCase(cboTeam.Text) Then
            MsgBox lstTo.List(i) & " already exists on the Email list", _
                        vbExclamation, "Duplicate Recipient..."
            Exit Sub
        End If
    Next i
    lstTo.AddItem cboTeam.Text
    
    If CheckIfReady Then
        cmdSend.Enabled = True: cmdPreview.Enabled = True
    Else
        cmdSend.Enabled = False: cmdPreview.Enabled = False
    End If
    
'''    If txtEmailAddress.Text = "" Then
'''        txtEmailAddress.Text = cboTeam.Text
'''    Else
'''        If Trim(txtEmailAddress.Text) <> "" Then
'''            ''CHECK IF ALREADY PRESENT''
'''            If InStr(1, txtEmailAddress.Text, cboTeam.Text) = 0 Then
'''                ''CHECK FOR COMMA''
'''                If Right(txtEmailAddress.Text, 1) = "," Then
'''                    txtEmailAddress.Text = txtEmailAddress.Text & " " & cboTeam.Text
'''                ElseIf Right(Trim(txtEmailAddress.Text), 1) = "," Then
'''                    txtEmailAddress.Text = txtEmailAddress.Text & cboTeam.Text
'''                Else
'''                    txtEmailAddress.Text = txtEmailAddress.Text & ", " & cboTeam.Text
'''                End If
'''            End If
'''        Else
'''            txtEmailAddress.Text = cboTeam.Text
'''        End If
'''    End If
End Sub

Private Sub chkMessage_Click(Index As Integer)
    Select Case Index
        Case 1
            If chkMessage(1).Value = 0 Then
                txtCustMessage.Enabled = False
            Else
                txtCustMessage.Enabled = True
            End If
        Case 0
            If chkMessage(0).Value = 0 Then
                txtDefMessage.Enabled = False
            Else
                txtDefMessage.Enabled = True
            End If
    End Select
End Sub

Private Sub cmdPreview_Click()
    Call SendEmail_Chilkat(0)
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    For i = lstTo.ListCount - 1 To 0 Step -1
        If lstTo.Selected(i) Then lstTo.RemoveItem (i)
    Next i
    cmdRemove.Visible = False
    If CheckIfReady Then
        cmdSend.Enabled = True: cmdPreview.Enabled = True
    Else
        cmdSend.Enabled = False: cmdPreview.Enabled = False
    End If
End Sub

Private Sub cmdSend_Click()
    Screen.MousePointer = 11
    'Call SendEmail(1)
    Call SendEmail_Chilkat(1)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim sMess As String, strSelect As String, sChk As String
    Dim rst As ADODB.Recordset
    
    bLoading = True
    
'''    txtSubject.Text = "GPJ Space Plans:  " & tSHYR & " - " & tSHNM
    Select Case UCase(tFrom)
        Case "FRMOSP", "FRMPHOTO"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Floorplans\FacilPho\"
            txtSubject.Text = tHDR
            Call GetFacilPhotos(pFCCD, pGID)
            ''DEFAULT MESSAGE''
            sMess = "The attached JPG files are from the GPJ Annotator Facilities Interface." & vbNewLine
            txtDefMessage.Text = sMess
            
        Case "FRMFACIL-PDF"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Floorplans\"
            txtSubject.Text = "Facility PDFs:  " & tHDR
            strSelect = "SELECT DF.DWFID, DF.DWFDESC, DF.DWFPATH, UPPER(DF.DWFDESC) AS UDESC, DF.DWFSTATUS " & _
                        "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                        "Where DM.AN8_CUNO = " & pFCCD & " " & _
                        "AND DM.DWGID = DS.DWGID " & _
                        "AND DS.DWGID = DF.DWGID " & _
                        "AND DS.SHTID = DF.SHTID " & _
                        "ORDER BY DF.DWFSTATUS DESC, UDESC ASC"
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sChk = Left(Trim(rst.Fields("DWFPATH")), Len(Trim(rst.Fields("DWFPATH"))) - 3) & "pdf"
                If Dir(sChk, vbNormal) <> "" Then
                    lstPlans.AddItem Trim(rst.Fields("DWFDESC"))
                    lstPlans.ItemData(lstPlans.NewIndex) = rst.Fields("DWFID")
                    lstPaths.AddItem sCopyPath & rst.Fields("DWFID") & ".pdf"
                    lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                End If
                
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            ''DEFAULT MESSAGE''
            sMess = "The attached PDF files are from the GPJ Annotator Facilities Interface." & vbNewLine
            txtDefMessage.Text = sMess
            
        Case "FRMFACIL-DWF"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Floorplans\"
            txtSubject.Text = "Facility DWFs:  " & tHDR
            strSelect = "SELECT DF.DWFID, DF.DWFDESC, DF.DWFPATH, UPPER(DF.DWFDESC) AS UDESC, DF.DWFSTATUS " & _
                        "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                        "Where DM.AN8_CUNO = " & pFCCD & " " & _
                        "AND DM.DWGID = DS.DWGID " & _
                        "AND DS.DWGID = DF.DWGID " & _
                        "AND DS.SHTID = DF.SHTID " & _
                        "ORDER BY DF.DWFSTATUS DESC, UDESC ASC"
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sChk = Trim(rst.Fields("DWFPATH"))
                If Dir(sChk, vbNormal) <> "" Then
                    lstPlans.AddItem Trim(rst.Fields("DWFDESC"))
                    lstPlans.ItemData(lstPlans.NewIndex) = rst.Fields("DWFID")
                    lstPaths.AddItem Trim(rst.Fields("DWFPATH"))
                    lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                End If
                
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            ''DEFAULT MESSAGE''
            sMess = "The attached DWF files are from the GPJ Annotator Facilities Interface." & vbNewLine
            txtDefMessage.Text = sMess
            
        Case "FRMDIL-MULTI"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Graphics\TempCopy\"
            txtSubject.Text = "DIL Images: " & tHDR
            For i = 0 To frmDIL.chkMulti.Count - 1
                If frmDIL.chkMulti(i).Value = 1 Then
                    lstPlans.AddItem frmDIL.imx0(i).ToolTipText
                    lstPlans.ItemData(lstPlans.NewIndex) = CLng(frmDIL.imx0(i).Tag)
                    lstPlans.Selected(lstPlans.NewIndex) = 1
                    lstPaths.AddItem GetGPath(CLng(frmDIL.imx0(i).Tag))
                    lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
'                    lstPaths.AddItem frmDIL.imx0(i).FileName
                End If
            Next i
            
            ''DEFAULT MESSAGE''
            sMess = "The attached images are from the GPJ Digital Image Library." & vbNewLine
            txtDefMessage.Text = sMess
            
        Case "FRMDIL-SINGLE"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Graphics\TempCopy\"
            txtSubject.Text = "DIL Image: " & tHDR
            lstPlans.AddItem sCGDesc
            lstPlans.ItemData(lstPlans.NewIndex) = lGID
            lstPlans.Selected(lstPlans.NewIndex) = 1
            lstPaths.AddItem GetGPath(lGID)
            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
            
            ''DEFAULT MESSAGE''
            sMess = "The attached image is from the GPJ Digital Image Library." & vbNewLine
            txtDefMessage.Text = sMess
            
        Case "FRMGRAPHICS-MULTI"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Graphics\TempCopy\"
            txtSubject.Text = "Gfx Handler Images: " & tFBCN & " Images"
            
            Select Case tTAB
                Case 0
                    For i = 0 To frmGraphics.chk0.Count - 1
                        If frmGraphics.chk0(i).Value = 1 Then
                            lstPlans.AddItem GetGDesc(frmGraphics.lbl0(i).ToolTipText)
                            lstPlans.ItemData(lstPlans.NewIndex) = CLng(frmGraphics.lbl0(i).Tag)
                            lstPlans.Selected(lstPlans.NewIndex) = 1
                            lstPaths.AddItem GetGPath(CLng(frmGraphics.lbl0(i).Tag))
                            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                        End If
                    Next i
                Case 1
                    For i = 0 To frmGraphics.chk1.Count - 1
                        If frmGraphics.chk1(i).Value = 1 Then
                            lstPlans.AddItem GetGDesc(frmGraphics.lbl1(i).ToolTipText)
'''                            lstPlans.AddItem frmGraphics.imx1(i).ToolTipText
                            lstPlans.ItemData(lstPlans.NewIndex) = CLng(frmGraphics.lbl1(i).Tag)
                            lstPlans.Selected(lstPlans.NewIndex) = 1
                            lstPaths.AddItem GetGPath(CLng(frmGraphics.lbl1(i).Tag))
                            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                        End If
                    Next i
                Case 2
                    For i = 0 To frmGraphics.chk2.Count - 1
                        If frmGraphics.chk2(i).Value = 1 Then
                            lstPlans.AddItem GetGDesc(frmGraphics.lbl2(i).ToolTipText)
                            lstPlans.ItemData(lstPlans.NewIndex) = CLng(frmGraphics.lbl2(i).Tag)
                            lstPlans.Selected(lstPlans.NewIndex) = 1
                            lstPaths.AddItem GetGPath(CLng(frmGraphics.lbl2(i).Tag))
                            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                        End If
                    Next i
                Case 3
                    For i = 0 To frmGraphics.chk3.Count - 1
                        If frmGraphics.chk3(i).Value = 1 Then
                            lstPlans.AddItem GetGDesc(frmGraphics.lbl3(i).ToolTipText)
                            lstPlans.ItemData(lstPlans.NewIndex) = CLng(frmGraphics.lbl3(i).Tag)
                            lstPlans.Selected(lstPlans.NewIndex) = 1
                            lstPaths.AddItem GetGPath(CLng(frmGraphics.lbl3(i).Tag))
                            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                        End If
                    Next i
                Case 4
                    For i = 0 To frmGraphics.chk4.Count - 1
                        If frmGraphics.chk4(i).Value = 1 Then
                            lstPlans.AddItem frmGraphics.flxApprove.TextMatrix(i + 1, 3)
                            lstPlans.ItemData(lstPlans.NewIndex) = frmGraphics.flxApprove.TextMatrix(i + 1, 0)
                            lstPlans.Selected(lstPlans.NewIndex) = 1
                            lstPaths.AddItem GetGPath(frmGraphics.flxApprove.TextMatrix(i + 1, 0))
                            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
                        End If
                    Next i
                End Select
            
            ''DEFAULT MESSAGE''
            sMess = "The attached images are from the GPJ Annotator's Graphic Handler." & vbNewLine & vbNewLine & _
                        vbTab & "Client:" & vbTab & tFBCN & vbNewLine & _
                        vbTab & "Tab:    " & vbTab & UCase(frmGraphics.sst1.TabCaption(tTAB)) & vbNewLine & _
                        vbTab & "Folder:" & vbTab & tHDR
            txtDefMessage.Text = sMess
        
        Case "FRMGRAPHICS-SINGLE"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Graphics\TempCopy\"
            txtSubject.Text = "Gfx Handler Image: " & tFBCN & " Image"
            lstPlans.AddItem sCGDesc
            lstPlans.ItemData(lstPlans.NewIndex) = lGID
            lstPlans.Selected(lstPlans.NewIndex) = 1
            lstPaths.AddItem GetGPath(lGID)
            lstPaths.ItemData(lstPaths.NewIndex) = CLng(FileLen(lstPaths.List(lstPaths.NewIndex)))
            
            ''DEFAULT MESSAGE''
            sMess = "The attached image is from the GPJ Annotator's Graphic Handler." & vbNewLine & vbNewLine & _
                        vbTab & "Client:" & vbTab & tFBCN & vbNewLine & _
                        vbTab & "Tab:    " & vbTab & UCase(frmGraphics.sst1.TabCaption(tTAB)) & vbNewLine & _
                        vbTab & "Folder:" & vbTab & tHDR
            txtDefMessage.Text = sMess
        
        Case "FRMANNOTATOR", "FRMSHOW"
            sCopyPath = "\\DETMSFS01\GPJAnnotator\Floorplans\TempCopy\"
            If tDWGID = 0 Then
                txtSubject.Text = "GPJ Show Plans:  " & tSHYR & " - " & tSHNM
                Call GetShowPlans(frmEmailFile, "PDF", tSHYR, tSHCD)
            Else
                txtSubject.Text = "GPJ Space Plans:  " & tSHYR & " - " & tSHNM
                Call GetDrawings(frmEmailFile, "PDF", tDWGID, tSHYR, tSHCD)
                For i = 0 To lstPlans.ListCount - 1
                    If UCase(lstPlans.List(i)) = "FLOORPLAN" Then lstPlans.Selected(i) = True
                Next i
            End If
    
            ''DEFAULT MESSAGE''
            sMess = "The attached drawings are for the following Show:" & vbNewLine
            If tDWGID <> 0 Then sMess = sMess & Space(4) & "Client:        " & vbTab & tFBCN & vbNewLine
            sMess = sMess & Space(4) & "Show Year:" & vbTab & tSHYR & vbNewLine
            sMess = sMess & Space(4) & "Show Name:" & vbTab & tSHNM & vbNewLine
            sMess = sMess & Space(4) & "Show Dates:" & vbTab & UCase(tSHDT)
            txtDefMessage.Text = sMess
            
            If tTTT <> "" Then fraFiles.Caption = tTTT
            
            
    End Select
    
    If Left(UCase(tFrom), 8) <> "FRMFACIL" _
                And UCase(tFrom) <> "FRMOSP" _
                And UCase(tFrom) <> "FRMPHOTO" Then
        Call GetEmailList(UserID, CLng(tBCC))
    End If
    lblSize.Caption = "Size of All Attachments:  " & GetTotalSize
    bLoading = False
End Sub


Public Sub GetAttachments()
    Dim i As Integer, iCnt As Integer, iSlash As Integer
    Dim sFile As String, sPath As String, sNewFile As String, sFormat As String
    
    iCnt = 0
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) = True Then
            iCnt = iCnt + 1
            ReDim Preserve FileAttach(iCnt)
            sFile = lstPaths.List(i)
            ''GET PATH''
            For iSlash = Len(sFile) To 1 Step -1
                If Mid(sFile, iSlash, 1) = "." Then
                    sFormat = Mid(sFile, iSlash + 1)
                End If
                If Mid(sFile, iSlash, 1) = "\" Then
                    sPath = Left(sFile, iSlash) & "TempCopy\"
                    Exit For
                End If
            Next iSlash
            ''SET FULL NAME''
            Select Case UCase(tFrom)
                Case "FRMSHOW"
                    sNewFile = sPath & CStr(tSHYR) & " " & DashForSlash(tSHNM) & _
                                " [" & DashForSlash(lstPlans.List(i)) & "].pdf"
                Case "FRMANNOTATOR"
                    sNewFile = sPath & DashForSlash(Trim(tFBCN)) & " - " & CStr(tSHYR) & _
                                " " & DashForSlash(Trim(tSHNM)) & _
                                " [" & Legalize(lstPlans.List(i)) & "].pdf"
                    ''sNewFile = Legalize(sNewFile)
                Case "FRMDIL-MULTI", "FRMDIL-SINGLE", "FRMGRAPHICS-MULTI", "FRMGRAPHICS-SINGLE"
                    sNewFile = sPath & DashForSlash(lstPlans.List(i)) & "." & sFormat
                Case "FRMFACIL-PDF"
                    sNewFile = sPath & Legalize(tHDR) & " [" & lstPlans.List(i) & "].pdf"
                Case "FRMFACIL-DWF"
                    sNewFile = sPath & Legalize(tHDR) & " [" & lstPlans.List(i) & "].dwf"
                Case "FRMOSP"
                    sNewFile = sPath & Legalize(Mid(tHDR, 12)) & " [" & lstPlans.List(i) & "].jpg"
                Case "FRMPHOTO"
                    sNewFile = sPath & Legalize(tHDR) & " [" & lstPlans.List(i) & "].jpg"
            End Select
            ''COPY FILE''
            FileCopy sFile, sNewFile
            FileAttach(iCnt) = sNewFile
            
            ''WRITE TO DELETE LIST''
            lstDelete.AddItem sNewFile
        End If
    Next i
    
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    If lstDelete.ListCount > 0 Then
        On Error Resume Next
        For i = lstDelete.ListCount - 1 To 0 Step -1
            Kill lstDelete.List(i)
            lstDelete.RemoveItem (i)
        Next i
    End If
End Sub

Private Sub imgEmailEdit_Click()
    frmEmailEdit.PassUID = UserID
    frmEmailEdit.PassBCC = tBCC
    frmEmailEdit.Show 1, Me
    
    Call GetEmailList(UserID, CLng(tBCC))
End Sub

Private Sub lstPlans_Click()
    If CheckIfReady Then
        cmdSend.Enabled = True: cmdPreview.Enabled = True
    Else
        cmdSend.Enabled = False: cmdPreview.Enabled = False
    End If
End Sub

Private Sub lstPlans_ItemCheck(Item As Integer)
    If CheckIfReady Then
        cmdSend.Enabled = True: cmdPreview.Enabled = True
    Else
        cmdSend.Enabled = False: cmdPreview.Enabled = False
    End If
    If Not bLoading Then lblSize.Caption = "Size of All Attachments:  " & GetTotalSize
End Sub


Private Sub lstPlans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = vbRightButton Then
        i = Int(Y / 225)
        If i + lstPlans.TopIndex < lstPlans.ListCount Then
            iPDFRow = i + lstPlans.TopIndex
            Me.PopupMenu Me.mnuPreview
'            MsgBox "Selected " & i + lstPlans.TopIndex & " row"
        End If
    End If
End Sub

Private Sub lstTo_Click()
    Dim i As Integer
    
    cmdRemove.Visible = False
    For i = 0 To lstTo.ListCount - 1
        If lstTo.Selected(i) Then
            cmdRemove.Visible = True
            Exit Sub
        End If
    Next i
End Sub

Private Sub lstTo_Validate(Cancel As Boolean)
    Debug.Print "Validating"
End Sub

Private Sub mnuPDFPreview_Click()
'''    MsgBox "Preview PDF of '" & lstPlans.List(iPDFRow) & "'" & vbNewLine & _
'''                "Path:  " & lstPaths.List(iPDFRow)
    With frmHTMLViewer
        .PassFile = lstPaths.List(iPDFRow)
        .PassHDR = txtSubject & "  (" & lstPlans.List(iPDFRow) & ")"
        .PassFrom = Me.Name
        .Show 1, Me
    End With
End Sub

Private Sub txtEmailAddress_Change()
    Dim i As Integer
    
    If Text1.Text = "" Then
        If txtEmailAddress.Text <> "" Then
            txtEmailAddress.SelStart = Len(txtEmailAddress.Text)
        End If
        Exit Sub
    End If
    
    For i = 0 To lstEmail.ListCount - 1
        If UCase(Left(lstEmail.List(i), Len(Text1.Text))) = UCase(Text1.Text) Then
            txtEmailAddress.Text = lstEmail.List(i)
            Exit For
        End If
    Next i
    
    If TPos > Len(txtEmailAddress.Text) Then TPos = Len(txtEmailAddress.Text)
    txtEmailAddress.SelStart = TPos
    txtEmailAddress.SelLength = Len(txtEmailAddress.Text) - TPos
    
    If CheckIfReady Then
        cmdSend.Enabled = True: cmdPreview.Enabled = True
    Else
        cmdSend.Enabled = False: cmdPreview.Enabled = False
    End If
End Sub

Public Function CheckIfReady() As Boolean
    Dim i As Integer
    Dim bReady As Boolean
    
    bReady = False
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) = True Then
            bReady = True
            Exit For
        End If
    Next i
    
    If bReady Then
        If lstTo.ListCount = 0 Then bReady = False
    End If
    
    CheckIfReady = bReady
End Function

Public Function DashForSlash(sName As String) As String
    Dim Pos As Integer
    '///// DeSLASH \\\\\
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(Pos, sName, "/")
        If Pos > 0 Then
            sName = Left(sName, Pos - 1) & "-" & Mid(sName, Pos + 1)
        End If
    Loop
    DashForSlash = sName
End Function


Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
'    bEnd = False
    If txtEmailAddress.SelStart < TPos Then
        Text1.Text = Left(txtEmailAddress.Text, txtEmailAddress.SelStart)
'    If Len(txtEmailAddress.Text) <= Len(Text1.Text) Then
'        Text1.Text = txtEmailAddress.Text
        TPos = Len(Text1.Text)
    End If
    
    
    If KeyAscii = 8 Then
        If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    ElseIf KeyAscii = 13 Then
        For i = 0 To lstTo.ListCount - 1
            If lstTo.List(i) = txtEmailAddress.Text Then
                MsgBox lstTo.List(i) & " already exists on the Email list", _
                            vbExclamation, "Duplicate Recipient..."
                Text1.Text = ""
                txtEmailAddress.Text = ""
                Exit Sub
            End If
        Next i
        lstTo.AddItem txtEmailAddress.Text
        Text1.Text = ""
        txtEmailAddress.Text = ""
        If CheckIfReady Then
            cmdSend.Enabled = True: cmdPreview.Enabled = True
        Else
            cmdSend.Enabled = False: cmdPreview.Enabled = False
        End If
        Exit Sub
    Else
        Text1.Text = Text1.Text & Chr(KeyAscii)
    End If
    TPos = Len(Text1.Text)
    If Len(Text1.Text) = 0 Then txtEmailAddress.Text = ""
End Sub

Public Sub AddToANO_Email_Address(tUID As Long, tCUNO As Long, tAddress As String)
    Dim strInsert As String
    Dim rstL As ADODB.Recordset
    Dim tEID As Long
    
'''    If tAddress Is Null Then Exit Sub
    
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    tEID = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing
            
    strInsert = "INSERT INTO ANNOTATOR.ANO_EMAIL_ADDRESS " & _
                "(USER_SEQ_ID, EMAIL_ID, AN8_CUNO, EMAIL_ADDRESS, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                "VALUES " & _
                "(" & tUID & ", " & tEID & ", " & tCUNO & ", " & _
                "'" & DeGlitch(Left(tAddress, 50)) & "', " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
    
    
End Sub

Public Sub GetEmailList(tUID As Long, tCUNO As Long)
    Dim strSelect As String, sChk As String
    Dim rst As ADODB.Recordset
    
    
    
    lstEmail.Clear
    cboTeam.Clear
    ''GET EMAIL ADDRESS LIST''
    strSelect = "select iucr.user_seq_id, iu.name_logon, iu.employer, iu.email_address " & _
                "from IGLPROD.IGL_user_cuno_r iucr, IGLPROD.IGL_user iu " & _
                "where iucr.cuno_group_id in " & _
                "(select cuno_group_id from IGLPROD.IGL_cuno_group_r where an8_cuno = " & tCUNO & ") " & _
                "and iucr.user_seq_id = iu.user_seq_id " & _
                "and iu.user_status = 1 " & _
                "order by iu.email_address"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Do While Not rst.EOF
            sChk = UCase(Trim(rst.Fields("employer")))
            If InStr(1, UCase(tFBCN), sChk) <> 0 Then
                If cboTeam.ListCount = 0 Then
                    cboTeam.AddItem "--------------- Client List ---------------"
                End If
                cboTeam.AddItem LCase(Trim(rst.Fields("EMAIL_ADDRESS")))
                lstEmail.AddItem LCase(Trim(rst.Fields("EMAIL_ADDRESS")))
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    ''ADD HISTORY''
    strSelect = "SELECT EMAIL_ADDRESS FROM ANNOTATOR.ANO_EMAIL_ADDRESS " & _
                "WHERE USER_SEQ_ID = " & tUID & " " & _
                "AND AN8_CUNO = " & tBCC & " " & _
                "AND EMAIL_ADDRESS IS NOT NULL"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        cboTeam.AddItem "--------------- History List ---------------"
        Do While Not rst.EOF
            cboTeam.AddItem LCase(Trim(rst.Fields("EMAIL_ADDRESS")))
            lstEmail.AddItem LCase(Trim(rst.Fields("EMAIL_ADDRESS")))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
End Sub

Public Function GetGPath(tGID As Long) As String
    Dim strSelect As String, tPath As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT GPATH " & _
                "FROM " & GFXMas & " " & _
                "WHERE GID = " & tGID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        tPath = Trim(rst.Fields("GPATH"))
    Else
        tPath = ""
    End If
    rst.Close: Set rst = Nothing
    
    GetGPath = tPath
End Function

Public Function GetTotalSize() As String
    Dim i As Integer
    Dim lTotal As Long
    
    lTotal = 0
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) Then
            lTotal = lTotal + lstPaths.ItemData(i)
        End If
    Next i
    GetTotalSize = Format(lTotal / 1000, "#,##0") & " KB"
'    Select Case lTotal
'        Case Is < 1000: GetTotalSize = format(lTotal, "##0") & " bytes"
'        Case Is < 1000000: GetTotalSize = format(lTotal / 1000, "##0") & " KB"
'        Case Else: GetTotalSize = format(lTotal / 1000000, "##0.00") & " mb"
'    End Select
End Function

Public Function GetGDesc(sText As String) As String
    Dim iDash As Integer
    Dim sDesc As String
    
    For iDash = 1 To Len(sText)
        If Mid(sText, iDash, 1) = "-" Then
            GetGDesc = Trim(Mid(sText, iDash + 2))
            Exit Function
        End If
    Next iDash
    
    GetGDesc = Trim(sText)
End Function

Public Sub GetFacilPhotos(tFCCD As Long, tGID As Long)
    Dim strSelect As String, sThumbPath As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iDel As Integer
    
'    i = -1
'    sThumbPath = "\\DETMSFS01\GPJAnnotator\Floorplans\FacilPho\Thumbs\"
    
    strSelect = "SELECT AB.ABALPH AS FACIL, " & _
                "GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT " & _
                "From ANNOTATOR.GFX_MASTER GM, " & F0101 & " AB " & _
                "Where GM.GID > 0 " & _
                "AND GM.AN8_CUNO = " & tFCCD & " " & _
                "AND GM.GTYPE = 66 " & _
                "AND GM.GSTATUS = 66 " & _
                "AND GM.AN8_CUNO = AB.ABAN8 " & _
                "ORDER BY GM.GDESC"
                
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close: Set rst = Nothing
        On Error Resume Next
        Me.Caption = "No Facility photos were found..."
        Screen.MousePointer = 0
        Exit Sub
    Else
'        sHDR = "Facility:  " & Trim(rst.Fields("FACIL"))
'        lblCaption.Caption = Trim(rst.Fields("FACIL"))
    End If
    Do While Not rst.EOF
        lstPlans.AddItem Trim(rst.Fields("GDESC"))
        lstPlans.ItemData(lstPlans.NewIndex) = rst.Fields("GID")
        lstPaths.AddItem Trim(rst.Fields("GPATH"))
        lstPaths.ItemData(lstPaths.NewIndex) = FileLen(Trim(rst.Fields("GPATH")))
        
        lstPlans.Selected(lstPlans.NewIndex) = CBool(rst.Fields("GID") = tGID)
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing

End Sub

