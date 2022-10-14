VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Floorplan Downloader"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Download Site"
      Height          =   435
      Left            =   3660
      TabIndex        =   8
      Top             =   7980
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Check Floorplan drawings you wish to download"
      Height          =   3795
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   3300
         Width           =   2235
      End
      Begin VB.ListBox lstPaths 
         Height          =   1425
         Left            =   4500
         MultiSelect     =   1  'Simple
         TabIndex        =   9
         Top             =   300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download Drawings"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2580
         TabIndex        =   7
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ListBox lstPlans 
         Height          =   1185
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   300
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Drive:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   3060
         Width           =   435
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   1515
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   4215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Viewing downloaded Floorplans"
      Height          =   3795
      Left            =   5100
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdViewer 
         Caption         =   "Viewer Download Link"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   3120
         Width           =   3435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   2595
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   3675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tSHYR As Integer
Dim tBCC As String
Dim tSHCD As Long, tDWGID As Long
Dim tFBCN As String
Dim tSHNM As String
Dim tFileType As String


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


'''Private Sub cmdClose_Click()
'''    Me.Height = 4485
'''    webDownload.Visible = False
'''End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdDownload_Click()
    Dim sTemp As String, sFolder As String, sChk As String, sPath As String, _
                sDrive As String
    Dim i As Integer
    
    Screen.MousePointer = 11
    On Error GoTo ErrorTrap
'''    sTemp = txtDrive.Text & ":\"
    sTemp = Left(Drive1.Drive, 1) & ":\"
    sChk = Dir(sTemp, vbDirectory)
    If sChk = "" Then
        MsgBox "Designated Drive was not found.", vbExclamation, "Drive not found..."
        Screen.MousePointer = 0
        frmDownloadError.Show 1
        Exit Sub
    End If
    sFolder = sTemp & "GPJ Annotator Download"
    sChk = Dir(sFolder, vbDirectory)
    If sChk = "" Then MkDir sFolder
    sFolder = sFolder & "\" & "Floorplans - " & Legalize(Trim(tFBCN))
    sChk = Dir(sFolder, vbDirectory)
    If sChk = "" Then MkDir sFolder
    
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) = True Then
            sPath = sFolder & "\" & Legalize(lblCaption) & " [" & lstPlans.List(i) & "]." & tFileType
            FileCopy lstPaths.List(i), sPath
            lstPlans.Selected(i) = False
        End If
    Next i
    
    Screen.MousePointer = 0
    MsgBox "File(s) Copied", vbInformation, "File Download Successful..."
Exit Sub
ErrorTrap:
    Screen.MousePointer = 0
    MsgBox "Error:  " & Err.Description, vbExclamation, "File Not Copied..."
    frmDownloadError.Show 1
    
End Sub

Private Sub cmdViewer_Click()
    Dim MessHdr As String, MessBody As String
    Dim myNotes As New Domino.NotesSession
    Dim myDB As New Domino.NotesDatabase
    Dim myItem  As Object ''' NOTESITEM
    Dim myDoc As Object ''' NOTESDOCUMENT
    Dim myRichText As Object ' NOTESRICHTEXTITEM
    Dim myReply  As Object ''' NOTESITEM
    Dim Address As String
    
    Screen.MousePointer = 11
    MessHdr = "AutoDesk Volo Viewer Express"
    MessBody = "http://www3.autodesk.com/adsk/item/0,,837421-123112,00.html"
        
'    myNotes.Initialize
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
    Set myDoc = myDB.CreateDocument
    If sNOTESID = "GANNOTAT" Then Call myDoc.ReplaceItemValue("Principal", LogName)
    Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
    If sNOTESID = "GANNOTAT" Then Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
    Set myRichText = myDoc.CreateRichTextItem("Body")
    myRichText.AppendText MessBody
    myDoc.AppendItemValue "SENDTO", LogAddress
'''    myDoc.SaveMessageOnSend = True
    
    Call myDoc.Send(False, LogAddress)
    
    Set myReply = Nothing
    Set myRichText = Nothing
    Set myItem = Nothing
    Set myDoc = Nothing
    Set myDB = Nothing
    Set myNotes = Nothing
    
    Screen.MousePointer = 0
'''
'''    Me.Height = 8815
'''    webDownload.Navigate "http://support01.autodesk.com/knowledgebase/html/148490.htm"
'''    webDownload.Visible = True
'''
End Sub

Private Sub Drive1_Change()
'    Debug.Print Drive1.Drive
    If UCase(Left(Drive1.Drive, 1)) = "U" Or UCase(Left(Drive1.Drive, 1)) = "V" Then
        MsgBox "You've chosen an invalid drive.  Please, make another selection.", vbExclamation, "Sorry..."
        Exit Sub
    End If
    cmdDownload.Enabled = True
End Sub

Private Sub Form_Load()
    Dim strSelect As String, sCheck As String, sPDF As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    If tDWGID = 0 Then
        lblCaption.Caption = tSHYR & "  " & tSHNM
        Call GetShowPlans(frmDownload, tFileType, tSHYR, tSHCD)
    Else
        lblCaption.Caption = tFBCN & "  -  " & tSHYR & "  " & tSHNM
        Call GetDrawings(frmDownload, tFileType, tDWGID, tSHYR, tSHCD)
    End If
    
    lblPrompt = "Selected drawings will be downloaded.  " & _
                "They will be placed in a folder named 'GPJ Annotator Download' " & _
                "on the drive noted below." & vbNewLine & vbNewLine & _
                "To proceed, click the 'Download Drawings' button." & _
                vbNewLine & vbNewLine & _
                "The download option will not work on a Mac."
'''    webDownload.Top = 120: webDownload.Left = 180
    Me.Height = 4485
    Select Case tFileType
        Case "DWF"
            Me.Width = 5160 '''10095
            Me.Caption = Me.Caption & ": DWF Files"
        Case "PDF"
            Me.Width = 5160
            Me.Caption = Me.Caption & ": PDF Files"
    End Select
    
'''    Call GetDrawings(frmDownload, tFileType, tDWGID, tSHYR, tSHCD)
            
    Label2.Caption = "To view downloaded Floorplans, the AutoDesk Volo Viewer is required.  " & _
                "The viewer is free, and is available from AutoDesk's web site.  To recieve an emailed link " & _
                "to the AutoDesk site, click the button below."
    
    For i = 0 To Drive1.ListCount - 1
        Debug.Print Drive1.List(i)
        If UCase(Left(Drive1.List(i), 1)) = "C" Then Drive1.Drive = Drive1.List(i)
    Next i
    
End Sub

Private Sub lstPlans_ItemCheck(Item As Integer)
    Dim i As Integer
    Dim bChecked As Boolean
    
    bChecked = False
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) = True Then bChecked = True
    Next i
    If Drive1.Drive = "" Then bChecked = False
    cmdDownload.Enabled = bChecked
End Sub

'''Private Sub txtDrive_KeyPress(KeyAscii As Integer)
'''    Debug.Print KeyAscii
'''    If KeyAscii >= 97 And KeyAscii <= 122 Then
'''        KeyAscii = KeyAscii - 32
'''    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then
'''        KeyAscii = KeyAscii
'''    ElseIf KeyAscii <> 8 Then
'''        KeyAscii = 0
'''    End If
'''End Sub

Public Function Legalize(sName As String) As String
    sName = Replace(sName, "\", "-")
    sName = Replace(sName, "/", "-")
    sName = Replace(sName, ":", "-")
    sName = Replace(sName, "*", "")
    sName = Replace(sName, """", "'")
    sName = Replace(sName, "|", "-")
    sName = Replace(sName, "<", "")
    sName = Replace(sName, ">", "")
    sName = Replace(sName, "?", "")
    
    Legalize = sName
End Function
