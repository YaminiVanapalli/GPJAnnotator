VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComments.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":0EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":20B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":298C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":2CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":2FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComments.frx":32DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   300
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picAttach 
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   720
      ScaleHeight     =   555
      ScaleWidth      =   3615
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   3675
      Begin VB.TextBox txtAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   20
         ToolTipText     =   "Click to Cancel Attachment"
         Top             =   60
         Width           =   3015
      End
      Begin VB.Image imgFormat 
         Height          =   495
         Left            =   60
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Cancel Attachment"
         Top             =   30
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAttach 
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmComments.frx":35F4
      MousePointer    =   99  'Custom
      Picture         =   "frmComments.frx":38FE
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Click to Attach Document"
      Top             =   3120
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxComments 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   4140
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      MouseIcon       =   "frmComments.frx":3C08
   End
   Begin VB.PictureBox Picture2 
      Height          =   3435
      Left            =   4440
      ScaleHeight     =   3375
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   300
      Width           =   4155
      Begin VB.ListBox lstNotifyEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Index           =   1
         ItemData        =   "frmComments.frx":3F22
         Left            =   2340
         List            =   "frmComments.frx":3F24
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ListBox lstShortname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Index           =   1
         ItemData        =   "frmComments.frx":3F26
         Left            =   3420
         List            =   "frmComments.frx":3F28
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin TabDlg.SSTab sstEmail 
         Height          =   2115
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3731
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Team Members"
         TabPicture(0)   =   "frmComments.frx":3F2A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lstOptNotify(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "GPJ Personnel"
         TabPicture(1)   =   "frmComments.frx":3F46
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lstOptNotify(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.ListBox lstOptNotify 
            Height          =   1635
            Index           =   1
            ItemData        =   "frmComments.frx":3F62
            Left            =   60
            List            =   "frmComments.frx":3F69
            Style           =   1  'Checkbox
            TabIndex        =   15
            Top             =   360
            Width           =   3975
         End
         Begin VB.ListBox lstOptNotify 
            Height          =   1635
            Index           =   0
            ItemData        =   "frmComments.frx":3F7B
            Left            =   -74940
            List            =   "frmComments.frx":3F82
            Style           =   1  'Checkbox
            TabIndex        =   14
            Top             =   390
            Width           =   3975
         End
      End
      Begin VB.ListBox lstShortname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Index           =   0
         ItemData        =   "frmComments.frx":3F94
         Left            =   1140
         List            =   "frmComments.frx":3F96
         TabIndex        =   12
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdSaveComment 
         Caption         =   "Save Comment"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2820
         Width           =   1785
      End
      Begin VB.ListBox lstNotifyEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   225
         Index           =   0
         ItemData        =   "frmComments.frx":3F98
         Left            =   60
         List            =   "frmComments.frx":3F9A
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdClearComment 
         Caption         =   " Clear Comment"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2220
         Width           =   1785
      End
   End
   Begin VB.TextBox txtCommentAdd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2715
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   4275
   End
   Begin VB.Image imgPrint 
      Height          =   240
      Left            =   8280
      Picture         =   "frmComments.frx":3F9C
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label lblMess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<-- Click to Attach a Document to a Comment Entry.  Once selected, click ""Save Comment"" to attach.  A Text Comment is optional."
      Height          =   585
      Left            =   840
      TabIndex        =   21
      Top             =   3150
      Width           =   3555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H80000002&
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   8475
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   4440
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   2820
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   2475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTeam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Notification Team:"
      Height          =   195
      Left            =   4440
      TabIndex        =   7
      Top             =   60
      Width           =   1710
   End
   Begin VB.Label lblCommEditor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Comment:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1065
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "mnuDelete"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRefID As Long, tSHCD As Long, lRef As Long
Dim iRef As Integer
Dim CommentSaved As Boolean
Dim bComm As Boolean
Dim iType As Integer, tSHYR As Integer
Dim tBCC As String, tFBCN As String, sMessPath As String, sMessSub As String, _
            sGPath As String, sDPath As String
Dim sForm As String, sTable As String
Dim bCheckForAttachment As Boolean


Public Property Get PassForm() As String
    PassForm = sForm
End Property
Public Property Let PassForm(ByVal vNewValue As String)
    sForm = vNewValue
End Property

Public Property Get PassREFID() As String
    PassREFID = lRefID
End Property
Public Property Let PassREFID(ByVal vNewValue As String)
    lRefID = vNewValue
End Property

Public Property Get PassTable() As String
    PassTable = sTable
End Property
Public Property Let PassTable(ByVal vNewValue As String)
    sTable = vNewValue
End Property

Public Property Get PassIType() As Integer
    PassIType = iType
End Property
Public Property Let PassIType(ByVal vNewValue As Integer) '/// 0=Floorplan,1=Graphics,2=Const Dwg \\\
    iType = vNewValue
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

Public Property Get PassMessPath() As String
    PassMessPath = sMessPath
End Property
Public Property Let PassMessPath(ByVal vNewValue As String)
    sMessPath = vNewValue
End Property

Public Property Get PassMessSub() As String
    PassMessSub = sMessSub
End Property
Public Property Let PassMessSub(ByVal vNewValue As String)
    sMessSub = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = tSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    tSHCD = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = tSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    tSHYR = vNewValue
End Property

Public Property Get PassGPath() As String
    PassGPath = sGPath
End Property
Public Property Let PassGPath(ByVal vNewValue As String)
    sGPath = vNewValue
End Property

Public Property Get PassDPath() As String
    PassDPath = sDPath
End Property
Public Property Let PassDPath(ByVal vNewValue As String)
    sDPath = vNewValue
End Property


Private Sub cmdAttach_Click()
    Dim sFormat As String, sDesc As String
    Dim iDot As Integer, i As Integer
    
TryAgain:
    cdl1.FileName = ""
    cdl1.DefaultExt = ".doc"
'''    cdl1.Filter = "Support Files (*.htm;*.html;*.pdf)|*.htm;*.html;*.pdf"
    cdl1.Filter = "Document Files (*.doc;*.xls;*.pdf;*.txt;*.htm;*.html;*.jpg)|*.doc;*.xls;*.pdf;*.txt;*.htm;*.html;*.jpg"
    cdl1.Flags = cdlOFNHideReadOnly
    cdl1.Flags = cdlOFNFileMustExist
    
    
'''    If def_Path <> "" Then cdl1.InitDir = def_Path
    
    cdl1.DialogTitle = "Select Attachment Document..."
    
    cdl1.ShowOpen
    If cdl1.FileName = "" Then Exit Sub
    
    Select Case UCase(Right(cdl1.FileName, 3))
        Case "DOC": imgFormat.Picture = ImageList1.ListImages(1).Picture
        Case "XLS": imgFormat.Picture = ImageList1.ListImages(2).Picture
        Case "PDF": imgFormat.Picture = ImageList1.ListImages(3).Picture
        Case "HTM": imgFormat.Picture = ImageList1.ListImages(4).Picture
        Case "TML": imgFormat.Picture = ImageList1.ListImages(5).Picture
        Case "TXT": imgFormat.Picture = ImageList1.ListImages(6).Picture
        Case "JPG": imgFormat.Picture = ImageList1.ListImages(7).Picture
        Case "BMP": imgFormat.Picture = ImageList1.ListImages(8).Picture
    End Select
    
    txtAttach.Text = cdl1.FileName
    picAttach.Visible = True
    cmdSaveComment.Enabled = True
    cmdClearComment.Enabled = True
    CommentSaved = False
    
'''    sDesc = cdl1.FileTitle
'''    For i = Len(sDesc) - 1 To 0 Step -1
'''        If Mid(sDesc, i, 1) = "." Then
'''            txtFormat.Text = UCase(Mid(sDesc, i + 1))
'''            txtDesc.Text = Left(sDesc, i - 1)
'''            web1.Navigate2 cdl1.FileName
'''            web1.Visible = True
'''            cmdSave.Tag = "0"
'''            Exit For
'''        End If
'''    Next i

End Sub

Private Sub cmdClearComment_Click()
    Dim Resp As VbMsgBoxResult
    If Trim(txtCommentAdd.Text) <> "" Then
        If txtAttach.Text = "" Then
            Resp = MsgBox("Are you certain you want to clear this unsaved Comment?", _
                        vbExclamation + vbYesNoCancel, "Just Checking...")
            If Resp = vbYes Then
                txtCommentAdd.Text = ""
                cmdSaveComment.Enabled = False
                cmdClearComment.Enabled = False
                CommentSaved = True
            End If
        Else
            Resp = MsgBox("Do you want to clear the Attachment also?", _
                        vbExclamation + vbYesNoCancel, "Just Checking...")
            If Resp = vbYes Then
                txtCommentAdd.Text = ""
                picAttach.Visible = False
                txtAttach.Text = ""
                imgFormat.Picture = LoadPicture("")
                cmdSaveComment.Enabled = False
                cmdClearComment.Enabled = False
                CommentSaved = True
            ElseIf Resp = vbNo Then
                txtCommentAdd.Text = ""
            End If
        End If
    Else
        txtCommentAdd.Text = ""
        If txtAttach.Text <> "" Then
            Resp = MsgBox("Do you want to clear the Attachment?", _
                        vbExclamation + vbYesNo, "Just Checking...")
            If Resp = vbYes Then
                picAttach.Visible = False
                txtAttach.Text = ""
                imgFormat.Picture = LoadPicture("")
                cmdSaveComment.Enabled = False
                cmdClearComment.Enabled = False
                CommentSaved = True
            End If
        End If
    End If
    txtCommentAdd.SetFocus
End Sub

Private Sub cmdSaveComment_Click()
    Dim sNewComm As String, strInsert As String, sRefSource As String
    Dim rstL As ADODB.Recordset
    Dim lCOMMID As Long, lSDID As Long
    Dim i As Integer, iCnt As Integer, iDot As Integer, iSlash As Integer
    Dim sFormat As String, sDesc As String
    
    If txtCommentAdd.Text <> "" Or txtAttach.Text <> "" Then
        '///// GET NEW COMMID \\\\\
        Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
        lCOMMID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        
        Conn.BeginTrans
        If txtAttach.Text <> "" Then
            For iDot = Len(txtAttach.Text) To 0 Step -1
                If Mid(txtAttach.Text, iDot, 1) = "." Then
                    sFormat = UCase(Mid(txtAttach.Text, iDot + 1))
                    Exit For
                End If
            Next iDot
                
            For iSlash = iDot To 0 Step -1
                If Mid(txtAttach.Text, iSlash, 1) = "\" Then
                    sDesc = Mid(txtAttach.Text, iSlash + 1, iDot - iSlash - 1)
                    Exit For
                End If
            Next iSlash
            
            Set rstL = Conn.Execute("SELECT " & GFXSeq & ".NEXTVAL FROM DUAL")
            lSDID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            
            strInsert = "INSERT INTO ANNOTATOR.GFX_SUPDOC " & _
                        "(SUPDOC_ID, AN8_CUNO, SUPDOCDESC, SUPDOCFORMAT, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lSDID & ", " & CLng(tBCC) & ", " & _
                        "'" & DeGlitch(Left(sDesc, 50)) & "', " & _
                        "'" & DeGlitch(sFormat) & "', " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            
            FileCopy txtAttach.Text, sSupDocPath & lSDID & "." & LCase(sFormat)
        Else
            lSDID = 0
        End If
        
        If txtCommentAdd.Text = "" Then
            sNewComm = "[Attachment Posting: " & sDesc & "." & LCase(sFormat) & "]"
        Else
            If txtAttach.Text <> "" Then
                sNewComm = DeGlitch(txtCommentAdd.Text) & "  [Attachment Posting: " & _
                            sDesc & "." & LCase(sFormat) & "]"
            Else
                sNewComm = DeGlitch(txtCommentAdd.Text)
            End If
        End If
        
        strInsert = "INSERT INTO " & ANOComment & " " & _
                    "(COMMID, REFID, SUPDOC_ID, REFSOURCE, ANO_COMMENT, " & _
                    "COMMUSER, COMMDATE, COMMSTATUS, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lCOMMID & ", " & lRefID & ", " & lSDID & ", '" & sTable & "', '" & sNewComm & "', " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
        Conn.CommitTrans
        
        lblHeader = " LAST COMMENT BY: " & LogName & " on " & _
                    Format(Now, "DD-MMM-YYYY") & " (" & _
                    Format(Now, "h:mm am/pm") & " Detroit Time)"
        lblCommEditor.Caption = "Add Comment:"
        CommentSaved = True
        
        iCnt = 0
        For i = 0 To lstOptNotify(0).ListCount - 1
            If lstOptNotify(0).Selected(i) = True Then
                iCnt = iCnt + 1
            End If
        Next i
        If sstEmail.TabVisible(1) Then
            For i = 0 To lstOptNotify(1).ListCount - 1
                If lstOptNotify(1).Selected(i) = True Then
                    iCnt = iCnt + 1
                End If
            Next i
        End If
        
        Select Case sForm
            Case "frmAnnotator"
                frmAnnotator.imgComm.Picture = frmAnnotator.imgMail(1).Picture
                frmAnnotator.imgComm.Visible = True
                If iCnt > 0 Then Call CommAlert("shyr=" & CStr(tSHYR) & "&shcd=" & CStr(tSHCD) & "&cuno=" & CStr(CLng(tBCC)), sNewComm)
            Case "frmGraphics"
                frmGraphics.imgComm.Picture = frmGraphics.imgMail(1).Picture
                frmGraphics.imgComm.Visible = True
                If iCnt > 0 Then Call CommAlert("gid=" & CStr(lRefID), sNewComm)
            Case "frmConst"
                frmConst.imgComm.Picture = frmConst.imgMail(1).Picture
                frmConst.imgComm.Visible = True
                If iCnt > 0 Then Call CommAlert("", sNewComm)
        End Select
        
        picAttach.Visible = False
        txtAttach.Text = ""
        imgFormat.Picture = LoadPicture("")
        
        LoadRTX
    End If
End Sub

Private Sub flxComments_Click()
    Dim strSelect As String, sFile As String, sHDR As String, sDFile As String
    Dim rst As ADODB.Recordset
    Dim lSDID As Long
    
    If Not bCheckForAttachment Then Exit Sub
    If flxComments.Rows = 1 Then Exit Sub
    
    lSDID = CLng(flxComments.TextMatrix(flxComments.RowSel, 5))
    If lSDID > 0 Then
        Screen.MousePointer = 11
        strSelect = "SELECT SUPDOCDESC, SUPDOCFORMAT " & _
                    "FROM ANNOTATOR.GFX_SUPDOC " & _
                    "WHERE SUPDOC_ID = " & lSDID
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sFile = sSupDocPath & CStr(lSDID) & "." & LCase(Trim(rst.Fields("SUPDOCFORMAT")))
            sHDR = "Comment Attachment:  " & rst.Fields("SUPDOCDESC")
            sDFile = Trim(rst.Fields("SUPDOCDESC")) & "." & LCase(Trim(rst.Fields("SUPDOCFORMAT")))
            rst.Close: Set rst = Nothing
            
            frmHTMLViewer.PassDFile = sDFile
            frmHTMLViewer.PassFile = sFile
            frmHTMLViewer.PassHDR = sHDR
            frmHTMLViewer.PassFrom = Me.Name
            frmHTMLViewer.Show 1, Me
            
        Else
            rst.Close: Set rst = Nothing
            MsgBox "File Not Found", vbExclamation, "Sorry..."
        End If
        
        Screen.MousePointer = 0
    End If
End Sub

Private Sub flxComments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bCheckForAttachment = False
    If X > flxComments.ColPos(2) And X < flxComments.ColPos(3) Then
        bCheckForAttachment = True
    ElseIf Button = vbLeftButton And bTeamMember And flxComments.RowSel > 0 Then
        Select Case iType
            Case 0: If Not bPerm(19) Then Exit Sub
            Case 1: If Not bPerm(28) Then Exit Sub
            Case 2: If Not bPerm(36) Then Exit Sub
        End Select
        lRef = flxComments.TextMatrix(flxComments.RowSel, 4)
        iRef = flxComments.RowSel
        Me.PopupMenu mnuDelete
    
    End If
End Sub

Private Sub flxComments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > flxComments.ColPos(2) And X < flxComments.ColPos(3) Then
        flxComments.MousePointer = flexCustom
    Else
        flxComments.MousePointer = flexDefault
    End If
End Sub

Private Sub Form_Load()
    Dim strSelect As String, sEmployer As String, sColumn As String
    Dim rst As ADODB.Recordset
    Dim iCheck As Integer
    
    bTeamMember = False
    Select Case iType
        Case 0: sColumn = "RECIPIENT_FLAG0"
        Case 1: sColumn = "RECIPIENT_FLAG1"
        Case 2: sColumn = "RECIPIENT_FLAG2"
    End Select
    
    '///// FIRST, GET TEAM \\\\\
    '///// SEE IF CLIENT-SHOW TEAM EXISTS \\\\\
    strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, " & _
                "U.EMPLOYER, R." & sColumn & " " & _
                "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                "WHERE T.AN8_CUNO = " & CLng(tBCC) & " " & _
                "AND T.AN8_SHCD = " & tSHCD & " " & _
                "AND T.MCU IS NULL " & _
                "AND T.TEAM_ID = R.TEAM_ID " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, " & _
                    "U.EMPLOYER, R." & sColumn & " " & _
                    "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                    "WHERE T.AN8_CUNO = " & CLng(tBCC) & " " & _
                    "AND T.AN8_SHCD IS NULL " & _
                    "AND T.MCU IS NULL " & _
                    "AND T.TEAM_ID = R.TEAM_ID " & _
                    "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                    "AND U.USER_STATUS > 0 " & _
                    "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Set rst = Conn.Execute(strSelect)
        If rst.EOF Then
            rst.Close
            Set rst = Nothing
            MsgBox "No Email Notification Team has been setup for " & tFBCN & ".", vbExclamation, "Sorry..."
            lstOptNotify(0).Clear
            lstNotifyEmail(0).Clear
'''''            Unload Me
            Exit Sub
        Else
            lblTeam = "Client-based Email Notification Team:"
        End If
    Else
        lblTeam = "Client/Show-based Email Notification Team:"
    End If
    
    lstOptNotify(0).Clear
    lstNotifyEmail(0).Clear
    lstShortname(0).Clear
    
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) <> "GPJ" Then
            sEmployer = " (" & Trim(rst.Fields("EMPLOYER")) & ")"
        Else
            sEmployer = ""
        End If
        If StrConv(Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")), vbProperCase) = LogName Then
            bTeamMember = True
        End If
        lstOptNotify(0).AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & _
                    UCase(Trim(rst.Fields("NAME_LAST"))) & sEmployer
        lstOptNotify(0).ItemData(lstOptNotify(0).NewIndex) = rst.Fields(sColumn)
        lstOptNotify(0).Selected(lstOptNotify(0).NewIndex) = CBool(rst.Fields(sColumn) * -1)
        lstNotifyEmail(0).AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstShortname(0).AddItem LCase(Trim(rst.Fields("NAME_LOGON")))
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    Call UnselectSelf(LogName, Me.lstOptNotify(0))
    
    '///// GET GPJ PERSONNEL \\\\\
    If bGPJ Then
        sstEmail.TabVisible(1) = True
        Call GetGPJEmail
    Else
        sstEmail.TabVisible(1) = False
    End If
    
    flxComments.ColWidth(0) = (flxComments.Width - 240) * 0.15
    flxComments.ColWidth(1) = (flxComments.Width - 240) * 0.2
    flxComments.ColWidth(2) = (flxComments.Width - 240) * 0.06
    flxComments.ColWidth(3) = (flxComments.Width - 240) * 0.59
    flxComments.ColWidth(4) = 0
    flxComments.ColWidth(5) = 0
    flxComments.ColAlignment(0) = 3
    flxComments.ColAlignment(1) = 3
    flxComments.ColAlignment(2) = 3
    flxComments.ColAlignment(3) = 0: flxComments.FixedAlignment(3) = 3
    flxComments.TextMatrix(0, 0) = "Date"
    flxComments.TextMatrix(0, 1) = "Poster"
    flxComments.Col = 2: flxComments.Row = 0
    Set flxComments.CellPicture = ImageList1.ListImages(9).Picture
    
'    flxComments.TextMatrix(0, 2) = "Attach"
    flxComments.TextMatrix(0, 3) = "Comment"
    flxComments.WordWrap = True
    lblCheck(0).Width = flxComments.ColWidth(0)
    lblCheck(1).Width = flxComments.ColWidth(1)
    lblCheck(2).Width = flxComments.ColWidth(3)
    lblCheck(0).Font.Size = flxComments.Font.Size
    lblCheck(1).Font.Size = flxComments.Font.Size
    lblCheck(2).Font.Size = flxComments.Font.Size
    
    If lRefID <> 0 Then LoadRTX
    
    Select Case iType
        Case 0
            Me.Caption = frmAnnotator.lblWelcome & " Comments"
            If Not bPerm(18) Then
                lblHeader.Top = 30
                flxComments.Top = 300
                flxComments.Height = 6915
            End If
''            If sDPath <> "" Then
''                Call PopImage(App.Path & "\DefaultDWF.jpg")
''            End If
        Case 1
'            Me.Caption = frmGraphics.lblWelcome & " Comments"
            Me.Caption = sMessPath & " Comments"
            If Not bPerm(27) Then
                lblHeader.Top = 30
                flxComments.Top = 330
                flxComments.Height = 6885
            End If
''            If sGPath <> "" Then Call PopImage(sGPath)
            
        Case 2
            Me.Caption = frmConst.lblWelcome & " Comments"
            If Not bPerm(35) Then
                lblHeader.Top = 30
                flxComments.Top = 330
                flxComments.Height = 6885
            End If
    End Select
    
    
    
End Sub

Public Function DeDblApost(txt As String) As String
    Dim Pos As Integer
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(1, txt, "''")
        If Pos <> 0 Then txt = Left(txt, Pos - 1) & Chr(34) & Mid(txt, Pos + 2)
    Loop
    DeDblApost = txt
End Function

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    flxComments.Height = Me.ScaleHeight - flxComments.Top - flxComments.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sNewComm As String, strInsert As String
    Dim rstL As ADODB.Recordset
    Dim lCOMMID As Long
    Dim Resp As VbMsgBoxResult
    
    bCommentsOpen = False
    Select Case iType
        Case 0: If Not bPerm(18) Then Exit Sub
        Case 1: If Not bPerm(27) Then Exit Sub
        Case 2: If Not bPerm(35) Then Exit Sub
    End Select
    If CommentSaved = False Then
        Resp = MsgBox("Would you like to save your Comment?", vbQuestion + vbYesNoCancel, "Comment was Changed...")
        If Resp = vbYes Then
            sNewComm = DeGlitch(txtCommentAdd.Text)
        
            '///// GET NEW COMMID \\\\\
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lCOMMID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            
            strInsert = "INSERT INTO " & ANOComment & " " & _
                        "(COMMID, REFID, REFSOURCE, ANO_COMMENT, " & _
                        "COMMUSER, COMMDATE, COMMSTATUS, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lCOMMID & ", " & lRefID & ", '" & sTable & "', '" & sNewComm & "', " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            CommentSaved = True
        ElseIf Resp = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

'''Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.FontBold = False
'''End Sub

'''Private Sub lblClose_Click()
'''    Unload Me
'''End Sub

'''Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.FontBold = True
'''End Sub

'''Private Sub lstOptNotify_ItemCheck(Item As Integer)
''''''    Dim i As Integer
''''''    Debug.Print lstOptNotify(0).Selected(Item)
''''''    If lstOptNotify(0).ItemData(Item) = 1 And lstOptNotify(0).Selected(Item) = False Then
''''''        lstOptNotify(0).Selected(Item) = True
''''''        MsgBox "The person selected is a Mandatory Recipient.", _
''''''                    vbExclamation, "Member cannot be Unchecked..."
''''''    End If
''''''    If txtCommentAdd.Text <> "" Then
''''''        For i = 0 To lstOptNotify(0).ListCount - 1
''''''            If lstOptNotify(0).Selected(i) = True Then
''''''                cmdSaveComment.Enabled = True
''''''                GoTo FoundOne
''''''            End If
''''''        Next
''''''        cmdSaveComment.Enabled = False
''''''FoundOne:
''''''    Else
''''''        cmdSaveComment.Enabled = False
''''''    End If
'''End Sub

Private Sub imgFormat_Click()
    Dim Resp As VbMsgBoxResult
    Resp = MsgBox("Are you certain you want to clear this Attachment?", vbYesNo, "Just Checking...")
    If Resp = vbYes Then
        picAttach.Visible = False
        txtAttach.Text = ""
        imgFormat.Picture = LoadPicture("")
        If txtCommentAdd.Text = "" Then
            cmdSaveComment.Enabled = False
            CommentSaved = True
        End If
    End If
End Sub

Private Sub imgPrint_Click()
    frmPrintOut.PassHDR = Me.Caption
    frmPrintOut.PassPATH = ConvertToHTML(Me.Caption, Me.flxComments)
    frmPrintOut.Show 1, Me
End Sub

Private Sub lstOptNotify_ItemCheck(Index As Integer, Item As Integer)
    Dim i As Integer
    
    If Index <> 0 Then Exit Sub
    
    If lstOptNotify(Index).ItemData(Item) = 1 And lstOptNotify(Index).Selected(Item) = False Then
        lstOptNotify(Index).Selected(Item) = True
        MsgBox "The person selected is a Mandatory Recipient.", _
                    vbExclamation, "Member cannot be Unchecked..."
    End If
    If txtCommentAdd.Text <> "" Then
        For i = 0 To lstOptNotify(Index).ListCount - 1
            If lstOptNotify(Index).Selected(i) = True Then
                cmdSaveComment.Enabled = True
                GoTo FoundOne
            End If
        Next
        cmdSaveComment.Enabled = False
FoundOne:
    Else
        cmdSaveComment.Enabled = False
    End If

End Sub

Private Sub mnuDeleteRow_Click()
    Dim strUpdate As String
    On Error Resume Next
    strUpdate = "UPDATE " & ANOComment & " " & _
                "SET COMMSTATUS = -1, " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE COMMID = " & lRef
    Conn.Execute (strUpdate)
    If Err = 0 Then
        If flxComments.Rows = 2 Then flxComments.Rows = 1 Else flxComments.RemoveItem (iRef)
        If flxComments.Rows = 1 Then
            ''RESET ICON''
            Select Case sForm
                Case "frmAnnotator"
                    frmAnnotator.imgComm.Picture = frmAnnotator.imgMail(0).Picture
                    frmAnnotator.imgComm.ToolTipText = "There are no saved Comments."
                Case "frmGraphics"
                    frmGraphics.imgComm.Picture = frmGraphics.imgMail(0).Picture
                    frmGraphics.imgComm.ToolTipText = "There are no saved Comments."
                Case "frmConst"
                    frmConst.imgComm.Picture = frmConst.imgMail(0).Picture
                    frmConst.imgComm.ToolTipText = "There are no saved Comments."
            End Select
        End If
    Else
        MsgBox "Error:  " & Err.Description, vbExclamation, "Error Encountered..."
    End If
End Sub

Private Sub txtAttach_Click()
    Call imgFormat_Click
End Sub


'''Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.FontBold = False
'''End Sub

Private Sub txtCommentAdd_Change()
    If txtCommentAdd.Text <> "" Then
        cmdClearComment.Enabled = True
        cmdSaveComment.Enabled = True
        CommentSaved = False
    Else
        If txtAttach.Text = "" Then
            cmdClearComment.Enabled = False
            cmdSaveComment.Enabled = False
            CommentSaved = True
        End If
    End If
End Sub

Public Sub CommAlert(sPass As String, sComment As String)
    Dim MessBody As String, MessHdr As String
    Dim i As Integer, iAdd As Integer
    Dim sList As String, sIntro As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim Address() As String
    Dim Shortname() As String
    Dim sPreLink As String, sPostLink As String, sFullLink As String
    
    
    Dim MailMan As New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
    
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
    
    Dim Email As New ChilkatEmail2
    
    Email.FromAddress = LogAddress
    Email.fromName = LogName
    
    
    
    
    
'    sPreLink = "http://gpjapps02.gpjco.com/LinksToAnno.asp?name_logon="
'    sPostLink = "&" & sPass
    Select Case iType
        Case 0
            sIntro = "A Floor Plan Annotator"
        Case 1
            sIntro = "A Graphics"
        Case 2
            sIntro = "A Construction Drawing"
    End Select
    
    
    sList = "": iAdd = -1
    For i = 0 To lstOptNotify(0).ListCount - 1
        If lstOptNotify(0).Selected(i) = True Then
            iAdd = iAdd + 1
            ReDim Preserve Address(iAdd)
            Address(iAdd) = lstNotifyEmail(0).List(i)
            
            Email.AddTo Address(iAdd), Address(iAdd)
            
            ReDim Preserve Shortname(iAdd)
            Shortname(iAdd) = lstShortname(0).List(i)
            sList = sList & vbTab & lstOptNotify(0).List(i) & vbNewLine
        End If
    Next i
    If sstEmail.TabVisible(1) Then
        For i = 0 To lstOptNotify(1).ListCount - 1
            If lstOptNotify(1).Selected(i) = True Then
                iAdd = iAdd + 1
                ReDim Preserve Address(iAdd)
                Address(iAdd) = lstNotifyEmail(1).List(i)
                
                Email.AddTo Address(iAdd), Address(iAdd)
                
                ReDim Preserve Shortname(iAdd)
                Shortname(iAdd) = lstShortname(1).List(i)
                sList = sList & vbTab & lstOptNotify(1).List(i) & vbNewLine
            End If
        Next i
    End If
    
    MessHdr = sMessPath & " Comment Entry"
    MessBody = sIntro & " comment has been saved by " & _
                LogName & " for " & sMessPath & " (" & sMessSub & ")." & vbNewLine & vbNewLine & _
                "The following Team members are being alerted through this email:" & _
                vbNewLine & vbNewLine & sList & vbNewLine & _
                "Below is the saved dialog:" & vbNewLine & String(75, "=") & vbNewLine & vbNewLine & _
                sComment & vbNewLine & vbNewLine & String(75, "=") & vbNewLine & _
                vbNewLine & _
                vbNewLine & _
                LogName
        
    Email.subject = MessHdr
    Email.Body = MessBody
    
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
    
    
    
    '///// EXECUTE E-MAIL \\\\\
''''    Dim myNotes As New Domino.NotesSession ''NotesSession '' Domino.NotesSession
''''    Dim myDB As New Domino.NotesDatabase '' NotesDatabase '' Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    
    
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
    
    
    
    
    
'''    On Error Resume Next
'''    If bDebug Then MsgBox "sNOTESID = " & sNOTESID
'''    If sNOTESID = "GANNOTAT" Then
'''        If bDebug Then MsgBox "About to init MyNotes"
'''        myNotes.Initialize (sNOTESPASSWORD)
'''        If bDebug Then MsgBox "MyNotes initialized (err=" & Err.Number & ")"
'''    Else
'''        If sNOTESPASSWORD = "" Then
'''            ''GET PASSWORD''
'''TryPWAgain:
'''            frmGetPassword.Show 1, Me
'''            Select Case sNOTESPASSWORD
'''                Case "_CANCEL"
'''                    sNOTESPASSWORD = ""
'''                    MsgBox "No email will be sent", vbExclamation, "User Canceled..."
'''                    Set myNotes = Nothing
'''                    Set myDB = Nothing
'''                Case Else
'''                    Err.Clear
'''                    myNotes.Initialize (sNOTESPASSWORD)
'''                    If Err Then
'''                        Err.Clear
'''                        GoTo TryPWAgain
'''                    End If
'''            End Select
'''        Else
'''            myNotes.Initialize (sNOTESPASSWORD)
'''        End If
'''    End If
    
    
''''/// ACTIVATE FOR CITRIX \\\
'''    If bDebug Then MsgBox "MyDB about to be opened"
'''    Set myDB = myNotes.GetDatabase(strMailSrvr, strMailFile)
'''    If bDebug Then MsgBox "MyDB opened (err=" & Err.Number & ")"
''''''    Set myDB = myNotes.GETDATABASE(strMailSrvr, strMailFile)
    
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
''''''        If sNOTESID = "GANNOTAT" Then Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
''''''        If sNOTESID = "GANNOTAT" Then Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'
'        Set myRichText = myDoc.CreateRichTextItem("Body")
'        Select Case iType
'            Case 0, 1
'                myRichText.AppendText MessBody '' & vbNewLine & _
''                            "Click for direct access to file:  " & _
''                            sPreLink & Shortname(i) & sPostLink & _
''                            vbNewLine & vbNewLine & sLink_Disclaimer
'            Case Else
'                myRichText.AppendText MessBody & vbNewLine & vbNewLine & sLink
'        End Select
'        myRichText.AddNewLine 3
'        myRichText.AppendText LogName
'
'        myDoc.AppendItemValue "SENDTO", Address(i)
'
'        Call myDoc.Send(False, Address(i))
'
'        Set myReply = Nothing
'        Set myRichText = Nothing
'        Set myItem = Nothing
'        Set myDoc = Nothing
'    Next i
'    If Err Then
'        MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                    vbExclamation, "Error Encountered"
'        Err = 0
'        GoTo GetOut
'    End If
'
'GetOut:
'    If bCitrix Then
'        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'        If Not myDom Is Nothing Then Set myDom = Nothing
'    Else
'        If Not myDB Is Nothing Then Set myDB = Nothing
'        If Not myNotes Is Nothing Then Set myNotes = Nothing
'    End If
    

End Sub

Public Function DeQuotate(txt As String) As String
    Dim i As Integer
    Dim strCheck As String
    DeQuotate = ""
    i = Len(txt)
    strCheck = Trim(txt)
    Do While i > 0
        If Mid(strCheck, i, 1) = """" Then
            strCheck = Left(strCheck, i - 1) & "''" & Mid(strCheck, i + 1)
        End If
        i = i - 1
    Loop
    DeQuotate = strCheck
End Function

Public Sub LoadRTX()
    Dim strSelect As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim i As Integer, iHgt As Integer, iC As Integer
    Dim dDate As Date
    
    flxComments.redraw = False
    flxComments.Rows = 1
    strSelect = "SELECT COMMID, ANO_COMMENT, SUPDOC_ID, NVL(COMMUSER, 'UNKNOWN') AS COMMUSER, COMMDATE, COMMSTATUS " & _
                "FROM " & ANOComment & " " & _
                "WHERE REFID = " & lRefID & " " & _
                "AND COMMSTATUS > 0 " & _
                "ORDER BY COMMDATE DESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        i = 1
        dDate = rst.Fields("COMMDATE")
        lblHeader = " LAST COMMENT BY: " & Trim(rst.Fields("COMMUSER")) & " on " & _
                    Format(dDate, "DD-MMM-YYYY") & " (" & _
                    Format(dDate, "h:mm") & " (DET))"
        Do While Not rst.EOF
            With flxComments
                .Rows = i + 1
                .TextMatrix(i, 0) = Format(rst.Fields("COMMDATE"), "DD-MMM-YYYY") & vbCr & _
                            Format(rst.Fields("COMMDATE"), "h:mm") & " (DET)"
                lblCheck(0).Caption = .TextMatrix(i, 0)
                .TextMatrix(i, 1) = Trim(rst.Fields("COMMUSER"))
                lblCheck(1).Caption = .TextMatrix(i, 1)
                .TextMatrix(i, 3) = Trim(rst.Fields("ANO_COMMENT"))
                lblCheck(2).Caption = Trim(rst.Fields("ANO_COMMENT"))
                .TextMatrix(i, 4) = rst.Fields("COMMID")
                .TextMatrix(i, 5) = rst.Fields("SUPDOC_ID")
                iHgt = 0
                For iC = 0 To 2
                    If lblCheck(iC).Height > iHgt Then iHgt = lblCheck(iC).Height
                Next iC
                .RowHeight(i) = iHgt
                If rst.Fields("SUPDOC_ID") > 0 Then
                    strSelect = "SELECT SUPDOCFORMAT FROM ANNOTATOR.GFX_SUPDOC " & _
                                "WHERE SUPDOC_ID = " & rst.Fields("SUPDOC_ID")
                    Set rstL = Conn.Execute(strSelect)
                    If Not rstL.EOF Then
                        .Col = 2: .Row = i
                        Select Case UCase(Trim(rstL.Fields("SUPDOCFORMAT")))
                            Case "DOC": Set .CellPicture = ImageList1.ListImages(1).Picture
                            Case "XLS": Set .CellPicture = ImageList1.ListImages(2).Picture
                            Case "PDF": Set .CellPicture = ImageList1.ListImages(3).Picture
                            Case "HTM": Set .CellPicture = ImageList1.ListImages(4).Picture
                            Case "HTML": Set .CellPicture = ImageList1.ListImages(5).Picture
                            Case "TXT": Set .CellPicture = ImageList1.ListImages(6).Picture
                            Case "JPG": Set .CellPicture = ImageList1.ListImages(7).Picture
                            Case "BMP": Set .CellPicture = ImageList1.ListImages(8).Picture
                        End Select
                    End If
                    rstL.Close: Set rstL = Nothing
                    If .RowHeight(i) < 480 Then .RowHeight(i) = 480
                End If
            End With
            i = i + 1
            rst.MoveNext
        Loop
    Else
        lblHeader.Caption = " NO ASSOCIATED COMMENTS"
    End If
    rst.Close
    Set rst = Nothing
    txtCommentAdd.Text = ""
    lblCommEditor.Caption = "Add Comment:"
    CommentSaved = True
    cmdSaveComment.Enabled = False
    flxComments.redraw = True
End Sub

Public Sub PopComments()

End Sub

Public Sub PopImage(tPath As String)
''    imx1.Update = False
''    imx1.FileName = tPath
''    imx1.Update = True
''    imx1.Visible = True
''    imx1.Refresh

End Sub

Private Sub vol1_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    MsgBox "Help"
End Sub

Public Sub GetGPJEmail()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    lstOptNotify(1).Clear: lstNotifyEmail(1).Clear: lstShortname(1).Clear
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
        lstOptNotify(1).AddItem Trim(rst.Fields("NAME_FIRST")) & " " & _
                    Trim(rst.Fields("NAME_LAST"))
        lstNotifyEmail(1).AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstShortname(1).AddItem LCase(Trim(rst.Fields("NAME_LOGON")))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

Public Function ConvertToHTML(pHdr As String, pflx As MSFlexGrid) As String
    Dim iRow As Integer, iCol As Integer
    
    Dim sHTML As String, strHPath As String, tFile1 As String
    Dim i As Integer
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f3O As String, f3C As String, f2O As String, f2C As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tbl2O As String, tbl3O As String, tblC As String
    Dim trO_A As String, trO_B As String, trCO As String, trC As String, trMO As String, trO As String
    Dim hr As String, br As String, sp As String
    
    Dim th_A As String, th_B As String, th_C As String
    Dim td_A As String, td_B As String, td_C As String
    Dim td_A2 As String, td_B2 As String, td_C2 As String
    Dim thcO_A As String, thcO_B As String, thcO_C As String
    Dim sTab As String, sColor As String, sClient As String
    Dim sBold As String
    
    
    strHPath = strHTMLPath & pHdr & " - " & Format(Now, "dd-mmm-yy") & ".htm"
    sColor = "#FFFFFF"
    
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""middle"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trO_A = "<TR VALIGN=""top"" BGCOLOR=": trO_B = ">": trC = "</TR>"
    hr = "<HR>": br = "<BR>": sp = "&nbsp;"
    
    
    th_A = "<TH WIDTH="""
    th_B = "><FONT SIZE=2 FACE=""Times New Roman""><I><B>"
    th_C = "</B></I></FONT></TH>"
    thcO_A = "<TH WIDTH=""": thcO_B = """ ALIGN=""right"" COLSPAN=": thcO_C = "><FONT SIZE=2 FACE=""Times New Roman""><B><I>"

    td_A = "<TD WIDTH="""
    td_B = "><FONT SIZE=2 FACE=""Tahoma"">"
    td_C = "</FONT></TD>"
    
    td_A2 = "<TD WIDTH="""
    td_B2 = "><FONT SIZE=2 FACE=""Tahoma"">"
    td_C2 = "</FONT></TD>"
    
    thcO_A = "<TH WIDTH=""": thcO_B = """ ALIGN=""center"" COLSPAN=": thcO_C = "><FONT SIZE=2 FACE=""Times New Roman""><B><I>"
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY BGCOLOR=""#FFFFFF"">": bodC = "</BODY>"
'''    f4O = "<FONT SIZE=4 FACE=""Times New Roman""><B><I>": f4C = "</I></B></FONT>"
    f3O = "<FONT SIZE=3 FACE=""Tahoma""><B>": f3C = "</B></FONT>"
    f2O = "<FONT SIZE=2 FACE=""Tahoma"">": f2C = "</FONT>"
    
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=1 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
'''    trO_A = "<TR VALIGN=""top"" BGCOLOR=": trO_B = ">": trC = "</TR>"
'''    trO = "<TR VALIGN=""top"" height=""30"">": trC = "</TR>"
    
    hr = "<HR>": br = "<BR>": sp = "&nbsp;"
    
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & pHdr & " - " & Format(Now, "dd-mmm-yyyy") & tiC & hdC & vbNewLine

    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & pHdr & br & f3C & vbNewLine
    sHTML = sHTML & f2O & "Print Date:  " & Format(Now, "dd-mmm-yyyy") & f2C & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & trO_A & sColor & trO_B & vbNewLine
    sHTML = sHTML & th_A & "20%"" ALIGN=CENTER" & th_B & pflx.TextMatrix(0, 0) & th_C & vbNewLine
    sHTML = sHTML & th_A & "20%"" ALIGN=CENTER" & th_B & pflx.TextMatrix(0, 1) & th_C & vbNewLine
    sHTML = sHTML & th_A & "60%"" ALIGN=CENTER" & th_B & pflx.TextMatrix(0, 3) & th_C & vbNewLine
    
    sHTML = sHTML & trC & vbNewLine
    sHTML = sHTML & tblC & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & tblO & vbNewLine
    
    
    sClient = ""
    For iRow = 1 To pflx.Rows - 1
        sHTML = sHTML & trO & vbNewLine
        For iCol = 0 To 3
            Select Case iCol
                Case 0
                    sHTML = sHTML & td_A & "20%"" ALIGN=CENTER" & td_B & pflx.TextMatrix(iRow, iCol) & td_C & vbNewLine
                Case 1
                    sHTML = sHTML & td_A & "20%"" ALIGN=CENTER" & td_B & pflx.TextMatrix(iRow, iCol) & td_C & vbNewLine
                Case 3
                    sHTML = sHTML & td_A & "60%"" ALIGN=LEFT" & td_B & pflx.TextMatrix(iRow, iCol) & td_C & vbNewLine
            End Select
'''            MsgBox "Row " & iRow & " : Col " & iCol & vbNewLine & vbNewLine & flx1.TextMatrix(iRow, iCol)
        Next iCol
        sHTML = sHTML & trC & vbNewLine
    Next iRow
    
    
    sHTML = sHTML & tblC & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = Replace(strHPath, ": ", " -")
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    ConvertToHTML = tFile1
    
''    frmLog.flxLog.Rows = 2
    
'''    web1.Navigate tFile1
'''    web1.Visible = True
End Function

