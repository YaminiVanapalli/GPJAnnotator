VERSION 5.00
Begin VB.Form frmKeywordEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyword Editor..."
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKeywordEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   3360
      TabIndex        =   12
      Text            =   "test"
      Top             =   2100
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clear All Associations"
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2475
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save Associations"
      Height          =   435
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2835
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   300
      ScaleHeight     =   495
      ScaleWidth      =   3945
      TabIndex        =   8
      Top             =   120
      Width           =   3945
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Associate Keywords, either type in New Keyword, or select Existing Keywords from list below:"
         Height          =   465
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   3885
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraKeyword 
      Height          =   4275
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   5415
      Begin VB.CommandButton cmdApply 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2400
         Picture         =   "frmKeywordEdit.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to Remove Selections"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdApply 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2400
         Picture         =   "frmKeywordEdit.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Click to Add Selections"
         Top             =   1980
         Width           =   615
      End
      Begin VB.ListBox lstKeyApply 
         Height          =   3375
         ItemData        =   "frmKeywordEdit.frx":114E
         Left            =   3180
         List            =   "frmKeywordEdit.frx":1150
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2115
      End
      Begin VB.ListBox lstKeyAvail 
         Height          =   2985
         ItemData        =   "frmKeywordEdit.frx":1152
         Left            =   120
         List            =   "frmKeywordEdit.frx":1154
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1110
         Width           =   2115
      End
      Begin VB.TextBox txtKeyword 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label lblHdr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Associated Keywords:"
         Height          =   195
         Index           =   1
         Left            =   3180
         TabIndex        =   11
         Top             =   480
         Width           =   1590
      End
      Begin VB.Label lblHdr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Keywords:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuKeywordNameEdit 
         Caption         =   "Edit Keyword..."
      End
      Begin VB.Menu mnuKeywordDelete 
         Caption         =   "Delete Keyword"
      End
   End
End
Attribute VB_Name = "frmKeywordEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bNeedToSave As Boolean
Dim iEditRow As Integer, iEditList As Integer ''0=lstGfxAvail;1=lstGfxApply''
Dim lGID As Long
Dim tFrom As String

Public Property Get PassFrom() As String
    PassFrom = tFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    tFrom = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = lGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    lGID = vNewValue
End Property




Private Sub cmdApply_Click(Index As Integer)
    Dim i As Integer
    Dim bFound As Boolean
    
    bNeedToSave = True
    bFound = False
    Select Case Index
        Case 0
            For i = lstKeyAvail.ListCount - 1 To 0 Step -1
                If lstKeyAvail.Selected(i) = True Then
                    lstKeyApply.AddItem lstKeyAvail.List(i)
                    lstKeyApply.ItemData(lstKeyApply.NewIndex) = lstKeyAvail.ItemData(i)
                    bFound = True
                    lstKeyAvail.RemoveItem (i)
                End If
            Next i
            If Not bFound Then
                If Trim(txtKeyword.Text) <> "" Then
                    lstKeyApply.AddItem txtKeyword.Text
                    txtKeyword.Text = ""
                    txtKeyword.SetFocus
                Else
                    MsgBox "No Keyword entered", vbExclamation, "Sorry..."
                    Exit Sub
                End If
            End If
                    
        Case 1
            For i = lstKeyApply.ListCount - 1 To 0 Step -1
                If lstKeyApply.Selected(i) = True Then
                    lstKeyAvail.AddItem lstKeyApply.List(i)
                    lstKeyAvail.ItemData(lstKeyAvail.NewIndex) = lstKeyApply.ItemData(i)
                    lstKeyApply.RemoveItem (i)
                End If
            Next i
    End Select
End Sub

Private Sub cmdCancel_Click()
    Dim strDelete As String
    Dim i As Integer
    
    ''DELETE EXISTING ASSOCIATIONS''
    strDelete = "DELETE FROM ANNOTATOR.GFX_METADATA_R " & _
                "WHERE GID = " & lGID
    Conn.Execute (strDelete)
    
    ''CLEAR LIST''
    For i = lstKeyApply.ListCount - 1 To 0 Step -1
        lstKeyAvail.AddItem lstKeyApply.List(i)
        lstKeyApply.RemoveItem (i)
    Next i
    
    bNeedToSave = False
End Sub

Private Sub cmdOK_Click()
    Dim bSuccessful As Boolean
    
    bSuccessful = SaveKeywords(lGID)
    bNeedToSave = False
    If bSuccessful Then Unload Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    bNeedToSave = False
    Call GetAllKeywords
    Call GetGraphicKeywords(lGID)
    Call CompareLists
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Resp As VbMsgBoxResult
    Dim bSuccessful As Boolean
    
    If bNeedToSave Then
        Resp = MsgBox("You have changed the Keyword Associations.  " & _
                    "Do you want to Save the changes?", _
                    vbExclamation + vbYesNo, "Save before Closing?")
        If Resp = vbYes Then
            bSuccessful = SaveKeywords(lGID)
            If Not bSuccessful Then
                MsgBox "The Keyword Interface is closing without a successful Save", _
                            vbExclamation, "Closing..."
            End If
        End If
    End If
        
End Sub


Private Sub lstKeyApply_GotFocus()
    cmdApply(1).Default = True
End Sub

Private Sub lstKeyApply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        iEditList = 1
        iEditRow = Int(Y / 197)
        txtEdit.Text = lstKeyApply.List(iEditRow)
        txtEdit.Height = 225
        txtEdit.Top = fraKeyword.Top + lstKeyApply.Top + _
                    (iEditRow * 197) - ((txtEdit.Height - 197) / 2)
        txtEdit.Left = fraKeyword.Left + lstKeyApply.Left
        txtEdit.Tag = CStr(lstKeyApply.ItemData(iEditRow))
        PopupMenu mnuEdit
    End If
End Sub

Private Sub lstKeyAvail_GotFocus()
    cmdApply(0).Default = True
End Sub

Private Sub lstKeyAvail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        iEditList = 0
        fraKeyword.Enabled = False
        iEditRow = Int(Y / 197)
        txtEdit.Text = lstKeyAvail.List(iEditRow)
        txtEdit.Height = 225
        txtEdit.Top = fraKeyword.Top + lstKeyAvail.Top + _
                    (iEditRow * 197) - ((txtEdit.Height - 197) / 2)
        txtEdit.Left = fraKeyword.Left + lstKeyAvail.Left
        txtEdit.Tag = CStr(lstKeyAvail.ItemData(iEditRow))
        PopupMenu mnuEdit
    End If
End Sub

'''Private Sub lstKeyAvail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    PopupMenu mnuEdit
'''End Sub

Private Sub mnuKeywordNameEdit_Click()
    txtEdit.Visible = True
    fraKeyword.Enabled = False
End Sub

Private Sub txtEdit_GotFocus()
    cmdApply(0).Default = False
    cmdApply(1).Default = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    ElseIf KeyAscii = 13 Then
'''        MsgBox "Edit Time"
        Call EditKeyword(CLng(txtEdit.Tag), Trim(txtEdit.Text), iEditRow, iEditList)
        txtEdit.Visible = False
        fraKeyword.Enabled = True
    End If
End Sub

Private Sub txtKeyword_Change()
    Dim i As Integer, iCnt As Integer, iTop As Integer
    
    iCnt = 0: iTop = 0
    lstKeyAvail.Visible = False
    For i = lstKeyAvail.ListCount - 1 To 0 Step -1
        If Left(lstKeyAvail.List(i), Len(txtKeyword.Text)) = txtKeyword.Text Then
            lstKeyAvail.Selected(i) = True
            iTop = i
            iCnt = iCnt + 1
        Else
            lstKeyAvail.Selected(i) = False
        End If
        lstKeyAvail.Selected(i) = Left(lstKeyAvail.List(i), Len(txtKeyword.Text)) = txtKeyword.Text
    Next i
    If iTop - 2 >= 0 Then iTop = iTop - 2
    lstKeyAvail.TopIndex = iTop
'    lblCnt = "Selected Files:  " & CStr(iCnt)
    lstKeyAvail.Visible = True
End Sub

Private Sub txtKeyword_GotFocus()
    cmdApply(0).Default = True
End Sub

Private Sub txtKeyword_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Public Sub GetAllKeywords()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    Select Case tFrom
        Case "GH"
'''            strSelect = "SELECT GFX_METADATA_ID AS KEYID, " & _
'''                        "UPPER(VALUECHAR) AS KEYWORD " & _
'''                        "FROM GFX_METADATA " & _
'''                        "WHERE TYPE_CD = 305 " & _
'''                        "ORDER BY KEYWORD"
            strSelect = "SELECT DISTINCT " & _
                        "GMX.GFX_METADATA_ID AS KEYID, " & _
                        "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                        "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GMX.TYPE_CD = 305 " & _
                        "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                        "AND GMR.GID = GM.GID " & _
                        "AND GM.AN8_CUNO <> 40579 " & _
                        "ORDER BY KEYWORD"
        Case "DIL"
            strSelect = "SELECT DISTINCT " & _
                        "GMX.GFX_METADATA_ID AS KEYID, " & _
                        "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                        "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GMX.TYPE_CD = 305 " & _
                        "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                        "AND GMR.GID = GM.GID " & _
                        "AND GM.AN8_CUNO = 27129 " & _
                        "ORDER BY KEYWORD"
    End Select
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstKeyAvail.AddItem Trim(rst.Fields("KEYWORD"))
        lstKeyAvail.ItemData(lstKeyAvail.NewIndex) = rst.Fields("KEYID")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

Public Sub GetGraphicKeywords(tGID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT GX.GFX_METADATA_ID AS KEYID, " & _
                "UPPER(GX.VALUECHAR) AS KEYWORD " & _
                "FROM ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_METADATA GX " & _
                "Where GMR.GID = " & tGID & " " & _
                "AND GMR.GFX_METADATA_ID = GX.GFX_METADATA_ID " & _
                "ORDER BY KEYWORD"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstKeyApply.AddItem Trim(rst.Fields("KEYWORD"))
        lstKeyApply.ItemData(lstKeyApply.NewIndex) = rst.Fields("KEYID")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

Public Sub CompareLists()
    Dim i As Integer, iA As Integer
    Dim lCheck As Long
    
    For i = 0 To lstKeyApply.ListCount - 1
        lCheck = lstKeyApply.ItemData(i)
        For iA = lstKeyAvail.ListCount - 1 To 0 Step -1
            If lstKeyAvail.ItemData(iA) = lCheck Then
                lstKeyAvail.RemoveItem (iA)
            End If
        Next iA
    Next i
        
End Sub

Public Function GetKeywordID(sKeyword As String) As Long
    Dim strSelect As String, strInsert As String
    Dim rstL As ADODB.Recordset
    Dim lNewID As Long
    
    Set rstL = Conn.Execute("SELECT " & GFXSeq & ".NEXTVAL FROM DUAL")
    lNewID = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing
    
    strInsert = "INSERT INTO ANNOTATOR.GFX_METADATA " & _
                "(GFX_METADATA_ID, TYPE_CD, TYPEDESC, VALUECHAR, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                "VALUES " & _
                "(" & lNewID & ", 305, 'GFX_KEYWORD', " & _
                "'" & DeGlitch(Trim(UCase(sKeyword))) & "', '" & Left(DeGlitch(LogName), 24) & "', " & _
                "SYSDATE, '" & Left(DeGlitch(LogName), 24) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
    
    GetKeywordID = lNewID
    
End Function


Public Function SaveKeywords(lGID As Long) As Boolean
    Dim strDelete As String, strInsert As String
    Dim i As Integer, i1 As Integer
    
    On Error GoTo ErrorTrap
    Conn.BeginTrans
    ''FIRST, DELETE EXISTING ASSOCIATIONS''
    strDelete = "DELETE FROM ANNOTATOR.GFX_METADATA_R " & _
                "WHERE GID = " & lGID
    Conn.Execute (strDelete)
    
    ''NEXT, CHECK FOR NEW KEYWORDS''
    For i = 0 To lstKeyApply.ListCount - 1
        If lstKeyApply.ItemData(i) = 0 Then
            lstKeyApply.ItemData(i) = GetKeywordID(lstKeyApply.List(i))
        End If
    Next i
    
    ''NOW, SAVE NEW LIST''
    If lstKeyApply.ListCount > 0 Then
        For i1 = 0 To lstKeyApply.ListCount - 1
            strInsert = "INSERT INTO ANNOTATOR.GFX_METADATA_R " & _
                        "(GID, GFX_METADATA_ID, ADDUSER, ADDDTTM) " & _
                        "VALUES " & _
                        "(" & lGID & ", " & lstKeyApply.ItemData(i1) & ", " & _
                        "'" & Left(DeGlitch(LogName), 24) & "', SYSDATE)"
            Conn.Execute (strInsert)
        Next i1
    End If
    
    Conn.CommitTrans
    SaveKeywords = True
Exit Function

ErrorTrap:
    Conn.RollbackTrans
    MsgBox "Error:  " & Err.Description, vbExclamation, "Unable to Associate Keywords..."
    SaveKeywords = False
End Function

Public Sub EditKeyword(lXID As Long, sKeyword As String, iRow As Integer, iList As Integer)
    Dim strUpdate As String
    
    On Error Resume Next
    strUpdate = "UPDATE ANNOTATOR.GFX_METADATA " & _
                "SET VALUECHAR = '" & DeGlitch(Trim(sKeyword)) & "', " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE GFX_METADATA_ID = " & lXID
    Conn.Execute (strUpdate)
    
    If Err Then
        MsgBox "Error:  " & Err.Description, vbCritical, "Error Encountered..."
    Else
        Select Case iList
            Case 0: lstKeyAvail.List(iRow) = sKeyword
            Case 1: lstKeyApply.List(iRow) = sKeyword
        End Select
    End If

End Sub
