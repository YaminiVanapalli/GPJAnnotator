VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for Folder"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1515
   End
   Begin VB.ListBox lstPlans 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   6075
   End
   Begin VB.ListBox lstPaths 
      Height          =   1035
      Left            =   2940
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create New Folder..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   3435
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1035
   End
   Begin VB.DirListBox Dir1 
      Height          =   5265
      Left            =   6840
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   5355
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":045E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":08B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":0D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1154
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":229C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":26EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":2C88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   4995
      Left            =   240
      TabIndex        =   1
      Top             =   540
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   8811
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   8520
      Width           =   6135
   End
   Begin VB.Label lblPath 
      Alignment       =   1  'Right Justify
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
      Left            =   6255
      TabIndex        =   11
      Top             =   5640
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select file(s) to download:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Folder to download Graphic into:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNode As String
Dim tSHYR As Integer
Dim tBCC As String
Dim tSHCD As Long, tDWGID As Long, pFCCD As Long, pGID As Long, pEID As Long
Dim tFBCN As String
Dim tSHNM As String
Dim tFileType As String
Public pFrom As String, pFacil As String

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

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property

Public Property Get PassFCCD() As Long
    PassFCCD = pFCCD
End Property
Public Property Let PassFCCD(ByVal vNewValue As Long)
    pFCCD = vNewValue
End Property

Public Property Get PassFacil() As String
    PassFacil = pFacil
End Property
Public Property Let PassFacil(ByVal vNewValue As String)
    pFacil = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property

Public Property Get PassEID() As Long
    PassEID = pEID
End Property
Public Property Let PassEID(ByVal vNewValue As Long)
    pEID = vNewValue
End Property





Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    Dim sNewPath As String, sNewFolder As String, sNewTry As String
    Dim nodX As Node
    Dim i As Integer
    Dim sChk As String
    
    If sNode <> "" And UCase(Left(sNode, 1)) = "F" Then tvw.Nodes(sNode).Image = 10
    
    sNewFolder = "New Folder"
    sNewPath = Dir1.Path & "\" & sNewFolder
    
    
    ''LOOK FOR EXISTING NEW FOLDER''
    If Dir(sNewPath, vbDirectory) <> "" Then
        i = 0
        Do
            i = i + 1
            sNewTry = sNewPath & " (" & i & ")"
            sNewFolder = "New Folder" & " (" & i & ")"
        Loop Until Dir(sNewTry, vbDirectory) = ""
        sNewPath = sNewTry
    End If
    
    On Error Resume Next
    MkDir sNewPath
    
    If Err = 0 Then
        Set nodX = tvw.Nodes.Add(sNode, tvwChild, "f" & sNewPath, sNewFolder, 11)
        sNode = nodX.Key
        nodX.Selected = True
        tvw.StartLabelEdit
    Else
        sNode = ""
        MsgBox "Unable to create a folder at the specified location " & _
                    "[" & sNewPath & "]" & _
                    vbNewLine & vbNewLine & _
                    "Error: " & Err.Description, vbCritical, "Sorry..."
    End If
    
End Sub

Private Sub cmdOK_Click()
'''    MsgBox Mid(sNode, 2)
    If sNode = "" Then
        cmdOK.Enabled = False
        MsgBox "No folder location has been selected", vbExclamation, "Sorry..."
        Exit Sub
    End If
    
    Select Case UCase(pFrom)
        Case "FRMDIL"
            frmDIL.PassDLPath = Mid(sNode, 2)
        Case "FRMSEARCHRESULTS"
            frmSearchResults.PassDLPath = Mid(sNode, 2)
        Case "FRMDOWNLOADCART"
            frmDownloadCart.PassDLPath = Mid(sNode, 2)
        Case "FRMGRAPHICS"
            frmGraphics.PassDLPath = Mid(sNode, 2)
        Case "FRMANNOTATOR", "FRMSHOW", "FRMFACIL-PDF", "FRMFACIL-DWF", "FRMOSP", "FRMPHOTO", "FRMCONST", "FRMPHOTO-ELEM"
            Call DoDownload(Mid(sNode, 2))
        Case "FRMHTMLVIEWER"
            frmHTMLViewer.pDownloadPath = Mid(sNode, 2)
        
            
    End Select
    Unload Me
End Sub


Private Sub Dir1_Change()
    Dim i As Integer
    Dim nodX As Node
    Dim sFolder As String
    
    Debug.Print Dir1.Path
    
    For i = 0 To Dir1.ListCount - 1
        Select Case Len(Dir1.Path)
            Case 3: sFolder = Mid(Dir1.List(i), Len(Dir1.Path) + 1)
            Case Else: sFolder = Mid(Dir1.List(i), Len(Dir1.Path) + 2)
        End Select
        Set nodX = tvw.Nodes.Add(sNode, tvwChild, "f" & Dir1.List(i), sFolder, 10)
'        MsgBox Dir1.List(i)
    Next i
    tvw.Nodes(sNode).Expanded = True
End Sub

Private Sub Dir1_Click()
    Debug.Print Dir1.List(Dir1.ListIndex)
End Sub

'Private Sub Drive1_Change()
'    Debug.Print Drive1.Drive
'    Dir1.Path = Drive1.Drive
'    Debug.Print Dir1.Path
'End Sub

Private Sub Form_Load()
    Dim fs, d, s
    Dim i As Integer, iIcon As Integer
    Dim nodX As Node
    Dim sDrive As String
    
    lblPath.Caption = ""
    sNode = ""
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    tvw.ImageList = ImageList1
    
    For i = 0 To Drive1.ListCount - 1
'''        If InStr(1, UCase(Drive1.List(i)), "DETCTX") = 0 _
'''                    And UCase(Left(Drive1.List(i), 1)) <> "U" _
'''                    And UCase(Left(Drive1.List(i), 1)) <> "V" Then
        If (iAppConn = 1 And InStr(1, UCase(Drive1.List(i)), "CTX") = 0 _
                    And UCase(Left(Drive1.List(i), 1)) <> "C") _
                    Or _
                    (iAppConn = 2 And UCase(Left(Drive1.List(i), 1)) <> "C") _
                    Or iAppConn = 0 Then
            sDrive = Drive1.List(i)
            Set d = fs.GetDrive(fs.GetDriveName(sDrive))
            iIcon = d.DriveType
            If iIcon = 4 Or Not d.IsReady Then
'''                MsgBox "Skipping " & sDrive & vbNewLine & _
'''                            "Drive Name = " & sDrive & vbNewLine & _
'''                            "Drive Type = " & iIcon & vbNewLine & _
'''                            "IsReady Status = " & d.IsReady, vbInformation, "fyi..."
                GoTo SkipDrive
            End If
            If InStr(1, sDrive, "\\Client\") Then
                If InStr(1, sDrive, "A$") Then
                    sDrive = Left(sDrive, 1) & ": LOCAL CLIENT (A:)"
                ElseIf InStr(1, sDrive, "B$") Then
                    sDrive = Left(sDrive, 1) & ": LOCAL CLIENT (B:)"
                ElseIf InStr(1, sDrive, "C$") Then
                    sDrive = Left(sDrive, 1) & ": LOCAL CLIENT (C:)"
                ElseIf InStr(1, sDrive, "D$") Then
                    sDrive = Left(sDrive, 1) & ": LOCAL CLIENT (D:)"
                ElseIf InStr(1, sDrive, "E$") Then
                    sDrive = Left(sDrive, 1) & ": LOCAL CLIENT (E:)"
                End If
            End If
            
            Set nodX = tvw.Nodes.Add(, , "d" & Left(Drive1.List(i), 1) & ":\", sDrive, iIcon)
SkipDrive:
        End If
    Next i
    Set d = Nothing
    Set fs = Nothing
    
    Select Case UCase(pFrom)
        Case "FRMANNOTATOR"
            lblCaption.Caption = tFBCN & "  -  " & tSHYR & "  " & tSHNM
            Call GetDrawings(frmBrowse, tFileType, tDWGID, tSHYR, tSHCD)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
            
        Case "FRMSHOW"
            lblCaption.Caption = tSHYR & "  " & tSHNM
            Call GetShowPlans(frmBrowse, tFileType, tSHYR, tSHCD)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
            
        Case "FRMFACIL-PDF"
            lblCaption.Caption = pFacil
            Call GetFCCDpdfs(pFCCD)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
        
        Case "FRMFACIL-DWF"
            lblCaption.Caption = pFacil
            Call GetFCCDdwfs(pFCCD)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
            
        Case "FRMOSP", "FRMPHOTO"
            lblCaption.Caption = pFacil
            Call GetFacilPhotos(pFCCD, pGID)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
            
        Case "FRMPHOTO-ELEM"
            
            
        Case "FRMCONST"
            lblCaption.Caption = tFBCN
            Call GetElementPDFs(tDWGID)
            lstPlans.Visible = True
            lblFiles.Visible = True
            cmdOK.Enabled = False
            
        Case Else
            lstPlans.Visible = False
            lblFiles.Visible = False
            cmdOK.Enabled = True
    End Select
    
End Sub

'Private Sub lstPlans_Click()
'    Dim i As Integer
'End Sub

Private Sub lstPlans_ItemCheck(Item As Integer)
    cmdOK.Enabled = CheckForChecked
End Sub

Private Sub mnuRename_Click()
    tvw.SelectedItem.Selected = True
    tvw.StartLabelEdit
End Sub

Private Sub tvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim sNewKey As String, sOldKey As String
    
    On Error Resume Next
    sOldKey = tvw.SelectedItem.Key
    sNewKey = tvw.SelectedItem.Parent.Key & "\" & NewString
    If UCase(Left(sNewKey, 1)) = "D" Then sNewKey = "f" & Mid(sNewKey, 2)
    tvw.SelectedItem.Key = sNewKey
    If Err Then
        MsgBox "Cannot rename folder as entered.  Folder already exists.", vbCritical, "Sorry..."
        Cancel = 1
        Exit Sub
    End If
    sNode = sNewKey
    Name Mid(sOldKey, 2) As Mid(sNewKey, 2)
    If Err Then
        MsgBox "Cannot rename folder as entered", vbCritical, "Sorry..."
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub tvw_DblClick()
    If UCase(Left(tvw.SelectedItem.Key, 1)) = "D" Then
        MsgBox "Unable to edit the name of a drive", vbExclamation, "Sorry..."
        Exit Sub
    End If
    tvw.SelectedItem.Selected = True
    tvw.StartLabelEdit
End Sub

'''Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'''    If Button = vbRightButton Then
'''        If UCase(Left(tvw.SelectedItem.key, 1)) = "D" Then
'''            MsgBox "Unable to edit the name of a drive", vbExclamation, "Sorry..."
'''            Exit Sub
'''        End If
'''        tvw.SelectedItem.Selected = True
'''        tvw.StartLabelEdit
'''    End If
'''End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sFolder As String
    
    Me.MousePointer = 11
    
    If sNode <> "" And UCase(Left(sNode, 1)) = "F" Then tvw.Nodes(sNode).Image = 10
    
    Debug.Print Node.Key
    sNode = Node.Key
    If sNode <> "" And UCase(Left(Node.Key, 1)) = "F" Then Node.Image = 11
    
    If Node.Children = 0 Then
        Select Case UCase(Left(Node.Key, 1))
            Case "D": Dir1.Path = Mid(Node.Key, 2) ''& ":\"
            Case "F": Dir1.Path = Mid(Node.Key, 2)
        End Select
    End If
    
    cmdCreate.Enabled = True
    Select Case UCase(pFrom)
        Case "FRMANNOTATOR", "FRMSHOW"
            cmdOK.Enabled = CheckForChecked
        Case Else
            cmdOK.Enabled = True
    End Select
    
    Me.MousePointer = 0
End Sub

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

Public Sub DoDownload(sFolder As String)
    Dim i As Integer
    Dim sPath As String
    
    Screen.MousePointer = 11
    
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) = True Then
            sPath = sFolder & "\" & Legalize(lblCaption) & " [" & Legalize(lstPlans.List(i)) & "]." & tFileType
            FileCopy lstPaths.List(i), sPath
            lstPlans.Selected(i) = False
        End If
    Next i
    
    Screen.MousePointer = 0
    MsgBox "File(s) Copied", vbInformation, "File Download Successful..."
End Sub

Public Function CheckForChecked() As Boolean
    Dim i As Integer
    Dim bFound As Boolean
    bFound = False
    For i = 0 To lstPlans.ListCount - 1
        If lstPlans.Selected(i) Then
            bFound = True
            Exit For
        End If
    Next i
    If bFound Then bFound = CBool(sNode <> "")
    CheckForChecked = bFound
End Function

Public Sub GetFCCDPhotos(tFCCD As Long)
    Dim strSelect As String, sThumbPath As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iDel As Integer
    
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
        cmdOK.Enabled = False
        Screen.MousePointer = 0
        Exit Sub
    Else
'        sHDR = "Facility:  " & Trim(rst.Fields("FACIL"))
        lstPlans.Tag = Trim(rst.Fields("FACIL"))
    End If
    Do While Not rst.EOF
        lstPlans.AddItem Trim(rst.Fields("GDESC"))
        lstPlans.ItemData(lstPlans.NewIndex) = rst.Fields("GID")
        lstPaths.AddItem Trim(rst.Fields("GPATH"))
        lstPaths.ItemData(lstPaths.NewIndex) = rst.Fields("GID")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Public Sub GetFCCDpdfs(tFCCD As Long)
    Dim sCopyPath As String, strSelect As String, sChk As String
    Dim rst As ADODB.Recordset
    
    sCopyPath = "\\DETMSFS01\GPJAnnotator\Floorplans\"
    strSelect = "SELECT DF.DWFID, DF.DWFDESC, DF.DWFPATH, UPPER(DF.DWFDESC) AS UDESC, DF.DWFSTATUS " & _
                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                "Where DM.AN8_CUNO = " & tFCCD & " " & _
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
    
End Sub

Public Sub GetFCCDdwfs(tFCCD As Long)
    Dim strSelect As String, sChk As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT DF.DWFID, DF.DWFDESC, DF.DWFPATH, UPPER(DF.DWFDESC) AS UDESC, DF.DWFSTATUS " & _
                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                "Where DM.AN8_CUNO = " & tFCCD & " " & _
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
    
End Sub


Public Sub GetFacilPhotos(tFCCD As Long, tGID As Long)
    Dim strSelect As String, sThumbPath As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iDel As Integer
    
    i = -1
    sThumbPath = "\\DETMSFS01\GPJAnnotator\Floorplans\FacilPho\Thumbs\"
    
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
        lblCaption.Caption = Trim(rst.Fields("FACIL"))
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

Public Sub GetElementPDFs(pDWGID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sPDF As String, sDesc As String
    
    lstPlans.Clear
    lstPaths.Clear
    strSelect = "SELECT NVL(EE.PRGID, 0) AS PRGID, " & _
                "DM.DWGNUM, DS.SHTDESC, DS.SHTSEQ, DF.DWFPATH " & _
                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF, ANNOTATOR.ENG_ELEMENT EE " & _
                "Where DM.DWGID = " & pDWGID & " " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DF.DWGID " & _
                "AND DS.SHTID = DF.SHTID " & _
                "AND DM.DWGID = EE.DWGID (+) " & _
                "ORDER BY DS.SHTSEQ"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        sPDF = Left(Trim(rst.Fields("DWFPATH")), Len(Trim(rst.Fields("DWFPATH"))) - 3) & "pdf"
        If Dir(sPDF, vbNormal) <> "" Then
            If rst.Fields("PRGID") <> 0 Then
                sDesc = rst.Fields("PRGID") & "-" & Right("00" & rst.Fields("DWGNUM"), 2) & _
                            Trim(rst.Fields("SHTSEQ")) & " -- " & Trim(rst.Fields("SHTDESC"))
                lstPlans.AddItem sDesc
                lstPaths.AddItem sPDF
            End If
        End If
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub
