VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{23319180-2253-11D7-BD2E-08004608C318}#3.0#0"; "XpdfViewerCtrl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSearchResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Results..."
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSearch 
      Height          =   8835
      Left            =   120
      ScaleHeight     =   8775
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdDownloadAdd 
         Caption         =   "Add to Download Cart"
         Height          =   375
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download Result..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.ListBox lstPaths 
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox lstFiles 
         Height          =   255
         Left            =   3780
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export Result to Approval List"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   2595
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":0A24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":0B7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":0CD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":0E32
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":13CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":1526
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchResults.frx":1AC0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Open in Full Viewer..."
         Enabled         =   0   'False
         Height          =   555
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5340
         Width           =   1275
      End
      Begin VB.CommandButton cmdRevise 
         Caption         =   "Revise Search..."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin MSComctlLib.TreeView tvw1 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   7646
         _Version        =   393217
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.PictureBox picPDF 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   5835
         TabIndex        =   4
         Top             =   5280
         Visible         =   0   'False
         Width           =   5835
         Begin XpdfViewerCtl.XpdfViewer xpdf1 
            Height          =   3375
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   5835
            showScrollbars  =   -1  'True
            showBorder      =   -1  'True
            showPasswordDialog=   -1  'True
         End
      End
      Begin IMAGXPR5LibCtl.ImagXpress imx1 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   5953
         ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
         ErrCode         =   1235318547
         ErrInfo         =   -886549320
         Persistence     =   -1  'True
         _cx             =   132055552
         _cy             =   1
         FileName        =   ""
         MousePointer    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         AutoSize        =   4
         BorderType      =   0
         ShowDithered    =   1
         ScrollBars      =   0
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         Multitask       =   0   'False
         CancelMode      =   0
         CancelLoad      =   0   'False
         CancelRemove    =   0   'False
         Palette         =   0
         ShowHourglass   =   0   'False
         LZWPassword     =   ""
         PlaceHolder     =   ""
         PFileName       =   ""
         PICPassword     =   ""
         PrinterBanding  =   0   'False
         UndoEnabled     =   0   'False
         Update          =   -1  'True
         CropX           =   0
         CropY           =   0
         SaveGIFType     =   0
         SaveTIFCompression=   0
         SavePNGInterlaced=   0   'False
         SaveGIFInterlaced=   0   'False
         SaveGIFTransparent=   0   'False
         SaveJPGProgressive=   0   'False
         SaveJPGGrayscale=   0   'False
         SaveGIFTColor   =   0
         TwainProductName=   ""
         TwainProductFamily=   ""
         TwainManufacturer=   ""
         TwainVersionInfo=   ""
         Notify          =   0   'False
         NotifyDelay     =   0
         SavePBMType     =   0
         SavePGMType     =   0
         SavePPMType     =   0
         PageNbr         =   0
         ProgressEnabled =   0   'False
         ManagePalette   =   -1  'True
         PictureEnabled  =   -1  'True
         SaveJPGLumFactor=   25
         SaveJPGChromFactor=   35
         DisplayMode     =   0
         DrawStyle       =   1
         DrawWidth       =   1
         DrawFillColor   =   0
         DrawFillStyle   =   1
         DrawMode        =   13
         PICThumbnail    =   0
         PICCropEnabled  =   0   'False
         PICCropX        =   0
         PICCropY        =   0
         PICCropWidth    =   1
         PICCropHeight   =   1
         Antialias       =   0
         SaveJPGSubSampling=   2
         OLEDropMode     =   0
         CompressInMemory=   0
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5880
         TabIndex        =   6
         Top             =   4980
         Width           =   45
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove the selection"
      End
   End
End
Attribute VB_Name = "frmSearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pTBL As String, pSQL As String, pFrom As String
Public pDownloadPath As String


Public Property Get PassDLPath() As String
    PassDLPath = pDownloadPath
End Property
Public Property Let PassDLPath(ByVal vNewValue As String)
    pDownloadPath = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property

Public Property Get PassTBL() As String
    PassTBL = pTBL
End Property
Public Property Let PassTBL(ByVal vNewValue As String)
    pTBL = vNewValue
End Property

Public Property Get PassSQL() As String
    PassSQL = pSQL
End Property
Public Property Let PassSQL(ByVal vNewValue As String)
    pSQL = vNewValue
End Property



Private Sub cmdDownload_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, sChk As String, sPath As String
    Dim rst As ADODB.Recordset
    Dim strGID As String
    Dim i As Integer
    
    strGID = ""
    For i = 1 To tvw1.Nodes.Count
        If UCase(Left(tvw1.Nodes(i).Key, 1)) = "G" Then
            If strGID = "" Then
                strGID = Mid(tvw1.Nodes(i).Key, 2)
            Else
                strGID = strGID & ", " & Mid(tvw1.Nodes(i).Key, 2)
            End If
        End If
    Next i
    
    
pDownloadPath = ""
    frmBrowse.PassFrom = Me.Name
    frmBrowse.Show 1, Me
    
'''    If shlShell Is Nothing Then
'''        Set shlShell = New Shell32.Shell
'''    End If
'''
'''    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
'''                "Select Folder to download Graphic into:", _
'''                BIF_RETURNONLYFSDIRS)
'''
'''    If shlFolder Is Nothing Then

    If pDownloadPath = "" Then
        Exit Sub
    Else
        Screen.MousePointer = 11
        On Error GoTo BadFile
        sFolder = pDownloadPath '' shlFolder.Items.Item.Path
        
        If UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        Err = 0
        On Error GoTo ErrorTrap
        i = 0
        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                    "FROM " & GFXMas & " " & _
                    "WHERE GID IN (" & strGID & ")"
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            sPath = sFolder & "\" & Trim(rst.Fields("GDESC")) & _
                        "." & Trim(rst.Fields("GFORMAT"))
            FileCopy Trim(rst.Fields("GPATH")), sPath
            
            i = i + 1
            rst.MoveNext
        Loop
        rst.Close: Set rst = Nothing
        
                
        Screen.MousePointer = 0
        Select Case i
            Case 0
                MsgBox "No files were copied", vbExclamation, "No Files..."
            Case 1
                MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
            Case Else
                MsgBox "(" & i & ") Files Copied to " & sFolder, vbInformation, "File Download Successful..."
        End Select
    End If
    
Exit Sub
ErrorTrap:
    Screen.MousePointer = 0
    rst.Close: Set rst = Nothing
    MsgBox "Error:  " & Err.Description, vbExclamation, "File Not Copied..."
Exit Sub
BadFile:
    Screen.MousePointer = 0
    MsgBox "The location you chose is not valid.  If you chose " & _
                "'DESKTOP', be aware the desktop in this interface is the " & _
                "Citrix Servers Desktop and you do not have rights to " & _
                "download files there." & vbNewLine & vbNewLine & _
                "Please, choose another folder.", vbExclamation, "Invalid Location..."
    Err.Clear
End Sub

Private Sub cmdDownloadAdd_Click()
    Call AddToCart
End Sub

Private Sub cmdExport_Click()
    Dim i As Integer
    Dim sList As String, sStat As String
    Dim bMulti As Boolean
    
    For i = 1 To tvw1.Nodes.Count
        If UCase(Left(tvw1.Nodes(i).Key, 1)) = "G" Then
            If sList = "" Then
                sList = Mid(tvw1.Nodes(i).Key, 2)
            Else
                sList = sList & ", " & Mid(tvw1.Nodes(i).Key, 2)
            End If
        End If
    Next i
    Debug.Print sList
    frmGraphics.optApproverView(2).Value = True
    frmGraphics.sSearchList = sList
    frmGraphics.txtNoShows.Text = "Search Result for the Client above"
    frmGraphics.txtNoShows.Visible = True
    frmGraphics.cboCUNO(4).Enabled = False
    Call frmGraphics.GetApprovalGraphics(frmGraphics.cboCUNO(4).ItemData(frmGraphics.cboCUNO(4).ListIndex), _
                frmGraphics.sOrder, 0, 0, 0)
    
    frmGraphics.picOuter(4).Visible = True
    
    sStat = "": bMulti = False
    If frmSearch.chkFilter(2).Value = 1 Then
        If frmSearch.chkGSTATUS(10).Value = 1 Then sStat = "I"
        If frmSearch.chkGSTATUS(20).Value = 1 Then
            If sStat = "" Then
                sStat = "C"
            Else
                bMulti = True
                GoTo FoundMulti
            End If
        End If
        If frmSearch.chkGSTATUS(27).Value = 1 Then
            If sStat = "" Then
                sStat = "R"
            Else
                bMulti = True
                GoTo FoundMulti
            End If
        End If
FoundMulti:
    Else
        bMulti = True
    End If
    
'''    For i = 0 To frmGraphics.flxApprove.Rows - 2
'''        If sStat = "" Then
'''            sStat = frmGraphics.lblStat(i).Caption
'''        Else
'''            If sStat <> frmGraphics.lblStat(i).Caption Then
'''                bMulti = True
'''                Exit For
'''            End If
'''        End If
'''    Next i
    frmGraphics.bResetting = True
    If bMulti Then
        Call frmGraphics.lblFilterAll_Click
    Else
        Select Case UCase(Left(sStat, 1))
            Case "I": Call frmGraphics.imgStatus_Click(1)
            Case "C": Call frmGraphics.imgStatus_Click(2)
            Case "R": Call frmGraphics.imgStatus_Click(3)
        End Select
    End If
    frmGraphics.bResetting = False
    
    Unload Me
    Unload frmSearch
End Sub

Private Sub cmdLoad_Click()
    Dim i As Integer
    
    Select Case pFrom
        Case "GH", "GA"
            lOpenInViewer = CLng(Mid(tvw1.SelectedItem.Key, 2))
            With frmGraphics
                .PassFBCN = tvw1.SelectedItem.Parent.Text
                .PassBCC = Mid(tvw1.SelectedItem.Parent.Key, 2)
                .PassGNode = Mid(tvw1.SelectedItem.Key, 2)
            End With
            Unload Me
            Unload frmSearch
        Case "DIL"
'''            lOpenInViewer = CLng(Mid(tvw1.SelectedItem.key, 2))
            
            Call frmDIL.LoadGraphic(0, Mid(tvw1.SelectedItem.Key, 2), _
                        tvw1.SelectedItem.Text, "Search Result...")
            
            frmDIL.lstResult.Clear: frmDIL.lstResultPath.Clear
            For i = 0 To lstFiles.ListCount - 1
                frmDIL.lstResult.AddItem lstFiles.List(i)
                frmDIL.lstResult.ItemData(frmDIL.lstResult.NewIndex) = lstFiles.ItemData(i)
                frmDIL.lstResultPath.AddItem lstPaths.List(i)
            Next i
            If frmDIL.lstResult.ListCount > 0 Then
                If frmDIL.lstResult.ListCount <= 25 Then
                    frmDIL.lstResult.Height = (frmDIL.lstResult.ListCount * 195) + 60
                Else
                    frmDIL.lstResult.Height = (25 * 195) + 60
                End If
                frmDIL.picResult.Height = frmDIL.lstResult.Top + _
                            frmDIL.lstResult.Height + 120
                frmDIL.picResult.Top = frmDIL.picMenu2.Top + frmDIL.picMenu2.Height
                frmDIL.picResult.Left = 0 ''frmDIL.ScaleWidth - 240 - frmDIL.picResult.Width
                frmDIL.picResult.Visible = True
            End If
            
            Unload Me
            Unload frmSearch
    End Select
    
'    Unload Me
End Sub

Private Sub cmdRevise_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    If pFrom <> "GA" Then
        cmdExport.Visible = False
    End If
'''    strSQL = "SELECT GID FROM " & pTBL & " " & pSQL
    strSQL = "SELECT GM.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT, " & _
                "GM.GID, GM.GDESC, GM.GFORMAT, GM.GPATH " & _
                "FROM ANNOTATOR.GFX_MASTER GM, " & F0101 & " AB " & _
                "WHERE GM.GID > 0 " & _
                "AND GM.GID IN (" & pSQL & ") " & _
                "AND GM.AN8_CUNO = AB.ABAN8 " & _
                "ORDER BY CLIENT, CUNO, UPPER(GM.GDESC)"
    Call PopTree(strSQL)
    If tvw1.Nodes.Count > 0 Then
        cmdExport.Enabled = True
        cmdDownload.Enabled = True
'''        cmdDownloadAdd.Enabled = True
    Else
        cmdExport.Enabled = False
        cmdDownload.Enabled = False
'''        cmdDownloadAdd.Enabled = False
    End If
End Sub

Public Sub PopTree(sSQL As String)
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sCNode As String, sDNode As String, sTNode As String, _
                sSNode As String, sGNode As String
    Dim sDesc As String
    Dim iImage As Integer
    Dim lCnt As Long
    
    lCnt = 0
    tvw1.ImageList = ImageList1
    tvw1.Nodes.Clear
    lstFiles.Clear: lstPaths.Clear
    Set rst = Conn.Execute(sSQL)
    Do While Not rst.EOF
        lCnt = lCnt + 1
        If sCNode <> "c" & rst.Fields("CUNO") Then
            sCNode = "c" & rst.Fields("CUNO")
            sDesc = Trim(rst.Fields("CLIENT"))
            iImage = 1
            Set nodX = tvw1.Nodes.Add(, , sCNode, sDesc, iImage)
            sDNode = ""
        End If

        sGNode = "g" & rst.Fields("GID")
        sDesc = Trim(rst.Fields("GDESC"))
        Select Case UCase(Trim(rst.Fields("GFORMAT")))
            Case "JPG": iImage = 2
            Case "BMP": iImage = 3
            Case "PDF": iImage = 4
            Case "PPT": iImage = 5
            Case "AVI": iImage = 6
            Case "MPG": iImage = 7
            Case "MOV": iImage = 8
        End Select
'        Set nodX = tvw1.Nodes.Add(sSNode, tvwChild, sGNode, sDesc)
        Set nodX = tvw1.Nodes.Add(sCNode, tvwChild, sGNode, sDesc, iImage)
        
        lstFiles.AddItem Trim(rst.Fields("GDESC"))
        lstFiles.ItemData(lstFiles.NewIndex) = rst.Fields("GID")
        lstPaths.AddItem Trim(rst.Fields("GPATH"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    lblCount.Caption = lCnt & " Matches"
    If tvw1.Nodes.Count > 0 Then tvw1.Nodes(1).Expanded = True
End Sub

Public Sub AddToCart()
    Dim i As Integer, iLoc As Integer, iCnt As Integer
    Dim sMess As String, sLoc As String, sPath As String, sFile As String, sFormat As String
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lDLID As Long, lGID As Long
    
    iCnt = 0
    For i = 1 To tvw1.Nodes.Count
        If UCase(Left(tvw1.Nodes(i).Key, 1)) = "G" Then
            lGID = CLng(Mid(tvw1.Nodes(i).Key, 2))
            Select Case tvw1.Nodes(i).Image
                Case 2: sFormat = "JPG"
                Case 3: sFormat = "BMP"
                Case 4: sFormat = "PDF"
                Case 5: sFormat = "PPT"
                Case 6: sFormat = "AVI"
                Case 7: sFormat = "MPG"
                Case 8: sFormat = "MOV"
            End Select
            ''CHECK IF GID EXISTS IN USER'S CART''
            strSelect = "SELECT DLID FROM ANNOTATOR.ANO_DOWNLOAD " & _
                        "WHERE USER_SEQ_ID = " & UserID & " " & _
                        "AND FILE_TYPE = 'DIL' " & _
                        "AND FILE_ID = " & lGID
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                rst.Close: Set rst = Nothing
                MsgBox "'" & tvw1.Nodes(i).Text & "." & sFormat & "' already exists in your Download Cart", _
                            vbExclamation, "Skipping Selection..."
                GoTo SkipIt
            End If
            rst.Close: Set rst = Nothing
            
            sLoc = "DIL Search Result (" & tvw1.Nodes(i).Parent.Text & ")"

            strSelect = "SELECT GPATH FROM ANNOTATOR.GFX_MASTER WHERE GID = " & lGID
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sPath = Trim(rst.Fields("GPATH"))
            Else
                rst.Close: Set rst = Nothing
                MsgBox "The source file for '" & tvw1.Nodes(i).Text & "." & sFormat & "' " & _
                            "could not be found", vbExclamation, "Skipping file..."
                GoTo SkipIt
            End If
            rst.Close: Set rst = Nothing
            
            sFile = tvw1.Nodes(i).Text & "." & LCase(sFormat)
            
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lDLID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            
            strInsert = "INSERT INTO ANNOTATOR.ANO_DOWNLOAD " & _
                        "(DLID, USER_SEQ_ID, FILE_TYPE, DLSTATUS, FILE_ID, " & _
                        "SOURCE_PATH, FILE_NAME, SOURCE_DESC, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lDLID & ", " & UserID & ", 'DIL', 1, " & lGID & ", " & _
                        "'" & DeGlitch(sPath) & "', " & _
                        "'" & DeGlitch(sFile) & "', " & _
                        "'" & DeGlitch(sLoc) & "', " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            
            frmDIL.lblDLCnt.Caption = CInt(frmDIL.lblDLCnt.Caption) + 1
            frmDIL.picDLCart.Visible = True
            iCnt = iCnt + 1
        End If
SkipIt:
    Next i
    
    If iCnt = 1 Then
        MsgBox "A new file has been added to your Download Cart.  It is available " & _
                    "to download by clicking on the Cart icon above.", _
                    vbInformation, "New Download Cart file..."
    ElseIf iCnt <> 0 Then
        MsgBox "( " & iCnt & " ) new files have been added to your Download Cart.  They are available " & _
                    "to download by clicking on the Cart icon above.", _
                    vbInformation, "New Download Cart files..."
    End If
End Sub


Private Sub mnuRemove_Click()
    If imx1.Visible Then
        imx1.FileName = ""
        imx1.Visible = False
'    ElseIf pdf1.Visible Then
    ElseIf xpdf1.Visible Then
        'pdf1.src = ""
        ''pdf1.LoadFile ("")
'        pdf1.Visible = False
        xpdf1.Visible = False
    End If
    tvw1.Nodes.Remove (tvw1.SelectedItem.Key)
    cmdLoad.Enabled = False
    
    If UCase(Left(tvw1.SelectedItem.Key, 1)) = "G" Then
        Call tvw1_NodeClick(tvw1.SelectedItem)
    End If
    lblCount.Caption = GetCount & " Matches"
End Sub

'''Private Sub tvw1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'''    If Button = vbRightButton Then
'''        Debug.Print tvw1.SelectedItem.key
'''    End If
'''End Sub

Private Sub tvw1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And UCase(Left(tvw1.SelectedItem.Key, 1)) = "G" Then
        Me.PopupMenu mnuEdit
    End If
End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If UCase(Left(Node.Key, 1)) = "G" Then
        strSelect = "SELECT GPATH, GFORMAT " & _
                    "FROM ANNOTATOR.GFX_MASTER " & _
                    "WHERE GID = " & Mid(Node.Key, 2)
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            imx1.Visible = False
            picPDF.Visible = False
            Select Case UCase(Trim(rst.Fields("GFORMAT")))
                Case "JPG", "BMP"
                    imx1.FileName = Trim(rst.Fields("GPATH"))
                    imx1.Refresh
                    imx1.Visible = True
                    cmdLoad.Enabled = True
                Case "PDF"
'                    pdf1.src = Trim(rst.Fields("GPATH"))
                    xpdf1.loadFile (Trim(rst.Fields("GPATH")))
                    xpdf1.Zoom = xpdf1.zoomWidth
                    xpdf1.Visible = True
'                    pdf1.LoadFile (Trim(rst.Fields("GPATH")))
'                    pdf1.Visible = True
                    picPDF.Visible = True
                    cmdLoad.Enabled = True
                Case Else
                    MsgBox "Search Result Viewer is unable to display this format", _
                                vbExclamation, "Sorry..."
            End Select
        End If
        rst.Close: Set rst = Nothing
    End If
End Sub


Public Function GetCount() As Long
    Dim i As Long, iCnt As Integer
    
    iCnt = 0
    For i = 1 To tvw1.Nodes.Count
        If UCase(Left(tvw1.Nodes(i).Key, 1)) = "G" Then iCnt = iCnt + 1
    Next i
    If iCnt = 0 Then
        cmdDownload.Enabled = False
'''        cmdDownloadAdd.Enabled = False
    End If
    GetCount = iCnt
    
End Function

