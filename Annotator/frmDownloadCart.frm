VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{A0369ABE-B6D8-11D3-901D-00207816FA15}#3.0#0"; "aghypertext.ocx"
Begin VB.Form frmDownloadCart 
   Caption         =   "Download Cart..."
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12255
   Icon            =   "frmDownloadCart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloadCart.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloadCart.frx":0B44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download Selections..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click this to Download all ""Green Lighted"" entries above..."
      Top             =   6420
      Width           =   3615
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Selections from Download Cart"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click this to delete all ""Green Lighted"" entries above..."
      Top             =   6420
      Width           =   3615
   End
   Begin VB.PictureBox picOuter 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1440
      ScaleHeight     =   1815
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   1260
      Width           =   1335
      Begin VB.PictureBox picInner 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         ScaleHeight     =   1575
         ScaleWidth      =   1275
         TabIndex        =   2
         Top             =   0
         Width           =   1275
         Begin IMAGXPR5LibCtl.ImagXpress imx 
            Height          =   960
            Index           =   0
            Left            =   60
            TabIndex        =   3
            ToolTipText     =   "Click to view Full-Size..."
            Top             =   60
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   1693
            ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
            ErrCode         =   1000956018
            ErrInfo         =   762537149
            Persistence     =   -1  'True
            _cx             =   132383232
            _cy             =   1
            FileName        =   ""
            MouseIcon       =   "frmDownloadCart.frx":0C9E
            MousePointer    =   99
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
            BackColor       =   -2147483643
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
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx 
      Height          =   3795
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   6694
      _Version        =   393216
      Rows            =   25
      Cols            =   11
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      GridLines       =   0
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
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
   Begin AgHyperText.AgHyperTxt lnkSelect 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HyperLink       =   "Clear All"
   End
   Begin AgHyperText.AgHyperTxt lnkSelect 
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HyperLink       =   "Select All"
   End
   Begin VB.Label lblFileCnt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10500
      TabIndex        =   12
      Top             =   68
      Width           =   1635
   End
   Begin VB.Label lblCntHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Files Selected in my Download Cart:  "
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
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Width           =   3525
   End
   Begin VB.Label lblTotalSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7380
      TabIndex        =   8
      Top             =   5700
      Width           =   1635
   End
   Begin VB.Label lblTotalHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected Download Size:  "
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
      Left            =   4080
      TabIndex        =   7
      Top             =   5940
      Width           =   2625
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files in my Download Cart:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2145
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   480
      Picture         =   "frmDownloadCart.frx":0FB8
      Top             =   5220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmDownloadCart.frx":1542
      Top             =   5460
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDownloadCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sSelPath(0 To 1) As String

Public pDLType As String
Public pDownloadPath As String

Public Property Get PassDLType() As String
    PassDLType = pDLType
End Property
Public Property Let PassDLType(ByVal vNewValue As String)
    pDLType = vNewValue
End Property

Public Property Get PassDLPath() As String
    PassDLPath = pDownloadPath
End Property
Public Property Let PassDLPath(ByVal vNewValue As String)
    pDownloadPath = vNewValue
End Property



Private Sub cmdDelete_Click()
    Dim i As Integer, iCnt As Integer
    Dim Resp As VbMsgBoxResult
    Dim strDelete As String
    
    On Error GoTo ResetPointer
    Screen.MousePointer = 11
    iCnt = 0
    For i = 1 To flx.Rows - 1
        If flx.TextMatrix(i, 1) = "1" Then iCnt = iCnt + 1
    Next i
    
    Select Case iCnt
        Case Is = 0
            Screen.MousePointer = 0
            MsgBox "No files have been selected for removal", _
                        vbExclamation, "Nothing is being removed..."
            Exit Sub
        Case Is = 1
            Resp = MsgBox("Are you certain you would like to remove the one selected file " & _
                        "from your Download Cart?", vbQuestion + vbYesNoCancel, "Just Checking...")
        Case Else
            Resp = MsgBox("Are you certain you would like to remove the (" & iCnt & ") selected files " & _
                        "from your Download Cart?", vbQuestion + vbYesNoCancel, "Just Checking...")
    End Select
    
    If Resp = vbYes Then
        For i = flx.Rows - 1 To 1 Step -1
            If flx.TextMatrix(i, 1) = "1" Then
                strDelete = "DELETE FROM ANNOTATOR.ANO_DOWNLOAD " & _
                            "WHERE DLID = " & flx.TextMatrix(i, 0) & " " & _
                            "AND USER_SEQ_ID = " & UserID
                Conn.Execute (strDelete)
                
            End If
        Next i
        
        Call GetDLFiles(pDLType, 0)
    Else
        Screen.MousePointer = 0
    End If
    
    Screen.MousePointer = 0
Exit Sub
ResetPointer:
    Screen.MousePointer = 0
End Sub

Private Sub cmdDownload_Click()
    Dim i As Integer
    Dim strUpdate As String
    Dim RetVal
    
    pDownloadPath = ""
    frmBrowse.PassFrom = Me.Name
    frmBrowse.Show 1, Me
    
'''    MsgBox pDownloadPath
    If pDownloadPath <> "" Then
        For i = 1 To flx.Rows - 1
            If flx.TextMatrix(i, 1) = "1" Then
                strUpdate = "UPDATE ANNOTATOR.ANO_DOWNLOAD SET " & _
                            "DL_PATH = '" & DeGlitch(pDownloadPath) & "\', " & _
                            "DLSTATUS = 5, " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                            "WHERE DLID = " & flx.TextMatrix(i, 0)
                Conn.Execute (strUpdate)
            End If
        Next i
        RetVal = Shell(App.Path & "\Downloader.exe", vbNormalFocus)
        Call RemoveSelections
    End If
    
    
    
    
End Sub

Private Sub flx_Click()
    Dim i As Integer
    Dim lColor As Long
    
    If flx.Rows = 1 Then Exit Sub
    
    If flx.ColSel = 2 Then
        i = Abs(CInt(flx.TextMatrix(flx.RowSel, 1)) - 1)
        flx.TextMatrix(flx.RowSel, 1) = i
        If Dir(sSelPath(i), vbNormal) <> "" Then
            Set flx.CellPicture = LoadPicture(sSelPath(i), 0) '' img(iSelected).Picture '' ImageList1.ListImages(2).Picture  '' LoadPicture(App.Path & "\Check-Yes4.bmp", 0) '' img(iSelected).Picture
        Else
            Set flx.CellPicture = img(i).Picture
        End If
'''        Set flx.CellPicture = img(i).Picture
        Select Case i
            Case 0: lColor = RGB(180, 180, 180)
            Case 1: lColor = 0
        End Select
        For i = 2 To flx.Cols - 1
            flx.Col = i: flx.CellForeColor = lColor
        Next i
        Call ReTotalSize
    End If
End Sub

Private Sub flx_Scroll()
    picInner.Top = ((flx.TopRow - 1) * 1080) * -1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    
'''    For i = 1 To flx.Rows - 1
'''
'''
'''
'''        flx.Row = i: flx.Col = 2
'''        Set flx.CellPicture = img(0).Picture
'''        flx.CellPictureAlignment = 4
'''        flx.TextMatrix(i, 1) = 0
'''
'''        flx.TextMatrix(i, 3) = "Digital Image Library file"
'''    Next i

    sSelPath(0) = App.Path & "\Check-No6.bmp"
    sSelPath(1) = App.Path & "\Check-Yes6.bmp"
    
    flx.ColWidth(0) = 0
    flx.ColWidth(1) = 0
    flx.ColWidth(2) = 900: flx.ColAlignment(2) = 4: flx.TextMatrix(0, 2) = "Include"
    flx.ColWidth(3) = 1500: flx.ColAlignment(3) = 4: flx.TextMatrix(0, 3) = "File Type"
    flx.ColWidth(4) = 1400: flx.ColAlignment(4) = 4: flx.TextMatrix(0, 4) = "File Image"
    flx.ColAlignment(5) = 1: flx.TextMatrix(0, 5) = "File Source Location": flx.FixedAlignment(5) = 4
    flx.ColAlignment(6) = 1: flx.TextMatrix(0, 6) = "File Name": flx.FixedAlignment(6) = 4
    flx.ColWidth(7) = 1200: flx.ColAlignment(7) = 7: flx.TextMatrix(0, 7) = "File Size": flx.FixedAlignment(7) = 4
    flx.ColWidth(8) = 240: flx.ColAlignment(8) = 4
    flx.ColWidth(9) = 0
    flx.ColWidth(10) = 0
    
    picOuter.Top = flx.Top + flx.RowHeight(0)
    picOuter.Left = flx.Left + flx.ColPos(4) - 5
    picOuter.Width = flx.ColWidth(4) - 10
    picInner.Height = 1080 * (flx.Rows - 1)
    picInner.Width = picOuter.ScaleWidth
    
    Call GetDLFiles(pDLType, 1)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.ScaleWidth < 7200 Or Me.ScaleHeight < 2000 Then Exit Sub
    
    flx.Width = Me.ScaleWidth - (flx.Left * 2)
    flx.Height = Me.ScaleHeight - flx.Top - 1200
    flx.ColWidth(5) = (flx.Width - 5240) * 0.5
    flx.ColWidth(6) = (flx.Width - 5240) * 0.5
    
    picOuter.Height = flx.Height - flx.RowHeight(0)
    
    lnkSelect(0).Top = flx.Top + flx.Height + 60
    lnkSelect(1).Top = lnkSelect(0).Top
    
    lblTotalHdr.Top = lnkSelect(0).Top + (lnkSelect(0).Height / 2) - (lblTotalHdr.Height / 2)
    lblTotalHdr.Left = flx.Left + flx.ColPos(7) - lblTotalHdr.Width
    lblTotalSize.Top = lblTotalHdr.Top + (lblTotalHdr.Height / 2) - (lblTotalSize.Height / 2)
    lblTotalSize.Width = flx.ColWidth(7)
    lblTotalSize.Left = flx.Left + flx.ColPos(7)
    
    cmdDelete.Top = Me.ScaleHeight - 120 - cmdDelete.Height
    cmdDownload.Top = cmdDelete.Top
    cmdDelete.Left = (Me.ScaleWidth / 2) - 120 - cmdDelete.Width
    cmdDownload.Left = (Me.ScaleWidth / 2) + 120
    
    lblFileCnt.Width = lblTotalSize.Width
    lblFileCnt.Left = lblTotalSize.Left
    lblCntHdr.Left = lblFileCnt.Left - lblCntHdr.Width
    
    picInner.Top = ((flx.TopRow - 1) * 1080) * -1
End Sub

Private Sub imx_Click(Index As Integer)
    frmHTMLViewer.PassFile = flx.TextMatrix(Index, 9)
    frmHTMLViewer.PassFrom = Me.Name
    frmHTMLViewer.PassDFile = flx.TextMatrix(Index, 6)
    frmHTMLViewer.PassHDR = flx.TextMatrix(Index, 6)
'''    frmHTMLViewer.PassGID = lGID
    frmHTMLViewer.Show 1, Me
End Sub

Private Sub lnkSelect_Click(Index As Integer)
    Dim i As Integer, iCol As Integer
    Dim lColor As Long
    
    Screen.MousePointer = 11
    flx.Visible = False
    For i = 1 To flx.Rows - 1
        If flx.TextMatrix(i, 1) <> Index Then
            flx.TextMatrix(i, 1) = Index
            flx.Row = i: flx.Col = 2
            If Dir(sSelPath(Index), vbNormal) <> "" Then
                Set flx.CellPicture = LoadPicture(sSelPath(Index), 0) '' img(iSelected).Picture '' ImageList1.ListImages(2).Picture  '' LoadPicture(App.Path & "\Check-Yes4.bmp", 0) '' img(iSelected).Picture
            Else
                Set flx.CellPicture = img(Index).Picture
            End If
        End If
        Select Case Index
            Case 0: lColor = RGB(180, 180, 180)
            Case 1: lColor = 0
        End Select
        For iCol = 2 To flx.Cols - 1
            flx.Col = iCol: flx.CellForeColor = lColor
        Next iCol
    Next i
    Call ReTotalSize
    flx.Visible = True
    Screen.MousePointer = 0
End Sub

Public Sub GetDLFiles(pType As String, iSelected As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Long
    Dim lTotal As Long, lFileLen As Long
    
    lTotal = 0
    flx.Rows = 1
    i = 0
    
    Select Case pType
        Case "ALL"
            strSelect = "SELECT * FROM ANNOTATOR.ANO_DOWNLOAD " & _
                        "WHERE USER_SEQ_ID = " & UserID & " " & _
                        "AND DLSTATUS = 1 " & _
                        "ORDER BY FILE_TYPE, SOURCE_DESC, FILE_NAME"
        Case Else
            strSelect = "SELECT * FROM ANNOTATOR.ANO_DOWNLOAD " & _
                        "WHERE USER_SEQ_ID = " & UserID & " " & _
                        "AND FILE_TYPE = '" & pType & "' " & _
                        "AND DLSTATUS = 1 " & _
                        "ORDER BY FILE_TYPE, SOURCE_DESC, FILE_NAME"
    End Select
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        i = i + 1
        flx.Rows = i + 1
        flx.RowHeight(i) = 1080
        flx.TextMatrix(i, 0) = rst.Fields("DLID")
        
        flx.TextMatrix(i, 1) = iSelected
        flx.Row = i: flx.Col = 2
        If Dir(sSelPath(1), vbNormal) <> "" Then
            Set flx.CellPicture = LoadPicture(sSelPath(iSelected), 0) '' img(iSelected).Picture '' ImageList1.ListImages(2).Picture  '' LoadPicture(App.Path & "\Check-Yes4.bmp", 0) '' img(iSelected).Picture
        Else
            Set flx.CellPicture = img(iSelected).Picture
        End If
        flx.CellPictureAlignment = 4
        
        Select Case rst.Fields("FILE_TYPE")
            Case "DIL": flx.TextMatrix(i, 3) = "Digital Image Library file"
            Case "GFX": flx.TextMatrix(i, 3) = "Graphic file"
            Case "FP": flx.TextMatrix(i, 3) = "Floorplan file"
            Case "ENG": flx.TextMatrix(i, 3) = "Engineering Drawing file"
        End Select
        
        If imx.Count <= i Then
            Load imx(i)
            Set imx(i).Container = picInner
        End If
        imx(i).Top = ((i - 1) * 1080) + 60
        imx(i).Left = 60
        If Trim(rst.Fields("FILE_TYPE")) = "DIL" Or Trim(rst.Fields("FILE_TYPE")) = "GFX" Then
            
            imx(i).FileName = GetImageFile(rst.Fields("FILE_ID"), Trim(rst.Fields("SOURCE_PATH")))
            
        Else
            
        End If
        imx(i).Visible = True
        imx(i).Update = True
        
        flx.TextMatrix(i, 5) = Trim(rst.Fields("SOURCE_DESC"))
        flx.TextMatrix(i, 6) = Trim(rst.Fields("FILE_NAME"))
        
        lFileLen = FileLen(Trim(rst.Fields("SOURCE_PATH")))
        flx.TextMatrix(i, 7) = Format(lFileLen / 1000, "#,##0") & "KB"
        lTotal = lTotal + lFileLen
        
        flx.TextMatrix(i, 9) = Trim(rst.Fields("SOURCE_PATH"))
        flx.TextMatrix(i, 10) = imx(i).FileName
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    picInner.Height = (flx.Rows - 1) * 1080
    
    Call ReTotalSize
'''    lblTotalSize.Caption = Format(lTotal / 1000, "#,##0") & "KB"
'''    lblFileCnt.Caption = flx.Rows - 1 & " of " & flx.Rows - 1
End Sub

Public Sub ReTotalSize()
    Dim i As Integer, iCnt As Integer
    Dim lTotal As Long
    
    lTotal = 0: iCnt = 0
    For i = 1 To flx.Rows - 1
        If flx.TextMatrix(i, 1) = "1" Then
            iCnt = iCnt + 1
            lTotal = lTotal + CLng(Left(flx.TextMatrix(i, 7), Len(flx.TextMatrix(i, 7)) - 2))
        End If
    Next i
    lblTotalSize.Caption = Format(lTotal, "#,##0") & "KB"
    lblFileCnt.Caption = iCnt & " of " & flx.Rows - 1
    
    cmdDownload.Enabled = CBool(iCnt)
    cmdDelete.Enabled = CBool(iCnt)
    
End Sub

Public Function GetImageFile(pGID As Long, pPath As String) As String
    Dim sChk As String, sPath As String
    
    If InStr(1, UCase(pPath), ".JPG") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\Thumbs\thb_" & pGID & ".jpg"
    ElseIf InStr(1, UCase(pPath), ".PDF") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\pdf_" & pGID & ".bmp"
        If Dir(sPath, vbNormal) = "" Then
            sPath = "\\DETMSFS01\GPJAnnotator\Graphics\acrobatid.bmp"
        End If
    ElseIf InStr(1, UCase(pPath), ".AVI") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\avi.bmp"
    ElseIf InStr(1, UCase(pPath), ".MOV") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\mov.bmp"
    ElseIf InStr(1, UCase(pPath), ".MPG") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\mpg.bmp"
    ElseIf InStr(1, UCase(pPath), ".PPS") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\pps.bmp"
    ElseIf InStr(1, UCase(pPath), ".PPT") > 0 Then
        sPath = "\\DETMSFS01\GPJAnnotator\Graphics\ppt.bmp"
    End If
    If Dir(sPath, vbNormal) <> "" Then
        GetImageFile = sPath
    Else
        GetImageFile = pPath
    End If
    
End Function

Public Sub RemoveSelections()
    Dim i As Integer
    
    ''CLEAN-UP FLX ROWS''
    For i = flx.Rows - 1 To 1 Step -1
        If flx.TextMatrix(i, 1) = "1" Then
            If flx.Rows > 2 Then flx.RemoveItem (i) Else flx.Rows = 1
        End If
    Next i
    
    ''RESET THUMBNAILS''
    For i = 1 To flx.Rows - 1
        imx(i).FileName = flx.TextMatrix(i, 10)
        imx(i).Update = True
    Next i
    picInner.Height = (flx.Rows - 1) * 1080
    
    Call ReTotalSize
End Sub
