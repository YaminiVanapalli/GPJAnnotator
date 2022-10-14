VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPhoto 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhoto.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   9000
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Show Graphics Also"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   7140
      Width           =   1815
   End
   Begin IMAGXPR5LibCtl.ImagXpress imxPhoto 
      Height          =   7095
      Left            =   2100
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   12515
      ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
      ErrCode         =   1235493135
      ErrInfo         =   -443604675
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
   Begin VB.VScrollBar vsc1 
      Height          =   6855
      Left            =   1740
      SmallChange     =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   180
      ScaleHeight     =   6795
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   180
      Width           =   1575
      Begin VB.PictureBox picInner 
         BackColor       =   &H80000005&
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
         Height          =   2400
         Left            =   0
         ScaleHeight     =   2400
         ScaleWidth      =   1515
         TabIndex        =   1
         Top             =   0
         Width           =   1515
         Begin IMAGXPR5LibCtl.ImagXpress imx1 
            Height          =   960
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   1693
            ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
            ErrCode         =   1235493135
            ErrInfo         =   -443604675
            Persistence     =   -1  'True
            _cx             =   132055184
            _cy             =   1
            FileName        =   ""
            MouseIcon       =   "frmPhoto.frx":08CA
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Shape shp1 
            BackColor       =   &H8000000C&
            BorderColor     =   &H8000000D&
            BorderWidth     =   6
            Height          =   1200
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   1515
         End
      End
   End
   Begin VB.Label lblNone 
      AutoSize        =   -1  'True
      Caption         =   "Sorry... There are no Associated Images."
      Height          =   195
      Left            =   2100
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuPPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuPDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPEmail 
         Caption         =   "Email Photo..."
      End
      Begin VB.Menu mnuPDownload 
         Caption         =   "Download Photo..."
      End
      Begin VB.Menu mnuPDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCurrIndex As Integer
Dim lBCC As Long, pEID As Long, pFCCD As Long
Dim sElt As String
Dim tLink As String, tIn As String
Dim sHDR As String, tFBCN As String

Public Property Get PassLink() As String
    PassLink = tLink
End Property
Public Property Let PassLink(ByVal vNewValue As String)
    tLink = vNewValue
End Property

Public Property Get PassIn() As String
    PassIn = tIn
End Property
Public Property Let PassIn(ByVal vNewValue As String)
    tIn = vNewValue
End Property

Public Property Get PassEID() As Long
    PassEID = pEID
End Property
Public Property Let PassEID(ByVal vNewValue As Long)
    pEID = vNewValue
End Property

Public Property Get PassFCCD() As Long
    PassFCCD = pFCCD
End Property
Public Property Let PassFCCD(ByVal vNewValue As Long)
    pFCCD = vNewValue
End Property


Private Sub chkAll_Click()
    Dim sIN As String
    Dim i As Integer
    
    Screen.MousePointer = 11
    
    Select Case chkAll.Value
        Case 1
            sIN = "1, 2, 3, 4"
        Case 0
            sIN = tIn
    End Select
    
    i = GetPhotos(lBCC, sElt, sIN, pEID)
    
    If i > 0 Then
        Call imx1_Click(0)
        picInner.Height = imx1(i - 1).Top + imx1(i - 1).Height + 150
    Else
        lblNone.Visible = True
        picInner.Height = 0
        imxPhoto.FileName = ""
        shp1.Visible = False
    End If
    
'''    picInner.Height = imx1(i - 1).Top + imx1(i - 1).Height + 150
    
    Call SetScroll
    
'''    If picInner.Height > picOuter.ScaleHeight Then
'''        vsc1.Max = (picInner.Height / 100) - (picOuter.ScaleHeight / 100)
'''        vsc1.Visible = True
'''        vsc1.Value = 0
'''        vsc1.LargeChange = picOuter.ScaleHeight / 100
'''    End If
    
    
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim sIN As String
    
    iCurrIndex = -1
    If pFCCD = 0 Then
        If pEID = 0 Then
            lBCC = CLng(Left(tLink, 8))
            sElt = Mid(tLink, 10)
        Else
            sElt = tLink
        End If
        
        sIN = tIn
        Select Case sIN
            Case "1": chkAll.Caption = "Show Graphics Also"
            Case "2, 3": chkAll.Caption = "Show Photos Also"
        End Select
                
        i = GetPhotos(lBCC, sElt, sIN, pEID)
            
        
        
    ElseIf pFCCD > 0 Then
        chkAll.Visible = False
        i = GetFCCDPhotos(pFCCD)
        
    End If
    
    If i > 0 Then
        Call imx1_Click(0)
        picInner.Height = imx1(i - 1).Top + imx1(i - 1).Height + 150
        picInner.Visible = True
    Else
        lblNone.Visible = True
        picInner.Height = 0
    End If
        
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 3000 Then Me.Height = 3000
        
    chkAll.Top = Me.ScaleHeight - 300
    picOuter.Height = Me.ScaleHeight - 405 - 180
    vsc1.Height = picOuter.Height
    Call SetScroll
    imxPhoto.Width = Me.ScaleWidth - imxPhoto.Left - 180
    imxPhoto.Height = Me.ScaleHeight - 360

End Sub

Private Sub imx1_Click(Index As Integer)
    Dim sJPGfile As String
    
    If Me.WindowState = 1 Then Exit Sub
    If iCurrIndex = Index Then Exit Sub
    
'    If imx1(Index).FileName = imxPhoto.FileName Then Exit Sub
    
    imxPhoto.Update = False
    If InStr(1, imx1(Index).FileName, "thb_") = 0 Then
        imxPhoto.FileName = imx1(Index).FileName
    Else
        sJPGfile = "\\DETMSFS01\GPJAnnotator\Floorplans\FacilPho\"
        sJPGfile = sJPGfile & Mid(imx1(Index).FileName, InStr(1, imx1(Index).FileName, "thb_") + 4)
        If Dir(sJPGfile, vbNormal) <> "" Then
            imxPhoto.FileName = sJPGfile
        Else
            MsgBox "JPG file not found - loading Thumbnail into viewer", vbExclamation, "Missing file..."
            imxPhoto.FileName = imx1(Index).FileName
        End If
    End If
    imxPhoto.Update = True
    If imxPhoto.FileName <> "" Then
        imxPhoto.Visible = True
    Else
        imxPhoto.Visible = False
    End If
    imxPhoto.Refresh
    
    
    shp1.Top = Index * 1200
    shp1.Visible = True
    
    If imxPhoto.FileName <> "" Then
        Me.Caption = sHDR & "  ( " & imx1(Index).ToolTipText & " )"
        iCurrIndex = Index
    Else
        Me.Caption = ""
        iCurrIndex = -1
    End If
    
End Sub


Private Sub mnuDownload_Click()

End Sub

Private Sub imxPhoto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If pFCCD > 0 And Button = vbRightButton Then
    
    If Button = vbRightButton Then
        mnuPEmail.Enabled = CBool(pFCCD > 0)
        mnuPDownload.Enabled = CBool(pFCCD > 0)
        Me.PopupMenu mnuRightClick
    End If
End Sub

Private Sub mnuPDownload_Click()
    Select Case pFCCD
        Case 0
            With frmBrowse
                .PassFrom = UCase(Me.Name) & "-ELEM"
                .PassGID = CLng(imx1(iCurrIndex).Tag)
                .PassFILETYPE = "jpg"
                .Show 1
            End With
        Case Else
            With frmBrowse
                .PassFrom = UCase(Me.Name)
                .PassFacil = Mid(frmFacil.lblWelcome.Caption, 12) ''  pHdr '' Mid(lblWelcome.Caption, 12)
                .PassFCCD = pFCCD
                .PassGID = CLng(imx1(iCurrIndex).Tag)
                .PassFILETYPE = "jpg"
                .Show 1
            End With
    End Select
End Sub

Private Sub mnuPEmail_Click()
    frmEmailFile.PassHDR = Mid(frmFacil.lblWelcome.Caption, 12)
    frmEmailFile.PassFrom = UCase(Me.Name)
    frmEmailFile.PassFCCD = pFCCD
    frmEmailFile.PassGID = CLng(imx1(iCurrIndex).Tag)
    frmEmailFile.Show 1, Me
End Sub

Private Sub mnuPPrint_Click()
    Dim lWidth As Long, lHgt As Long, lTop As Long, lLeft As Long
    Dim pWidth As Long, pHgt As Long
    
    
    cdl1.CancelError = True
'    cdl1.Flags = cdlPDPrintSetup
'    cdl1.FileName = sFile
    Printer.TrackDefault = True
    On Error Resume Next
    cdl1.ShowPrinter
    
    If Err = cdlCancel Then Exit Sub
    
    Me.MousePointer = 11
    
    Printer.Orientation = cdl1.Orientation
    Printer.Copies = cdl1.Copies
    
    pWidth = Printer.Width - 2160 ''1440
    pHgt = Printer.Height - 2160 ''1440
    If pWidth > imxPhoto.Width And pHgt > imxPhoto.Height Then
        lWidth = imxPhoto.Width
        lHgt = imxPhoto.Height
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    ElseIf pHgt > imxPhoto.Height Then
        lWidth = pWidth
        lHgt = (pWidth / imxPhoto.Width) * imxPhoto.Height
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    ElseIf pWidth > imxPhoto.Width Then
        lHgt = pHgt
        lWidth = (pHgt / imxPhoto.Height) * imxPhoto.Width
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    Else
        Select Case (imxPhoto.Width / imxPhoto.Height)
            Case Is > (Printer.Width / Printer.Height)
                ''USE WIDTH''
                lWidth = pWidth
                lHgt = (pWidth / imxPhoto.Width) * imxPhoto.Height
            Case Else
                ''USER HEIGHT''
                lHgt = pHgt
                lWidth = (pHgt / imxPhoto.Height) * imxPhoto.Width
        End Select
        
'        If (pWidth / pHgt) < (imxphoto.Width / imxphoto.Height) Then
'            lWidth = pWidth
'            lHgt = (pWidth / imxphoto.Width) * imxphoto.Height
''            lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
''            lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
'        Else
'            lHgt = pHgt
'            lWidth = (pHgt / imxphoto.Height) * imxphoto.Width
''            lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
''            lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
'        End If
    End If

    lTop = 1080 - ((Printer.Height - Printer.ScaleHeight) / 2)
    lLeft = 1080 - ((Printer.Width - Printer.ScaleWidth) / 2) + ((pWidth - lWidth) / 2)

    Printer.PaintPicture imxPhoto.Picture, lLeft, lTop, lWidth, lHgt
    
'    Printer.PaintPicture imxphoto.Picture, 0, 0, Printer.Width, Printer.Height
    Printer.EndDoc
    Me.MousePointer = 0
    
End Sub

Private Sub vsc1_Change()
    picInner.Top = CLng(vsc1.Value) * (-100)
End Sub

Private Sub vsc1_Scroll()
    picInner.Top = CLng(vsc1.Value) * (-100)
End Sub

Public Function GetPhotos(lBCC As Long, sElt As String, sIN As String, tEID As Long) As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim lEID As Long
    Dim i As Integer, iDel As Integer
    
    Dim sPath As String, sFile As String
    sPath = "\\DETWEB05\GPJWS\ElementPhotos\"
    
    i = -1
    If tEID = 0 Then
        strSelect = "SELECT AB.ABALPH AS CLIENT, " & _
                    "GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM, " & F0101 & " AB " & _
                    "WHERE GE.ELTID IN " & _
                    "(SELECT E.ELTID " & _
                    "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K " & _
                    "WHERE K.AN8_CUNO = " & lBCC & " " & _
                    "AND K.KITID = E.KITID " & _
                    "AND E.ELTFNAME = '" & sElt & "') " & _
                    "AND GE.GID = GM.GID " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "AND GM.GTYPE IN (" & sIN & ") " & _
                    "ORDER BY GM.GTYPE, GM.GDESC"
    Else
        strSelect = "SELECT AB.ABALPH AS CLIENT, " & _
                    "GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM, " & F0101 & " AB " & _
                    "WHERE GE.ELTID = " & tEID & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "AND GM.GTYPE IN (" & sIN & ") " & _
                    "ORDER BY GM.GTYPE, GM.GDESC"
    End If
    
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close: Set rst = Nothing
        On Error Resume Next
        Me.Caption = "Element:  " & sElt
        tFBCN = Trim(rst.Fields("CLIENT"))
        Screen.MousePointer = 0
        Exit Function
    Else
        sHDR = "Element:  " & sElt
    End If
    Do While Not rst.EOF
        i = i + 1
        If i + 1 > imx1.Count Then
            Load imx1(i)
            imx1(i).Top = (1200 * i) + 120
        End If
        If UCase(Trim(rst.Fields("GFORMAT"))) <> "PDF" Then
            sFile = sPath & rst.Fields("GID") & "." & rst.Fields("GFORMAT")
            imx1(i).Update = False
            imx1(i).PICThumbnail = 2
            imx1(i).FileName = sFile ''Trim(rst.Fields("GPATH"))
            imx1(i).ToolTipText = Trim(rst.Fields("GDESC"))
            imx1(i).Tag = rst.Fields("GID") '' Trim(rst.Fields("GDESC"))
            imx1(i).Update = True
            imx1(i).Buttonize 1, 1, 50
            imx1(i).Visible = True
            imx1(i).Refresh
        End If
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    If imx1.Count > i + 1 Then
        For iDel = i + 1 To imx1.Count - 1
            Unload imx1(iDel)
        Next iDel
    End If
    
    If i < 0 Then lblNone.Visible = True Else lblNone.Visible = False
    
    GetPhotos = i + 1
End Function

Public Function GetFCCDPhotos(tFCCD As Long) As Integer
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
        Exit Function
    Else
        sHDR = "Facility:  " & Trim(rst.Fields("FACIL"))
    End If
    Do While Not rst.EOF
        i = i + 1
        If i + 1 > imx1.Count Then
            Load imx1(i)
            imx1(i).Top = (1200 * i) + 120
        End If

        imx1(i).Update = False
        imx1(i).PICThumbnail = 2
        If Dir(sThumbPath & "thb_" & rst.Fields("GID") & ".jpg", vbNormal) <> "" Then
            imx1(i).FileName = sThumbPath & "thb_" & rst.Fields("GID") & ".jpg"
        Else
            imx1(i).FileName = Trim(rst.Fields("GPATH"))
        End If
        imx1(i).ToolTipText = Trim(rst.Fields("GDESC"))
        imx1(i).Tag = rst.Fields("GID") '' Trim(rst.Fields("GDESC"))
        imx1(i).Update = True
        imx1(i).Buttonize 1, 1, 50
        imx1(i).Visible = True
        imx1(i).Refresh
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    If imx1.Count > i + 1 Then
        For iDel = i + 1 To imx1.Count - 1
            Unload imx1(iDel)
        Next iDel
    End If
    
    If i < 0 Then lblNone.Visible = True Else lblNone.Visible = False
    
    GetFCCDPhotos = i + 1
End Function

Public Sub SetScroll()
    If picInner.Height > picOuter.ScaleHeight Then
        vsc1.Max = (picInner.Height / 100) - (picOuter.ScaleHeight / 100)
        vsc1.Visible = True
        vsc1.Value = 0
        vsc1.SmallChange = 1200 / 100
        vsc1.LargeChange = picOuter.ScaleHeight / 100
    Else
        picInner.Top = 0
        vsc1.Visible = False
    End If
End Sub
