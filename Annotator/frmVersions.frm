VERSION 5.00
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{23319180-2253-11D7-BD2E-08004608C318}#3.0#0"; "XpdfViewerCtrl.ocx"
Begin VB.Form frmVersions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmVersions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Highlighted Version"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   6060
      Width           =   3075
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Set Highlighted Version as Current Version"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3300
      TabIndex        =   5
      Top             =   6060
      Width           =   3915
   End
   Begin VB.VScrollBar vsc1 
      Height          =   5785
      Left            =   6960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      BackColor       =   &H80000005&
      Height          =   5785
      Left            =   120
      ScaleHeight     =   5730
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox picInner 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5720
         Left            =   0
         ScaleHeight     =   5715
         ScaleWidth      =   6795
         TabIndex        =   2
         Top             =   0
         Width           =   6795
         Begin XpdfViewerCtl.XpdfViewer pdf1 
            Height          =   1335
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
            showScrollbars  =   0   'False
            showBorder      =   -1  'True
            showPasswordDialog=   -1  'True
         End
         Begin IMAGXPR5LibCtl.ImagXpress imx1 
            Height          =   2400
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   4233
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
         Begin VB.Label lbl1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   1680
            TabIndex        =   4
            Top             =   2580
            Width           =   60
         End
         Begin VB.Shape shp1 
            BackColor       =   &H001CAF6F&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H001CAF6F&
            Height          =   2800
            Left            =   60
            Top             =   60
            Visible         =   0   'False
            Width           =   3320
         End
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuExpandedView 
         Caption         =   "Expanded View..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Details..."
      End
   End
End
Attribute VB_Name = "frmVersions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iCurrent As Integer, iNew As Integer

Dim pGID As Long, pVID As Long
Dim pHdr As String
Dim pIndex As Integer, iRight As Integer

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property

Public Property Get PassVID() As Long
    PassVID = pVID
End Property
Public Property Let PassVID(ByVal vNewValue As Long)
    pVID = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = pHdr
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    pHdr = vNewValue
End Property

Public Property Get PassIndex() As Integer
    PassIndex = pIndex
End Property
Public Property Let PassIndex(ByVal vNewValue As Integer)
    pIndex = vNewValue
End Property




Private Sub cmdDelete_Click()
    Dim strUpdate As String
    Dim i As Integer
    
    strUpdate = "UPDATE ANNOTATOR.GFX_VERSION SET " & _
                "V_STATUS = 0, " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDCNT = UPDCNT + 1, " & _
                "UPDDTTM = SYSDATE " & _
                "WHERE VERSION_ID = " & CLng(imx1(iNew).Tag)
    Conn.Execute (strUpdate)
    
    For i = 0 To imx1.Count - 1
        imx1(i).Visible = False
        imx1(i).BackColor = vbWindowBackground
        lbl1(i).Caption = ""
    Next i
    
    Call Form_Load
End Sub

Private Sub cmdSave_Click()
    Dim strUpdate As String, sNote As String, sFormat As String, sDesc As String
    Dim iErr As Integer, iStr As Integer
    
    If imx1(iNew).Visible = True Then ''NON-PDF''
        sFormat = Right(imx1(iNew).FileName, 3)
        iStr = InStr(lbl1(iNew).Caption, "[") + 1
        sDesc = Mid(lbl1(iNew).Caption, iStr, Len(lbl1(iNew).Caption) - iStr)
        
        strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                    "GDESC = '" & DeGlitch(Left(sDesc, 50)) & "', " & _
                    "VERSION_ID = " & CLng(imx1(iNew).Tag) & ", " & _
                    "GPATH = '" & "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat & "', " & _
                    "GFORMAT = '" & sFormat & "', " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDCNT = UPDCNT + 1, " & _
                    "UPDDTTM = SYSDATE " & _
                    "WHERE GID = " & pGID
        Conn.Execute (strUpdate)
        
        FileCopy imx1(iNew).FileName, "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat
        
'        frmGraphics.imx4(pIndex).FileName = "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat
'        frmGraphics.imx4(pIndex).Tag = imx1(iNew).Tag
        
        
    
    ElseIf pdf1(iNew).Visible = True Then ''PDF''
        sFormat = Right(pdf1(iNew).GetFileName, 3)
        iStr = InStr(lbl1(iNew).Caption, "[") + 1
        sDesc = Mid(lbl1(iNew).Caption, iStr, Len(lbl1(iNew).Caption) - iStr)
        
        strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                    "GDESC = '" & DeGlitch(Left(sDesc, 50)) & "', " & _
                    "VERSION_ID = " & CLng(pdf1(iNew).Tag) & ", " & _
                    "GPATH = '" & "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat & "', " & _
                    "GFORMAT = '" & sFormat & "', " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDCNT = UPDCNT + 1, " & _
                    "UPDDTTM = SYSDATE " & _
                    "WHERE GID = " & pGID
        Conn.Execute (strUpdate)
        
        On Error Resume Next
        Err = 0
        FileCopy pdf1(iNew).GetFileName, "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat
        If Err > 0 Then MsgBox "Unable to copy file.  The orignal file may be in use.", vbExclamation, "Sorry..."
        
    End If
    
    
    Select Case UCase(sFormat)
        Case "PDF"
            frmGraphics.imx4(pIndex).FileName = CheckForThumbPath(CLng(pdf1(iNew).Tag), "PDF", 1)
            frmGraphics.imx4(pIndex).Tag = pGID '' pdf1(iNew).Tag
        Case Else
            frmGraphics.imx4(pIndex).FileName = "\\DETMSFS01\GPJAnnotator\Graphics\" & CStr(pGID) & "." & sFormat ''' sVPath & imx1(iNew).Tag & _
                "." & Right(imx1(iNew).FileName, 3)
            frmGraphics.imx4(pIndex).Tag = pGID '' imx1(iNew).Tag
    End Select
    
    ''WRITE NEW DESC TO FLX''
    frmGraphics.flxApprove.TextMatrix(pIndex + 1, 3) = sDesc
    
    sNote = "Current Version set to " & lbl1(iNew).Caption & " by " & _
                StrConv(LogName, vbProperCase)
    iErr = InsertGfxComment(pGID, sNote)
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.Caption = "Graphic File:  " & pHdr
    
    Call GetVersions(pGID)
    For i = 0 To imx1.Count - 1
        If CLng(imx1(i).Tag) = pVID Then
            shp1.Top = imx1(i).Top - 60
            shp1.Left = imx1(i).Left - 60
            imx1(i).BackColor = lGeo_Bright ''vbActiveTitleBar
            shp1.Visible = True
            iCurrent = i
            Exit For
        ElseIf CLng(pdf1(i).Tag) = pVID Then
            shp1.Top = pdf1(i).Top - 60
            shp1.Left = pdf1(i).Left - 60
            pdf1(i).matteColor = lGeo_Bright ''vbActiveTitleBar
            shp1.Visible = True
            iCurrent = i
            Exit For
        End If
    Next i
    
    If picInner.Height - picOuter.ScaleHeight > 0 Then
        vsc1.Max = picInner.Height - picOuter.ScaleHeight
        vsc1.LargeChange = picInner.Height - picOuter.ScaleHeight
        vsc1.SmallChange = 2920
        vsc1.Enabled = True
    Else
        vsc1.Enabled = False
    End If
    
    If Not bPerm(60) Then cmdDelete.Visible = False
End Sub

Public Sub GetVersions(tGID As Long)
    Dim strSelect As String, sPath As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    i = -1
    strSelect = "SELECT VERSION_ID, V_FORMAT, V_NUMBER, " & _
                "NVL(VDESC, '-') AS VDESC " & _
                "From ANNOTATOR.GFX_VERSION " & _
                "Where GID = " & tGID & " " & _
                "AND V_STATUS = 1 " & _
                "ORDER BY V_NUMBER DESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        i = i + 1
        If i >= imx1.Count Then
            Load imx1(i)
            imx1(i).Top = 120 + (Int(i / 2) * (2800))
            imx1(i).Left = 120 + ((i Mod 2) * 3320)
            
            Load pdf1(i)
            pdf1(i).Top = 120 + (Int(i / 2) * (2800))
            pdf1(i).Left = 120 + ((i Mod 2) * 3320)
            pdf1(i).Width = imx1(i).Width
            pdf1(i).Height = imx1(i).Height
            pdf1(i).enableMouseEvents = True
        Else
            pdf1(i).Width = imx1(i).Width
            pdf1(i).Height = imx1(i).Height
            pdf1(i).enableMouseEvents = True
        End If
        
        If i >= lbl1.Count Then
            Load lbl1(i)
            lbl1(i).Top = 2580 + (Int(i / 2) * 2800)
            lbl1(i).Left = imx1(i).Left + ((imx1(i).Width - lbl1(i).Width) / 2)
            lbl1(i).ZOrder
        End If
        
        sPath = sVPath & CStr(rst.Fields("VERSION_ID")) & "." & Trim(rst.Fields("V_FORMAT"))
        Select Case Trim(rst.Fields("V_FORMAT"))
            Case "PDF"
                imx1(i).Visible = False
                imx1(i).Tag = 0
                pdf1(i).loadFile (sPath)
                pdf1(i).Tag = rst.Fields("VERSION_ID")
                pdf1(i).Visible = True
            Case Else
                pdf1(i).Visible = False
                pdf1(i).Tag = 0
                imx1(i).Update = False
                imx1(i).FileName = sPath
                imx1(i).Update = True
                imx1(i).Tag = rst.Fields("VERSION_ID")
                imx1(i).Visible = True
                imx1(i).Refresh
        End Select
        
        If Trim(rst.Fields("VDESC")) = "-" Then
            lbl1(i).Caption = "Version " & rst.Fields("V_NUMBER")
        Else
            lbl1(i).Caption = "v" & rst.Fields("V_NUMBER") & " [" & Trim(rst.Fields("VDESC")) & "]"
        End If
        lbl1(i).Visible = True
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    picInner.Height = 120 + (Int(i / 2) * 2800) + 2800
    
    
    
'''            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
'''                ''LOOK FOR BMP HERE''
'''                sFile = sGPath & "pdf_" & rst.Fields("GID") & ".bmp"
'''                If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.BMP''
'''                    ''PDF.BMP NOT FOUND''
'''                    sFile = sGPath & "pdf_" & rst.Fields("GID") & ".jpg"
'''                    If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.JPG''
'''                        ''NO THUMBNAIL AT ALL''
'''                        .PICThumbnail = THUMB_64
'''                        .FileName = sGPath & "pdf.bmp"
'''                    Else
'''                        ''DISPLAY PDF.JPG''
'''                        Select Case FileLen(sFile)
'''                            Case Is < 10000: .PICThumbnail = THUMB_None
'''                            Case Is < 25000: .PICThumbnail = THUMB_4
'''                            Case Is < 50000: .PICThumbnail = THUMB_16
'''                            Case Else: .PICThumbnail = THUMB_64
'''                        End Select
'''                        .FileName = sFile
'''                    End If
'''                Else
'''                    ''DISPLAY PDF.BMP''
'''                    Select Case FileLen(sFile)
'''                        Case Is < 10000: .PICThumbnail = THUMB_None
'''                        Case Is < 25000: .PICThumbnail = THUMB_4
'''                        Case Is < 50000: .PICThumbnail = THUMB_16
'''                        Case Else: .PICThumbnail = THUMB_64
'''                    End Select
'''                    .FileName = sFile
'''                End If
'''            Else
'''                sFile = sGPath & "Thumbs\thb_" & rst.Fields("GID") & ".jpg"
'''                If Dir(sFile, vbNormal) = "" Then ''OPEN FULL FILE''
'''                    Select Case FileLen(Trim(rst.Fields("GPATH")))
'''                        Case Is < 10000: .PICThumbnail = THUMB_None
'''                        Case Is < 25000: .PICThumbnail = THUMB_4
'''                        Case Is < 50000: .PICThumbnail = THUMB_16
'''                        Case Else: .PICThumbnail = THUMB_64
'''                    End Select
'''                    .FileName = Trim(rst.Fields("GPATH"))
'''                Else
'''                    Select Case FileLen(sFile)
'''                        Case Is < 10000: .PICThumbnail = THUMB_None
'''                        Case Is < 25000: .PICThumbnail = THUMB_4
'''                        Case Is < 50000: .PICThumbnail = THUMB_16
'''                        Case Else: .PICThumbnail = THUMB_64
'''                    End Select
'''                    .FileName = sFile
'''                End If
'''
'''            End If
            
End Sub

Private Sub imx1_Click(Index As Integer)
    Dim i As Integer
    
    iNew = Index
    
    shp1.Visible = False
    
    For i = 0 To imx1.Count - 1
        imx1(i).BackColor = vbWindowBackground
        pdf1(i).matteColor = vbWindowBackground
    Next i
    
    shp1.Top = imx1(Index).Top - 60
    shp1.Left = imx1(Index).Left - 60
    imx1(Index).BackColor = lGeo_Bright ''vbActiveTitleBar
    
    shp1.Visible = True
    
    If iNew <> iCurrent Then
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub imx1_DblClick(Index As Integer)
    frmHTMLViewer.PassFile = imx1(Index).FileName
    frmHTMLViewer.PassHDR = Me.Caption & "  [" & lbl1(Index).Caption & "]"
    frmHTMLViewer.PassFrom = Me.Name
    frmHTMLViewer.Show 1, Me
End Sub

Private Sub imx1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        iRight = Index
        Me.PopupMenu mnuRightClick
    End If
End Sub

Private Sub mnuDetails_Click()
    Dim strSelect As String
    
    strSelect = "SELECT ('" & sVPath & "'||GV.VERSION_ID||'.'||TRIM(GV.V_FORMAT))GPATH, " & _
                "(TRIM(GM.GDESC)||' (Version '||GV.V_NUMBER||')')GDESC, " & _
                "('GID-'||GV.GID||' ('||'VID-'||GV.VERSION_ID||')')GID, GV.V_FORMAT AS GFORMAT, " & _
                "GM.GTYPE, GM.GSTATUS, GV.ADDUSER, GV.ADDDTTM, GV.UPDUSER, GV.UPDDTTM " & _
                "FROM ANNOTATOR.GFX_VERSION GV, ANNOTATOR.GFX_MASTER GM " & _
                "WHERE GV.VERSION_ID = " & Trim(imx1(iRight).Tag) & " " & _
                "AND GV.GID = GM.GID"
    
    Call GetGFXData(strSelect, "msgbox")
End Sub

Private Sub mnuExpandedView_Click()
    If imx1(iRight).Visible = True Then
        frmHTMLViewer.PassFile = imx1(iRight).FileName
    Else
        frmHTMLViewer.PassFile = pdf1(iRight).GetFileName
    End If
    frmHTMLViewer.PassHDR = Me.Caption & "  [" & lbl1(iRight).Caption & "]"
    frmHTMLViewer.PassFrom = Me.Name
    frmHTMLViewer.Show 1, Me
End Sub

Private Sub pdf1_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim i As Integer
    
    If Button = vbRightButton Then
        iRight = Index
        Me.PopupMenu mnuRightClick
    
    Else
        iNew = Index
        
        shp1.Visible = False
        
        For i = 0 To pdf1.Count - 1
            pdf1(i).matteColor = vbWindowBackground
            imx1(i).BackColor = vbWindowBackground
        Next i
        
        shp1.Top = pdf1(Index).Top - 60
        shp1.Left = pdf1(Index).Left - 60
        pdf1(Index).matteColor = lGeo_Bright ''vbActiveTitleBar
        
        shp1.Visible = True
        
        If iNew <> iCurrent Then
            cmdSave.Enabled = True
            cmdDelete.Enabled = True
        Else
            cmdSave.Enabled = False
            cmdDelete.Enabled = False
        End If
        
    End If
    
        
End Sub

Private Sub vsc1_Change()
    picInner.Top = CLng(vsc1.Value) * (-1)
End Sub

Private Sub vsc1_Scroll()
    picInner.Top = CLng(vsc1.Value) * (-1)
End Sub
