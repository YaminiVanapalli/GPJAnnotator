VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIGLAssignment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Information Screen"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIGLAssignment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4755
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   5985
      ExtentX         =   10557
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.OptionButton optInfo 
      Height          =   555
      Index           =   1
      Left            =   660
      Picture         =   "frmIGLAssignment.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click to View List of Attending Supervisors"
      Top             =   5220
      Width           =   555
   End
   Begin VB.OptionButton optInfo 
      Height          =   555
      Index           =   0
      Left            =   120
      Picture         =   "frmIGLAssignment.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click to View Show Information"
      Top             =   5220
      Value           =   -1  'True
      Width           =   555
   End
   Begin VB.Frame fraIGL 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   6180
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      Begin VB.CheckBox chkIncludeDNS 
         Alignment       =   1  'Right Justify
         Caption         =   "Include 'Do Not Ships' in View"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtLogisticsNote 
         Height          =   2055
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3975
         Width           =   5505
      End
      Begin MSComctlLib.TreeView tvwIGL 
         Height          =   2835
         Left            =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   375
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   5001
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin VB.Image imgKey 
         Height          =   435
         Index           =   4
         Left            =   1680
         Stretch         =   -1  'True
         ToolTipText     =   "Denotes Item is a ""Dependency Ship"""
         Top             =   3255
         Width           =   435
      End
      Begin VB.Image imgKey 
         Height          =   435
         Index           =   3
         Left            =   600
         Stretch         =   -1  'True
         ToolTipText     =   "Denotes Item is marked as ""Do Not Ship"""
         Top             =   3255
         Width           =   435
      End
      Begin VB.Image imgKey 
         Height          =   435
         Index           =   2
         Left            =   1140
         Stretch         =   -1  'True
         ToolTipText     =   "Denotes Item is to be ""Shipped, but Not Used"""
         Top             =   3255
         Width           =   435
      End
      Begin VB.Image imgKey 
         Height          =   435
         Index           =   1
         Left            =   60
         Stretch         =   -1  'True
         ToolTipText     =   "Denotes Item is ""Shipping"""
         Top             =   3255
         Width           =   435
      End
      Begin VB.Label lblIGLHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGL Assignment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   135
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logistics Note"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   11
         Top             =   3735
         Width           =   1350
      End
      Begin VB.Image imgCube 
         Height          =   495
         Left            =   5100
         Picture         =   "frmIGLAssignment.frx":0EDE
         Stretch         =   -1  'True
         ToolTipText     =   "Click to View Cubing Diagram"
         Top             =   3300
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   780
      Top             =   8340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIGLAssignment.frx":11E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIGLAssignment.frx":1502
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIGLAssignment.frx":181C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIGLAssignment.frx":1B36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxSupers 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   6180
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3519
      _Version        =   393216
      Rows            =   7
      Cols            =   20
      FixedRows       =   2
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483640
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
   End
   Begin RichTextLib.RichTextBox rtbPrint 
      Height          =   2175
      Left            =   2280
      TabIndex        =   13
      Top             =   9120
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmIGLAssignment.frx":1E50
   End
   Begin VB.Label lblPrint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIGLAssignment.frx":1ECB
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   1320
      TabIndex        =   15
      Top             =   5220
      Visible         =   0   'False
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgFPS 
      Height          =   495
      Left            =   5640
      Picture         =   "frmIGLAssignment.frx":1F74
      Stretch         =   -1  'True
      ToolTipText     =   "Click to View Floorplan Status"
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label lblRTB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Info/FPS/Show Regulation Abstract"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblSupers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor Schedule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   5940
      Width           =   1980
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Close Show Information Screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8760
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Click to Return to Annotator"
      Top             =   8280
      Width           =   2790
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   7860
      TabIndex        =   6
      Top             =   7980
      Width           =   4095
   End
End
Attribute VB_Name = "frmIGLAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iDOff As Integer, iDStart As Integer, iDEnd As Integer
Dim dStart As Date
Dim iLHdr As Integer, iLFPS As Integer, iLSRA As Integer, i As Integer
Dim iSHdr As Long, iSFPS As Long, iSSRA As Long
Dim tSHYR As Integer
Dim tBCC As String
Dim tSHCD As Long
Dim tFBCN As String
Dim tSHNM As String
Dim SuperStuff As Boolean, bNoDates As Boolean, bSurfing As Boolean



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

Private Sub chkIncludeDNS_Click()
    Screen.MousePointer = 11
    tvwIGL.Visible = False
    tvwIGL.Nodes.Clear
    Call GenIGLList(tBCC, tSHYR, tSHCD)
    tvwIGL.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub flxSupers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < flxSupers.ColWidth(0) And flxSupers.Row > 1 Then
        MsgBox "Cell Number for " & flxSupers.TextMatrix(flxSupers.Row, 0), vbInformation, "Supervisor Cell..."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sChk As String
    
    On Error Resume Next
    sChk = Dir(strHTMLPath & sIFile, vbNormal)
    If sChk <> "" Then Kill strHTMLPath & sIFile
    sChk = Dir(strHTMLPath & sSFile, vbNormal)
    If sChk <> "" Then Kill strHTMLPath & sSFile
End Sub

Private Sub imgCube_Click()
    frmCubing.Show 1
End Sub

Private Sub imgFPS_Click()
    With frmFPStatus
        .PassBCC = tBCC
        .PassFBCN = tFBCN
        .Show 1
    End With
End Sub

'''Private Sub imgPrintInfo_Click()
'''    Dim sAdd As String
'''    Dim iAdd As Integer, iFrom As Integer
'''
'''    If optInfo(0).Value = True Then iFrom = 0 Else iFrom = 1
'''
'''    With rtbPrint
'''        .Text = ""
'''        If iFrom = 0 Then
'''            sAdd = vbNewLine & vbNewLine & "GPJ Annotator Show Information" & vbNewLine & vbNewLine
'''            iAdd = Len(sAdd)
'''            .Text = sAdd & rtbInfo.Text
'''        Else
'''            sAdd = vbNewLine & vbNewLine & "GPJ Annotator Attending Supervisors List" & vbNewLine & vbNewLine
'''            iAdd = Len(sAdd)
'''            .Text = sAdd & rtbSupers.Text
'''        End If
'''        .SelStart = 0
'''        .SelLength = Len(.Text)
'''        .SelBold = False
'''        .SelStart = 0
'''        .SelLength = iAdd
'''        .SelFontSize = 11
'''        .SelBold = True
'''        .SelAlignment = 2
'''
'''        .SelStart = iAdd
'''        .SelLength = Len(.Text) - iAdd
'''        .SelFontSize = 10
'''
'''        .SelStart = iSHdr + iAdd
'''        .SelLength = iLHdr
'''        .SelBold = True
'''        .SelFontSize = 10
'''
'''        If iFrom = 0 Then
'''            .SelStart = iSFPS + iAdd
'''            .SelLength = iLFPS
'''            .SelBold = True
'''            .SelFontSize = 10
'''
'''            .SelStart = iSSRA + iAdd
'''            .SelLength = iLSRA
'''            .SelBold = True
'''            .SelFontSize = 10
'''
'''            .SelStart = 0
'''            .SelLength = Len(.Text)
'''            .SelIndent = 1250
'''            .SelRightIndent = 1250
'''            .SelTabCount = 3
'''            .SelTabs(0) = 0
'''            .SelTabs(1) = 2000
'''            .SelTabs(2) = 3500
'''        Else
'''            .SelStart = 0
'''            .SelLength = Len(.Text)
'''            .SelIndent = 1250
'''            .SelRightIndent = 1250
'''            .SelTabCount = 3
'''            .SelTabs(0) = 0
'''            .SelTabs(1) = 2000
'''            .SelTabs(2) = 6000
'''        End If
'''
'''        .SelLength = 0
'''
'''        dlgPrint.Flags = cdlPDReturnDC + cdlPDNoPageNums
'''        If .SelLength = 0 Then
'''           dlgPrint.Flags = dlgPrint.Flags + cdlPDAllPages
'''        Else
'''           dlgPrint.Flags = dlgPrint.Flags + cdlPDSelection
'''        End If
'''
'''        On Error Resume Next
'''        dlgPrint.CancelError = True
'''        dlgPrint.ShowPrinter
'''        If Err = 0 Then .SelPrint dlgPrint.hdC
'''
'''    End With
'''End Sub

Private Sub lblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblClose.FontBold = False
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sFile As String
    
    Screen.MousePointer = 11

    Me.Caption = frmAnnotator.lblWelcome.Caption & " Information Screen"
    
'''    strHTMLPath = "\\DETMSFS01\GPJAnnotator\Support\HTML\"
    
    If bPerm(9) Then
        Call GenIGLList(tBCC, tSHYR, tSHCD)
        Call GetLogistics(tBCC, tSHYR, tSHCD)
        fraIGL.Visible = True
        imgCube.Visible = False
    Else
        fraIGL.Visible = False
    End If
    If bPerm(8) Then
        sIFile = "SI-" & CStr((CLng(Format(Now, "h")) * 60 * 60) + (CLng(Format(Now, "n")) * 60) + _
                    (CLng(Format(Now, "s")))) & ".htm"
        sFile = PopShowInfo(CLng(tBCC), tSHYR, tSHCD, tSHNM, tFBCN)
        web1.Navigate sFile
        optInfo(0).Visible = True
        lblRTB.Visible = True
    Else
        optInfo(0).Visible = False
        lblRTB.Visible = False
    End If
    If bPerm(11) Then
        Call SupersSetup(tBCC, tSHYR, tSHCD)
        If SuperStuff Then
            Call GetSupers(tBCC, tSHYR, tSHCD)
            lblSupers.Visible = True
            flxSupers.Visible = True
        End If
        If bPerm(12) Then
            sSFile = "AS-" & CStr((CLng(Format(Now, "h")) * 60 * 60) + (CLng(Format(Now, "n")) * 60) + _
                        (CLng(Format(Now, "s")))) & ".htm"
            sFile = GetAttSupers(tSHYR, tSHCD, tSHNM)
            optInfo(1).Visible = True
        Else
            optInfo(1).Visible = False
        End If
    Else
        lblSupers.Visible = False
        flxSupers.Visible = False
        Me.Height = 6540
        optInfo(1).Visible = False
    End If
    If bPerm(22) Then imgFPS.Visible = True Else imgFPS.Visible = False
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(True)
    If Not bENABLE_PRINTERS Then lblPrint.Visible = True
    '\\\\\ -------------------------------------------------------- /////
    
'''    If bPerm(8) Or bPerm(12) Then
'''        imgPrintInfo.Visible = True
'''        If bPerm(8) Then
'''            imgPrintInfo.ToolTipText = "Click to Print Show Information"
'''        ElseIf bPerm(12) Then
'''            imgPrintInfo.ToolTipText = "Click to Print Attending Supervisor Info"
'''        End If
'''    Else
'''        imgPrintInfo.Visible = False
'''    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub imgKey_Click(Index As Integer)
    Dim sMess As String, sHDR As String
    Select Case Index
        Case 1
            sMess = "The Element or Part is scheduled to SHIP."
            sHDR = "SHIP Description"
        Case 2
            sMess = "The Element or Part is scheduled to SHIP, but it will not be used on the Show Floor."
            sHDR = "SHIP, BUT NOT USED Description"
        Case 3
            sMess = "The Element or Part is NOT scheduled to SHIP."
            sHDR = "DO NOT SHIP Description"
        Case 4
            sMess = "Dependency Ship implies some, but not all, of an Elements Parts are Shipping," & _
                        vbCr & vbCr & _
                        "BUT...  It is possible to assign all of an Elements Parts as SHIP, and the Element" & _
                        vbCr & "will retain a DEPENDENCY SHIP status."
            sHDR = "DEPENDENCY SHIP Description"
    End Select
    MsgBox sMess, vbInformation, sHDR
End Sub


Public Sub GenIGLList(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long)
    Dim rst As ADODB.Recordset
    Dim strSelect As String, sNodePar As String, sNodeKey As String, sNodeDesc As String, _
                sPlusMins As String, sClientPref As String
    Dim nodX As Node
    Dim iShip As Integer, i As Integer
    Dim lKitUID As Long, lPrimeKID As Long
    
    tvwIGL.ImageList = ImageList1
    
    For i = 1 To 4
        imgKey(i).Picture = ImageList1.ListImages(i).Picture
    Next i
   
    strSelect = "SELECT KITUSEID FROM " & AQUAKitU & " " & _
                "WHERE SHYR = " & tmpSHYR & " " & _
                "AND AN8_SHCD = " & tmpSHCD & " " & _
                "AND AN8_CUNO = " & CLng(tmpBCC)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        lKitUID = rst.Fields("KITUSEID")
        rst.Close
        Set rst = Nothing
        tvwIGL.Nodes.Clear
        strSelect = "SELECT KITID, KITFNAME, PLUSMINS, " & _
                    "SHSTATUS, UPDUSER, UPDDTTM " & _
                    "FROM " & AQUAKitU & " " & _
                    "WHERE KITUSEID = " & lKitUID & " " & _
                    "AND SHYR = " & tmpSHYR & " " & _
                    "AND AN8_SHCD = " & tmpSHCD & " " & _
                    "AND AN8_CUNO = " & CLng(tmpBCC)
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            Do While Not rst.EOF
                sNodeKey = "k" & lKitUID
                lPrimeKID = rst.Fields("KITID")
                Select Case rst.Fields("PLUSMINS")
                    Case 1: sPlusMins = " +"
                    Case 2: sPlusMins = " -"
                    Case 3: sPlusMins = " /"
                End Select
                sNodeDesc = Trim(rst.Fields("KITFNAME")) & sPlusMins & "     [Last updated by " & _
                            Trim(rst.Fields("UPDUSER")) & " on " & Format(rst.Fields("UPDDTTM"), "dd-mmm-yyyy") & "]"
                iShip = rst.Fields("SHSTATUS")
                Set nodX = tvwIGL.Nodes.Add(, , sNodeKey, sNodeDesc, iShip)
                rst.MoveNext
            Loop
            rst.Close
            Set rst = Nothing

            If chkIncludeDNS.Value = 0 Then
                strSelect = "SELECT EU.ELTUSEID, EU.ELTCODE, " & _
                            "EU.ELSUFFIX, EU.ELTFNAME, EU.ELTDESC, EU.SHSTATUS, " & _
                            "K.AN8_CUNO, K.KITREF, K.KITFNAME " & _
                            "FROM " & AQUAEltU & " EU, " & IGLKit & " K " & _
                            "WHERE EU.KITUSEID = " & lKitUID & " " & _
                            "AND EU.SHYR = " & tmpSHYR & " " & _
                            "AND EU.AN8_SHCD = " & tmpSHCD & " " & _
                            "AND EU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                            "AND EU.SHSTATUS <> 3 " & _
                            "AND EU.KITID = K.KITID " & _
                            "ORDER BY K.AN8_CUNO, K.KITREF, EU.ELTCODE, EU.ELSUFFIX"
            Else
                strSelect = "SELECT EU.ELTUSEID, EU.ELTCODE, " & _
                            "EU.ELSUFFIX, EU.ELTFNAME, EU.ELTDESC, EU.SHSTATUS, " & _
                            "K.AN8_CUNO, K.KITREF, K.KITFNAME " & _
                            "FROM " & AQUAEltU & " EU, " & IGLKit & " K " & _
                            "WHERE EU.KITUSEID = " & lKitUID & " " & _
                            "AND EU.SHYR = " & tmpSHYR & " " & _
                            "AND EU.AN8_SHCD = " & tmpSHCD & " " & _
                            "AND EU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                            "AND EU.KITID = K.KITID " & _
                            "ORDER BY K.AN8_CUNO, K.KITREF, EU.ELTCODE, EU.ELSUFFIX"
            End If

            Set rst = Conn.Execute(strSelect)
            'msgbox "Got Element Recordset"
            If Not rst.EOF Then
                Do While Not rst.EOF
                    sNodePar = "k" & lKitUID
                    sNodeKey = "e" & rst.Fields("ELTUSEID")
                    If rst.Fields("AN8_CUNO") <> CLng(tmpBCC) Then
                        sClientPref = "{" & rst.Fields("AN8_CUNO") & "} "
                    Else: sClientPref = ""
                    End If
                    sNodeDesc = sClientPref & Trim(rst.Fields("KITFNAME")) & " - " & Trim(rst.Fields("ELTFNAME")) & " ---> [" & _
                                UCase(Trim(rst.Fields("ELTDESC"))) & "]"
                    iShip = rst.Fields("SHSTATUS")
                    Set nodX = tvwIGL.Nodes.Add(sNodePar, tvwChild, sNodeKey, sNodeDesc, iShip)
                    rst.MoveNext
                Loop
                rst.Close
                Set rst = Nothing

                If chkIncludeDNS.Value = 0 Then
                    strSelect = "SELECT TU.ELTUSEID, TU.PRTUSEID, TU.FABLOC, TU.YRBUILT, " & _
                                "TU.PNUMBER, TU.PARTDESC, TU.SHSTATUS " & _
                                "FROM " & AQUAPartU & " TU, " & AQUAEltU & " EU " & _
                                "WHERE TU.KITUSEID = " & lKitUID & " " & _
                                "AND TU.SHYR = " & tmpSHYR & " " & _
                                "AND TU.AN8_SHCD = " & tmpSHCD & " " & _
                                "AND TU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                                "AND TU.SHYR = EU.SHYR " & _
                                "AND TU.AN8_SHCD = EU.AN8_SHCD " & _
                                "AND TU.AN8_CUNO = EU.AN8_CUNO " & _
                                "AND TU.KITUSEID = EU.KITUSEID " & _
                                "AND TU.ELTUSEID = EU.ELTUSEID " & _
                                "AND EU.SHSTATUS <> 3"
                Else
                    strSelect = "SELECT TU.ELTUSEID, TU.PRTUSEID, TU.FABLOC, TU.YRBUILT, " & _
                            "TU.PNUMBER, TU.PARTDESC, TU.SHSTATUS " & _
                            "FROM " & AQUAPartU & " TU, " & AQUAEltU & " EU " & _
                            "WHERE TU.KITUSEID = " & lKitUID & " " & _
                            "AND TU.SHYR = " & tmpSHYR & " " & _
                            "AND TU.AN8_SHCD = " & tmpSHCD & " " & _
                            "AND TU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                            "AND TU.SHYR = EU.SHYR " & _
                            "AND TU.AN8_SHCD = EU.AN8_SHCD " & _
                            "AND TU.AN8_CUNO = EU.AN8_CUNO " & _
                            "AND TU.KITUSEID = EU.KITUSEID " & _
                            "AND TU.ELTUSEID = EU.ELTUSEID"
                End If

                Set rst = Conn.Execute(strSelect)
                i = 0
                Do While Not rst.EOF
                    i = i + 1
                    sNodePar = "e" & rst.Fields("ELTUSEID")
                    sNodeKey = "t" & rst.Fields("PRTUSEID")
                    sNodeDesc = CStr(rst.Fields("FABLOC")) & Format(rst.Fields("YRBUILT"), "YY") & "-" & _
                                CStr(rst.Fields("PNUMBER")) & " ---> [" & UCase(Trim(rst.Fields("PARTDESC"))) & "]"
                    iShip = rst.Fields("SHSTATUS")
                    On Error Resume Next
                    Set nodX = tvwIGL.Nodes.Add(sNodePar, tvwChild, sNodeKey, sNodeDesc, iShip)
                    If Err Then Err.Clear
                    On Error GoTo ErrorTrap
                    rst.MoveNext
                Loop
                rst.Close
                Set rst = Nothing
                On Error GoTo 0
            Else
                rst.Close
                Set rst = Nothing
                On Error GoTo 0
            End If
        End If
    Else
        rst.Close
        Set rst = Nothing
        On Error GoTo 0
        MsgBox "No entry is found in the IGL KitUse Table.", vbExclamation, "Uh, oh..."
    End If
Exit Sub
ErrorTrap:
    MsgBox "Error Encountered:  " & Err.Description & vbCr & vbCr & _
                "Property Assignment for this show must be viewed in IGL.", _
                vbExclamation, "Error while accessing IGL Kit Assignment..."
    Err.Clear
    Screen.MousePointer = 0
    Unload Me

End Sub

Public Sub GetLogistics(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long)
    Dim strSelect As String, sSizeUnit As String, sWtUnit As String
    Dim rst As ADODB.Recordset
    Dim rVolMult As Double, rVol As Double, rWT As Double, TruckVol As Double
    Dim WZF As Boolean, SZF As Boolean
    Dim sVNote As String, sVDisc As String, sWNote As String, sWDisc As String
    Dim iParts As Integer
    
    rVol = 0: rWT = 0: WZF = False: SZF = False: iParts = 0
    strSelect = "SELECT TU.WEIGHT, TU.LENGTH, TU.WIDTH, TU.HEIGHT, " & _
                "TU.SIZEUNIT, TU.WTUNIT " & _
                "FROM " & AQUAKitU & " KU, " & AQUAPartU & " TU " & _
                "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                "AND KU.SHYR = " & tmpSHYR & " " & _
                "AND KU.AN8_CUNO = TU.AN8_CUNO " & _
                "AND KU.AN8_SHCD = TU.AN8_SHCD " & _
                "AND KU.SHYR = TU.SHYR " & _
                "AND KU.KITUSEID = TU.KITUSEID " & _
                "AND TU.SHSTATUS < 3" '/// ALL SHIP & SBNU \\\
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Select Case rst.Fields("SIZEUNIT")
            Case 1 '/// INCHES \\\
                rVolMult = CDbl(12) * 12 * 12
                sSizeUnit = " cuft"
                TruckVol = 3000
            Case 5 '/// CM \\\
                rVolMult = CDbl(100) * 100 * 100
                sSizeUnit = " cuM"
                TruckVol = 49.161192
        End Select
        Select Case rst.Fields("WTUNIT")
            Case 1 '/// LBS \\\
                sWtUnit = " lbs"
            Case 2 '/// KG \\\
                sWtUnit = " kg"
        End Select
        Do While Not rst.EOF
            iParts = iParts + 1
            rVol = rVol + (rst.Fields("LENGTH") * rst.Fields("WIDTH") * rst.Fields("HEIGHT") / rVolMult)
            rWT = rWT + rst.Fields("WEIGHT")
            If SZF = False And (rst.Fields("LENGTH") = 0 Or _
                        rst.Fields("WIDTH") = 0 Or _
                        rst.Fields("HEIGHT") = 0) Then SZF = True
            If WZF = False And rst.Fields("WEIGHT") = 0 Then WZF = True
            rst.MoveNext
        Loop
        
        '///// TIME TO ASSEMBLE LOGISTICS NOTE \\\\\
        If SZF Then
            sVDisc = "  (NOTE: Undimensioned Parts were found.  Actual volume may be different.)"
        Else
            sVDisc = ""
        End If
        If WZF Then
            sWDisc = "  (NOTE: Unweighed Parts were found.  Actual total weight may be different.)"
        Else
            sWDisc = ""
        End If
        sVNote = "Total volume of shipping IGL-assigned properties is " & _
                    Format(rVol, "#,##0") & sSizeUnit & ", requiring an estimated ( " & _
                    Format(rVol / TruckVol, "0.0") & " ) trucks to transport.  " & _
                    "Total number of parts: ( " & iParts & " )." & sVDisc
        sWNote = "Total shipping weight is " & Format(rWT, "#,##0") & sWtUnit & "." & sWDisc
        txtLogisticsNote.Text = sWNote & vbNewLine & vbNewLine & sVNote

    Else
        txtLogisticsNote.Text = ""
    End If
    rst.Close
    Set rst = Nothing
End Sub

Public Sub GetSupers(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long)
    Dim rst As ADODB.Recordset
    Dim strSelect As String, strPhone As String
    Dim lSuper As Long
    Dim iCol As Integer
    iCol = 1
    lSuper = 0
    flxSupers.Row = 1
    
'''    strSelect = "SELECT S.AN8_EMNO, S.TASKTYPE, X.WPPHTP, " & _
'''                    "X.WPAR1, X.WPPH1, E.YAALPH " & _
'''                    "FROM " & IGLEmpTask & " S, " & F0115 & " X, " & _
'''                    "" & F060116 & " E, " & IGLEmpDay & " ED " & _
'''                    "WHERE S.AN8_SHCD = " & tmpSHCD & " " & _
'''                    "AND S.SHYR = " & tmpSHYR & " " & _
'''                    "AND S.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''                    "AND S.AN8_EMNO = E.YAAN8 " & _
'''                    "AND S.AN8_EMNO = X.WPAN8 " & _
'''                    "AND X.WPPHTP IN ('BPR', 'CEL') " & _
'''                    "ORDER BY S.TASKTYPE" 'CHANGE - WPPHTP & OUTER JOIN IF NO PHONES
    
    strSelect = "SELECT ED.AN8_EMNO, E.YAALPH, ED.EADATE, R.VALUE, EP.CEL, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'DD-MON-YYYY')BEG_DATE " & _
                "FROM " & IGLEmpDay & " ED, " & F060116_View & " E, " & IGLRef & " R, " & _
                "(SELECT ABC.WPAN8, DECODE (TRIM (WPAR1), NULL, TRIM (WPPH1), " & _
                "TRIM (WPAR1) || ' ' || TRIM (WPPH1)) CEL " & _
                "FROM " & F0115 & " ABC, " & F0101 & " AB " & _
                "WHERE AB.ABAN8 = ABC.WPAN8 " & _
                "AND UPPER (TRIM (ABC.WPPHTP)) = 'CEL' " & _
                "AND TRIM (ABC.WPPH1) IS NOT NULL) EP, " & F5601 & " SM " & _
                "WHERE ED.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND ED.AN8_SHCD = " & tmpSHCD & " " & _
                "AND ED.SHYR = " & tmpSHYR & " " & _
                "AND ED.AN8_EMNO = E.YAAN8 " & _
                "AND ED.TASKTYPE = R.REF_ID " & _
                "AND R.TYPE_CD = 3 " & _
                "AND ED.AN8_EMNO = EP.WPAN8 (+) " & _
                "AND ED.AN8_SHCD = SM.SHY56SHCD " & _
                "AND ED.SHYR = SM.SHY56SHYR " & _
                "ORDER BY E.YAALPH"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If lSuper <> rst.Fields("AN8_EMNO") Then
            If flxSupers.Row + 1 = flxSupers.Rows Then
                MsgBox "The Supervisor View in this interface is restricted to a maximum of five Supervisors.", _
                            vbInformation, "Sorry..."
                GoTo TooManySupers
            End If
            flxSupers.Row = flxSupers.Row + 1
            If Not IsNull(rst.Fields("CEL")) Then strPhone = "  [" & Trim(rst.Fields("CEL")) & "]" Else strPhone = ""
            flxSupers.TextMatrix(flxSupers.Row, 0) = UCase(Trim(rst.Fields("YAALPH"))) & strPhone
            lSuper = rst.Fields("AN8_EMNO")
        End If
        If CInt(Format(rst.Fields("EADATE"), "Y")) >= iDStart Then
            flxSupers.Col = CInt(Format(rst.Fields("EADATE"), "Y")) - iDOff
        Else
            flxSupers.Col = CInt(Format(rst.Fields("EADATE"), "Y")) - iDOff + _
                        CInt(Format(DateValue("12/31/" & Format(DateValue(rst.Fields("BEG_DATE")), "YYYY")), "Y"))
        End If
        If flxSupers.Col > iCol Then iCol = flxSupers.Col
        flxSupers.CellBackColor = vbRed
        flxSupers.Text = UCase(Trim(rst.Fields("VALUE")))
        rst.MoveNext
    Loop
TooManySupers:
    rst.Close
    Set rst = Nothing
    flxSupers.Cols = iCol + 2
End Sub

'''''Public Sub SupersSetup(tmpSHYR As Integer, tmpSHCD As Long)
'''''    Dim iDays As Integer, i As Integer
'''''    Dim rst As ADODB.Recordset
'''''    Dim strSelect As String
'''''
'''''    bNoDates = False
'''''    On Error GoTo DataProblem
''''''''    strSelect = "SELECT EADATE FROM " & IGLEmpDay & " " & _
''''''''                "WHERE AN8_SHCD = " & tmpSHCD & " " & _
''''''''                "AND SHYR = " & tmpSHYR & " " & _
''''''''                "ORDER BY EADATE"
'''''    strSelect = "SELECT IGL_JDEDATE_TOCHAR(SHY56SBEDT, 'DD-MON-YYYY')BEG_DATE, SHY56SBEDT, " & _
'''''                "IGL_JDEDATE_TOCHAR(SHY56TEDDT, 'DD-MON-YYYY')END_DATE, SHY56TEDDT " & _
'''''                "FROM " & F5601 & " " & _
'''''                "WHERE SHY56SHCD = " & tmpSHCD & " " & _
'''''                "AND SHY56SHYR = " & tmpSHYR
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''        If rst.FIELDS("SHY56SBEDT") <> 0 And rst.FIELDS("SHY56TEDDT") <> 0 Then
'''''            dStart = DateAdd("d", -2, DateValue(rst.FIELDS("BEG_DATE")))
'''''            dStart = dStart - CInt(format(dStart, "w")) + 1
'''''            iDStart = CInt(format(dStart, "y"))
'''''            iDEnd = CInt(format(DateValue(rst.FIELDS("END_DATE")), "y")) + _
'''''                        (14 - CInt(format(DateValue(rst.FIELDS("END_DATE")), "w")))
'''''            If iDEnd < iDStart Then
'''''                iDEnd = iDEnd + CInt(format(DateValue("12/31/" & _
'''''                            format(DateValue(rst.FIELDS("BEG_DATE")), "YYYY")), "Y"))
'''''            End If
'''''            iDOff = iDStart - 1
'''''
'''''            With flxSupers
'''''                .Cols = iDEnd - iDStart + 1
'''''                .Col = 0: .Row = 1: .CellAlignment = 4: .Text = "Supervisor [Cell Number]"
'''''                For i = 0 To .Cols - 1
'''''                    Select Case i
'''''                        Case 0
'''''                            .ColWidth(i) = 3500
'''''                        Case Else
'''''                            .ColWidth(i) = 500
'''''                            .TextMatrix(0, i) = format(DateAdd("d", i - 1, dStart), "ddd")
'''''                            .TextMatrix(1, i) = format(DateAdd("d", i - 1, dStart), "m/d")
'''''                            .ColAlignment(i) = 4
'''''                    End Select
'''''                Next i
'''''            End With
'''''            SuperStuff = True
'''''        Else
'''''            GoTo DataProblem
'''''        End If
'''''    End If
'''''    rst.Close
'''''    Set rst = Nothing
'''''Exit Sub
'''''DataProblem:
'''''    rst.Close
'''''    Set rst = Nothing
'''''    Me.Height = 6540
'''''    flxSupers.Visible = False
'''''    lblSupers.Visible = False
'''''    SuperStuff = False
'''''End Sub

Public Sub SupersSetup(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long)
    Dim iDays As Integer, i As Integer
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    
    bNoDates = False
    On Error GoTo DataProblem
    strSelect = "SELECT EADATE FROM " & IGLEmpDay & " " & _
                "WHERE AN8_SHCD = " & tmpSHCD & " " & _
                "AND SHYR = " & tmpSHYR & " " & _
                "AND AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "ORDER BY EADATE"
'''    strSelect = "SELECT IGL_JDEDATE_TOCHAR(SHY56SBEDT, 'DD-MON-YYYY')BEG_DATE, SHY56SBEDT, " & _
'''                "IGL_JDEDATE_TOCHAR(SHY56TEDDT, 'DD-MON-YYYY')END_DATE, SHY56TEDDT " & _
'''                "FROM " & F5601 & " " & _
'''                "WHERE SHY56SHCD = " & tmpSHCD & " " & _
'''                "AND SHY56SHYR = " & tmpSHYR
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        dStart = DateAdd("d", -2, DateValue(rst.Fields("EADATE")))
        dStart = dStart - CInt(Format(dStart, "w")) + 1
        iDStart = CInt(Format(dStart, "y"))
        Do While Not rst.EOF
            iDEnd = CInt(Format(DateValue(rst.Fields("EADATE")), "y")) + _
                        (14 - CInt(Format(DateValue(rst.Fields("EADATE")), "w")))
            rst.MoveNext
        Loop
        iDOff = iDStart - 1
        
'''        If rst.FIELDS("SHY56SBEDT") <> 0 And rst.FIELDS("SHY56TEDDT") <> 0 Then
'''            dStart = DateAdd("d", -2, DateValue(rst.FIELDS("BEG_DATE")))
'''            dStart = dStart - CInt(format(dStart, "w")) + 1
'''            iDStart = CInt(format(dStart, "y"))
'''            iDEnd = CInt(format(DateValue(rst.FIELDS("END_DATE")), "y")) + _
'''                        (14 - CInt(format(DateValue(rst.FIELDS("END_DATE")), "w")))
'''            If iDEnd < iDStart Then
'''                iDEnd = iDEnd + CInt(format(DateValue("12/31/" & _
'''                            format(DateValue(rst.FIELDS("BEG_DATE")), "YYYY")), "Y"))
'''            End If
'''            iDOff = iDStart - 1
            
            With flxSupers
                .Cols = iDEnd - iDStart + 1
                .Col = 0: .Row = 1: .CellAlignment = 4: .Text = "Supervisor [Cell Number]"
                For i = 0 To .Cols - 1
                    Select Case i
                        Case 0
                            .ColWidth(i) = 3500
                        Case Else
                            .ColWidth(i) = 500
                            .TextMatrix(0, i) = Format(DateAdd("d", i - 1, dStart), "ddd")
                            .TextMatrix(1, i) = Format(DateAdd("d", i - 1, dStart), "m/d")
                            .ColAlignment(i) = 4
                    End Select
                Next i
            End With
            SuperStuff = True
        
'''        End If
    Else
        GoTo DataProblem
    End If
    rst.Close
    Set rst = Nothing
Exit Sub
DataProblem:
    rst.Close
    Set rst = Nothing
    Me.Height = 6540
    flxSupers.Visible = False
    lblSupers.Visible = False
    SuperStuff = False
End Sub

'''''Public Sub PopShowInfo(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long, tmpSHNM As String, tmpFBCN As String)
'''''    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
'''''    Dim strSelect As String, stext As String, sAdd As String, sPad As String, sDate1 As String, sDate2 As String
'''''
'''''    rtbInfo.Visible = False
'''''    rtbInfo.Text = ""
'''''
'''''    '///// FIRST, GET SHOW INFORMATION \\\\\
''''''''''    strSelect = "SELECT SM.SHY56TENDT, IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MM/DD/YYYY') AS SBED, SM.SHY56SBETT, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MM/DD/YYYY') AS SEND, SM.SHY56SENTT, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MM/DD/YYYY') AS TBED, SM.SHY56TBETT, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MM/DD/YYYY') AS TEDD, SM.SHY56TENTT, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MM/DD/YYYY') AS VBED, SM.SHY56VBETT, " & _
''''''''''                "IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MM/DD/YYYY') AS VEND, SM.SHY56VENTT, " & _
''''''''''                "SM.SHY56FCCDT, S1.NAME AS FCNM, " & _
''''''''''                "S1.ADDR1 AS FCA1, S1.ADDR2 AS FCA2, S1.ADDR3 AS FCA3, S1.ADDR4 AS FCA4, " & _
''''''''''                "S1.CITY AS FCA5, S1.STATE AS FCA6, S1.ZIP AS FCA7, S1.PHONE AS FCPH, " & _
''''''''''                "SM.SHY56SMGRT, S2.NAME AS SMNM, " & _
''''''''''                "S2.ADDR1 AS SMA1, S2.ADDR2 AS SMA2, S2.ADDR3 AS SMA3, S2.ADDR4 AS SMA4, " & _
''''''''''                "S2.CITY AS SMA5, S2.STATE AS SMA6, S2.ZIP AS SMA7, S2.PHONE AS SMPH, " & _
''''''''''                "S2.FAX, SM.SHY56DRAIT, SM.SHY56CARIT, SM.SHY56VACIT " & _
''''''''''                "FROM " & F5601 & " SM, " & ANOSuppAll & " S1, " & ANOSuppAll & " S2 " & _
''''''''''                "WHERE SM.SHY56FCCDT = S1.SUPPLIERID (+) " & _
''''''''''                "AND SM.SHY56SMGRT = S2.SUPPLIERID (+) " & _
''''''''''                "AND SM.SHY56SHCD = " & tmpSHCD & " " & _
''''''''''                "AND SM.SHY56SHYR = " & tmpSHYR
'''''    strSelect = "SELECT SM.SHY56TENDT, " & _
'''''                "SM.SHY56BEGDT, IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, " & _
'''''                "SM.SHY56ENDDT, IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, " & _
'''''                "SM.SHY56SBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MM/DD/YYYY') AS SBED, SM.SHY56SBETT, " & _
'''''                "SM.SHY56SENDT, IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MM/DD/YYYY') AS SEND, SM.SHY56SENTT, " & _
'''''                "SM.SHY56TBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MM/DD/YYYY') AS TBED, SM.SHY56TBETT, " & _
'''''                "SM.SHY56TEDDT, IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MM/DD/YYYY') AS TEDD, SM.SHY56TENTT, " & _
'''''                "SM.SHY56VBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MM/DD/YYYY') AS VBED, SM.SHY56VBETT, " & _
'''''                "SM.SHY56VENDT, IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MM/DD/YYYY') AS VEND, SM.SHY56VENTT, " & _
'''''                "SM.SHY56FCCDT, SM.SHY56SMGRT, SM.SHY56DRAIT, SM.SHY56CARIT, SM.SHY56VACIT " & _
'''''                "FROM " & F5601 & " SM " & _
'''''                "WHERE SM.SHY56SHCD = " & tmpSHCD & " " & _
'''''                "AND SM.SHY56SHYR = " & tmpSHYR
'''''
''''''''''    strSelect = "SELECT SM.SHY56TENDT, " & _
''''''''''                "SM.SHY56BEGDT, IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, SM.SHY56BEGTT, " & _
''''''''''                "SM.SHY56ENDDT, IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, SM.SHY56ENDTT, " & _
''''''''''                "CS.CSY56FARDT, IGL_JDEDATE_TOCHAR(CS.CSY56FARDT, 'MM/DD/YYYY') AS FRAD, CS.CSY56FARTT, " & _
''''''''''                "SM.SHY56SBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MM/DD/YYYY') AS SBED, SM.SHY56SBETT, " & _
''''''''''                "SM.SHY56SENDT, IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MM/DD/YYYY') AS SEND, SM.SHY56SENTT, " & _
''''''''''                "CS.CSY56VMVDT, IGL_JDEDATE_TOCHAR(CS.CSY56VMVDT, 'MM/DD/YYYY') AS VMVD, CS.CSY56VMVTT, " & _
''''''''''                "SM.SHY56PBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'MM/DD/YYYY') AS PBED, SM.SHY56PBETT, " & _
''''''''''                "SM.SHY56PENDT, IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'MM/DD/YYYY') AS PEND, SM.SHY56PENTT, " & _
''''''''''                "SM.SHY56VBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MM/DD/YYYY') AS VBED, SM.SHY56VBETT, " & _
''''''''''                "SM.SHY56VENDT, IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MM/DD/YYYY') AS VEND, SM.SHY56VENTT, " & _
''''''''''                "SM.SHY56TBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MM/DD/YYYY') AS TBED, SM.SHY56TBETT, " & _
''''''''''                "SM.SHY56TEDDT, IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MM/DD/YYYY') AS TEDD, SM.SHY56TENTT, " & _
''''''''''                "SM.SHY56FCCDT , SM.SHY56SMGRT, SM.SHY56DRAIT, SM.SHY56CARIT, SM.SHY56VACIT " & _
''''''''''                "FROM " & F5601 & " SM, " & F5611 & " CS " & _
''''''''''                "WHERE SM.SHY56SHCD = 12861 " & _
''''''''''                "AND SM.SHY56SHYR = 2001 " & _
''''''''''                "AND SM.SHY56SHCD =CS.CSY56SHCD " & _
''''''''''                "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
''''''''''                "AND CS.CSY56CUNO = 1190"
'''''
'''''
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''        stext = ""
'''''        sAdd = tmpSHYR & " " & tmpSHNM
'''''        iLHdr = Len(sAdd)
'''''        iSHdr = 0
'''''        stext = stext & sAdd & vbNewLine '''''& vbNewLine
'''''        If rst.FIELDS("SHY56TENDT") = 1 Then
'''''            sAdd = "NOTE:  All Dates are TENTATIVE" & vbNewLine
'''''            stext = stext & sAdd
'''''        End If
'''''        If rst.FIELDS("SHY56BEGDT") <> 0 And rst.FIELDS("SHY56ENDDT") <> 0 Then 'CHANGE
'''''            sDate1 = format(DateValue(rst.FIELDS("BEGD")), "dddd mmm d")
'''''            sDate2 = format(DateValue(rst.FIELDS("ENDD")), "dddd mmm d")
'''''            sAdd = "Show Dates:" & vbTab & sDate1 & " - " & sDate2 & vbNewLine & vbNewLine
'''''        Else
'''''            sDate1 = "": sDate2 = ""
'''''            sAdd = "Show Dates:" & vbNewLine & vbNewLine
'''''        End If
'''''        stext = stext & sAdd
'''''
'''''        If rst.FIELDS("SHY56SBEDT") <> 0 Then
'''''            sAdd = "Setup Begin Date:" & vbTab & format(DateValue(rst.FIELDS("SBED")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56SBETT")) & vbNewLine
'''''        Else
'''''            sAdd = "Setup Begin Date:" & vbNewLine
'''''        End If
'''''        stext = stext & sAdd
'''''
'''''        If rst.FIELDS("SHY56SENDT") <> 0 Then
'''''            sAdd = "Setup End Date:" & vbTab & format(DateValue(rst.FIELDS("SEND")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56SENTT")) & vbNewLine & vbNewLine
'''''        Else
'''''            sAdd = "Setup End Date:" & vbNewLine & vbNewLine
'''''        End If
'''''        stext = stext & sAdd
'''''
'''''        If rst.FIELDS("SHY56TBEDT") <> 0 Then
'''''            sAdd = "Dismantle Begin:" & vbTab & format(DateValue(rst.FIELDS("TBED")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56TBETT")) & vbNewLine
'''''        Else
'''''            sAdd = "Dismantle Begin:" & vbNewLine
'''''        End If
'''''        stext = stext & sAdd
'''''
'''''        If rst.FIELDS("SHY56TEDDT") <> 0 Then
'''''            sAdd = "Dismantle End:" & vbTab & format(DateValue(rst.FIELDS("TEDD")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56TENTT")) & vbNewLine & vbNewLine
'''''        Else
'''''            sAdd = "Dismantle End:" & vbNewLine & vbNewLine
'''''        End If
'''''        stext = stext & sAdd
'''''
'''''        If rst.FIELDS("SHY56VBEDT") <> 0 Then
'''''            sAdd = "Vehicle Begin:" & vbTab & format(DateValue(rst.FIELDS("VBED")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56VBETT")) & vbNewLine
'''''            stext = stext & sAdd
'''''        End If
'''''
'''''        If rst.FIELDS("SHY56VENDT") <> 0 Then
'''''            sAdd = "Vehicle End:" & vbTab & format(DateValue(rst.FIELDS("VEND")), "dddd mmm d, yyyy") & _
'''''                        " @ " & ConvertTime(rst.FIELDS("SHY56VENTT")) & vbNewLine & vbNewLine
'''''            stext = stext & sAdd
'''''        End If
'''''    End If
'''''
'''''    If rst.FIELDS("SHY56FCCDT") <> 0 Then
'''''        strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
'''''                    "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
'''''                    "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
'''''                    "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
'''''                    "WHERE AB.ABAN8 = " & rst.FIELDS("SHY56FCCDT") & " " & _
'''''                    "AND AB.ABAN8 = AL.ALAN8 " & _
'''''                    "AND AL.ALAN8 = WP.WPAN8"
'''''        Set rstX = Conn.Execute(strSelect)
'''''        If Not rstX.EOF Then
'''''            sAdd = "Facility:     " & vbTab & UCase(Trim(rstX.FIELDS("ABALPH"))) & vbNewLine
'''''            stext = stext & sAdd
'''''            If Trim(rstX.FIELDS("ALADD1")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD1"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD2")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD2"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD3")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD3"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD4")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD4"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALCTY1")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALCTY1"))) & ", " & _
'''''                            UCase(Trim(rstX.FIELDS("ALADDS"))) & "  " & _
'''''                            UCase(Trim(rstX.FIELDS("ALADDZ"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            Do While Not rstX.EOF
'''''                Select Case Trim(rstX.FIELDS("WPPHTP"))
'''''                    Case ""
'''''                        sAdd = "Facility Phone:" & vbTab & UCase(Trim(rstX.FIELDS("WPAR1"))) & _
'''''                                    " " & UCase(Trim(rstX.FIELDS("WPPH1"))) & vbNewLine
'''''                        stext = stext & sAdd
'''''                    Case "FAX"
'''''                        sAdd = "Facility Fax:" & vbTab & UCase(Trim(rstX.FIELDS("WPAR1"))) & _
'''''                                    " " & UCase(Trim(rstX.FIELDS("WPPH1"))) & vbNewLine
'''''                        stext = stext & sAdd
'''''                End Select
'''''                rstX.MoveNext
'''''            Loop
'''''            stext = stext & vbNewLine
'''''        End If
'''''        rstX.Close: Set rstX = Nothing
'''''    End If
'''''
'''''    If rst.FIELDS("SHY56SMGRT") <> 0 Then
'''''        strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
'''''                    "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
'''''                    "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
'''''                    "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
'''''                    "WHERE AB.ABAN8 = " & rst.FIELDS("SHY56SMGRT") & " " & _
'''''                    "AND AB.ABAN8 = AL.ALAN8 " & _
'''''                    "AND AL.ALAN8 = WP.WPAN8"
'''''        Set rstX = Conn.Execute(strSelect)
'''''        If Not rstX.EOF Then
'''''            sAdd = "Show Manager:" & vbTab & UCase(Trim(rstX.FIELDS("ABALPH"))) & vbNewLine
'''''            stext = stext & sAdd
'''''            If Trim(rstX.FIELDS("ALADD1")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD1"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD2")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD2"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD3")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD3"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALADD4")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALADD4"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            If Trim(rstX.FIELDS("ALCTY1")) <> "" Then
'''''                sAdd = vbTab & vbTab & UCase(Trim(rstX.FIELDS("ALCTY1"))) & ", " & _
'''''                            UCase(Trim(rstX.FIELDS("ALADDS"))) & "  " & _
'''''                            UCase(Trim(rstX.FIELDS("ALADDZ"))) & vbNewLine
'''''                stext = stext & sAdd
'''''            End If
'''''            Do While Not rstX.EOF
'''''                Select Case Trim(rstX.FIELDS("WPPHTP"))
'''''                    Case ""
'''''                        sAdd = "Show Mgr Phone:" & vbTab & UCase(Trim(rstX.FIELDS("WPAR1"))) & _
'''''                                    " " & UCase(Trim(rstX.FIELDS("WPPH1"))) & vbNewLine
'''''                        stext = stext & sAdd
'''''                    Case "FAX"
'''''                        sAdd = "Show Mgr Fax:" & vbTab & UCase(Trim(rstX.FIELDS("WPAR1"))) & _
'''''                                    " " & UCase(Trim(rstX.FIELDS("WPPH1"))) & vbNewLine
'''''                        stext = stext & sAdd
'''''                End Select
'''''                rstX.MoveNext
'''''            Loop
'''''            stext = stext & vbNewLine
'''''        End If
'''''        rstX.Close: Set rstX = Nothing
'''''    End If
'''''
'''''    rst.Close: Set rst = Nothing
'''''
'''''
'''''    '///// NEXT, GET FLOORPLAN STATUS \\\\\
'''''    stext = stext & vbNewLine
'''''    strSelect = "SELECT KU.FPSTATUS, KU.FPSTATBY, " & _
'''''                "TO_CHAR(KU.FPSTATDT, 'MM/DD/YYYY') AS FPSTATDT, R.VAL_DESC " & _
'''''                "FROM " & IGLKitU & " KU, " & IGLRef & " R " & _
'''''                "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''''                "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'''''                "AND KU.SHYR = " & tmpSHYR & " " & _
'''''                "AND KU.FPSTATUS = R.REF_ID " & _
'''''                "AND R.TYPE_CD = 12"
'''''    strSelect = "SELECT DM.DSTATUS, DM.UPDUSER, " & _
'''''                "TO_CHAR(DM.UPDDTTM, 'MM/DD/YYYY') AS FPSTATDT, R.VAL_DESC " & _
'''''                "FROM " & DWGMas & " DM, " & DWGShow & " DS, " & IGLRef & " R " & _
'''''                "WHERE DS.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''''                "AND DS.AN8_SHCD = " & tmpSHCD & " " & _
'''''                "AND DS.SHYR = " & tmpSHYR & " " & _
'''''                "AND DS.DWGID = DM.DWGID " & _
'''''                "AND DM.DWGTYPE = 0 " & _
'''''                "AND DM.DSTATUS = R.REF_ID " & _
'''''                "AND R.TYPE_CD = 12"
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''        iSFPS = Len(stext)
'''''        sAdd = tmpFBCN & " FPS" & vbNewLine
'''''        iLFPS = Len(sAdd)
'''''        stext = stext & sAdd
'''''        sAdd = vbTab & rst.FIELDS("DSTATUS") & " - " & UCase(Trim(rst.FIELDS("VAL_DESC"))) & _
'''''                    "  (LAST EDIT: " & Trim(rst.FIELDS("FPSTATDT")) & " BY " & _
'''''                    UCase(Trim(rst.FIELDS("UPDUSER"))) & ")" & vbNewLine
'''''        stext = stext & sAdd
'''''    End If
'''''
'''''    rst.Close
'''''    Set rst = Nothing
'''''
'''''    '///// LAST, GET SHOW REG ABSTRACT DATA \\\\\
'''''    stext = stext & vbNewLine
'''''    strSelect = "SELECT CH.HALLID, HM.AN8_FCCD, CH.AN8_SHCD, SU.ABALPH, " & _
'''''                "HM.HALLDESC, HM.CLGHGT, HM.CLGUNIT, HM.CLGNOTE, HM.HALLNOTE, " & _
'''''                "SHR.HGTRES, SHR.RESUNIT, SHR.RESNOTE, SHR.SHOWNOTE, " & _
'''''                "EA.EASENAME , EA.EASEVAL, EA.EASEUNIT, EA.EASEDESC " & _
'''''                "FROM " & SRACliHall & " CH, " & SRAHallMas & " HM, " & SRAEase & " EA, " & _
'''''                "" & F0101 & " SU, " & SRAHallRes & " SHR " & _
'''''                "WHERE CH.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''''                "AND CH.SHYR = " & tmpSHYR & " " & _
'''''                "AND CH.AN8_SHCD = " & tmpSHCD & " " & _
'''''                "AND CH.HALLID = HM.HALLID " & _
'''''                "AND HM.AN8_FCCD = SU.ABAN8 " & _
'''''                "AND CH.HALLID =SHR.HALLID " & _
'''''                "AND CH.AN8_SHCD = SHR.AN8_SHCD " & _
'''''                "AND CH.HALLID = EA.HALLID " & _
'''''                "AND CH.AN8_SHCD = EA.AN8_SHCD"
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''        iSSRA = Len(stext)
'''''        sAdd = "SHOW REGULATION ABSTRACT" & vbNewLine
'''''        iLSRA = Len(sAdd)
'''''        stext = stext & sAdd
'''''        sAdd = "Facility:            " & vbTab & UCase(Trim(rst.FIELDS("ABALPH"))) & vbNewLine
'''''        stext = stext & sAdd
'''''        sAdd = "Hall:                 " & vbTab & UCase(Trim(rst.FIELDS("HALLDESC"))) & vbNewLine
'''''        stext = stext & sAdd
'''''        sAdd = "Hall Ceiling Hgt:   " & vbTab & ConvertDims(CDbl(rst.FIELDS("CLGHGT")), rst.FIELDS("CLGUNIT")) & vbNewLine
'''''        stext = stext & sAdd
'''''
'''''        sAdd = "Show Restriction:   " & vbTab & ConvertDims(CDbl(rst.FIELDS("HGTRES")), rst.FIELDS("RESUNIT")) & vbNewLine
'''''        stext = stext & sAdd
'''''
'''''        i = 1
'''''        If Trim(rst.FIELDS("EASENAME")) <> "" Then
'''''            sAdd = "Easements:"
'''''            stext = stext & sAdd
'''''            Do While Not rst.EOF
'''''                Select Case i
'''''                    Case 1
'''''                        sPad = vbTab
'''''                        i = 2
'''''                    Case 2
'''''                        sPad = vbTab & vbTab
'''''                End Select
'''''                sAdd = sPad & UCase(Trim(rst.FIELDS("EASENAME"))) & "  (" & _
'''''                            ConvertDims(CDbl(rst.FIELDS("EASEVAL")), rst.FIELDS("EASEUNIT")) & ")" & vbNewLine
'''''                stext = stext & sAdd
'''''                rst.MoveNext
'''''            Loop
'''''        End If
'''''    End If
'''''
'''''    rst.Close
'''''    Set rst = Nothing
'''''
'''''    With rtbInfo
'''''        .Text = stext
'''''
'''''        .SelStart = iSHdr
'''''        .SelLength = iLHdr
'''''        .SelBold = True
'''''        .SelFontSize = 10
'''''
'''''        .SelStart = iLHdr + 1
'''''        .SelLength = Len(stext) - iLHdr
'''''        .SelBold = False
'''''        .SelFontSize = 8
'''''
'''''        .SelStart = iSFPS
'''''        .SelLength = iLFPS
'''''        .SelBold = True
'''''        .SelFontSize = 10
'''''
'''''        .SelStart = iSSRA
'''''        .SelLength = iLSRA
'''''        .SelBold = True
'''''        .SelFontSize = 10
'''''
'''''        .SelStart = iSHdr
'''''        .SelLength = Len(stext)
'''''        .SelTabCount = 3
'''''        .SelTabs(0) = 0
'''''        .SelTabs(1) = 500
'''''        .SelTabs(2) = 1550
'''''        .SelIndent = 80
'''''        .SelRightIndent = 80
'''''        .SelLength = 0
'''''    End With
'''''    rtbInfo.Visible = True
'''''End Sub

'''Public Function PopShowInfo(tmpBCC As Long, tmpSHYR As Integer, tmpSHCD As Long, _
'''            tmpSHNM As String, tmpFBCN As String) As String
'''    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
'''    Dim strSelect As String, sHTML As String, sDate1 As String, sDate2 As String, tFile1 As String
'''    Dim i As Integer
'''    Dim htmO As String, htmC As String
'''    Dim hdO As String, hdC As String
'''    Dim tiO As String, tiC As String
'''    Dim bodO As String, bodC As String
'''    Dim f1O As String, f2O As String, f3O As String, fC As String
'''    Dim bolO As String, bolC As String
'''    Dim tblO As String, tblC As String
'''    Dim trO As String, trC As String
'''    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, tdOa As String, tdOb As String, tdC As String
'''    Dim tdNO As String, tdNC As String
'''    Dim hr As String, br As String
'''
'''
'''
'''    htmO = "<HTML>": htmC = "</HTML>"
'''    hdO = "<HEAD>": hdC = "</HEAD>"
'''    tiO = "<TITLE>": tiC = "</TITLE>"
'''    bodO = "<BODY>": bodC = "</BODY>"
'''    f2O = "<FONT SIZE=2 FACE=""Arial"">"
'''    f3O = "<FONT SIZE=3 FACE=""Arial"">"
'''    fC = "</FONT>"
'''    bolO = "<B>": bolC = "</B>"
'''    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
'''    trO = "<TR VALIGN=""top"">": trC = "</TR>"
'''    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
'''    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
'''    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
'''    tdcC = "</B></FONT></DIV></TD>"
'''    tdNO = "<TD WIDTH=""100%"" colspan=2><DIV align=left><FONT SIZE=2 COLOR=""#FF0000 "" FACE=""Arial"">"
'''    tdNC = "</FONT></DIV></TD>"
'''    tdOa = "<TD WIDTH=""": tdOb = "%"" VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
'''    hr = "<HR>": br = "<BR>"
'''
'''
'''    strSelect = "SELECT SM.SHY56SHTP, SM.SHY56TENDT, " & _
'''                "SM.SHY56BEGDT, IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, SM.SHY56BEGTT, " & _
'''                "SM.SHY56ENDDT, IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, SM.SHY56ENDTT, " & _
'''                "CS.CSY56FARDT, IGL_JDEDATE_TOCHAR(CS.CSY56FARDT, 'MM/DD/YYYY') AS FRAD, CS.CSY56FARTT, " & _
'''                "SM.SHY56SBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MM/DD/YYYY') AS SBED, SM.SHY56SBETT, " & _
'''                "SM.SHY56SENDT, IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MM/DD/YYYY') AS SEND, SM.SHY56SENTT, " & _
'''                "CS.CSY56VMVDT, IGL_JDEDATE_TOCHAR(CS.CSY56VMVDT, 'MM/DD/YYYY') AS VMVD, CS.CSY56VMVTT, " & _
'''                "SM.SHY56PBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'MM/DD/YYYY') AS PBED, SM.SHY56PBETT, " & _
'''                "SM.SHY56PENDT, IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'MM/DD/YYYY') AS PEND, SM.SHY56PENTT, " & _
'''                "SM.SHY56VBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MM/DD/YYYY') AS VBED, SM.SHY56VBETT, " & _
'''                "SM.SHY56VENDT, IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MM/DD/YYYY') AS VEND, SM.SHY56VENTT, " & _
'''                "SM.SHY56TBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MM/DD/YYYY') AS TBED, SM.SHY56TBETT, " & _
'''                "SM.SHY56TEDDT, IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MM/DD/YYYY') AS TEDD, SM.SHY56TENTT, " & _
'''                "SM.SHY56FCCDT , SM.SHY56SMGRT, SM.SHY56DRAIT, SM.SHY56CARIT, SM.SHY56VACIT " & _
'''                "FROM " & F5601 & " SM, " & F5611 & " CS " & _
'''                "WHERE SM.SHY56SHCD = " & tmpSHCD & " " & _
'''                "AND SM.SHY56SHYR = " & tmpSHYR & " " & _
'''                "AND SM.SHY56SHCD =CS.CSY56SHCD " & _
'''                "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
'''                "AND CS.CSY56CUNO = " & tmpBCC
'''
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        sHTML = htmO & vbNewLine
'''        sHTML = sHTML & hdO & tiO & tmpFBCN & " - " & tmpSHYR & " " & tmpSHNM & tiC & hdC & vbNewLine
'''        sHTML = sHTML & bodO & vbNewLine
'''        sHTML = sHTML & f3O & bolO & tmpFBCN & " - " & tmpSHYR & " " & tmpSHNM & bolC & fC & vbNewLine
'''        sHTML = sHTML & hr & vbNewLine
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & tdc2O & "Show Dates" & tdcC & trC & vbNewLine
'''        If rst.Fields("SHY56TENDT") = 1 Then
'''            sHTML = sHTML & trO & tdNO & bolO & "Note:  " & bolC & "All Dates are TENTATIVE" & tdNC & trC & vbNewLine
'''        End If
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Dates:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56BEGDT") <> 0 And rst.Fields("SHY56ENDDT") <> 0 Then 'CHANGE
'''            sDate1 = format(DateValue(rst.Fields("BEGD")), "dddd mmm d")
'''            sDate2 = format(DateValue(rst.Fields("ENDD")), "dddd mmm d")
'''            sHTML = sHTML & tdOa & "63" & tdOb & sDate1 & " - " & sDate2 & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Freight Arrive Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("CSY56FARDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("FRAD")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("CSY56FARTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Setup Begin Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56SBEDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("SBED")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56SBETT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Setup End Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56SENDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("SEND")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56SENTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        If UCase(Trim(rst.Fields("SHY56SHTP"))) = "S1" Then
'''            sHTML = sHTML & tblO & vbNewLine
'''            sHTML = sHTML & trO & vbNewLine
'''            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Vehicle Move-In:" & bolC & tdC & vbNewLine
'''            If rst.Fields("CSY56VMVDT") <> 0 Then
'''                sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("VMVD")), "dddd mmm d, yyyy") & _
'''                            " @ " & ConvertTime(rst.Fields("CSY56VMVTT")) & tdC & vbNewLine
'''            End If
'''            sHTML = sHTML & trC & vbNewLine
'''            sHTML = sHTML & tblC & vbNewLine
'''
'''            sHTML = sHTML & br & vbNewLine
'''        End If
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Press Begin Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56PBEDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("PBED")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56PBETT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Press End Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56PENDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("PEND")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56PENTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Preview Begin Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56VBEDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("VBED")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56VBETT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Preview End Date:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56VENDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("VEND")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56VENTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Open to Public:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56BEGDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("BEGD")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56BEGTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Close to Public:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56ENDDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("ENDD")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56ENDTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
'''        sHTML = sHTML & br & vbNewLine
'''
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Dismantle Begin:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56TBEDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("TBED")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56TBETT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Dismantle End:" & bolC & tdC & vbNewLine
'''        If rst.Fields("SHY56TEDDT") <> 0 Then
'''            sHTML = sHTML & tdOa & "63" & tdOb & format(DateValue(rst.Fields("TEDD")), "dddd mmm d, yyyy") & _
'''                        " @ " & ConvertTime(rst.Fields("SHY56TENTT")) & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''
''''''        sHTML = sHTML & br & vbNewLine
'''
'''
'''        sHTML = sHTML & hr & vbNewLine
'''
'''        If rst.Fields("SHY56FCCDT") <> 0 Then
'''            sHTML = sHTML & tblO & vbNewLine
'''            sHTML = sHTML & trO & tdc2O & "Facility" & tdcC & trC & vbNewLine
'''            strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
'''                        "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
'''                        "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
'''                        "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
'''                        "WHERE AB.ABAN8 = " & rst.Fields("SHY56FCCDT") & " " & _
'''                        "AND AB.ABAN8 = AL.ALAN8 " & _
'''                        "AND AL.ALAN8 = WP.WPAN8"
'''            Set rstX = Conn.Execute(strSelect)
'''            If Not rstX.EOF Then
'''                sHTML = sHTML & trO & vbNewLine
'''                sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility:" & bolC & tdC & vbNewLine
'''                sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
'''                sHTML = sHTML & UCase(Trim(rstX.Fields("ABALPH"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD1")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD1"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD2")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD2"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD3")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD3"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD4")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD4"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALCTY1")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALCTY1"))) & ", " & _
'''                            UCase(Trim(rstX.Fields("ALADDS"))) & "  " & _
'''                            Trim(rstX.Fields("ALADDZ")) & br & vbNewLine
'''                sHTML = sHTML & tdC & vbNewLine
'''                sHTML = sHTML & trC & vbNewLine
'''                sHTML = sHTML & tblC & vbNewLine
'''                sHTML = sHTML & tblO & vbNewLine
'''
'''                Do While Not rstX.EOF
'''                    Select Case Trim(rstX.Fields("WPPHTP"))
'''                        Case ""
'''                            sHTML = sHTML & trO & vbNewLine
'''                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility Phone:" & bolC & tdC & vbNewLine
'''                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
'''                                        " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
'''                            sHTML = sHTML & trC & vbNewLine
'''                        Case "FAX"
'''                            sHTML = sHTML & trO & vbNewLine
'''                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility Fax:" & bolC & tdC & vbNewLine
'''                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
'''                                    " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
'''                            sHTML = sHTML & trC & vbNewLine
'''                    End Select
'''                    rstX.MoveNext
'''                Loop
'''                sHTML = sHTML & tblC & vbNewLine
'''                sHTML = sHTML & hr & vbNewLine
'''            End If
'''            rstX.Close: Set rstX = Nothing
'''        End If
'''
'''
'''        If rst.Fields("SHY56SMGRT") <> 0 Then
'''            sHTML = sHTML & tblO & vbNewLine
'''            sHTML = sHTML & trO & tdc2O & "Show Manager" & tdcC & trC & vbNewLine
'''            strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
'''                        "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
'''                        "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
'''                        "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
'''                        "WHERE AB.ABAN8 = " & rst.Fields("SHY56SMGRT") & " " & _
'''                        "AND AB.ABAN8 = AL.ALAN8 " & _
'''                        "AND AL.ALAN8 = WP.WPAN8"
'''            Set rstX = Conn.Execute(strSelect)
'''            If Not rstX.EOF Then
'''                sHTML = sHTML & trO & vbNewLine
'''                sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Manager:" & bolC & tdC & vbNewLine
'''                sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
'''                sHTML = sHTML & UCase(Trim(rstX.Fields("ABALPH"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD1")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD1"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD2")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD2"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD3")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD3"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALADD4")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD4"))) & br & vbNewLine
'''                If Trim(rstX.Fields("ALCTY1")) <> "" Then _
'''                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALCTY1"))) & ", " & _
'''                            UCase(Trim(rstX.Fields("ALADDS"))) & "  " & _
'''                            Trim(rstX.Fields("ALADDZ")) & br & vbNewLine
'''                sHTML = sHTML & tdC & vbNewLine
'''                sHTML = sHTML & trC & vbNewLine
'''                sHTML = sHTML & tblC & vbNewLine
'''                sHTML = sHTML & tblO & vbNewLine
'''
'''                Do While Not rstX.EOF
'''                    Select Case Trim(rstX.Fields("WPPHTP"))
'''                        Case ""
'''                            sHTML = sHTML & trO & vbNewLine
'''                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Mgr Phone:" & bolC & tdC & vbNewLine
'''                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
'''                                        " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
'''                            sHTML = sHTML & trC & vbNewLine
'''                        Case "FAX"
'''                            sHTML = sHTML & trO & vbNewLine
'''                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Mgr Fax:" & bolC & tdC & vbNewLine
'''                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
'''                                    " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
'''                            sHTML = sHTML & trC & vbNewLine
'''                    End Select
'''                    rstX.MoveNext
'''                Loop
'''                sHTML = sHTML & tblC & vbNewLine
'''                sHTML = sHTML & hr & vbNewLine
'''            End If
'''            rstX.Close: Set rstX = Nothing
'''        End If
'''    End If
'''    rst.Close: Set rst = Nothing
'''
'''
'''    '///// GET SHOW REG ABSTRACT DATA \\\\\
'''    strSelect = "SELECT CH.HALLID, HM.AN8_FCCD, CH.AN8_SHCD, SU.ABALPH, " & _
'''                "HM.HALLDESC, HM.CLGHGT, HM.CLGUNIT, HM.CLGNOTE, HM.HALLNOTE, " & _
'''                "SHR.HGTRES, SHR.RESUNIT, SHR.RESNOTE, SHR.SHOWNOTE, " & _
'''                "EA.EASENAME , EA.EASEVAL, EA.EASEUNIT, EA.EASEDESC " & _
'''                "FROM IGLPROD.SRA_CLIENTHALL CH, IGLPROD.SRA_HALLMASTER HM, " & _
'''                "IGLPROD.SRA_EASEMENT EA, " & _
'''                "" & F0101 & " SU, IGLPROD.SRA_SHOWHALLRESTRICTION SHR " & _
'''                "WHERE CH.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''                "AND CH.SHYR = " & tmpSHYR & " " & _
'''                "AND CH.AN8_SHCD = " & tmpSHCD & " " & _
'''                "AND CH.HALLID = HM.HALLID " & _
'''                "AND HM.AN8_FCCD = SU.ABAN8 " & _
'''                "AND CH.HALLID =SHR.HALLID " & _
'''                "AND CH.AN8_SHCD = SHR.AN8_SHCD " & _
'''                "AND CH.HALLID = EA.HALLID " & _
'''                "AND CH.AN8_SHCD = EA.AN8_SHCD"
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & tdc2O & "Hall Information from Show Regulation Abstract" & tdcC & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & UCase(Trim(rst.Fields("ABALPH"))) & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Hall:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & UCase(Trim(rst.Fields("HALLDESC"))) & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Hall Ceiling Hgt:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & ConvertDims(CDbl(rst.Fields("CLGHGT")), rst.Fields("CLGUNIT")) & _
'''                    tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Restriction:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & ConvertDims(CDbl(rst.Fields("HGTRES")), rst.Fields("RESUNIT")) & _
'''                    tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        If Trim(rst.Fields("EASENAME")) <> "" Then
'''            sHTML = sHTML & trO & vbNewLine
'''            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Easements:" & bolC & tdC & vbNewLine
'''            sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
'''            Do While Not rst.EOF
'''                sHTML = sHTML & UCase(Trim(rst.Fields("EASENAME"))) & "  (" & _
'''                            ConvertDims(CDbl(rst.Fields("EASEVAL")), rst.Fields("EASEUNIT")) & ")" & _
'''                            br & vbNewLine
'''                rst.MoveNext
'''            Loop
'''            sHTML = sHTML & tdC & vbNewLine
'''            sHTML = sHTML & trC & vbNewLine
'''        End If
'''        sHTML = sHTML & tblC & vbNewLine
'''        sHTML = sHTML & hr & vbNewLine
'''    End If
'''    rst.Close: Set rst = Nothing
'''
'''
'''
'''    '///// GET FLOORPLAN STATUS \\\\\
'''    Dim sSQFT As String
'''    strSelect = "SELECT KU.FPSTATUS, KU.FPSTATBY, " & _
'''                "TO_CHAR(KU.FPSTATDT, 'MM/DD/YYYY') AS FPSTATDT, R.VAL_DESC " & _
'''                "FROM IGL_KIT_USE KU, IGL_REF R " & _
'''                "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''                "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'''                "AND KU.SHYR = " & tmpSHYR & " " & _
'''                "AND KU.FPSTATUS = R.REF_ID " & _
'''                "AND R.TYPE_CD = 12"
'''    strSelect = "SELECT DM.DSTATUS, DM.UPDUSER, " & _
'''                "TO_CHAR(DM.UPDDTTM, 'MON DD, YYYY') AS FPSTATDT, R.VAL_DESC, " & _
'''                "CS.CSY56BOOTT BONO, CS.CSY56SQFTT AS SQFT, CS.CSY56BOTPT AS BOPH " & _
'''                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHOW DS, IGL_REF R, " & _
'''                "" & F5611 & " CS " & _
'''                "WHERE DS.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''                "AND DS.AN8_SHCD = " & tmpSHCD & " " & _
'''                "AND DS.SHYR = " & tmpSHYR & " " & _
'''                "AND DS.DWGID = DM.DWGID " & _
'''                "AND DM.DWGTYPE = 0 " & _
'''                "AND DM.DSTATUS = R.REF_ID " & _
'''                "AND R.TYPE_CD = 12 " & _
'''                "AND DS.AN8_CUNO = CS.CSY56CUNO " & _
'''                "AND DS.AN8_SHCD = CS.CSY56SHCD " & _
'''                "AND DS.SHYR = CS.CSY56SHYR"
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        sHTML = sHTML & tblO & vbNewLine
'''        sHTML = sHTML & trO & tdc2O & "Space Information" & tdcC & trC & vbNewLine
'''
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Space Number:" & bolC & tdC & vbNewLine
'''        If Not IsNull(rst.Fields("BONO")) Then _
'''            sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("BONO") & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Floor Space Area:" & bolC & tdC & vbNewLine
'''        If Not IsNull(rst.Fields("SQFT")) Then
'''            sSQFT = format(rst.Fields("SQFT") / 1000, "#,##0")
'''            sHTML = sHTML & tdOa & "63" & tdOb & sSQFT & tdC & vbNewLine
'''        End If
'''        sHTML = sHTML & trC & vbNewLine
'''
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Booth Phone No:" & bolC & tdC & vbNewLine
'''        If Not IsNull(rst.Fields("BOPH")) Then _
'''            sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("BOPH") & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Is Floorplan Req'd?:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & "YES" & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Floorplan Status:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("DSTATUS") & _
'''                    " - " & UCase(Trim(rst.Fields("VAL_DESC"))) & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & trO & vbNewLine
'''        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Status By:" & bolC & tdC & vbNewLine
'''        sHTML = sHTML & tdOa & "63" & tdOb & "LAST EDIT ON " & UCase(Trim(rst.Fields("FPSTATDT"))) & _
'''                    " BY " & UCase(Trim(rst.Fields("UPDUSER"))) & tdC & vbNewLine
'''        sHTML = sHTML & trC & vbNewLine
'''        sHTML = sHTML & tblC & vbNewLine
'''        sHTML = sHTML & hr & vbNewLine
'''    End If
'''    rst.Close
'''    Set rst = Nothing
'''
'''
'''
'''
'''
'''
'''    sHTML = sHTML & bodC & vbNewLine
'''    sHTML = sHTML & htmC
'''
'''    tFile1 = strHTMLPath & sIFile
'''    Open tFile1 For Output As #1
'''    Print #1, sHTML
'''    Close #1
'''
'''    PopShowInfo = tFile1
'''End Function

Public Function ConvertDims(Num As Double, iUnit As Integer) As String
    Dim Feet As Integer, Inch As Integer, Numer As Integer
    Dim Frac As Currency
    Dim strFrac As String
    Select Case iUnit
        Case 1
            Feet = Int(Num / 12)
            Inch = Int(Num - (Feet * 12))
            Frac = CCur((((Num / 12) - Feet) _
                    * 12) - Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case 2
            Feet = Int(Num)
            Inch = (Num - Feet) * 12
            Frac = Inch - Int(Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case Else
            ConvertDims = "Soon!"
    End Select
End Function

'''''Public Sub GetAttSupers(tmpSHYR As Integer, tmpSHCD As Long, tmpSHNM As String)
'''''    Dim strSelect As String, sClient As String, sMess As String, sAdd As String, _
'''''                sFirst As String, sLast As String, sFull As String
'''''    Dim rst As ADODB.Recordset
'''''    Dim iSHdr As Integer, iLHdr As Integer, iCom As Integer
'''''
''''''''    rtbSupers.Visible = False
'''''    rtbSupers.Text = ""
'''''    sMess = tmpSHYR & " " & tmpSHNM & vbNewLine
'''''    iSHdr = 0: iLHdr = Len(sMess) - 1
'''''    sClient = ""
'''''
''''''''''SELECT DISTINCT ET.AN8_CUNO, UPPER(C.ABALPH) AS CLNM, ET.AN8_EMNO,
''''''''''UPPER(E.YAALPH) AS EMNM, EP.CEL
''''''''''From
''''''''''IGL_EMPLOYEE_TASK ET,
''''''''''" & F060116 & " E,
''''''''''(SELECT ABC.WPAN8, DECODE (TRIM (WPAR1), NULL, TRIM (WPPH1), TRIM (WPAR1) || ' ' || TRIM (WPPH1)) CEL
''''''''''FROM " & F0115 & " ABC, " & F0101 & " AB
''''''''''Where AB.ABAN8 = ABC.WPAN8
''''''''''AND UPPER (TRIM (ABC.WPPHTP)) = 'CEL'
''''''''''AND TRIM (ABC.WPPH1) IS NOT NULL) EP,
''''''''''" & F0101 & " C
''''''''''Where ET.AN8_SHCD = 12920
''''''''''AND ET.SHYR = 2001
''''''''''AND ET.AN8_CUNO IN (1011, 1009, 1024, 1025, 13837, 1017, 1034, 1032,
''''''''''13857, 13843, 13862, 13855, 13865, 13847, 13841, 13851, 13848, 13849,
''''''''''13856, 1070, 13878, 1065, 1067, 1061, 1076, 1012, 1208, 1237, 1238,
''''''''''1221, 1232, 1213, 1223, 1265, 1218, 1231, 1225, 1226, 1217, 1236,
''''''''''1277, 1233, 1230, 1216, 1249, 1211, 1253, 1234, 1244, 1212, 1210,
''''''''''1276, 1247, 1273, 1228, 1252, 1229, 1101, 1104, 1105, 13846, 1108,
''''''''''1117, 1120, 1133, 15271, 1143, 1151, 1148, 1046, 1161, 1165, 1166,
''''''''''1168, 1159, 1088, 13900, 1190, 1203)
''''''''''AND ET.AN8_EMNO = E.YAAN8
''''''''''AND ET.AN8_EMNO = EP.WPAN8 (+)
''''''''''AND ET.AN8_CUNO = C.ABAN8
''''''''''ORDER BY CLNM, UPPER(EMNM);
'''''
'''''
'''''    strSelect = "SELECT DISTINCT ET.AN8_CUNO, UPPER(C.ABALPH) AS CLNM, ET.AN8_EMNO, " & _
'''''                "UPPER(E.YAALPH) AS EMNM, EP.CEL " & _
'''''                "FROM " & IGLEmpTask & " ET, " & F060116 & " E, " & _
'''''                "(SELECT ABC.WPAN8, DECODE (TRIM (WPAR1), NULL, TRIM (WPPH1), " & _
'''''                "TRIM (WPAR1) || ' ' || TRIM (WPPH1)) CEL " & _
'''''                "FROM " & F0115 & " ABC, " & F0101 & " AB " & _
'''''                "WHERE AB.ABAN8 = ABC.WPAN8 " & _
'''''                "AND UPPER (TRIM (ABC.WPPHTP)) = 'CEL' " & _
'''''                "AND TRIM (ABC.WPPH1) IS NOT NULL) EP, " & F0101 & " C " & _
'''''                "WHERE ET.AN8_SHCD = " & tmpSHCD & " " & _
'''''                "AND ET.SHYR = " & tmpSHYR & " " & _
'''''                "AND ET.AN8_CUNO IN (" & CunoList & ") " & _
'''''                "AND ET.AN8_EMNO = E.YAAN8 " & _
'''''                "AND ET.AN8_EMNO = EP.WPAN8 (+) " & _
'''''                "AND ET.AN8_CUNO = C.ABAN8 " & _
'''''                "ORDER BY CLNM, UPPER(EMNM)"
'''''    Debug.Print strSelect
'''''    Set rst = Conn.Execute(strSelect)
'''''    Do While Not rst.EOF
'''''        If sClient <> Trim(rst.FIELDS("CLNM")) Then
'''''            sClient = Trim(rst.FIELDS("CLNM"))
'''''            sAdd = sClient & vbNewLine
'''''            sMess = sMess & vbNewLine & sAdd
'''''        End If
'''''        iCom = InStr(1, Trim(rst.FIELDS("EMNM")), ",")
'''''        sLast = Left(Trim(rst.FIELDS("EMNM")), iCom - 1)
'''''        sFirst = Trim(Mid(rst.FIELDS("EMNM"), iCom + 1))
'''''        sFull = Left(sFirst & " " & sLast & Space(35), 35)
'''''        If IsNull(Trim(rst.FIELDS("CEL"))) Then
'''''            sAdd = vbTab & sFull & vbNewLine
'''''        Else
'''''            sAdd = vbTab & sFull & vbTab & "[" & Trim(rst.FIELDS("CEL")) & "]" & vbNewLine
'''''        End If
'''''        sMess = sMess & sAdd
'''''        rst.MoveNext
'''''    Loop
'''''    rst.Close
'''''    Set rst = Nothing
'''''
'''''    With rtbSupers
'''''        .Text = sMess
'''''
'''''        .SelStart = iSHdr
'''''        .SelLength = iLHdr
'''''        .SelBold = True
'''''        .SelFontSize = 10
'''''
'''''        .SelStart = iLHdr + 1
'''''        .SelLength = Len(sMess) - iLHdr
'''''        .SelBold = False
'''''        .SelFontSize = 8
'''''
'''''        .SelStart = iSHdr
'''''        .SelLength = Len(sMess)
'''''        .SelTabCount = 4
'''''        .SelTabs(0) = 0
'''''        .SelTabs(1) = 300
'''''        .SelTabs(2) = 3000
'''''        .SelTabs(3) = 3600
'''''        .SelIndent = 80
'''''        .SelRightIndent = 80
'''''        .SelLength = 0
'''''    End With
''''''''''    rtbSupers.Visible = True
''''''''''    MsgBox sMess, vbInformation, "Attending Supervisors..."
'''''
'''''End Sub

Public Function GetAttSupers(tmpSHYR As Integer, tmpSHCD As Long, tmpSHNM As String) As String
    Dim strSelect As String, sClient As String, sMess As String, sAdd As String, _
                sFirst As String, sLast As String, sFull As String
    Dim tFile1 As String, sHTML As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iCom As Integer
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, tdOa As String, tdOb As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    
    
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY>": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=3><DIV align=left><FONT SIZE=2 FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdOb = "%"" VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & tmpSHYR & " " & tmpSHNM & tiC & hdC & vbNewLine
    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & bolO & tmpSHYR & " " & tmpSHNM & bolC & fC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & trO & tdc3O & "Attending Supervisors / Cell Phones" & tdcC & trC & vbNewLine
    
    sClient = ""
    strSelect = "SELECT DISTINCT ET.AN8_CUNO, UPPER(C.ABALPH) AS CLNM, ET.AN8_EMNO, " & _
                "UPPER(E.YAALPH) AS EMNM, EP.CEL " & _
                "FROM " & IGLEmpTask & " ET, " & F060116_View & " E, " & _
                "(SELECT ABC.WPAN8, DECODE (TRIM (ABC.WPAR1), NULL, TRIM (ABC.WPPH1), " & _
                "TRIM (ABC.WPAR1) || ' ' || TRIM (ABC.WPPH1)) CEL " & _
                "FROM " & F0115 & " ABC, " & F0101 & " AB " & _
                "WHERE AB.ABAN8 = ABC.WPAN8 " & _
                "AND UPPER (TRIM (ABC.WPPHTP)) = 'CEL' " & _
                "AND TRIM (ABC.WPPH1) IS NOT NULL) EP, " & F0101 & " C " & _
                "WHERE ET.AN8_SHCD = " & tmpSHCD & " " & _
                "AND ET.SHYR = " & tmpSHYR & " " & _
                "AND ET.AN8_CUNO IN (" & CunoList & ") " & _
                "AND ET.AN8_EMNO = E.YAAN8 " & _
                "AND ET.AN8_EMNO = EP.WPAN8 (+) " & _
                "AND ET.AN8_CUNO = C.ABAN8 " & _
                "ORDER BY CLNM, UPPER(EMNM)"
'''''    strSelect = "SELECT DISTINCT ET.AN8_CUNO, UPPER(C.ABALPH) AS CLNM, ET.AN8_EMNO, " & _
'''''                "UPPER(E.YAALPH) AS EMNM, EP.CEL " & _
'''''                "FROM IGL_EMPLOYEE_TASK ET, " & F060116 & " E, " & _
'''''                "(SELECT ABC.WPAN8, DECODE (TRIM (WPAR1), NULL, TRIM (WPPH1), " & _
'''''                "TRIM (WPAR1) || ' ' || TRIM (WPPH1)) CEL " & _
'''''                "FROM " & F0115 & " ABC, " & F0101 & " AB " & _
'''''                "WHERE AB.ABAN8 = ABC.WPAN8 " & _
'''''                "AND UPPER (TRIM (ABC.WPPHTP)) = 'CEL' " & _
'''''                "AND TRIM (ABC.WPPH1) IS NOT NULL) EP, " & F0101 & " C " & _
'''''                "WHERE ET.AN8_SHCD = " & tmpSHCD & " " & _
'''''                "AND ET.SHYR = " & tmpSHYR & " " & _
'''''                "AND ET.AN8_CUNO IN (" & CunoList & ") " & _
'''''                "AND ET.AN8_EMNO = E.YAAN8 " & _
'''''                "AND ET.AN8_EMNO = EP.WPAN8 (+) " & _
'''''                "AND ET.AN8_CUNO = C.ABAN8 " & _
'''''                "ORDER BY CLNM, UPPER(EMNM)"

    Debug.Print strSelect
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sClient <> Trim(rst.Fields("CLNM")) Then
            If sClient <> "" Then
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & br & vbNewLine
                sHTML = sHTML & tblO & vbNewLine
            End If
            sClient = Trim(rst.Fields("CLNM"))
            sHTML = sHTML & trO & tdNO & bolO & sClient & bolC & tdNC & trC & vbNewLine
        End If
        iCom = InStr(1, Trim(rst.Fields("EMNM")), ",")
        sLast = Left(Trim(rst.Fields("EMNM")), iCom - 1)
        sFirst = Trim(Mid(rst.Fields("EMNM"), iCom + 1))
        sFull = Left(sFirst & " " & sLast & Space(35), 35)
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "2" & tdOb & tdC & vbNewLine
        sHTML = sHTML & tdOa & "60" & tdOb & sFull & tdC & vbNewLine
        If Not IsNull(Trim(rst.Fields("CEL"))) Then
            sHTML = sHTML & tdOa & "38" & tdOb & Trim(rst.Fields("CEL")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath & sSFile
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    GetAttSupers = tFile1

End Function

Public Function ConvertTime(iTime As Long) As String
    Dim iHour As Long, iMin As Long
    Dim sAMPM As String
    Dim sTime As String, sHour As String, sMin As String
    
    
    Select Case iTime
        Case 0
            ConvertTime = "12:00 AM"
        Case Is < 120000
            sTime = Right("000000" & CStr(iTime), 6)
            sHour = Left(sTime, 2)
            sMin = Mid(sTime, 3, 2)
            If CInt(sHour) = 0 Then
                ConvertTime = "12:" & sMin & " AM"
            Else
                ConvertTime = CStr(CInt(sHour)) & ":" & sMin & " AM"
            End If
        Case 120000
            ConvertTime = "12:00 NOON"
        Case 240000
            ConvertTime = "12:00 MID"
        Case Else
            iTime = iTime - 120000
            sTime = Right("000000" & CStr(iTime), 6)
            sHour = Left(sTime, 2)
            sMin = Mid(sTime, 3, 2)
            If CInt(sHour) = 0 Then
                ConvertTime = "12:" & sMin & " PM"
            Else
                ConvertTime = CStr(CInt(sHour)) & ":" & sMin & " PM"
            End If
    End Select
    
'''    iMin = iTime Mod (iHour * 10000)
'''    If iHour = 0 Then iHour = 12
'''    Select Case iMin
'''        Case 0
'''            ConvertTime = iHour & ":00" & sAMPM
'''        Case Else
'''            ConvertTime = iHour & ":" & Right("00" & CStr(iMin / 100), 2) & sAMPM
'''    End Select
'''GotIt:
End Function

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblClose.FontBold = True
End Sub

Private Sub optInfo_Click(Index As Integer)
    If bSurfing Then Exit Sub
    Select Case Index
        Case 0
'''            lblRTB.Caption = "Show Info/FPS/Show Regulation Abstract"
            web1.Navigate strHTMLPath & sIFile
        Case 1
'''            lblRTB.Caption = "GPJ Attending Supervisors"
            web1.Navigate strHTMLPath & sSFile
    End Select
End Sub

Private Sub tvwIGL_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strSelect As String, sMess As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrorFound
    Debug.Print "Selected Node: " & Node.Key
    Select Case UCase(Left(Node.Key, 1))
        Case "T"
            strSelect = "SELECT TU.PARTDESC, TU.PKGTYPE, TU.WIDTH, TU.HEIGHT, " & _
                        "TU.LENGTH, TU.SIZEUNIT, TU.WEIGHT, " & _
                        "TU.WTUNIT, R.VALUE AS WTVAL " & _
                        "FROM " & AQUAPartU & " TU, " & IGLRef & " R " & _
                        "WHERE TU.SHYR = " & tSHYR & " " & _
                        "AND TU.AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND TU.AN8_SHCD = " & tSHCD & " " & _
                        "AND TU.PRTUSEID = " & Mid(Node.Key, 2) & " " & _
                        "AND TU.ELTUSEID = " & Mid(Node.Parent.Key, 2) & " " & _
                        "AND TU.KITUSEID = " & Mid(Node.Parent.Parent.Key, 2) & " " & _
                        "AND R.TYPE_CD = 104 " & _
                        "AND TU.WTUNIT = R.REF_ID"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sMess = "Container Type:  " & Trim(rst.Fields("PKGTYPE")) & vbCr & vbCr & _
                            "Part Dimensions:  " & rst.Fields("WIDTH") & " x " & rst.Fields("HEIGHT") & " x " & _
                            rst.Fields("LENGTH") & vbCr & vbCr & _
                            "Part Weight:  " & rst.Fields("WEIGHT") & " " & rst.Fields("WTVAL")
                MsgBox sMess, vbInformation, Node.Text & "..."
            End If
            rst.Close
            Set rst = Nothing
    End Select
Exit Sub
ErrorFound:
    rst.Close
    Set rst = Nothing
    MsgBox "Data is not available.", vbInformation, Node.Text & "..."
End Sub

Private Sub web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Dim i As String
    bSurfing = True
    If InStr(1, URL, "SI-") > 0 Then
        lblRTB.Caption = "Show Info/FPS/Show Regulation Abstract"
        optInfo(0).Value = True
    Else
        lblRTB.Caption = "GPJ Attending Supervisors"
        optInfo(1).Value = True
    End If
    bSurfing = False
End Sub

