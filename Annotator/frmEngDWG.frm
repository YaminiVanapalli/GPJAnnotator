VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmEngDWG 
   BackColor       =   &H00000000&
   Caption         =   "Engineering Drawings"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEngDWG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDWG 
      Height          =   645
      Left            =   7320
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cboDWG 
      BackColor       =   &H00505050&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   7095
   End
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BackColor       =   &H00666666&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   8100
      ScaleHeight     =   420
      ScaleWidth      =   2160
      TabIndex        =   0
      Top             =   90
      Width           =   2160
      Begin VB.Image imgNav 
         Height          =   300
         Index           =   1
         Left            =   1620
         MouseIcon       =   "frmEngDWG.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmEngDWG.frx":0BD4
         Top             =   60
         Width           =   300
      End
      Begin VB.Image imgNav 
         Height          =   300
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmEngDWG.frx":11FE
         MousePointer    =   99  'Custom
         Picture         =   "frmEngDWG.frx":1508
         Top             =   60
         Width           =   300
      End
      Begin VB.Label lblNavCnt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "24 of 35"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   60
         UseMnemonic     =   0   'False
         Width           =   1335
      End
   End
   Begin VOLOVIEWXLibCtl.AvViewX vol1 
      Height          =   3795
      Left            =   1140
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   5115
      _cx             =   9022
      _cy             =   6694
      Appearance      =   0
      BorderStyle     =   0
      BackgroundColor =   "DefaultColors"
      Enabled         =   -1  'True
      UserMode        =   "ZoomToRect"
      HighlightLinks  =   0   'False
      src             =   ""
      LayersOn        =   ""
      LayersOff       =   ""
      SrcTemp         =   ""
      SupportPath     =   $"frmEngDWG.frx":1B32
      FontPath        =   $"frmEngDWG.frx":1D2E
      NamedView       =   ""
      GeometryColor   =   "DefaultColors"
      PrintBackgroundColor=   "16777215"
      PrintGeometryColor=   "0"
      ShadingMode     =   "Gouraud"
      ProjectionMode  =   "Parallel"
      EnableUIMode    =   "DisableRightClickMenu"
      Layout          =   ""
      DisplayMode     =   -1
   End
   Begin MSComctlLib.ImageList imlNav 
      Left            =   10440
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEngDWG.frx":1F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEngDWG.frx":2564
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEngDWG.frx":2B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEngDWG.frx":31D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   315
      MouseIcon       =   "frmEngDWG.frx":3812
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   765
      UseMnemonic     =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   10395
   End
   Begin VB.Image imgMenu 
      Height          =   570
      Left            =   0
      Picture         =   "frmEngDWG.frx":3B1C
      Top             =   600
      Width           =   1080
   End
   Begin VB.Menu mnuVolo 
      Caption         =   "mnuVolo"
      Visible         =   0   'False
      Begin VB.Menu mnuPan 
         Caption         =   "Pan"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuZoomW 
         Caption         =   "Zoom Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuDisplayMain 
         Caption         =   "Display"
         Begin VB.Menu mnuDisplay 
            Caption         =   "Default Colors"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuDisplay 
            Caption         =   "Black on White"
            Index           =   1
         End
         Begin VB.Menu mnuDisplay 
            Caption         =   "Clear Scale"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyClip 
         Caption         =   "Copy to Clipboard"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPrintSet 
         Caption         =   "Print Set"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmEngDWG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iDisplay As Integer
Dim bViewSet As Boolean
Dim dLeft As Double, dRight As Double, dTop As Double, dBottom As Double

Dim pLink As String
Dim tCUNO As Long
Dim tELT As String

Public Property Get PassLink() As String
    PassLink = pLink
End Property
Public Property Let PassLink(ByVal vNewValue As String)
    pLink = vNewValue
End Property

Private Sub cboDWG_Click()
    lblNavCnt.Caption = CStr(cboDWG.ListIndex + 1) & " of " & cboDWG.ListCount
    Select Case cboDWG.ListIndex + 1
        Case Is = 1 And cboDWG.ListCount > 1
            imgNav(0).Picture = imlNav.ListImages(2).Picture: imgNav(0).Enabled = False
            imgNav(1).Picture = imlNav.ListImages(3).Picture: imgNav(1).Enabled = True
        Case Is = 1 And cboDWG.ListCount = 1
            imgNav(0).Picture = imlNav.ListImages(2).Picture: imgNav(0).Enabled = False
            imgNav(1).Picture = imlNav.ListImages(4).Picture: imgNav(1).Enabled = False
        Case Is < cboDWG.ListCount
            imgNav(0).Picture = imlNav.ListImages(1).Picture: imgNav(0).Enabled = True
            imgNav(1).Picture = imlNav.ListImages(3).Picture: imgNav(1).Enabled = True
        Case Is = cboDWG.ListCount
            imgNav(0).Picture = imlNav.ListImages(1).Picture: imgNav(0).Enabled = True
            imgNav(1).Picture = imlNav.ListImages(4).Picture: imgNav(1).Enabled = False
    End Select
    
    If lstDWG.List(cboDWG.ListIndex) <> "NA" Then
        bViewSet = False
        vol1.src = lstDWG.List(cboDWG.ListIndex)
        vol1.Visible = True
    Else
        vol1.Visible = False
        vol1.src = ""
        MsgBox "DWF file is not available", vbExclamation, "Sorry..."
    End If
End Sub

Private Sub Form_Load()
    tCUNO = CLng(Left(pLink, 8))
    tELT = UCase(Mid(pLink, 10))
    
    Call GetDWGList(tCUNO, tELT)
    If cboDWG.ListCount > 0 Then
        cboDWG.ListIndex = 0
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    shpHDR.Width = Me.ScaleWidth
    picNav.Left = Me.ScaleWidth - picNav.Width
    
    vol1.Width = Me.ScaleWidth - vol1.Left - 120
    vol1.Height = Me.ScaleHeight - vol1.Top - 120
    
End Sub

Public Sub GetDWGList(pCUNO As Long, pElt As String)
    Dim strSelect As String, sName As String
    Dim rst As ADODB.Recordset
    
    cboDWG.Clear
    lstDWG.Clear
    strSelect = "SELECT NVL(D.PRGID, 0) AS PRGID, NVL(F.DWFID, 0) AS DWFID, " & _
                "NVL(D.DWGNUM, 0) AS DWGNUM, S.SHTSEQ, S.SHTDESC, NVL(F.DWFPATH,'NA') AS DWFPATH " & _
                "FROM ANNOTATOR.DWG_MASTER D, ANNOTATOR.DWG_SHEET S, ANNOTATOR.DWG_DWF F " & _
                "WHERE D.DWGID IN (" & _
                    "SELECT DWGID " & _
                    "From ANNOTATOR.DWG_ELEMENT " & _
                    "WHERE INVID IN (" & _
                        "SELECT E.ELTID " & _
                        "FROM IGLPROD.IGL_KIT K, IGLPROD.IGL_ELEMENT E " & _
                        "Where K.AN8_CUNO = " & pCUNO & " " & _
                        "AND K.KITID = E.KITID " & _
                        "AND E.ELTFNAME = '" & pElt & "'" & _
                    ")" & _
                ") " & _
                "AND D.DSTATUS > 0 " & _
                "AND D.DWGID = S.DWGID " & _
                "AND S.SSTATUS > 0 " & _
                "AND S.DWGID = F.DWGID (+) " & _
                "AND S.SHTID = F.SHTID (+) " & _
                "ORDER BY D.DWGNUM, S.SHTSEQ"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If rst.Fields("PRGID") > 0 Then
            sName = CStr(rst.Fields("PRGID")) & "-" & Right("00" & rst.Fields("DWGNUM"), 2) & _
                        Trim(rst.Fields("SHTSEQ")) & " -- " & Trim(rst.Fields("SHTDESC"))
        Else
            sName = "SHEET " & Right("00" & rst.Fields("DWGNUM"), 2) & _
                        Trim(rst.Fields("SHTSEQ")) & " -- " & Trim(rst.Fields("SHTDESC"))
        End If
        cboDWG.AddItem sName
        cboDWG.ItemData(cboDWG.NewIndex) = rst.Fields("DWFID")
        
        lstDWG.AddItem Trim(rst.Fields("DWFPATH"))
            
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
        
'    Debug.Print strSelect
    
End Sub

Private Sub imgNav_Click(Index As Integer)
    Select Case Index
        Case 0 ''BACK''
            cboDWG.ListIndex = cboDWG.ListIndex - 1
        Case 1 ''FORWARD''
            cboDWG.ListIndex = cboDWG.ListIndex + 1
    End Select
        
End Sub

Private Sub lblMenu_Click()
    If vol1.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
    End If
End Sub

Public Sub ClearChecks()
    mnuPan.Checked = False
    mnuZoom.Checked = False
    mnuZoomW.Checked = False
End Sub

Private Sub mnuCopyClip_Click()
    SendKeys "^C"
End Sub

Private Sub mnuDisplay_Click(Index As Integer)
    Dim i As Integer
    
    iDisplay = Index
    
    For i = 0 To 2
        If i = Index Then mnuDisplay(i).Checked = True Else mnuDisplay(i).Checked = False
    Next i
    
    Select Case Index
        Case 0
            vol1.GeometryColor = "DefaultColors"
            vol1.BackgroundColor = "DefaultColors"
            Me.BackColor = vbBlack
        Case 1
            vol1.GeometryColor = vbBlack
            vol1.BackgroundColor = vbWhite
            Me.BackColor = vbWhite
        Case 2
            vol1.GeometryColor = "ClearScale"
            vol1.BackgroundColor = "ClearScale"
            Me.BackColor = vbWhite
    End Select
End Sub

Private Sub mnuFullView_Click()
    vol1.SetCurrentView dLeft, dRight, dBottom, dTop
End Sub

Private Sub mnuLayers_Click()
    vol1.ShowLayersDialog
End Sub

Private Sub mnuPan_Click()
    ClearChecks
    mnuPan.Checked = True
    vol1.UserMode = "Pan"
End Sub

Private Sub mnuPrint_Click()
    vol1.ShowPrintDialog
End Sub

Private Sub mnuZoom_Click()
    ClearChecks
    mnuZoom.Checked = True
    vol1.UserMode = "Zoom"
End Sub

Private Sub mnuZoomW_Click()
    ClearChecks
    mnuZoomW.Checked = True
    vol1.UserMode = "ZoomToRect"
End Sub

Private Sub vol1_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuVolo
    End If
End Sub

Private Sub vol1_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            InitialView
            bViewSet = True
        End If
    End If
End Sub

Public Function InitialView()
    vol1.GetCurrentView dLeft, dRight, dBottom, dTop
End Function
