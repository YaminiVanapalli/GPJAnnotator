VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOSP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOSP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   4440
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDtl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDtl"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image img1 
      Height          =   2235
      Left            =   120
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu mnuRC 
      Caption         =   "mnuRC"
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
      Begin VB.Menu mnuPCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmOSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPath As String, sFile As String
Dim pGID As Long
Dim pFile As String, pHdr As String
Dim pType As Integer

Public Property Get PassFile() As String
    PassFile = pFile
End Property
Public Property Let PassFile(ByVal vNewValue As String)
    pFile = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = pHdr
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    pHdr = vNewValue
End Property

Public Property Get PassType() As Integer
    PassType = pType
End Property
Public Property Let PassType(ByVal vNewValue As Integer)
    pType = vNewValue
End Property


Private Sub Form_Load()
    Dim rAsp As Double
    
    
    
    
    Select Case pType
        Case 0: sPath = "\\DETMSFS01\GPJAnnotator\Graphics\"
        Case 1: sPath = "\\DETMSFS01\GPJAnnotator\Floorplans\FacilPho\"
    End Select
    
    sFile = sPath & pFile & ".jpg"
    img1.Picture = LoadPicture(sFile)
    
    If (Me.Width - Me.ScaleWidth) + (img1.Left * 2) + img1.Width < (Screen.Width * 0.8) _
                And (Me.Height - Me.ScaleHeight) + (img1.Top * 2) + img1.Height < (Screen.Height * 0.8) Then
        Me.Width = (Me.Width - Me.ScaleWidth) + (img1.Left * 2) + img1.Width
        Me.Height = (Me.Height - Me.ScaleHeight) + (img1.Top * 2) + img1.Height
    Else
        rAsp = img1.Width / img1.Height
        If rAsp > Screen.Width / Screen.Height Then
            img1.Stretch = True
            Me.Width = Screen.Width * 0.8
            img1.Width = Me.ScaleWidth - (img1.Left * 2)
            img1.Height = img1.Width / rAsp
            Me.Height = (Me.Height - Me.ScaleHeight) + (img1.Top * 2) + img1.Height
            
        Else
            img1.Stretch = True
            Me.Height = Screen.Height * 0.8
            img1.Height = Me.ScaleHeight - (img1.Top * 2)
            img1.Width = img1.Height * rAsp
            Me.Width = (Me.Width - Me.ScaleWidth) + (img1.Left * 2) + img1.Width
        End If
    End If
    
    Me.Caption = pHdr
    
    If pType = 1 Then
        pGID = CLng(pFile)
        Call PopDetails(pGID)
        Me.Height = Me.Height + lblDtl.Height
        lblDtl.Top = Me.ScaleHeight - 255
    End If
End Sub

Public Sub PopDetails(tGID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT GDESC, ADDUSER, ADDDTTM " & _
                "FROM ANNOTATOR.GFX_MASTER " & _
                "WHERE GID = " & tGID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        lblDtl.Caption = "Photo posted by " & Trim(rst.Fields("ADDUSER")) & _
                    " on " & Format(rst.Fields("ADDDTTM"), "dddd, mmmm d, yyyy (h:nn ampm)")
        Me.Caption = Me.Caption & "  (" & Trim(rst.Fields("GDESC")) & ")"
        lblDtl.Visible = True
    Else
        lblDtl.Visible = False
    End If
    rst.Close: Set rst = Nothing
    
End Sub



Private Sub img1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pType = 1 And Button = vbRightButton Then Me.PopupMenu mnuRC
        
End Sub

Private Sub mnuPDownload_Click()
    With frmBrowse
        .PassFrom = UCase(Me.Name)
        .PassFacil = pHdr '' Mid(lblWelcome.Caption, 12)
        .PassFCCD = frmFacil.lFCCD
        .PassGID = pGID
        .PassFILETYPE = "jpg"
        .Show 1
    End With
End Sub

Private Sub mnuPEmail_Click()
    frmEmailFile.PassHDR = pHdr
    frmEmailFile.PassFrom = UCase(Me.Name)
    frmEmailFile.PassFCCD = frmFacil.lFCCD
    frmEmailFile.PassGID = pGID
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
    If pWidth > img1.Width And pHgt > img1.Height Then
        lWidth = img1.Width
        lHgt = img1.Height
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    ElseIf pHgt > img1.Height Then
        lWidth = pWidth
        lHgt = (pWidth / img1.Width) * img1.Height
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    ElseIf pWidth > img1.Width Then
        lHgt = pHgt
        lWidth = (pHgt / img1.Height) * img1.Width
'        lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
'        lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
    Else
        Select Case (img1.Width / img1.Height)
            Case Is > (Printer.Width / Printer.Height)
                ''USE WIDTH''
                lWidth = pWidth
                lHgt = (pWidth / img1.Width) * img1.Height
            Case Else
                ''USER HEIGHT''
                lHgt = pHgt
                lWidth = (pHgt / img1.Height) * img1.Width
        End Select
        
'        If (pWidth / pHgt) < (img1.Width / img1.Height) Then
'            lWidth = pWidth
'            lHgt = (pWidth / img1.Width) * img1.Height
''            lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
''            lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
'        Else
'            lHgt = pHgt
'            lWidth = (pHgt / img1.Height) * img1.Width
''            lTop = 720 - ((Printer.Height - Printer.ScaleHeight) / 2)
''            lLeft = 720 - ((Printer.Width - Printer.ScaleWidth) / 2) ''+ ((pWidth - lWidth) / 2)
'        End If
    End If

    lTop = 1080 - ((Printer.Height - Printer.ScaleHeight) / 2)
    lLeft = 1080 - ((Printer.Width - Printer.ScaleWidth) / 2) + ((pWidth - lWidth) / 2)

    Printer.PaintPicture img1.Picture, lLeft, lTop, lWidth, lHgt
    
'    Printer.PaintPicture img1.Picture, 0, 0, Printer.Width, Printer.Height
    Printer.EndDoc
    Me.MousePointer = 0
End Sub
